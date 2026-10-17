import logging
import os
from urllib.parse import urlparse

import httpx
import msal
from dotenv import load_dotenv

from schema import sharepointListschema

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    force=True,
)

logger = logging.getLogger(__name__)


class SharePointClient:
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        site_url: str,
        token_url: str = "https://login.microsoftonline.com/",
        graph_base_url: str = "https://graph.microsoft.com/v1.0",
        graph_scopes: str | list[str] = "https://graph.microsoft.com/.default",
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.site_url = site_url
        self.token_url = token_url.rstrip("/") + "/"
        self.graph_base_url = graph_base_url.rstrip("/")
        self.graph_scopes = (
            [graph_scopes] if isinstance(graph_scopes, str) else graph_scopes
        )
        self.access_token: str | None = None
        self.client: httpx.AsyncClient | None = None
        self._site_id: str | None = None

    @classmethod
    def from_env(cls) -> "SharePointClient":
        """Build a client from the variables in .env."""
        return cls(
            tenant_id=os.environ["MS_tenantId"],
            client_id=os.environ["MS_clientId"],
            client_secret=os.environ["MS_clientSecret"],
            site_url=os.environ["SHAREPOINT_siteUrl"],
            token_url=os.getenv("MS_token_url", "https://login.microsoftonline.com/"),
            graph_base_url=os.getenv("GRAPH_base_url", "https://graph.microsoft.com/v1.0"),
            graph_scopes=os.getenv("GRAPH_scopes", "https://graph.microsoft.com/.default"),
        )

    def require_client(self) -> httpx.AsyncClient:
        """
        Return the HTTP client if connected, else raise an error.
        """
        if self.client is None:
            raise RuntimeError("SharePoint client is not connected. Call connect() first.")
        return self.client

    def get_access_token(self) -> str:
        try:
            app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                authority=f"{self.token_url}{self.tenant_id}",
                client_credential=self.client_secret,
            )

            token_json = app.acquire_token_for_client(scopes=self.graph_scopes)

            if not token_json or "access_token" not in token_json:
                error_description = (
                    token_json.get("error_description") if token_json else "Unknown error"
                )
                raise RuntimeError(f"Failed to get access token: {error_description}")

            self.access_token = token_json["access_token"]
            return self.access_token

        except Exception as exc:
            logger.exception("Error fetching Microsoft Graph access token")
            raise RuntimeError(f"Unable to get access token: {exc}") from exc

    def connect(self) -> httpx.AsyncClient:
        self.access_token = self.get_access_token()
        self.client = httpx.AsyncClient(
            base_url=self.graph_base_url,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.access_token}",
            },
            timeout=30.0,
        )
        logger.info("Connected to Microsoft Graph API")
        return self.client

    async def disconnect(self) -> None:
        logger.info("Closing SharePoint client connection")
        if self.client:
            await self.client.aclose()  # AsyncClient must be closed with aclose()
            self.client = None
            self._site_id = None
            logger.info("SharePoint client connection closed")

    async def _get_all(self, endpoint: str) -> list[dict]:
        """GET an endpoint and follow @odata.nextLink pagination."""
        client = self.require_client()
        results: list[dict] = []
        url: str | None = endpoint
        while url:
            response = await client.get(url)
            response.raise_for_status()
            data = response.json()
            results.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return results

    async def get_site_id(self) -> str:
        """
        Resolve the Graph site id from the site URL
        (e.g. https://tenant.sharepoint.com/sites/Autominds).
        """
        if self._site_id:
            return self._site_id

        client = self.require_client()
        parsed = urlparse(self.site_url)
        hostname = parsed.netloc
        path = parsed.path.rstrip("/")
        endpoint = f"/sites/{hostname}:{path}" if path else f"/sites/{hostname}"

        response = await client.get(endpoint)
        response.raise_for_status()
        self._site_id = response.json()["id"]
        logger.info(f"Resolved site id for {self.site_url}")
        return self._site_id

    async def get_file_list(
        self, *, document_library_only: bool = True
    ) -> list[sharepointListschema]:
        """
        Get the lists on the site; by default only document libraries.
        """
        site_id = await self.get_site_id()
        raw_lists = await self._get_all(f"/sites/{site_id}/lists")
        lists = sharepointListschema.from_graph_collection(
            {"value": raw_lists}, parent_site_id=site_id
        )
        if document_library_only:
            lists = [item for item in lists if item.is_document_library]
        logger.info(f"Retrieved {len(lists)} lists from SharePoint site {self.site_url}")
        return lists

    async def get_list_drive_pair(
        self, lists: list[sharepointListschema]
    ) -> list[sharepointListschema]:
        """
        For each document library, look up its underlying drive and store
        the drive id on the schema object.
        """
        client = self.require_client()
        site_id = await self.get_site_id()

        for item in lists:
            response = await client.get(f"/sites/{site_id}/lists/{item.id}/drive")
            if response.status_code == 404:
                logger.warning(f"List {item.displayName} has no drive; skipping")
                continue
            response.raise_for_status()
            item.drive_id = response.json().get("id", "")

        logger.info(f"Resolved drives for {len(lists)} lists")
        return lists

    async def get_documents(self, drive_id: str) -> list[dict]:
        """
        List every item in a drive (recursing into folders) and return
        simple dicts with id, name, path and type ("file" or "folder").
        """
        documents: list[dict] = []

        async def walk(endpoint: str, parent_path: str) -> None:
            for item in await self._get_all(endpoint):
                path = f"{parent_path}/{item['name']}"
                if "folder" in item:
                    documents.append(
                        {
                            "id": item["id"],
                            "name": item["name"],
                            "path": path,
                            "type": "folder",
                            "child_count": item["folder"].get("childCount", 0),
                        }
                    )
                    await walk(f"/drives/{drive_id}/items/{item['id']}/children", path)
                else:
                    documents.append(
                        {
                            "id": item["id"],
                            "name": item["name"],
                            "path": path,
                            "type": "file",
                            "size": item.get("size", 0),
                            "web_url": item.get("webUrl", ""),
                        }
                    )

        await walk(f"/drives/{drive_id}/root/children", "")
        logger.info(f"Found {len(documents)} items in drive {drive_id}")
        return documents


if __name__ == "__main__":
    import asyncio

    async def main():
        sp = SharePointClient.from_env()
        sp.connect()
        try:
            libraries = await sp.get_file_list()
            libraries = await sp.get_list_drive_pair(libraries)
            for lib in libraries:
                print(f"\nLibrary: {lib.displayName} (list id: {lib.id}, drive id: {lib.drive_id})")
                if not lib.drive_id:
                    continue
                items = await sp.get_documents(lib.drive_id)
                folders = [i for i in items if i["type"] == "folder"]
                files = [i for i in items if i["type"] == "file"]

                print(f"  Folders: {len(folders)} | Documents: {len(files)}")

                print("  Folders:")
                if not folders:
                    print("    (none)")
                for folder in folders:
                    print(f"    - {folder['path']}  [folder id: {folder['id']}]")

                print("  Documents:")
                if not files:
                    print("    (none)")
                for doc in files:
                    print(f"    - {doc['name']}  [document id: {doc['id']}]")
        finally:
            await sp.disconnect()

    asyncio.run(main())