import logging
import os

from dotenv import load_dotenv
import httpx
import msal


load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    force=True,
)


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
        self.access_token = None
        self.client = None

    def get_access_token(self):
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
            logging.exception("Error fetching Microsoft Graph access token")
            raise RuntimeError(f"Unable to get access token: {exc}") from exc

    def connect(self):
        self.access_token = self.get_access_token()
        self.client = httpx.AsyncClient(
            base_url=self.graph_base_url,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.access_token}",
            },
        )
        logging.info("Connected to Microsoft Graph API")
        return self.client


def main():
    required_vars = [
        "MS_tenantId",
        "MS_clientId",
        "MS_clientSecret",
        "SHAREPOINT_siteUrl",
    ]

    missing = [name for name in required_vars if not os.getenv(name)]
    if missing:
        raise ValueError(f"Missing environment variables: {', '.join(missing)}")

    s1 = SharePointClient(
        tenant_id=os.getenv("MS_tenantId"),
        client_id=os.getenv("MS_clientId"),
        client_secret=os.getenv("MS_clientSecret"),
        site_url=os.getenv("SHAREPOINT_siteUrl"),
        token_url=os.getenv("MS_token_url") or "https://login.microsoftonline.com/",
    )

    token = s1.get_access_token()
    print("Token acquired:", token[:20], "...")
    client = s1.connect()
    print("Connected client:", client)


if __name__ == "__main__":
    main()