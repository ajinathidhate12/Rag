"""
Schemas to hold SharePoint JSON data returned by Microsoft Graph.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class sharepointListschema:
    """
    List / library metadata from the Graph API "/sites/{site-id}/lists" endpoint.
    """

    id: str
    name: str
    displayName: str = ""
    description: str = ""
    web_url: str = ""
    created_date_time: str = ""
    last_modified_date_time: str = ""
    template: str = "" 
    hidden: bool = False
    parent_site_id: str = ""
    drive_id: str = ""  # filled in by SharePointClient.get_list_drive_pair
    raw_json: dict[str, Any] = field(
        default_factory=dict, repr=False, compare=False, hash=False
    )

    @property
    def is_document_library(self) -> bool:
        return self.template == "documentLibrary"

    @classmethod
    def from_dict(
        cls, payload: dict[str, Any], parent_site_id: str = ""
    ) -> sharepointListschema:
        """
        Create a sharepointListschema from a single Graph list object.
        """
        list_info = payload.get("list") or {}

        return cls(
            id=payload.get("id", ""),
            name=payload.get("name", ""),
            displayName=payload.get("displayName", "") or payload.get("name", ""),
            description=payload.get("description", "") or "",
            web_url=payload.get("webUrl", "") or "",
            created_date_time=payload.get("createdDateTime", "") or "",
            last_modified_date_time=payload.get("lastModifiedDateTime", "") or "",
            template=list_info.get("template", "") or "",
            hidden=bool(list_info.get("hidden", False)),
            parent_site_id=parent_site_id,
            raw_json=payload,
        )

    @classmethod
    def from_graph_collection(
        cls, payload: dict[str, Any], parent_site_id: str = ""
    ) -> list[sharepointListschema]:  
        """
        Create a list of sharepointListschema from a Graph collection response.
        """
        items = payload.get("value", [])
        if not isinstance(items, list):
            return []
        return [
            cls.from_dict(item, parent_site_id)
            for item in items
            if isinstance(item, dict)
        ]
