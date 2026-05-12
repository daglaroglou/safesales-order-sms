from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class GroupResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def add(self, name: str) -> Any:
        return self._client.request("group/add", {"name": name})

    def add_contact(self, group_id: str, contact_id: str) -> Any:
        return self._client.request(
            "group/addContact",
            {"groupId": group_id, "contactId": contact_id},
        )

    def delete(self, group_id: str) -> Any:
        return self._client.request("group/delete", {"groupId": group_id})

    def delete_all_contacts(self, group_id: str) -> Any:
        return self._client.request(
            "group/deleteAllContacts",
            {"groupId": group_id},
        )

    def delete_contact(
        self,
        *,
        group_id: str | None = None,
        contact_id: str | None = None,
        contact_group_id: str | None = None,
    ) -> Any:
        return self._client.request(
            "group/deleteContact",
            {
                "groupId": group_id,
                "contactId": contact_id,
                "contactGroupId": contact_group_id,
            },
        )

    def get(self, group_id: str) -> Any:
        return self._client.request("group/get", {"groupId": group_id})

    def list(self) -> Any:
        return self._client.request("group/list")
