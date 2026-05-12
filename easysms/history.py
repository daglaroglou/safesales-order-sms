from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class HistoryResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def list_group(self) -> Any:
        return self._client.request("history/group/list")

    def list_single(self) -> Any:
        return self._client.request("history/single/list")
