from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class PurchaseResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def list(self) -> Any:
        return self._client.request("purchase/list")
