from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class HLRResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def lookup(self, mobile: str) -> Any:
        return self._client.request("hlr/lookup", {"mobile": mobile})
