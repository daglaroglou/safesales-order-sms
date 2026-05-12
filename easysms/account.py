from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class AccountResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def balance(self) -> Any:
        return self._client.request("me/balance")
