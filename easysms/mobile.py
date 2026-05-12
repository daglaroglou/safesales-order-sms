from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class MobileResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def check(self, mobile: str) -> Any:
        return self._client.request(
            "mobile/check",
            {"mobile": mobile},
            authenticate=False,
        )
