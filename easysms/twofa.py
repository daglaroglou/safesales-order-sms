from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class TwoFAResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def send(
        self,
        to: str,
        *,
        text: str | None = None,
        sender: str | None = None,
        wait: int | None = None,
        callback: str | None = None,
        ucs: bool | None = None,
    ) -> Any:
        return self._client.request(
            "2fa/send",
            {
                "to": to,
                "text": text,
                "from": sender,
                "wait": wait,
                "callback": callback,
                "ucs": ucs,
            },
        )

    def check(self, auth_id: str, code: str) -> Any:
        return self._client.request(
            "2fa/check",
            {"authId": auth_id, "code": code},
        )
