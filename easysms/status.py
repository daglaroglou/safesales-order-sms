from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class StatusResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def get(self) -> Any:
        return self._client.request("status/get")

    def sms(self, sms_id: str) -> Any:
        return self._client.request("status/sms", {"smsId": sms_id})
