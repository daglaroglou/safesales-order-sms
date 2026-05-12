from __future__ import annotations

from typing import TYPE_CHECKING, Any, Iterable

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class SMSResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def send(
        self,
        to: str,
        text: str,
        *,
        sender: str | None = None,
        ucs: bool | None = None,
        flash: bool | None = None,
        timestamp: int | None = None,
        callback: str | None = None,
    ) -> Any:
        params = {
            "to": to,
            "text": text,
            "from": sender,
            "ucs": ucs,
            "flash": flash,
            "timestamp": timestamp,
            "callback": callback,
        }
        fallback_url = self._client.base_url.replace("https://", "http://") + "/sms/send"
        return self._client.request(
            "sms/send",
            params,
            fallback_url=fallback_url,
        )

    def bulk(
        self,
        to: str | Iterable[str],
        text: str,
        *,
        sender: str | None = None,
        ucs: bool | None = None,
        flash: bool | None = None,
        timestamp: int | None = None,
    ) -> Any:
        recipients = ",".join(to) if not isinstance(to, str) else to
        return self._client.request(
            "sms/bulk",
            {
                "to": recipients,
                "text": text,
                "from": sender,
                "ucs": ucs,
                "flash": flash,
                "timestamp": timestamp,
            },
        )

    def cancel(self, sms_id: str) -> Any:
        return self._client.request("sms/cancel", {"smsId": sms_id})
