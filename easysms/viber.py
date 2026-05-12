from __future__ import annotations

from typing import TYPE_CHECKING, Any, Iterable

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class ViberResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def send(
        self,
        to: str | None = None,
        text: str | None = None,
        *,
        sender: str | None = None,
        button_caption: str | None = None,
        button_action: str | None = None,
        image: str | None = None,
        validity: int | None = None,
        all_devices: bool | None = None,
        sms_fallback: bool | None = None,
        sms_text: str | None = None,
        sms_from: str | None = None,
        ucs: bool | None = None,
        flash: bool | None = None,
        contact_id: str | None = None,
        group_id: str | None = None,
    ) -> Any:
        return self._client.request(
            "viber/send",
            self._message_params(
                to=to,
                text=text,
                sender=sender,
                button_caption=button_caption,
                button_action=button_action,
                image=image,
                validity=validity,
                all_devices=all_devices,
                sms_fallback=sms_fallback,
                sms_text=sms_text,
                sms_from=sms_from,
                ucs=ucs,
                flash=flash,
                contact_id=contact_id,
                group_id=group_id,
            ),
        )

    def bulk(
        self,
        to: str | Iterable[str],
        text: str | None = None,
        *,
        sender: str | None = None,
        button_caption: str | None = None,
        button_action: str | None = None,
        image: str | None = None,
        validity: int | None = None,
        all_devices: bool | None = None,
        sms_fallback: bool | None = None,
        sms_text: str | None = None,
        sms_from: str | None = None,
        ucs: bool | None = None,
        flash: bool | None = None,
        contact_id: str | None = None,
        group_id: str | None = None,
    ) -> Any:
        recipients = ",".join(to) if not isinstance(to, str) else to
        return self._client.request(
            "viber/bulk",
            self._message_params(
                to=recipients,
                text=text,
                sender=sender,
                button_caption=button_caption,
                button_action=button_action,
                image=image,
                validity=validity,
                all_devices=all_devices,
                sms_fallback=sms_fallback,
                sms_text=sms_text,
                sms_from=sms_from,
                ucs=ucs,
                flash=flash,
                contact_id=contact_id,
                group_id=group_id,
            ),
        )

    def balance(self) -> Any:
        return self._client.request("viber/balance")

    def _message_params(
        self,
        *,
        to: str | None,
        text: str | None,
        sender: str | None,
        button_caption: str | None,
        button_action: str | None,
        image: str | None,
        validity: int | None,
        all_devices: bool | None,
        sms_fallback: bool | None,
        sms_text: str | None,
        sms_from: str | None,
        ucs: bool | None,
        flash: bool | None,
        contact_id: str | None,
        group_id: str | None,
    ) -> dict[str, Any]:
        return {
            "to": to,
            "text": text,
            "from": sender,
            "button_caption": button_caption,
            "button_action": button_action,
            "image": image,
            "validity": validity,
            "all_devices": all_devices,
            "sms_fallback": sms_fallback,
            "sms_text": sms_text,
            "sms_from": sms_from,
            "ucs": ucs,
            "flash": flash,
            "contactId": contact_id,
            "groupId": group_id,
        }
