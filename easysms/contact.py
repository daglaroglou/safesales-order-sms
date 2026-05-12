from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class ContactResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def add(
        self,
        mobile: str,
        *,
        name: str | None = None,
        surname: str | None = None,
        fullname: str | None = None,
        vname: str | None = None,
        vsurname: str | None = None,
        birthday: str | None = None,
        nameday: str | None = None,
        custom1: str | None = None,
        custom2: str | None = None,
        custom3: str | None = None,
        custom4: str | None = None,
        custom5: str | None = None,
    ) -> Any:
        return self._client.request(
            "contact/add",
            self._fields(
                mobile=mobile,
                name=name,
                surname=surname,
                fullname=fullname,
                vname=vname,
                vsurname=vsurname,
                birthday=birthday,
                nameday=nameday,
                custom1=custom1,
                custom2=custom2,
                custom3=custom3,
                custom4=custom4,
                custom5=custom5,
            ),
        )

    def delete(self, contact_id: str) -> Any:
        return self._client.request("contact/delete", {"contactId": contact_id})

    def get(self, contact_id: str) -> Any:
        return self._client.request("contact/get", {"contactId": contact_id})

    def list(self) -> Any:
        return self._client.request("contact/list")

    def update(
        self,
        contact_id: str,
        *,
        mobile: str | None = None,
        name: str | None = None,
        surname: str | None = None,
        fullname: str | None = None,
        vname: str | None = None,
        vsurname: str | None = None,
        birthday: str | None = None,
        nameday: str | None = None,
        custom1: str | None = None,
        custom2: str | None = None,
        custom3: str | None = None,
        custom4: str | None = None,
        custom5: str | None = None,
    ) -> Any:
        params = self._fields(
            mobile=mobile,
            name=name,
            surname=surname,
            fullname=fullname,
            vname=vname,
            vsurname=vsurname,
            birthday=birthday,
            nameday=nameday,
            custom1=custom1,
            custom2=custom2,
            custom3=custom3,
            custom4=custom4,
            custom5=custom5,
        )
        params["contactId"] = contact_id
        return self._client.request("contact/update", params)

    def _fields(self, **kwargs: str | None) -> dict[str, str]:
        return {key: value for key, value in kwargs.items() if value is not None}
