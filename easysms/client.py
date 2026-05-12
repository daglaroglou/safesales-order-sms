from __future__ import annotations

import json
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from typing import Any, Literal, Mapping

ResponseFormat = Literal["json", "xml", "v2"]
DEFAULT_BASE_URL = "https://easysms.gr/api"


class EasySMSError(Exception):
    def __init__(
        self,
        message: str,
        *,
        error: str | int | None = None,
        remarks: str | None = None,
        response: Any = None,
    ) -> None:
        super().__init__(message)
        self.error = error
        self.remarks = remarks
        self.response = response


class EasySMSClient:
    def __init__(
        self,
        api_key: str,
        *,
        base_url: str = DEFAULT_BASE_URL,
        response_format: ResponseFormat = "json",
        timeout: float = 30.0,
        raise_on_error: bool = True,
        user_agent: str = "safesales-order-sms-easysms/1.0",
    ) -> None:
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")
        self.response_format = response_format
        self.timeout = timeout
        self.raise_on_error = raise_on_error
        self.user_agent = user_agent

        from easysms.account import AccountResource
        from easysms.contact import ContactResource
        from easysms.group import GroupResource
        from easysms.history import HistoryResource
        from easysms.hlr import HLRResource
        from easysms.key import KeyResource
        from easysms.mobile import MobileResource
        from easysms.purchase import PurchaseResource
        from easysms.sms import SMSResource
        from easysms.status import StatusResource
        from easysms.twofa import TwoFAResource
        from easysms.user import UserResource
        from easysms.viber import ViberResource

        self.mobile = MobileResource(self)
        self.key = KeyResource(self)
        self.sms = SMSResource(self)
        self.viber = ViberResource(self)
        self.account = AccountResource(self)
        self.twofa = TwoFAResource(self)
        self.contact = ContactResource(self)
        self.group = GroupResource(self)
        self.history = HistoryResource(self)
        self.hlr = HLRResource(self)
        self.purchase = PurchaseResource(self)
        self.status = StatusResource(self)
        self.user = UserResource(self)

    def request(
        self,
        path: str,
        params: Mapping[str, Any] | None = None,
        *,
        authenticate: bool = True,
        response_format: ResponseFormat | None = None,
        fallback_url: str | None = None,
    ) -> Any:
        payload = self._prepare_params(params or {})
        response_format = response_format or self.response_format
        payload["type"] = response_format

        if authenticate:
            if not self.api_key:
                raise ValueError("API key is required for this endpoint")
            payload["key"] = self.api_key

        primary_url = f"{self.base_url}/{path.lstrip('/')}"
        body = urllib.parse.urlencode(payload).encode("utf-8")
        request = urllib.request.Request(
            primary_url,
            data=body,
            method="POST",
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "User-Agent": self.user_agent,
            },
        )

        try:
            raw = self._read_response(request)
        except urllib.error.URLError:
            if not fallback_url:
                raise
            fallback_request = urllib.request.Request(
                fallback_url,
                data=body,
                method="POST",
                headers=request.headers,
            )
            raw = self._read_response(fallback_request)

        parsed = self._parse_response(raw, response_format)
        if self.raise_on_error:
            self._raise_for_status(parsed)
        return parsed

    def _read_response(self, request: urllib.request.Request) -> str:
        with urllib.request.urlopen(request, timeout=self.timeout) as response:
            return response.read().decode("utf-8")

    def _prepare_params(self, params: Mapping[str, Any]) -> dict[str, Any]:
        prepared: dict[str, Any] = {}
        for key, value in params.items():
            if value is None:
                continue
            if isinstance(value, bool):
                prepared[key] = "true" if value else "false"
            else:
                prepared[key] = value
        return prepared

    def _parse_response(self, raw: str, response_format: ResponseFormat) -> Any:
        if response_format == "json":
            return json.loads(raw)
        if response_format == "xml":
            return self._xml_to_dict(ET.fromstring(raw))
        return raw

    def _xml_to_dict(self, element: ET.Element) -> Any:
        children = list(element)
        if not children:
            text = (element.text or "").strip()
            return text

        grouped: dict[str, list[Any]] = {}
        for child in children:
            grouped.setdefault(child.tag, []).append(self._xml_to_dict(child))

        if len(grouped) == 1:
            tag, values = next(iter(grouped.items()))
            if len(values) == 1:
                return {tag: values[0]}
            return {tag: values}

        return {tag: values[0] if len(values) == 1 else values for tag, values in grouped.items()}

    def _raise_for_status(self, response: Any) -> None:
        if not isinstance(response, dict):
            return

        status = response.get("status")
        if status in (None, 1, "1"):
            return

        error = response.get("error")
        remarks = response.get("remarks")
        message = remarks or f"EasySMS API request failed with status {status!r}"
        if error not in (None, 0, "0"):
            message = f"{message} (error {error})"
        raise EasySMSError(message, error=error, remarks=remarks, response=response)
