from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class KeyResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def get(
        self,
        *,
        username: str | None = None,
        password: str | None = None,
        key: str | None = None,
    ) -> Any:
        return self._client.request(
            "key/get",
            self._credentials(username=username, password=password, key=key),
            authenticate=False,
        )

    def reset(
        self,
        *,
        username: str | None = None,
        password: str | None = None,
        key: str | None = None,
    ) -> Any:
        return self._client.request(
            "key/reset",
            self._credentials(username=username, password=password, key=key),
            authenticate=False,
        )

    def _credentials(
        self,
        *,
        username: str | None,
        password: str | None,
        key: str | None,
    ) -> dict[str, str]:
        if key:
            return {"key": key}
        if username and password:
            return {"username": username, "password": password}
        if self._client.api_key:
            return {"key": self._client.api_key}
        raise ValueError("Provide username and password, or an API key")
