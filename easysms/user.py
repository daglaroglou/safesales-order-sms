from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easysms.client import EasySMSClient


class UserResource:
    def __init__(self, client: EasySMSClient) -> None:
        self._client = client

    def add(self, email: str, password: str) -> Any:
        return self._client.request(
            "user/add",
            {"email": email, "password": password},
        )

    def list(self) -> Any:
        return self._client.request("user/list")

    def topup(self, user_id: str, sms: int, cost: int | float) -> Any:
        return self._client.request(
            "user/topup",
            {"userId": user_id, "sms": sms, "cost": cost},
        )

    def add_comment(self, user_id: str, comment: str) -> Any:
        return self._client.request(
            "user/comment/add",
            {"userId": user_id, "comment": comment},
        )

    def delete_comment(self, comment_id: str) -> Any:
        return self._client.request(
            "user/comment/delete",
            {"commentId": comment_id},
        )

    def list_comments(self, user_id: str) -> Any:
        return self._client.request(
            "user/comment/list",
            {"userId": user_id},
        )
