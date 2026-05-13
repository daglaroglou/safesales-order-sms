from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from typing import Any


class Carrier(str, Enum):
    ACS = "acs"
    BOX_EXPRESS = "box_express"


@dataclass(frozen=True, slots=True)
class Shipment:
    carrier: Carrier
    voucher: str
    recipient_name: str
    phone: str
    cod_amount: Decimal | None = None
    pickup_date: date | datetime | None = None
    order_reference: str | None = None
    notes: str | None = None
    source_file: str | None = None
    raw: dict[str, Any] | None = None
