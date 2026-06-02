from __future__ import annotations

from datetime import date
from itertools import chain
from pathlib import Path
from typing import Iterable

import openpyxl

from excel.models import Carrier, Shipment
from excel.utils import acs_output_path, normalize_phone, normalize_text, should_skip_phone_number, write_pair_workbook

ACS_SHEET_NAME = "ACS"
ACS_VOUCHER_COLUMN = "Αριθμός Αποδεικτικού"
ACS_PHONE_COLUMN = "Κωδ. Αποστ. Πελάτη"


def _iter_acs_pairs(path: Path) -> Iterable[tuple[str, str]]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook[ACS_SHEET_NAME]
        rows = worksheet.iter_rows(values_only=True)
        first_row = next(rows, None)
        if first_row is None:
            return

        headers = [normalize_text(value) for value in first_row]
        if headers and headers[0] == ACS_VOUCHER_COLUMN:
            voucher_index = headers.index(ACS_VOUCHER_COLUMN)
            phone_index = headers.index(ACS_PHONE_COLUMN)
            data_rows = rows
        else:
            voucher_index = 0
            phone_index = 1
            data_rows = chain((first_row,), rows)

        for row in data_rows:
            if not row or all(cell is None or str(cell).strip() == "" for cell in row):
                continue

            voucher = normalize_text(row[voucher_index] if voucher_index < len(row) else None)
            phone = normalize_text(row[phone_index] if phone_index < len(row) else None)
            if not voucher or not phone:
                continue
            if should_skip_phone_number(phone):
                continue

            yield voucher, phone
    finally:
        workbook.close()


def trim_acs_file(
    path: Path | str,
    output_path: Path | str | None = None,
    on_date: date | None = None,
) -> Path:
    source = Path(path)
    destination = Path(output_path) if output_path is not None else acs_output_path(source.parent, on_date)
    pairs = list(_iter_acs_pairs(source))
    return write_pair_workbook(destination, ACS_SHEET_NAME, pairs)


def parse_acs_file(
    path: Path | str,
    output_path: Path | str | None = None,
    on_date: date | None = None,
) -> list[Shipment]:
    source = Path(path)
    destination = Path(output_path) if output_path is not None else acs_output_path(source.parent, on_date)
    pairs = list(_iter_acs_pairs(source))
    write_pair_workbook(destination, ACS_SHEET_NAME, pairs)

    shipments: list[Shipment] = []
    for voucher, phone in pairs:
        normalized_phone = normalize_phone(phone)
        if not normalized_phone:
            continue

        shipments.append(
            Shipment(
                carrier=Carrier.ACS,
                voucher=voucher,
                recipient_name="",
                phone=normalized_phone,
                source_file=destination.name,
            )
        )

    return shipments
