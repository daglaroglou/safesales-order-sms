"""Business logic for the SafeSales WinUI app (EasySMS + Excel parsing)."""

from __future__ import annotations

import traceback
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

from easysms import EasySMSClient, EasySMSError
from excel.loader import parse_excel_file
from excel.models import Carrier, Shipment
from excel.utils import format_export_date
from ui import config as app_config

ACS_TRACKING_SMS_TEMPLATE = (
    "H αποστολή σας έχει γίνει με ACS Courier και αριθμό αποστολής [[custom1]]. "
    "Αναζήτηση αποστολής στο https://www.acscourier.net/el/track-and-trace/"
)
BOX_EXPRESS_TRACKING_SMS_TEMPLATE = (
    "H αποστολή σας έχει γίνει με BOX EXPRESS COURIER και αριθμό αποστολής [[custom1]]. "
    "Αναζήτηση αποστολής στο https://boxexpress.gr/tracking"
)


def get_api_key() -> str | None:
    return app_config.get_api_key()


def get_client() -> EasySMSClient | None:
    key = get_api_key()
    if not key:
        return None
    return EasySMSClient(api_key=key)


def build_client_for_key(api_key: str) -> EasySMSClient:
    key = (api_key or "").strip()
    if not key:
        raise ValueError("API key is required.")
    return EasySMSClient(api_key=key)


def validate_api_key(api_key: str) -> tuple[bool, str]:
    key = (api_key or "").strip()
    if not key:
        return False, "API key is required."
    client = build_client_for_key(key)
    try:
        client.account.balance()
    except EasySMSError as exc:
        return False, f"Invalid API key: {exc}"
    except OSError as exc:
        return False, f"Network error: {exc}"
    except Exception as exc:  # noqa: BLE001
        return False, f"Validation error: {exc}"
    return True, "API key validated."


def format_balance_response(data: object) -> str:
    if data is None:
        return "—"
    if isinstance(data, dict):
        if data.get("status") not in (None, 1, "1"):
            err = data.get("remarks") or data.get("error") or data.get("message")
            return f"Error: {err}" if err else "Unknown error"
        for key in ("amount", "balance", "credits", "sms", "value"):
            if key in data and data[key] not in (None, ""):
                return str(data[key])
        return str(data)
    return str(data)


def fetch_balance_display(client: EasySMSClient) -> str:
    try:
        return format_balance_response(client.account.balance())
    except EasySMSError as exc:
        return f"API error: {exc}"
    except OSError as exc:
        return f"Network: {exc}"


def is_easysms_online(client: EasySMSClient | None) -> bool:
    """True if the EasySMS service responds (reachable). Auth errors still count as online."""
    import urllib.error
    import urllib.request

    try:
        if client is not None:
            prev = client.raise_on_error
            client.raise_on_error = False
            try:
                client.account.balance()
            finally:
                client.raise_on_error = prev
            return True
        req = urllib.request.Request(
            "https://easysms.gr/",
            headers={"User-Agent": "safesales-order-sms/1.0"},
            method="GET",
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            resp.read(128)
        return True
    except (OSError, urllib.error.URLError):
        return False


def send_one_sms(
    client: EasySMSClient,
    to: str,
    text: str,
    *,
    sender: str | None = None,
) -> object:
    return client.sms.send(to, text, sender=sender or None)


@dataclass(slots=True)
class ProcessResult:
    shipments: list[Shipment]
    lines: list[str]


def process_excel_paths(
    paths: list[Path],
    *,
    export_directory: Path | None = None,
) -> ProcessResult:
    lines: list[str] = []
    all_shipments: list[Shipment] = []
    export_dir = export_directory or app_config.get_export_directory()
    lines.append(f"Export folder  ·  {export_dir}")

    for path in paths:
        if not path or not str(path).strip():
            continue
        p = Path(path)
        try:
            if not p.is_file():
                lines.append(f"{p.name}  ·  skipped (not a file)")
                continue
            shipments = parse_excel_file(p, export_directory=export_dir)
            all_shipments.extend(shipments)
            lines.append(f"{p.name}  ·  {len(shipments)} shipment{'s' if len(shipments) != 1 else ''} parsed")
        except ValueError as exc:
            lines.append(f"{p.name}  ·  unsupported or invalid format: {exc}")
        except OSError as exc:
            lines.append(f"{p.name}  ·  I/O error: {exc}")
        except Exception as exc:  # noqa: BLE001 — surface unexpected parse errors in UI log
            lines.append(f"{p.name}  ·  unexpected: {exc}")
            lines.append(traceback.format_exc())

    return ProcessResult(shipments=all_shipments, lines=lines)


def _carrier_label(s: Shipment) -> str:
    c = s.carrier.value if hasattr(s.carrier, "value") else str(s.carrier)
    return str(c)


def _shipment_courier_display(s: Shipment) -> str:
    """Short courier label for tables (box_express → BE)."""
    c = _carrier_label(s)
    if c == "box_express":
        return "B.E"
    if c == "acs":
        return "ACS"
    return c


def _clip_field(text: str, width: int) -> str:
    t = str(text).replace("\r", " ").replace("\n", " ")
    if len(t) <= width:
        return t.ljust(width)
    return t[: max(0, width - 1)] + "…"


def _center_field(text: str, width: int) -> str:
    t = str(text).replace("\r", " ").replace("\n", " ")
    if len(t) > width:
        t = t[: max(0, width - 1)] + "…"
    return t.center(width)


def _shipment_table_sep_line(wc: int, wi: int, wp: int) -> str:
    """Full-width dashed rule aligned with Courier | ID | Phone columns."""
    return "  | " + "-" * wc + " | " + "-" * wi + " | " + "-" * wp + " |"


def _column_width(header: str, values: list[str], cap: int) -> int:
    m = max(len(header), max((len(v) for v in values), default=0))
    return min(cap, max(m, len(header)))


def format_parse_lines_display(parse_lines: list[str]) -> str:
    """Readable summary of each workbook line (ASCII-friendly)."""
    if not parse_lines:
        return (
            "No files were processed in this run.\n\n"
            "Choose ACS and/or Box Express .xlsx files, then tap Process files."
        )
    out: list[str] = []
    out.append("WORKBOOKS")
    out.append("=" * 46)
    for raw in parse_lines:
        line = (raw or "").strip()
        if not line:
            continue
        if " · " in line:
            name, detail = line.split(" · ", 1)
            name, detail = name.strip(), detail.strip()
            out.append("")
            out.append(f"  *  {name}")
            out.append(f"     {detail}")
        else:
            warn = any(k in line.lower() for k in ("skip", "error", "invalid", "unsupported"))
            prefix = "[!] " if warn else "    "
            out.append(f"  {prefix}{line}")
    return "\n".join(out).rstrip() + "\n"


def render_shipments_table(shipments: list[Shipment], *, max_rows: int = 150) -> str:
    """Three columns (Courier | ID | Phone); dashed line when displayed courier changes."""
    if not shipments:
        return "No shipment rows parsed."
    rows = shipments[:max_rows]
    carriers = [_shipment_courier_display(s) for s in rows]
    ids_ = [str(s.voucher) for s in rows]
    phones = [str(s.phone) for s in rows]

    cap_c, cap_i, cap_p = 8, 28, 18
    wc = _column_width("Courier", carriers, cap_c)
    wi = _column_width("ID", ids_, cap_i)
    wp = _column_width("Phone", phones, cap_p)

    def data_line(courier: str, sid: str, phone: str) -> str:
        cells = (_center_field(courier, wc), _clip_field(sid, wi), _clip_field(phone, wp))
        return "  | " + " | ".join(cells) + " |"

    total = len(shipments)
    shown = len(rows)
    out: list[str] = []
    out.append(f"  Rows: {total} total   (showing {shown} below)")
    out.append("")
    out.append(data_line("Courier", "ID", "Phone"))
    prev: str | None = None
    for s in rows:
        cur = _shipment_courier_display(s)
        if prev is not None and cur != prev:
            out.append(_shipment_table_sep_line(wc, wi, wp))
        out.append(data_line(cur, str(s.voucher), str(s.phone)))
        prev = cur
    if total > max_rows:
        out.append("")
        out.append(f"  ... {total - max_rows} more row(s) not shown.")
    return "\n".join(out)


def format_excel_results_panel(parse_lines: list[str], shipments: list[Shipment], *, max_rows: int = 150) -> str:
    """Full report string (activity log / legacy); UI uses split panels instead."""
    return f"{format_parse_lines_display(parse_lines)}\n\n{render_shipments_table(shipments, max_rows=max_rows)}"


def render_shipments_summary(shipments: list[Shipment], *, max_rows: int = 80) -> str:
    """Compact table of shipments (legacy name; used for plain-text summaries)."""
    return render_shipments_table(shipments, max_rows=max_rows)


def format_message_template(template: str, shipment: Shipment) -> str:
    carrier = shipment.carrier.value if hasattr(shipment.carrier, "value") else str(shipment.carrier)
    return template.format(
        voucher=shipment.voucher,
        phone=shipment.phone,
        carrier=carrier,
        source_file=shipment.source_file or "",
    )


def tracking_template_for_carrier(carrier: Carrier) -> str:
    if carrier == Carrier.ACS:
        return ACS_TRACKING_SMS_TEMPLATE
    if carrier == Carrier.BOX_EXPRESS:
        return BOX_EXPRESS_TRACKING_SMS_TEMPLATE
    raise ValueError(f"unsupported carrier: {carrier}")


def format_tracking_message(template: str, voucher: str) -> str:
    return template.replace("[[custom1]]", voucher)


def carrier_group_name(carrier: Carrier, on_date: date | None = None) -> str:
    """Match trimmed export filenames: ``18-5-2026.xlsx`` / ``18-5-2026be.xlsx``."""
    base = format_export_date(on_date)
    if carrier == Carrier.BOX_EXPRESS:
        return f"{base}be"
    return base


def _unwrap_entity_list(node: Any, singular: str) -> list[dict[str, Any]]:
    if node is None:
        return []
    if isinstance(node, list):
        return [item for item in node if isinstance(item, dict)]
    if isinstance(node, dict):
        inner = node.get(singular)
        if inner is None:
            return [node]
        return _unwrap_entity_list(inner, singular)
    return []


def _groups_from_list_response(data: Any) -> list[dict[str, Any]]:
    if not isinstance(data, dict):
        return []
    return _unwrap_entity_list(data.get("groups"), "group")


def _contacts_from_list_response(data: Any) -> list[dict[str, Any]]:
    if not isinstance(data, dict):
        return []
    return _unwrap_entity_list(data.get("contacts"), "contact")


def _response_id(response: Any, key: str) -> str | None:
    if not isinstance(response, dict):
        return None
    value = response.get(key)
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _easysms_error_code(exc: EasySMSError) -> str | int | None:
    return exc.error


def _find_group_id_by_name(client: EasySMSClient, name: str) -> str | None:
    data = client.group.list()
    for group in _groups_from_list_response(data):
        if str(group.get("name") or "").strip() == name:
            gid = group.get("groupId")
            if gid is not None:
                return str(gid)
    return None


def ensure_group(client: EasySMSClient, name: str) -> str:
    existing = _find_group_id_by_name(client, name)
    if existing:
        return existing
    created = client.group.add(name)
    group_id = _response_id(created, "groupId")
    if not group_id:
        raise EasySMSError(f"group/add did not return groupId for {name!r}", response=created)
    return group_id


def _build_contact_index(client: EasySMSClient) -> dict[str, str]:
    data = client.contact.list()
    index: dict[str, str] = {}
    for contact in _contacts_from_list_response(data):
        mobile = contact.get("mobile")
        contact_id = contact.get("contactId")
        if mobile is None or contact_id is None:
            continue
        index[str(mobile).strip()] = str(contact_id)
    return index


def ensure_contact_id(
    client: EasySMSClient,
    mobile: str,
    contact_index: dict[str, str],
    *,
    custom1: str | None = None,
    name: str | None = None,
) -> str:
    phone = mobile.strip()
    cached = contact_index.get(phone)
    if cached:
        return cached

    try:
        created = client.contact.add(phone, name=name, custom1=custom1)
    except EasySMSError as exc:
        if str(_easysms_error_code(exc)) != "202":
            raise
        for contact in _contacts_from_list_response(client.contact.list()):
            if str(contact.get("mobile") or "").strip() == phone:
                contact_id = contact.get("contactId")
                if contact_id is not None:
                    contact_index[phone] = str(contact_id)
                    return str(contact_id)
        raise

    contact_id = _response_id(created, "contactId")
    if not contact_id:
        raise EasySMSError(f"contact/add did not return contactId for {phone!r}", response=created)
    contact_index[phone] = contact_id
    return contact_id


def add_contact_to_group(client: EasySMSClient, group_id: str, contact_id: str) -> None:
    try:
        client.group.add_contact(group_id, contact_id)
    except EasySMSError as exc:
        # 218: contact already in group — treat as success
        if str(_easysms_error_code(exc)) == "218":
            return
        raise


def populate_carrier_group(
    client: EasySMSClient,
    shipments: list[Shipment],
    carrier: Carrier,
    *,
    on_date: date | None = None,
) -> list[str]:
    """Create/reuse today's carrier group and add shipment mobiles as contacts."""
    name = carrier_group_name(carrier, on_date)
    log: list[str] = []
    try:
        group_id = ensure_group(client, name)
    except (EasySMSError, OSError) as exc:
        log.append(f"Group {name}: failed to create: {exc}")
        return log

    contact_index = _build_contact_index(client)
    added = 0
    failed = 0
    seen_phones: set[str] = set()

    for shipment in shipments:
        phone = shipment.phone.strip()
        if phone in seen_phones:
            continue
        seen_phones.add(phone)
        try:
            contact_id = ensure_contact_id(
                client,
                phone,
                contact_index,
                custom1=shipment.voucher,
                name=(shipment.recipient_name or None),
            )
            add_contact_to_group(client, group_id, contact_id)
            added += 1
        except (EasySMSError, OSError) as exc:
            failed += 1
            log.append(f"Group {name}: {phone}: {exc}")

    if failed == 0:
        log.insert(0, f"Group {name}: {added} contact(s) added.")
    else:
        log.insert(0, f"Group {name}: {added} added, {failed} failed.")
    return log


def send_shipments(
    client: EasySMSClient,
    shipments: list[Shipment],
    template: str,
    *,
    sender: str | None = None,
) -> list[str]:
    log: list[str] = []
    for idx, shipment in enumerate(shipments, start=1):
        try:
            text = format_message_template(template, shipment)
        except (KeyError, ValueError) as exc:
            log.append(f"Row {idx} ({shipment.voucher}): bad template: {exc}")
            continue
        try:
            send_one_sms(client, shipment.phone, text, sender=sender)
            log.append(f"Row {idx} ({shipment.voucher} → {shipment.phone}): sent.")
        except EasySMSError as exc:
            log.append(f"Row {idx} ({shipment.voucher} → {shipment.phone}): EasySMS error: {exc}")
        except OSError as exc:
            log.append(f"Row {idx} ({shipment.voucher} → {shipment.phone}): network: {exc}")
    return log


def send_carrier_shipments(
    client: EasySMSClient,
    shipments: list[Shipment],
    carrier: Carrier,
    *,
    sender: str | None = None,
) -> list[str]:
    """Send the predefined tracking SMS to rows matching *carrier* only."""
    template = tracking_template_for_carrier(carrier)
    filtered = [s for s in shipments if s.carrier == carrier]
    log: list[str] = []
    if not filtered:
        label = "ACS" if carrier == Carrier.ACS else "Box Express"
        log.append(f"No {label} rows in the last processed batch.")
        return log

    try:
        log.extend(populate_carrier_group(client, filtered, carrier))
    except Exception as exc:  # noqa: BLE001 — keep sending even if grouping fails
        log.append(f"Group setup error: {exc}")

    for idx, shipment in enumerate(filtered, start=1):
        text = format_tracking_message(template, shipment.voucher)
        try:
            send_one_sms(client, shipment.phone, text, sender=sender)
            log.append(f"Row {idx} ({shipment.voucher} → {shipment.phone}): sent.")
        except EasySMSError as exc:
            log.append(f"Row {idx} ({shipment.voucher} → {shipment.phone}): EasySMS error: {exc}")
        except OSError as exc:
            log.append(f"Row {idx} ({shipment.voucher} → {shipment.phone}): network: {exc}")
    return log


def send_selected_shipments(
    client: EasySMSClient,
    shipments: list[Shipment],
    carriers: set[Carrier],
    *,
    sender: str | None = None,
) -> list[str]:
    """Send tracking messages for all matching carriers using each row's own template."""
    if not carriers:
        return ["No courier selected."]

    filtered = [s for s in shipments if s.carrier in carriers]
    if not filtered:
        labels: list[str] = []
        if Carrier.ACS in carriers:
            labels.append("ACS")
        if Carrier.BOX_EXPRESS in carriers:
            labels.append("Box Express")
        selected = " + ".join(labels) if labels else "selected"
        return [f"No {selected} rows in the last processed batch."]

    log: list[str] = []
    for carrier in (Carrier.ACS, Carrier.BOX_EXPRESS):
        if carrier not in carriers:
            continue
        carrier_rows = [s for s in filtered if s.carrier == carrier]
        if not carrier_rows:
            continue
        try:
            log.extend(populate_carrier_group(client, carrier_rows, carrier))
        except Exception as exc:  # noqa: BLE001 — keep sending even if grouping fails
            log.append(f"Group setup error ({carrier.value}): {exc}")

    for idx, shipment in enumerate(filtered, start=1):
        text = format_tracking_message(tracking_template_for_carrier(shipment.carrier), shipment.voucher)
        try:
            send_one_sms(client, shipment.phone, text, sender=sender)
            log.append(f"Row {idx} ({shipment.voucher} → {shipment.phone}): sent.")
        except EasySMSError as exc:
            log.append(f"Row {idx} ({shipment.voucher} → {shipment.phone}): EasySMS error: {exc}")
        except OSError as exc:
            log.append(f"Row {idx} ({shipment.voucher} → {shipment.phone}): network: {exc}")
    return log
