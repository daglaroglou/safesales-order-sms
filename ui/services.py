"""Business logic for the SafeSales WinUI app (EasySMS + Excel parsing)."""

from __future__ import annotations

import os
import traceback
from dataclasses import dataclass
from pathlib import Path

try:
    import dotenv as _dotenv
except ModuleNotFoundError:  # pragma: no cover
    _dotenv = None

from easysms import EasySMSClient, EasySMSError
from excel.loader import parse_excel_file
from excel.models import Shipment


def load_dotenv_if_present() -> None:
    if _dotenv is not None:
        _dotenv.load_dotenv()


def get_api_key() -> str | None:
    load_dotenv_if_present()
    key = (os.environ.get("API_KEY") or "").strip()
    return key or None


def get_client() -> EasySMSClient | None:
    key = get_api_key()
    if not key:
        return None
    return EasySMSClient(api_key=key)


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


def process_excel_paths(paths: list[Path]) -> ProcessResult:
    lines: list[str] = []
    all_shipments: list[Shipment] = []

    for path in paths:
        if not path or not str(path).strip():
            continue
        p = Path(path)
        try:
            if not p.is_file():
                lines.append(f"{p.name}  ·  skipped (not a file)")
                continue
            shipments = parse_excel_file(p)
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
        return "BE"
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
