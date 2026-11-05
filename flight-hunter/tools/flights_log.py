"""Maintain flights.xlsx: one row per price check.

Commands:
    python tools/flights_log.py init
    python tools/flights_log.py append --json "<json>"      (or pipe JSON on stdin)
    python tools/flights_log.py status
    python tools/flights_log.py migrate
    python tools/flights_log.py daily-summary [--force] [--text] [--mark-sent]
    python tools/flights_log.py sources-health [--last 48]

append JSON fields (all optional except notes when nothing was found):
    site, airline, out_flight, out_times, ret_flight, ret_times,
    base_price, bag_fee, total_price, link, notes, usd_rate

If total_price is missing/null the row is recorded as "לא נמצא".
The script stamps the check date and time (Israel time and GMT) itself, fetches the
USD/ILS rate when usd_rate is not given, computes the change from the previous priced
row and reports whether this is a new lowest price.

If flights.xlsx is open in Excel (locked), the row is queued in pending_rows.jsonl
and written automatically on the next successful append.

A table with an older column layout is upgraded automatically (the original is kept
as flights_backup_<timestamp>.xlsx).
"""
import argparse
import json
import re
import shutil
import sys
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
CONFIG = json.loads((ROOT / "tools" / "config.json").read_text(encoding="utf-8"))
XLSX = ROOT / CONFIG["excel_file"]
PENDING = ROOT / "pending_rows.jsonl"
SUMMARY_STATE = ROOT / ".last_daily_summary"
SHEET = "כל הבדיקות"
TZ = ZoneInfo(CONFIG["timezone_display"])
SUMMARY_HOUR = CONFIG["daily_summary"]["hour_israel"]
SUMMARY_WINDOW = CONFIG["daily_summary"]["last_checks"]
SUMMARY_TOP = CONFIG["daily_summary"]["top"]

H_DATE = "תאריך בדיקה"
H_TIME = "שעת בדיקה (שעון ישראל)"
H_GMT = "זמן גריניץ'"
H_BASE = "מחיר בסיס (₪)"
H_USD = "שער הדולר (₪ ל-1$)"
H_TOTAL = "מחיר כולל לשניים (₪)"
H_CHANGE = "שינוי מהבדיקה הקודמת (₪)"
H_LINK = "קישור"
H_NOTES = "הערות"

HEADERS = [
    H_DATE,
    H_TIME,
    H_GMT,
    "אתר",
    "חברת תעופה",
    "מס' טיסה הלוך",
    "המראה/נחיתה הלוך",
    "מס' טיסה חזור",
    "המראה/נחיתה חזור",
    H_BASE,
    H_USD,
    "תוספת מזוודה (₪)",
    H_TOTAL,
    H_CHANGE,
    H_LINK,
    H_NOTES,
]
WIDTHS = [13, 12, 12, 16, 18, 14, 22, 14, 22, 14, 14, 14, 18, 18, 40, 45]
MONEY_COLUMNS = [H_BASE, "תוספת מזוודה (₪)", H_TOTAL, H_CHANGE]
COL = {name: i + 1 for i, name in enumerate(HEADERS)}

BLOCK_WORDS = ["חסום", "חסומ", "נחסם", "לא זמין", "blocked", "captcha", "fetch_failed"]


def _style_header(ws) -> None:
    for i, width in enumerate(WIDTHS, start=1):
        cell = ws.cell(row=1, column=i)
        ws.column_dimensions[cell.column_letter].width = width
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def _new_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.sheet_view.rightToLeft = True
    ws.append(HEADERS)
    _style_header(ws)
    return wb


def _gmt_from_israel(date_value, time_value) -> str | None:
    try:
        local = datetime.strptime(f"{date_value} {time_value}", "%d/%m/%Y %H:%M").replace(tzinfo=TZ)
    except (TypeError, ValueError):
        return None
    return local.astimezone(timezone.utc).strftime("%H:%M")


def _format_row(ws, row_idx: int, link: str | None) -> None:
    if link:
        cell = ws.cell(row=row_idx, column=COL[H_LINK])
        cell.hyperlink = link
        cell.font = Font(color="0563C1", underline="single")
    for name in MONEY_COLUMNS:
        ws.cell(row=row_idx, column=COL[name]).number_format = "#,##0"
    ws.cell(row=row_idx, column=COL[H_USD]).number_format = "0.000"


def _norm(header) -> str:
    text = str(header or "").strip().lower().replace("ץ", "צ")
    return re.sub(r"[\s'\"׳״`]", "", text)


def _match_header(header) -> str | None:
    """Map a column name from an older or hand-edited table to a current column (None = unknown)."""
    if header in HEADERS:
        return header
    norm = _norm(header)
    for name in HEADERS:
        if _norm(name) == norm:
            return name
    if "דולר" in norm or "usd" in norm:
        return H_USD
    if "גריני" in norm or "gmt" in norm or "utc" in norm:
        return H_GMT
    return None


def _upgrade(path: Path) -> dict:
    """Bring the table to the current column layout, matching columns by name.

    Columns this script does not know (e.g. added by hand) are kept, after the known columns.
    Other sheets in the workbook are left untouched. A backup copy is saved first.
    """
    wb = load_workbook(path)
    old_ws = wb[SHEET] if SHEET in wb.sheetnames else wb.worksheets[0]
    old_headers = [c.value for c in old_ws[1]]
    if old_headers[: len(HEADERS)] == HEADERS:
        return {"upgraded": False}

    backup = path.with_name(f"flights_backup_{datetime.now(TZ).strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(path, backup)

    mapping = {i: _match_header(h) for i, h in enumerate(old_headers) if h not in (None, "")}
    extras = []
    for i, target in mapping.items():
        if target is None and old_headers[i] not in extras:
            extras.append(old_headers[i])
    layout = HEADERS + extras

    position = wb.sheetnames.index(old_ws.title)
    ws = wb.create_sheet("upgrading", position)
    ws.sheet_view.rightToLeft = True
    ws.append(layout)
    _style_header(ws)
    for col in range(len(HEADERS) + 1, len(layout) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        ws.column_dimensions[cell.column_letter].width = 18

    moved = 0
    for old_row in old_ws.iter_rows(min_row=2):
        values, link = {}, None
        for i, target in mapping.items():
            if i >= len(old_row):
                continue
            cell = old_row[i]
            name = target or old_headers[i]
            if values.get(name) in (None, ""):
                values[name] = cell.value
            if target == H_LINK and cell.hyperlink:
                link = cell.hyperlink.target
        if all(v in (None, "") for v in values.values()):
            continue
        if values.get(H_LINK) and not link:
            link = values[H_LINK]
        if not values.get(H_GMT):
            values[H_GMT] = _gmt_from_israel(values.get(H_DATE), values.get(H_TIME))
        ws.append([values.get(name) for name in layout])
        _format_row(ws, ws.max_row, link)
        moved += 1

    wb.remove(old_ws)
    ws.title = SHEET
    wb.active = position
    wb.save(path)
    return {
        "upgraded": True,
        "rows": moved,
        "backup": backup.name,
        "old_columns": old_headers,
        "kept_extra_columns": extras,
    }


def _load() -> Workbook:
    if XLSX.exists():
        _upgrade(XLSX)
        return load_workbook(XLSX)
    return _new_workbook()


def _num(value):
    if value in (None, ""):
        return None
    try:
        return round(float(str(value).replace(",", "").replace("₪", "").strip()), 3)
    except ValueError:
        return None


def _rows(ws) -> list[dict]:
    rows = []
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(v in (None, "") for v in row):
            continue
        rows.append({"row": idx, **{name: row[COL[name] - 1] for name in HEADERS}})
    return rows


def _priced_totals(ws) -> list[float]:
    return [t for t in (_num(r[H_TOTAL]) for r in _rows(ws)) if t is not None]


def fetch_usd_rate() -> tuple[float | None, str]:
    """USD->ILS rate: Bank of Israel representative rate, falling back to Frankfurter (ECB)."""
    sources = [
        ("בנק ישראל", "https://boi.org.il/PublicApi/GetExchangeRate?key=USD",
         lambda d: d.get("currentExchangeRate")),
        ("Frankfurter/ECB", "https://api.frankfurter.app/latest?from=USD&to=ILS",
         lambda d: d.get("rates", {}).get("ILS")),
    ]
    for name, url, pick in sources:
        try:
            req = urllib.request.Request(url, headers={"Accept": "application/json", "User-Agent": "flight-hunter"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                rate = _num(pick(json.loads(resp.read().decode("utf-8"))))
            if rate:
                return rate, name
        except Exception:  # noqa: BLE001 - try the next source
            continue
    return None, ""


def _write_row(ws, data: dict, stamp: datetime) -> dict:
    totals = _priced_totals(ws)
    previous_total = totals[-1] if totals else None
    previous_best = min(totals) if totals else None

    total = _num(data.get("total_price"))
    notes = (data.get("notes") or "").strip()
    if total is None and "לא נמצא" not in notes:
        notes = ("לא נמצא. " + notes).strip()
    usd_rate = _num(data.get("usd_rate"))
    if usd_rate is None:
        notes = (notes + " שער הדולר לא זמין.").strip()

    change = None
    if total is not None and previous_total is not None:
        change = round(total - previous_total, 2)

    values = {
        H_DATE: stamp.strftime("%d/%m/%Y"),
        H_TIME: stamp.strftime("%H:%M"),
        H_GMT: stamp.astimezone(timezone.utc).strftime("%H:%M"),
        "אתר": data.get("site") or "",
        "חברת תעופה": data.get("airline") or "",
        "מס' טיסה הלוך": data.get("out_flight") or "",
        "המראה/נחיתה הלוך": data.get("out_times") or "",
        "מס' טיסה חזור": data.get("ret_flight") or "",
        "המראה/נחיתה חזור": data.get("ret_times") or "",
        H_BASE: _num(data.get("base_price")),
        H_USD: usd_rate,
        "תוספת מזוודה (₪)": _num(data.get("bag_fee")),
        H_TOTAL: total,
        H_CHANGE: change,
        H_LINK: data.get("link") or "",
        H_NOTES: notes,
    }
    ws.append([values[name] for name in HEADERS])
    row_idx = ws.max_row
    _format_row(ws, row_idx, data.get("link"))

    return {
        "row": row_idx,
        "checked_at": stamp.strftime("%d/%m/%Y %H:%M"),
        "checked_at_gmt": values[H_GMT],
        "total_price": total,
        "usd_rate": usd_rate,
        "previous_total": previous_total,
        "change": change,
        "previous_best": previous_best,
        "first_price": total is not None and previous_best is None,
        "new_low": total is not None and previous_best is not None and total < previous_best,
    }


def cmd_init() -> dict:
    if XLSX.exists():
        return {"created": False, "file": str(XLSX), **_upgrade(XLSX)}
    _new_workbook().save(XLSX)
    return {"created": True, "file": str(XLSX)}


def cmd_migrate() -> dict:
    if not XLSX.exists():
        return {"exists": False}
    return _upgrade(XLSX)


def cmd_append(data: dict) -> dict:
    stamp = datetime.now(TZ)
    if data.get("usd_rate") is None:
        data["usd_rate"], data["usd_rate_source"] = fetch_usd_rate()
    queued = []
    if PENDING.exists():
        queued = [json.loads(line) for line in PENDING.read_text(encoding="utf-8").splitlines() if line.strip()]

    try:
        wb = _load()
    except PermissionError:
        wb = None
    result = None
    if wb is not None:
        ws = wb[SHEET]
        for item in queued:
            _write_row(ws, item["data"], datetime.fromisoformat(item["stamp"]))
        result = _write_row(ws, data, stamp)
        try:
            wb.save(XLSX)
        except PermissionError:
            wb = None
    if wb is None:
        with PENDING.open("a", encoding="utf-8") as f:
            f.write(json.dumps({"stamp": stamp.isoformat(), "data": data}, ensure_ascii=False) + "\n")
        result = result or {"checked_at": stamp.strftime("%d/%m/%Y %H:%M"), "new_low": False, "first_price": False}
        result["saved"] = False
        result["warning"] = "flights.xlsx is open/locked; row queued in pending_rows.jsonl"
        return result
    if PENDING.exists():
        PENDING.unlink()
    result["saved"] = True
    result["flushed_pending_rows"] = len(queued)
    return result


def cmd_status() -> dict:
    if not XLSX.exists():
        return {"exists": False}
    ws = _load()[SHEET]
    totals = _priced_totals(ws)
    return {
        "exists": True,
        "checks": len(_rows(ws)),
        "priced_checks": len(totals),
        "last_total": totals[-1] if totals else None,
        "best_total": min(totals) if totals else None,
        "pending_rows": PENDING.exists(),
    }


def _fmt_money(value) -> str:
    return f"₪{value:,.0f}"


def build_summary(rows: list[dict]) -> tuple[str, list[dict]]:
    window = rows[-SUMMARY_WINDOW:]
    priced = [r for r in window if _num(r[H_TOTAL]) is not None]
    # One entry per distinct offer (same price, airline and flights), keeping its latest check.
    distinct = {}
    for r in priced:
        key = (_num(r[H_TOTAL]), r["חברת תעופה"], r["מס' טיסה הלוך"], r["מס' טיסה חזור"])
        distinct[key] = r
    best = sorted(distinct.values(), key=lambda r: (_num(r[H_TOTAL]), -r["row"]))[:SUMMARY_TOP]

    lines = ["✈️ צייד הטיסות: סיכום בוקר",
             "תל אביב ⇄ הרקליון | 15–21.8.2027 | ישירה"]
    if window:
        lines.append(f"{len(best) if best else 'אין'} המחירים הזולים מ-{len(window)} הבדיקות האחרונות "
                     f"({window[0][H_DATE]} {window[0][H_TIME]} עד {window[-1][H_DATE]} {window[-1][H_TIME]}):")
    if not best:
        lines.append("לא נמצא מחיר באף אחת מהבדיקות האלה.")
    for i, r in enumerate(best, start=1):
        legs = [
            (r["מס' טיסה הלוך"], r["המראה/נחיתה הלוך"]),
            (r["מס' טיסה חזור"], r["המראה/נחיתה חזור"]),
        ]
        flights = " / ".join(x for x in (f"{num or ''} {times or ''}".strip() for num, times in legs) if x)
        lines.append("")
        lines.append(f"{i}. 💰 {_fmt_money(_num(r[H_TOTAL]))} לשניים | {r['חברת תעופה'] or 'חברה לא צוינה'}")
        if flights:
            lines.append(f"   🛫 {flights}")
        lines.append(f"   🔎 {r['אתר'] or ''} | נבדק {r[H_DATE]} {r[H_TIME]} (ישראל) / {r[H_GMT] or '?'} GMT")
        if r[H_LINK]:
            lines.append(f"   🔗 {r[H_LINK]}")
    missing = len(window) - len(priced)
    if window and missing and best:
        lines.append("")
        lines.append(f"ב-{missing} מתוך {len(window)} בדיקות לא נמצא מחיר.")
    return "\n".join(lines), best


def cmd_daily_summary(force: bool, mark_sent: bool) -> dict:
    now = datetime.now(TZ)
    today = now.strftime("%Y-%m-%d")
    if mark_sent:
        SUMMARY_STATE.write_text(today, encoding="utf-8")
        return {"marked_sent": today}
    last_sent = SUMMARY_STATE.read_text(encoding="utf-8").strip() if SUMMARY_STATE.exists() else ""
    due = force or (now.hour >= SUMMARY_HOUR and last_sent != today)
    out = {"due": due, "last_sent": last_sent or None, "israel_time": now.strftime("%d/%m/%Y %H:%M")}
    if due:
        rows = _rows(_load()[SHEET]) if XLSX.exists() else []
        out["message"], best = build_summary(rows)
        out["prices"] = [_num(r[H_TOTAL]) for r in best]
    return out


def cmd_sources_health(last: int) -> dict:
    """Count, per source, how many of the last checks mention it as blocked/unavailable."""
    names = [s["name"] for s in CONFIG["sources"]] + [s["name"] for s in CONFIG.get("removed_sources", [])]
    rows = _rows(_load()[SHEET])[-last:] if XLSX.exists() else []
    report = {}
    for name in names:
        mentioned = blocked = 0
        for r in rows:
            notes = str(r[H_NOTES] or "")
            for part in re.split(r"[.;\n]", notes):
                if name.lower() in part.lower():
                    mentioned += 1
                    if any(w in part.lower() for w in BLOCK_WORDS):
                        blocked += 1
                    break
        found_on_site = sum(
            1 for r in rows
            if _num(r[H_TOTAL]) is not None and name.lower() in str(r["אתר"] or "").lower()
        )
        active = name in [s["name"] for s in CONFIG["sources"]]
        report[name] = {
            "active": active,
            "checks": len(rows),
            "blocked_mentions": blocked,
            "cheapest_found_here": found_on_site,
            "always_blocked": bool(rows) and blocked >= 0.9 * len(rows) and found_on_site == 0,
        }
    return report


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="cmd", required=True)
    sub.add_parser("init")
    sub.add_parser("status")
    sub.add_parser("migrate")
    ap = sub.add_parser("append")
    ap.add_argument("--json", help="row data as JSON; read from stdin when omitted")
    sp = sub.add_parser("daily-summary")
    sp.add_argument("--force", action="store_true", help="build the summary even if not due")
    sp.add_argument("--mark-sent", action="store_true", help="record that today's summary was sent")
    sp.add_argument("--text", action="store_true", help="print only the message text (empty when not due)")
    hp = sub.add_parser("sources-health")
    hp.add_argument("--last", type=int, default=48)
    args = parser.parse_args()

    if args.cmd == "init":
        out = cmd_init()
    elif args.cmd == "status":
        out = cmd_status()
    elif args.cmd == "migrate":
        out = cmd_migrate()
    elif args.cmd == "daily-summary":
        out = cmd_daily_summary(args.force, args.mark_sent)
        if args.text:
            print(out.get("message", ""))
            return 0
    elif args.cmd == "sources-health":
        out = cmd_sources_health(args.last)
    else:
        raw = args.json if args.json else sys.stdin.read()
        out = cmd_append(json.loads(raw))
    print(json.dumps(out, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
