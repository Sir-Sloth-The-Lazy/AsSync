"""
parse_rise.py — Normalize RISE Q3 Priority workbooks into a canonical record set.

Handles the schema drift observed across the six team files:
  * optional "Owner(s)" column that shifts week columns right
  * header row at row 5 or row 6
  * split priorities ("Priority 1a" / "Priority 1b")
  * merged milestone cells spanning multiple weeks
  * free-text meeting notes pasted below the tracker grid
  * #REF! / error values in Summary rollups

Output: two record streams
  MILESTONE  -> becomes a dated Asana task
  KPI        -> becomes an Asana task carrying RAG status + measurement history

Run standalone to inspect what will be synced:
    python parse_rise.py --input ./inbox --out ./build
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import logging
import re
import unicodedata
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Iterable

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

log = logging.getLogger("rise.parse")

# --------------------------------------------------------------------------
# Sprint calendar. Week N starts on the Monday listed in the sheet header and
# ends the following Sunday. Asana due dates use the week-end date.
# --------------------------------------------------------------------------
SPRINT_W1_MONDAY = dt.date(2026, 7, 6)
SPRINT_WEEKS = 13


def week_bounds(week_no: int) -> tuple[dt.date, dt.date]:
    start = SPRINT_W1_MONDAY + dt.timedelta(weeks=week_no - 1)
    return start, start + dt.timedelta(days=6)


# --------------------------------------------------------------------------
# Header vocabulary. Keys are canonical names; values are accepted labels.
# Matching is done on a normalized (lowercase, whitespace-collapsed) label so
# new files with cosmetic header edits keep working.
# --------------------------------------------------------------------------
COLUMN_SYNONYMS: dict[str, set[str]] = {
    "kpi_name": {"kpi (key performance indicator)", "kpi", "key performance indicator"},
    "kpi_owner": {"owner(s)", "owner", "owners"},
    "kpi_type": {"type", "kpi type"},
    "goal": {"goal (sg)", "goal", "stretch goal"},
    "green": {"green"},
    "yellow": {"yellow"},
    "red": {"red"},
    "kpi_status": {"status"},
    "adjustment": {
        "adjustment action if yellow / red",
        "adjustment action if yellow/red",
        "adjustment action",
    },
}

WEEK_HEADER_RE = re.compile(r"^w\s*(\d{1,2})\b")
PRIORITY_SHEET_RE = re.compile(r"^priority\s*([0-9]+[a-z]?)$", re.IGNORECASE)

# Rows whose first column matches these end the KPI grid.
GRID_TERMINATORS = (
    "cross-functional",
    "cross functional",
    "status key",
)

# Rows that are structural labels rather than KPIs.
TRACKER_ROW_LABELS = {"progress tracker", "actual", "target", "milestones", "milestone"}

# Cell contents that are measurements / placeholders / status words, never
# milestones. Matched on the whole normalized cell value, so a milestone that
# merely contains one of these words is unaffected.
NON_MILESTONE_TOKENS = {
    # placeholders
    "na", "n/a", "n.a.", "-", "--", "tbc", "tbd", "nil", "none", "no data",
    "not applicable", "x", "",
    # status words teams type into week cells instead of the Status column
    "on track", "on-track", "ontrack", "off track", "at risk", "behind",
    "complete", "completed", "done", "closed", "achieved", "met", "not met",
    "in progress", "ongoing", "in-progress", "not started", "pending",
    "delayed", "slipped", "carried over", "no change", "n.a", "yes", "no",
    # program-phase labels
    "start of program", "end of program", "start of programme",
    "end of programme", "mid program", "eop", "sop",
}
RAG_TOKENS = {"g", "y", "r", "green", "yellow", "red", "amber"}

# A week cell holding only a week label ("W4") is a header echo, not a task.
WEEK_LABEL_ONLY_RE = re.compile(r"^w\s*\d{1,2}$")

ERROR_VALUES = {"#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}


def norm(value: Any) -> str:
    """Lowercase, collapse whitespace, strip non-breaking spaces."""
    if value is None:
        return ""
    s = unicodedata.normalize("NFKC", str(value))
    return " ".join(s.split()).strip().lower()


def clean(value: Any) -> str | None:
    """Trim a cell value for storage; return None for blanks and Excel errors."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return str(value)
    s = unicodedata.normalize("NFKC", str(value)).strip()
    if not s or s in ERROR_VALUES:
        return None
    return s


def stable_id(*parts: Any) -> str:
    """Deterministic 16-char id used as the Asana idempotency key."""
    raw = "||".join(norm(p) for p in parts)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]


# --------------------------------------------------------------------------
# Records
# --------------------------------------------------------------------------
@dataclass
class Record:
    record_id: str
    record_kind: str            # "milestone" | "kpi"
    function: str
    function_lead: str | None
    source_file: str
    source_sheet: str
    source_cell: str
    priority_no: str
    priority_statement: str | None
    priority_owner: str | None
    priority_status: str | None
    kpi_name: str | None
    kpi_owner: str | None
    kpi_type: str | None
    goal: str | None
    thresholds: dict[str, str | None]
    kpi_status: str | None
    adjustment: str | None
    title: str
    notes: str | None = None
    week_no: int | None = None
    week_end: str | None = None
    start_on: str | None = None
    measurements: dict[str, str] = field(default_factory=dict)


@dataclass
class Issue:
    severity: str               # "error" | "warning"
    source_file: str
    source_sheet: str | None
    message: str


# --------------------------------------------------------------------------
# Cell classification
# --------------------------------------------------------------------------
def is_measurement(value: Any, column_header: str = "") -> bool:
    """True when a week cell holds a metric reading, status word, or header
    echo rather than an actionable milestone.

    ``column_header`` is the week column's own header text. Several sheets have
    a second header row pasted mid-grid, so a cell that simply repeats its
    column header ("W4 / Jul 27") is structure, not work.
    """
    if column_header and norm(value) == norm(column_header):
        return True
    if value is None:
        return True
    if isinstance(value, (int, float, dt.date, dt.datetime)):
        return True
    s = norm(value)
    if not s or s in NON_MILESTONE_TOKENS or s in RAG_TOKENS:
        return True
    if str(value).strip() in ERROR_VALUES:
        return True
    if WEEK_LABEL_ONLY_RE.fullmatch(s):
        return True
    # Punctuation-only junk (stray backticks, bullets left behind).
    if not re.search(r"[a-z0-9]", s):
        return True
    # A status word with trailing punctuation, e.g. "Completed." / "On track -"
    if s.rstrip(" .:;-–—*") in NON_MILESTONE_TOKENS:
        return True
    # "4.3", "0.83", "85%", "4.3 / 5", "≥4.2", "12 (of 20)"
    if re.fullmatch(r"[<>≥≤~=]{0,2}\s*\d+(\.\d+)?\s*%?(\s*/\s*\d+(\.\d+)?)?", s):
        return True
    # A bare number with a short unit, e.g. "33 responses", "2 workflows"
    if re.fullmatch(r"\d+(\.\d+)?\s*[a-z%]{0,12}", s) and len(s) <= 18:
        return True
    return False


# --------------------------------------------------------------------------
# Sheet structure discovery
# --------------------------------------------------------------------------
@dataclass
class SheetLayout:
    header_row: int
    col_of: dict[str, int]          # canonical name -> column index
    week_cols: dict[int, int]       # week_no -> column index
    first_data_row: int
    last_data_row: int


def find_layout(ws: Worksheet) -> SheetLayout | None:
    header_row = None
    for r in range(1, min(ws.max_row, 20) + 1):
        label = norm(ws.cell(r, 1).value)
        if label in COLUMN_SYNONYMS["kpi_name"]:
            header_row = r
            break
    if header_row is None:
        return None

    col_of: dict[str, int] = {}
    week_cols: dict[int, int] = {}
    for c in range(1, ws.max_column + 1):
        label = norm(ws.cell(header_row, c).value)
        if not label:
            continue
        m = WEEK_HEADER_RE.match(label)
        if m:
            wk = int(m.group(1))
            if 1 <= wk <= SPRINT_WEEKS:
                week_cols[wk] = c
            continue
        for canonical, aliases in COLUMN_SYNONYMS.items():
            if label in aliases and canonical not in col_of:
                col_of[canonical] = c
                break

    first = header_row + 1
    last = ws.max_row
    for r in range(first, ws.max_row + 1):
        label = norm(ws.cell(r, 1).value)
        if any(label.startswith(t) for t in GRID_TERMINATORS):
            last = r - 1
            break

    return SheetLayout(header_row, col_of, week_cols, first, last)


def merged_span(ws: Worksheet, row: int, col: int) -> tuple[int, int]:
    """Return (first_col, last_col) of the merged range anchored at row/col."""
    for rng in ws.merged_cells.ranges:
        if rng.min_row == row and rng.min_col == col:
            return rng.min_col, rng.max_col
    return col, col


# --------------------------------------------------------------------------
# Workbook parsing
# --------------------------------------------------------------------------
def parse_workbook(path: Path) -> tuple[list[Record], list[Issue]]:
    records: list[Record] = []
    issues: list[Issue] = []
    wb = openpyxl.load_workbook(path, data_only=True)
    fname = path.name

    # ---- Summary tab: function name + lead ----
    function = path.stem.replace("RISE_Q3_Priorities_", "").replace("_", " ")
    lead = None
    if "Summary" in wb.sheetnames:
        s = wb["Summary"]
        for r in range(1, min(s.max_row, 12) + 1):
            label = norm(s.cell(r, 1).value)
            if label.startswith("function"):
                function = clean(s.cell(r, 2).value) or function
            elif label.startswith("lead"):
                lead = clean(s.cell(r, 2).value)
        if lead and lead.startswith("["):
            issues.append(Issue("warning", fname, "Summary",
                                f"Function lead is still a placeholder: {lead!r}"))
            lead = None
        for row in s.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.strip() in ERROR_VALUES:
                    issues.append(Issue("error", fname, "Summary",
                                        f"Formula error {c.value} at {c.coordinate} "
                                        f"— Summary rollup is broken in the source file"))
    else:
        issues.append(Issue("warning", fname, None, "No Summary tab; function inferred from filename"))

    # ---- Priority tabs ----
    for ws in wb.worksheets:
        if ws.title == "Summary":
            continue
        m = PRIORITY_SHEET_RE.match(ws.title.strip())
        if not m:
            issues.append(Issue("warning", fname, ws.title, "Unrecognized tab name; skipped"))
            continue
        priority_no = m.group(1)

        layout = find_layout(ws)
        if layout is None:
            issues.append(Issue("error", fname, ws.title, "No KPI header row found; sheet skipped"))
            continue
        if not layout.week_cols:
            issues.append(Issue("error", fname, ws.title, "No week columns (W1..W13) found; sheet skipped"))
            continue

        # Cross-check the banner title against the tab name.
        banner = clean(ws.cell(1, 1).value) or ""
        bm = re.search(r"PRIORITY\s+([0-9]+[a-z]?)", banner, re.IGNORECASE)
        # "Priority 1a" under a "PRIORITY 1" banner is a deliberate split, not drift.
        banner_no = bm.group(1).lower() if bm else None
        split_of_banner = bool(banner_no) and priority_no.lower().startswith(banner_no)
        if bm and not split_of_banner:
            issues.append(Issue(
                "error", fname, ws.title,
                f"Tab is '{ws.title}' but the sheet banner says 'PRIORITY {bm.group(1)}'. "
                f"Likely a copy-paste of another team's tab — verify before syncing."))

        # Priority header block (rows above the grid). Labels sit in col A,
        # values in whichever column follows, so scan across.
        statement = owner = status = None
        for r in range(1, layout.header_row):
            for c in range(1, min(ws.max_column, 10) + 1):
                label = norm(ws.cell(r, c).value)
                if not label:
                    continue
                if label.startswith("we will be successful if"):
                    statement = _first_value_right(ws, r, c)
                elif label.startswith("owner"):
                    owner = _first_value_right(ws, r, c)
                elif label.startswith("overall status"):
                    status = _first_value_right(ws, r, c)

        # Quarantine anything pasted below the grid.
        stray = [r for r in range(layout.last_data_row + 1, ws.max_row + 1)
                 if any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1))]
        stray_below_key = [r for r in stray if r > layout.last_data_row + 6]
        if len(stray_below_key) > 3:
            issues.append(Issue(
                "warning", fname, ws.title,
                f"{len(stray_below_key)} rows of free text below the tracker grid "
                f"(rows {stray_below_key[0]}-{stray_below_key[-1]}) — excluded from sync"))

        ctx = dict(
            function=function, function_lead=lead, source_file=fname,
            source_sheet=ws.title, priority_no=priority_no,
            priority_statement=statement, priority_owner=owner,
            priority_status=_rag(status),
        )

        current_kpi: dict[str, Any] | None = None
        # Sheets legitimately repeat a KPI name across cohorts (DTCM P1a lists
        # "Average weekly content satisfaction" once per intake). An occurrence
        # counter keeps their IDs distinct without pinning IDs to row numbers,
        # which would break every time someone inserts a row.
        occurrence: dict[str, int] = {}

        for r in range(layout.first_data_row, layout.last_data_row + 1):
            name = clean(ws.cell(r, layout.col_of.get("kpi_name", 1)).value)
            label = norm(name)

            is_tracker_row = label in TRACKER_ROW_LABELS
            if name and not is_tracker_row:
                occurrence[label] = occurrence.get(label, 0) + 1
                current_kpi = {
                    "kpi_name": name,
                    "occurrence": occurrence[label],
                    "kpi_owner": _col(ws, r, layout, "kpi_owner") or owner,
                    "kpi_type": _kpi_type(_col(ws, r, layout, "kpi_type")),
                    "goal": _col(ws, r, layout, "goal"),
                    "thresholds": {
                        k: _col(ws, r, layout, k) for k in ("green", "yellow", "red")
                    },
                    "kpi_status": _rag(_col(ws, r, layout, "kpi_status")),
                    "adjustment": _col(ws, r, layout, "adjustment"),
                }
            elif not name and current_kpi is None:
                continue

            kpi = current_kpi or {
                "kpi_name": None, "occurrence": 0, "kpi_owner": owner,
                "kpi_type": None, "goal": None, "thresholds": {},
                "kpi_status": None, "adjustment": None,
            }

            measurements: dict[str, str] = {}

            for wk, c in sorted(layout.week_cols.items()):
                cell = ws.cell(r, c)
                raw = cell.value
                if raw is None:
                    continue
                if is_measurement(raw, norm(ws.cell(layout.header_row, c).value)):
                    v = clean(raw)
                    if v:
                        measurements[f"W{wk}"] = v
                    continue

                text = clean(raw)
                if not text:
                    continue

                # Merged milestone -> real start/end span for the Gantt view.
                first_col, last_col = merged_span(ws, r, c)
                end_wk = max(
                    (w for w, cc in layout.week_cols.items() if first_col <= cc <= last_col),
                    default=wk,
                )
                start_date, _ = week_bounds(wk)
                _, due_date = week_bounds(end_wk)

                title = _title_from(text)
                rid = stable_id(fname, ws.title, kpi["kpi_name"],
                                kpi["occurrence"], wk, text)
                records.append(Record(
                    record_id=rid, record_kind="milestone", **ctx,
                    source_cell=cell.coordinate,
                    kpi_name=kpi["kpi_name"], kpi_owner=kpi["kpi_owner"],
                    kpi_type=kpi["kpi_type"], goal=kpi["goal"],
                    thresholds=kpi["thresholds"], kpi_status=kpi["kpi_status"],
                    adjustment=kpi["adjustment"],
                    title=title,
                    notes=text if text != title else None,
                    week_no=wk,
                    week_end=due_date.isoformat(),
                    start_on=start_date.isoformat(),
                ))

            # Emit one KPI record per real KPI row, carrying its readings.
            if name and not is_tracker_row:
                rid = stable_id(fname, ws.title, "KPI", name, occurrence[label])
                records.append(Record(
                    record_id=rid, record_kind="kpi", **ctx,
                    source_cell=ws.cell(r, 1).coordinate,
                    kpi_name=kpi["kpi_name"], kpi_owner=kpi["kpi_owner"],
                    kpi_type=kpi["kpi_type"], goal=kpi["goal"],
                    thresholds=kpi["thresholds"], kpi_status=kpi["kpi_status"],
                    adjustment=kpi["adjustment"],
                    title=_title_from(name),
                    notes=name,
                    measurements=measurements,
                ))
            elif measurements and current_kpi:
                # Readings on a continuation row: fold into the KPI just emitted.
                for rec in reversed(records):
                    if rec.record_kind == "kpi" and rec.kpi_name == current_kpi["kpi_name"]:
                        rec.measurements.update(measurements)
                        break

    return records, issues


def _first_value_right(ws: Worksheet, row: int, col: int) -> str | None:
    for c in range(col + 1, min(ws.max_column, col + 8) + 1):
        v = clean(ws.cell(row, c).value)
        if v:
            return v
    return None


def _col(ws: Worksheet, row: int, layout: SheetLayout, key: str) -> str | None:
    c = layout.col_of.get(key)
    return clean(ws.cell(row, c).value) if c else None


def _kpi_type(value: Any) -> str | None:
    """Constrain the Type column to its real vocabulary.

    A few sheets have a stray second header row inside the grid, which would
    otherwise set a KPI's type to the literal string "Type".
    """
    s = norm(value)
    if s.startswith("lead"):
        return "Leading"
    if s.startswith("lag"):
        return "Lagging"
    return None


def _rag(value: Any) -> str | None:
    s = norm(value)
    if not s:
        return None
    if s in {"g", "green", "on track"}:
        return "Green"
    if s in {"y", "yellow", "amber", "at risk"}:
        return "Yellow"
    if s in {"r", "red", "off track"}:
        return "Red"
    return None


def _title_from(text: str, limit: int = 110) -> str:
    """Build a readable Asana task name from a multi-line milestone cell.

    Teams write cells like "Go live:\\n(1) New CSA Form\\n(2) Coaching Feedback".
    Taking only the first line yields the useless title "Go live:", so a stub
    first line absorbs the lines that follow until the title has substance.
    """
    lines = [re.sub(r"^\(?\d+[\).]\s*", "", l.strip(" -•\t")).strip()
             for l in str(text).splitlines()]
    lines = [l for l in lines if l]
    if not lines:
        return "Untitled milestone"

    title = lines[0]
    i = 1
    # A line ending in ':' or under ~18 chars is a heading, not a full task name.
    while i < len(lines) and (title.endswith(":") or len(title) < 18) and len(title) < limit:
        sep = " " if title.endswith(":") else "; "
        title = f"{title.rstrip(':')}{sep}{lines[i]}" if title.endswith(":") else f"{title}{sep}{lines[i]}"
        i += 1

    title = " ".join(title.split())
    if len(title) > limit:
        title = title[:limit].rsplit(" ", 1)[0] + "…"
    return title or "Untitled milestone"


def parse_folder(folder: Path) -> tuple[list[Record], list[Issue]]:
    records: list[Record] = []
    issues: list[Issue] = []
    files = sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise SystemExit(f"No .xlsx files found in {folder}")
    for p in files:
        try:
            r, i = parse_workbook(p)
        except Exception as exc:                      # noqa: BLE001
            issues.append(Issue("error", p.name, None, f"Failed to parse: {exc}"))
            continue
        log.info("%s -> %d records", p.name, len(r))
        records.extend(r)
        issues.extend(i)

    seen: dict[str, Record] = {}
    deduped: list[Record] = []
    for rec in records:
        if rec.record_id in seen:
            issues.append(Issue("warning", rec.source_file, rec.source_sheet,
                                f"Duplicate record_id {rec.record_id} — kept first occurrence"))
            continue
        seen[rec.record_id] = rec
        deduped.append(rec)
    return deduped, issues


def main() -> None:
    ap = argparse.ArgumentParser(description="Normalize RISE priority workbooks")
    ap.add_argument("--input", type=Path, default=Path("inbox"))
    ap.add_argument("--out", type=Path, default=Path("build"))
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    records, issues = parse_folder(args.input)
    args.out.mkdir(parents=True, exist_ok=True)
    (args.out / "records.json").write_text(
        json.dumps([asdict(r) for r in records], indent=2, ensure_ascii=False), encoding="utf-8")
    (args.out / "issues.json").write_text(
        json.dumps([asdict(i) for i in issues], indent=2, ensure_ascii=False), encoding="utf-8")

    ms = sum(1 for r in records if r.record_kind == "milestone")
    print(f"{len(records)} records  ({ms} milestones, {len(records)-ms} KPIs)")
    print(f"{sum(1 for i in issues if i.severity=='error')} errors, "
          f"{sum(1 for i in issues if i.severity=='warning')} warnings -> {args.out/'issues.json'}")


if __name__ == "__main__":
    main()
