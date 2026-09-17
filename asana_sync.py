"""
asana_sync.py — Idempotent upsert of parsed records into Asana.

The contract: running this twice in a row makes zero changes on the second
run. Every task carries its RISE Source ID in a custom field; the sync builds
an index of existing tasks from that field and then creates, updates, or
leaves each record alone. Nothing is ever deleted.

Handles three cases per record:
  new in Excel      -> create task
  changed in Excel  -> update only the fields that differ
  gone from Excel   -> flagged per config.stale_policy (never hard-deleted)
"""
from __future__ import annotations

import datetime as dt
import logging
from dataclasses import dataclass, field as dc_field
from typing import Any, Iterable

from asana_client import AsanaClient
from parse_rise import Record

log = logging.getLogger("rise.sync")

TASK_FIELDS = ("name,notes,completed,due_on,start_on,memberships.section.name,"
               "custom_fields.gid,custom_fields.name,custom_fields.display_value,"
               "custom_fields.text_value,custom_fields.number_value,"
               "custom_fields.enum_value.name,custom_fields.date_value")

STALE_SECTION = "⚠ No longer in source file"


@dataclass
class SyncStats:
    created: int = 0
    updated: int = 0
    unchanged: int = 0
    stale: int = 0
    skipped: int = 0
    errors: list[str] = dc_field(default_factory=list)

    def line(self) -> str:
        return (f"created={self.created} updated={self.updated} "
                f"unchanged={self.unchanged} stale={self.stale} "
                f"skipped={self.skipped} errors={len(self.errors)}")


class AsanaSync:
    def __init__(self, client: AsanaClient, state: dict, config: dict):
        self.client = client
        self.state = state
        self.config = config
        self.fields = state["custom_fields"]
        self.source_id_gid = self.fields["RISE Source ID"]["gid"]
        self.stats = SyncStats()
        self._section_cache: dict[tuple[str, str], str] = {}

    # ------------------------------------------------------------------
    # helpers
    # ------------------------------------------------------------------
    def _enum(self, field_name: str, value: str | None) -> str | None:
        """Resolve an enum option name to its gid, tolerating unknown values."""
        if not value:
            return None
        options = self.fields.get(field_name, {}).get("enum_options", {})
        gid = options.get(value)
        if gid is None:
            log.warning("enum option %r missing on field %r — left blank "
                        "(re-run setup_asana.py to add it)", value, field_name)
        return gid

    def _section(self, project_gid: str, name: str) -> str | None:
        key = (project_gid, name)
        if key in self._section_cache:
            return self._section_cache[key]
        for s in self.client.paginate(f"/projects/{project_gid}/sections",
                                      {"opt_fields": "name"}):
            self._section_cache[(project_gid, s["name"])] = s["gid"]
        if key not in self._section_cache:
            log.info("creating section %r", name)
            created = self.client.post(f"/projects/{project_gid}/sections", {"name": name})
            self._section_cache[key] = created["gid"]
        return self._section_cache[key]

    def _index(self, project_gid: str) -> dict[str, dict]:
        """Map RISE Source ID -> existing task, for one project."""
        index: dict[str, dict] = {}
        for task in self.client.paginate(f"/projects/{project_gid}/tasks",
                                         {"opt_fields": TASK_FIELDS}):
            for cf in task.get("custom_fields", []):
                if cf.get("gid") == self.source_id_gid and cf.get("text_value"):
                    index[cf["text_value"]] = task
                    break
        return index

    # ------------------------------------------------------------------
    # payload construction
    # ------------------------------------------------------------------
    def _notes(self, rec: Record) -> str:
        lines: list[str] = []
        if rec.priority_statement:
            lines.append(f"PRIORITY {rec.priority_no}: {rec.priority_statement}")
        if rec.kpi_name:
            lines.append(f"KPI: {rec.kpi_name}")
        if rec.kpi_type:
            lines.append(f"KPI type: {rec.kpi_type}")
        if rec.goal:
            lines.append(f"Goal: {rec.goal}")

        th = {k: v for k, v in (rec.thresholds or {}).items() if v}
        if th:
            lines.append("Thresholds: " + "  |  ".join(f"{k.title()}: {v}" for k, v in th.items()))

        if rec.record_kind == "milestone" and rec.notes and rec.notes != rec.title:
            lines += ["", "Full text from the sheet:", rec.notes]

        if rec.measurements:
            lines += ["", "Weekly readings:"]
            lines += [f"  {wk}: {val}" for wk, val in
                      sorted(rec.measurements.items(), key=lambda kv: int(kv[0][1:]))]

        if rec.adjustment:
            lines += ["", f"Adjustment if Yellow/Red: {rec.adjustment}"]

        lines += ["", "—",
                  f"Source: {rec.source_file} · {rec.source_sheet} · cell {rec.source_cell}",
                  f"Owner in sheet: {rec.kpi_owner or rec.priority_owner or 'unassigned'}",
                  "Managed by rise-asana-sync. Edits here are overwritten on the next run; "
                  "change the Excel file instead."]
        return "\n".join(lines)

    def _custom_fields(self, rec: Record) -> dict[str, Any]:
        cf: dict[str, Any] = {self.source_id_gid: rec.record_id}

        def put(name: str, value: Any) -> None:
            gid = self.fields.get(name, {}).get("gid")
            if gid and value is not None:
                cf[gid] = value

        put("Record Type", self._enum("Record Type",
                                      "Milestone" if rec.record_kind == "milestone" else "KPI"))
        put("Function", self._enum("Function", rec.function))
        put("Priority (RISE)", self._enum("Priority (RISE)", f"P{rec.priority_no}"))
        put("RAG Status", self._enum("RAG Status",
                                     rec.kpi_status or rec.priority_status or "Not set"))
        put("KPI Type", self._enum("KPI Type", rec.kpi_type.title() if rec.kpi_type else None))
        put("Sprint Week", rec.week_no)
        put("Last Synced", {"date": dt.date.today().isoformat()})
        return cf

    def _task_name(self, rec: Record) -> str:
        if rec.record_kind == "kpi":
            return f"KPI · {rec.title}"
        return rec.title

    def _desired(self, rec: Record, project_gid: str) -> dict[str, Any]:
        data: dict[str, Any] = {
            "name": self._task_name(rec),
            "notes": self._notes(rec),
            "projects": [project_gid],
            "custom_fields": self._custom_fields(rec),
        }
        if rec.record_kind == "milestone":
            data["due_on"] = rec.week_end
            if rec.start_on and rec.start_on != rec.week_end:
                data["start_on"] = rec.start_on
            data["resource_subtype"] = "milestone" if self.config.get(
                "milestones_as_asana_milestones") else "default_task"
        return data

    # ------------------------------------------------------------------
    # diffing
    # ------------------------------------------------------------------
    @staticmethod
    def _current_cf(task: dict) -> dict[str, Any]:
        out: dict[str, Any] = {}
        for cf in task.get("custom_fields", []):
            gid = cf.get("gid")
            if cf.get("enum_value"):
                out[gid] = cf["enum_value"].get("name")
            elif cf.get("text_value") is not None:
                out[gid] = cf["text_value"]
            elif cf.get("number_value") is not None:
                out[gid] = cf["number_value"]
            elif cf.get("date_value"):
                out[gid] = cf["date_value"].get("date")
        return out

    def _diff(self, task: dict, desired: dict) -> dict[str, Any]:
        """Return only the fields that actually changed."""
        patch: dict[str, Any] = {}
        for key in ("name", "notes", "due_on", "start_on"):
            if key in desired and (task.get(key) or None) != (desired[key] or None):
                patch[key] = desired[key]

        current = self._current_cf(task)
        cf_patch: dict[str, Any] = {}
        last_synced_gid = self.fields.get("Last Synced", {}).get("gid")
        for gid, value in desired.get("custom_fields", {}).items():
            if gid == last_synced_gid:
                continue                       # would force a write every run
            want = value
            if isinstance(value, dict):        # date field
                want = value.get("date")
            else:
                # Enum gids were resolved on the way in; compare by option name.
                for fname, meta in self.fields.items():
                    if meta.get("gid") == gid and meta.get("enum_options"):
                        want = next((n for n, g in meta["enum_options"].items() if g == value), value)
                        break
            if str(current.get(gid)) != str(want):
                cf_patch[gid] = value
        if cf_patch:
            if last_synced_gid:
                cf_patch[last_synced_gid] = desired["custom_fields"].get(last_synced_gid)
            patch["custom_fields"] = cf_patch
        return patch

    # ------------------------------------------------------------------
    # main entry point
    # ------------------------------------------------------------------
    def sync_function(self, function: str, records: list[Record]) -> None:
        project_gid = self.state["projects"].get(function)
        if not project_gid:
            self.stats.errors.append(
                f"No Asana project mapped for function {function!r} — run setup_asana.py")
            self.stats.skipped += len(records)
            return

        log.info("syncing %s (%d records) -> project %s", function, len(records), project_gid)
        index = self._index(project_gid)
        seen: set[str] = set()

        for rec in records:
            seen.add(rec.record_id)
            desired = self._desired(rec, project_gid)
            existing = index.get(rec.record_id)
            section_name = f"Priority {rec.priority_no}"

            try:
                if existing is None:
                    payload = dict(desired)
                    payload["workspace"] = self.state["workspace_gid"]
                    created = self.client.post("/tasks", payload)
                    section = self._section(project_gid, section_name)
                    if section and not self.client.dry_run:
                        self.client.post(f"/sections/{section}/addTask",
                                         {"task": created["gid"]})
                    self.stats.created += 1
                    log.debug("created %s", desired["name"][:60])
                else:
                    patch = self._diff(existing, desired)
                    if patch:
                        self.client.put(f"/tasks/{existing['gid']}", patch)
                        self.stats.updated += 1
                        log.debug("updated %s (%s)", desired["name"][:50],
                                  ", ".join(patch.keys()))
                    else:
                        self.stats.unchanged += 1
            except Exception as exc:                        # noqa: BLE001
                self.stats.errors.append(f"{rec.record_id} {rec.title[:40]}: {exc}")

        self._handle_stale(project_gid, index, seen)

    def _handle_stale(self, project_gid: str, index: dict[str, dict],
                      seen: set[str]) -> None:
        policy = self.config.get("stale_policy", "flag")
        orphans = [t for sid, t in index.items() if sid not in seen]
        if not orphans or policy == "ignore":
            return
        log.info("%d task(s) no longer in the source files (policy=%s)",
                 len(orphans), policy)
        for task in orphans:
            self.stats.stale += 1
            try:
                if policy == "complete":
                    if not task.get("completed"):
                        self.client.put(f"/tasks/{task['gid']}", {"completed": True})
                else:  # flag
                    section = self._section(project_gid, STALE_SECTION)
                    if section:
                        self.client.post(f"/sections/{section}/addTask", {"task": task["gid"]})
            except Exception as exc:                        # noqa: BLE001
                self.stats.errors.append(f"stale {task.get('gid')}: {exc}")

    def sync_all(self, records: Iterable[Record]) -> SyncStats:
        by_function: dict[str, list[Record]] = {}
        for rec in records:
            by_function.setdefault(rec.function, []).append(rec)
        for function, recs in sorted(by_function.items()):
            self.sync_function(function, recs)
        return self.stats
