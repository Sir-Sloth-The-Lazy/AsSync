"""
Offline verification of the sync engine against an in-memory fake Asana.

Proves the four behaviours the daily job depends on:
  1. first run creates every record
  2. second run over unchanged input writes nothing (idempotency)
  3. an edit in Excel produces exactly one targeted update
  4. a row deleted from Excel is flagged, never destroyed

Run:  python tests/test_sync_offline.py
"""
from __future__ import annotations

import sys
from itertools import count
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from asana_sync import AsanaSync                       # noqa: E402
from parse_rise import parse_folder                    # noqa: E402


class FakeAsana:
    """Implements just the surface asana_sync touches, backed by dicts."""

    def __init__(self) -> None:
        self.tasks: dict[str, dict] = {}
        self.sections: dict[str, dict] = {}
        self.membership: dict[str, str] = {}
        self.dry_run = False
        self.reads = 0
        self.writes = 0
        self._gid = count(1000)

    def _new_gid(self) -> str:
        return str(next(self._gid))

    def paginate(self, path: str, params=None, page_size: int = 100):
        self.reads += 1
        if path.endswith("/tasks"):
            project = path.split("/")[2]
            yield from (t for t in self.tasks.values() if project in t["projects"])
        elif path.endswith("/sections"):
            project = path.split("/")[2]
            yield from (s for s in self.sections.values() if s["project"] == project)

    def get(self, path: str, params=None):
        self.reads += 1
        return None

    def post(self, path: str, data: dict):
        self.writes += 1
        if path == "/tasks":
            gid = self._new_gid()
            task = {
                "gid": gid,
                "name": data.get("name"),
                "notes": data.get("notes"),
                "due_on": data.get("due_on"),
                "start_on": data.get("start_on"),
                "completed": False,
                "projects": data.get("projects", []),
                "custom_fields": self._cf_blocks(data.get("custom_fields", {})),
            }
            self.tasks[gid] = task
            return task
        if path.endswith("/sections"):
            project = path.split("/")[2]
            gid = self._new_gid()
            self.sections[gid] = {"gid": gid, "name": data["name"], "project": project}
            return self.sections[gid]
        if "/addTask" in path:
            self.membership[data["task"]] = path.split("/")[2]
            return {}
        return {"gid": self._new_gid()}

    def put(self, path: str, data: dict):
        self.writes += 1
        gid = path.split("/")[-1]
        task = self.tasks[gid]
        for key in ("name", "notes", "due_on", "start_on", "completed"):
            if key in data:
                task[key] = data[key]
        if "custom_fields" in data:
            merged = {b["gid"]: b for b in task["custom_fields"]}
            for b in self._cf_blocks(data["custom_fields"]):
                merged[b["gid"]] = b
            task["custom_fields"] = list(merged.values())
        return task

    # Mirror how Asana echoes custom field values back on reads.
    def _cf_blocks(self, cf: dict) -> list[dict]:
        blocks = []
        for gid, value in cf.items():
            block = {"gid": gid}
            if isinstance(value, dict):
                block["date_value"] = value
            elif isinstance(value, (int, float)):
                block["number_value"] = value
            elif isinstance(value, str) and value.startswith("enum:"):
                block["enum_value"] = {"name": value.split(":", 2)[2]}
            else:
                block["text_value"] = value
            blocks.append(block)
        return blocks


def build_state(records) -> dict:
    """Fabricate the gid map setup_asana.py would normally write."""
    functions = sorted({r.function for r in records})
    priorities = sorted({r.priority_no for r in records})

    def enum(name, options):
        return {"gid": f"cf_{name}",
                "enum_options": {o: f"enum:{name}:{o}" for o in options}}

    return {
        "workspace_gid": "ws_1",
        "portfolio_gid": "pf_1",
        "projects": {fn: f"proj_{i}" for i, fn in enumerate(functions)},
        "custom_fields": {
            "RISE Source ID": {"gid": "cf_srcid", "enum_options": {}},
            "Record Type": enum("rectype", ["Milestone", "KPI"]),
            "Function": enum("function", functions),
            "Priority": enum("priority", [f"P{p}" for p in priorities]),
            "RAG Status": enum("rag", ["Green", "Yellow", "Red", "Not set"]),
            "KPI Type": enum("kpitype", ["Leading", "Lagging"]),
            "Sprint Week": {"gid": "cf_week", "enum_options": {}},
            "Last Synced": {"gid": "cf_synced", "enum_options": {}},
        },
    }


CONFIG = {"stale_policy": "flag", "milestones_as_asana_milestones": True}


def check(label: str, ok: bool, detail: str = "") -> bool:
    print(f"  {'PASS' if ok else 'FAIL'}  {label}{'  — ' + detail if detail else ''}")
    return ok


def main() -> int:
    records, issues = parse_folder(ROOT / "inbox")
    state = build_state(records)
    fake = FakeAsana()
    passed = True

    print(f"\nparsed {len(records)} records from {len({r.source_file for r in records})} files")
    print(f"  {sum(1 for r in records if r.record_kind=='milestone')} milestones, "
          f"{sum(1 for r in records if r.record_kind=='kpi')} KPIs, "
          f"{sum(1 for i in issues if i.severity=='error')} parse errors\n")

    print("run 1 — empty Asana")
    s1 = AsanaSync(fake, state, CONFIG).sync_all(records)
    print(f"  {s1.line()}")
    passed &= check("every record created", s1.created == len(records),
                    f"{s1.created}/{len(records)}")
    passed &= check("no errors", not s1.errors, "; ".join(s1.errors[:2]))

    print("\nrun 2 — identical input (idempotency)")
    writes_before = fake.writes
    s2 = AsanaSync(fake, state, CONFIG).sync_all(records)
    print(f"  {s2.line()}")
    passed &= check("nothing created", s2.created == 0, str(s2.created))
    passed &= check("nothing updated", s2.updated == 0, str(s2.updated))
    passed &= check("no writes sent to Asana", fake.writes == writes_before,
                    f"{fake.writes - writes_before} writes")

    print("\nrun 3 — one milestone edited in Excel")
    target = next(r for r in records if r.record_kind == "milestone")
    target.title = target.title + " (revised)"
    target.week_no = 9
    target.week_end = "2026-09-06"
    s3 = AsanaSync(fake, state, CONFIG).sync_all(records)
    print(f"  {s3.line()}")
    passed &= check("exactly one task updated", s3.updated == 1, str(s3.updated))
    passed &= check("no duplicate created", s3.created == 0, str(s3.created))
    edited = next(t for t in fake.tasks.values() if "(revised)" in (t["name"] or ""))
    passed &= check("due date moved", edited["due_on"] == "2026-09-06", str(edited["due_on"]))

    print("\nrun 4 — a row deleted from Excel (stale_policy=flag)")
    dropped = records.pop()
    s4 = AsanaSync(fake, state, CONFIG).sync_all(records)
    print(f"  {s4.line()}")
    passed &= check("one task flagged stale", s4.stale == 1, str(s4.stale))
    passed &= check("task still exists (never deleted)",
                    any(t["name"] in (dropped.title, f"KPI · {dropped.title}")
                        for t in fake.tasks.values()))

    print("\nsanity checks on generated tasks")
    sample = next(t for t in fake.tasks.values() if t["due_on"])
    passed &= check("due dates land on a Sunday week-end",
                    all(t["due_on"][:2] == "20" for t in fake.tasks.values() if t["due_on"]))
    passed &= check("source id stamped on every task",
                    all(any(b["gid"] == "cf_srcid" and b.get("text_value")
                            for b in t["custom_fields"]) for t in fake.tasks.values()))
    passed &= check("notes cite the source cell",
                    "Source:" in sample["notes"] and "cell" in sample["notes"])

    print(f"\n{'ALL CHECKS PASSED' if passed else 'FAILURES PRESENT'}\n")
    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())
