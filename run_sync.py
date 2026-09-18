"""
run_sync.py — The command the scheduler runs. Parse, sync, report.

    python run_sync.py --dry-run      # parse and show what would change
    python run_sync.py                # parse and write to Asana
    python run_sync.py --parse-only   # no network at all

Exit codes: 0 clean · 1 parse errors present and fail_on_parse_error is true
            · 2 sync errors · 3 configuration problem
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import logging
import os
import sys
from dataclasses import asdict
from logging.handlers import RotatingFileHandler
from pathlib import Path

import yaml
from dotenv import load_dotenv

from asana_client import AsanaClient, AsanaError
from asana_sync import AsanaSync
from parse_rise import parse_folder


def setup_logging(log_dir: Path, verbose: bool) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    handler = RotatingFileHandler(log_dir / "sync.log", maxBytes=2_000_000,
                                  backupCount=5, encoding="utf-8")
    handler.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)-7s %(name)s %(message)s"))
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(logging.Formatter("%(levelname)-7s %(message)s"))
    logging.basicConfig(level=logging.DEBUG if verbose else logging.INFO,
                        handlers=[handler, console])


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", type=Path, default=Path(__file__).parent / "config.yaml")
    ap.add_argument("--dry-run", action="store_true",
                    help="read from Asana, log writes without sending them")
    ap.add_argument("--parse-only", action="store_true", help="no Asana calls at all")
    ap.add_argument("--only-function", help="sync a single function, for testing")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    root = args.config.parent
    cfg = yaml.safe_load(args.config.read_text(encoding="utf-8"))
    setup_logging(root / cfg.get("log_folder", "logs"), args.verbose)
    log = logging.getLogger("rise.run")
    started = dt.datetime.now()

    # ---------------- parse ----------------
    inbox = (root / cfg["input_folder"]).resolve()
    log.info("parsing %s", inbox)
    records, issues = parse_folder(inbox)
    errors = [i for i in issues if i.severity == "error"]
    warnings = [i for i in issues if i.severity == "warning"]

    build = root / cfg.get("build_folder", "build")
    build.mkdir(parents=True, exist_ok=True)
    (build / "records.json").write_text(
        json.dumps([asdict(r) for r in records], indent=2, ensure_ascii=False), encoding="utf-8")
    (build / "issues.json").write_text(
        json.dumps([asdict(i) for i in issues], indent=2, ensure_ascii=False), encoding="utf-8")

    milestones = sum(1 for r in records if r.record_kind == "milestone")
    log.info("parsed %d records (%d milestones, %d KPIs) from %d functions",
             len(records), milestones, len(records) - milestones,
             len({r.function for r in records}))
    for i in errors:
        log.error("%s | %s | %s", i.source_file, i.source_sheet, i.message)
    for i in warnings:
        log.warning("%s | %s | %s", i.source_file, i.source_sheet, i.message)

    if errors and cfg.get("fail_on_parse_error", False):
        log.error("aborting: %d parse error(s) and fail_on_parse_error is true", len(errors))
        return 1
    if args.parse_only:
        log.info("--parse-only: stopping before Asana")
        return 0

    if args.only_function:
        records = [r for r in records if r.function == args.only_function]
        log.info("--only-function %s -> %d records", args.only_function, len(records))

    # ---------------- sync ----------------
    load_dotenv(root / ".env")
    state_path = root / cfg.get("state_folder", "state") / "asana_ids.json"
    if not state_path.exists():
        log.error("missing %s — run setup_asana.py first", state_path)
        return 3
    state = json.loads(state_path.read_text(encoding="utf-8"))

    try:
        client = AsanaClient(os.getenv("ASANA_TOKEN", ""), dry_run=args.dry_run)
        stats = AsanaSync(client, state, cfg).sync_all(records)
    except AsanaError as exc:
        log.error("Asana error: %s", exc)
        return 2

    elapsed = (dt.datetime.now() - started).total_seconds()
    log.info("sync complete in %.1fs — %s", elapsed, stats.line())
    log.info("api calls: %d reads, %d writes%s", client.reads, client.writes,
             " (dry run, nothing sent)" if args.dry_run else "")
    for e in stats.errors[:20]:
        log.error("sync error: %s", e)

    report = {
        "run_at": started.isoformat(timespec="seconds"),
        "elapsed_seconds": round(elapsed, 1),
        "dry_run": args.dry_run,
        "records": len(records),
        "milestones": milestones,
        "parse_errors": [asdict(i) for i in errors],
        "parse_warnings": len(warnings),
        "sync": asdict(stats),
    }
    (build / "last_run.json").write_text(json.dumps(report, indent=2, ensure_ascii=False),
                                         encoding="utf-8")
    return 2 if stats.errors else 0


if __name__ == "__main__":
    sys.exit(main())
