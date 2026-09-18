"""
setup_asana.py — One-time scaffolding of the Asana workspace.

Creates (idempotently, by name):
  * 8 custom fields, including the RISE Source ID text field the sync uses
    as its idempotency key
  * one project per function, each with those fields attached
  * one portfolio holding all the projects, for the roll-up dashboard

Safe to re-run: anything that already exists by name is reused, not duplicated.
It writes every resolved gid to state/asana_ids.json, which asana_sync.py reads.

    python setup_asana.py --dry-run      # show what it would create
    python setup_asana.py                # create for real
"""
from __future__ import annotations

import argparse
import json
import logging
import os
from pathlib import Path

import yaml
from dotenv import load_dotenv

from asana_client import AsanaClient, AsanaError
from parse_rise import parse_folder

log = logging.getLogger("rise.setup")

# name -> (resource_subtype, enum options)
CUSTOM_FIELDS: dict[str, tuple[str, list[str]]] = {
    "RISE Source ID": ("text", []),
    "Record Type":    ("enum", ["Milestone", "KPI"]),
    "Function":       ("enum", []),           # filled from the workbooks
    "Priority (RISE)": ("enum", []),          # filled from the workbooks
    "RAG Status":     ("enum", ["Green", "Yellow", "Red", "Not set"]),
    "KPI Type":       ("enum", ["Leading", "Lagging"]),
    "Sprint Week":    ("number", []),
    "Last Synced":    ("date", []),
}

ENUM_COLORS = ["green", "yellow", "red", "blue", "purple", "orange",
               "aqua", "magenta", "cool-gray", "hot-pink", "indigo", "yellow-green"]


def find_by_name(items: list[dict], name: str) -> dict | None:
    target = name.strip().lower()
    return next((i for i in items if str(i.get("name", "")).strip().lower() == target), None)


def ensure_custom_field(client: AsanaClient, workspace: str, name: str,
                        subtype: str, options: list[str],
                        existing: list[dict]) -> dict:
    field = find_by_name(existing, name)
    if field:
        full = client.get(f"/custom_fields/{field['gid']}",
                          {"opt_fields": "name,resource_subtype,enum_options.name,enum_options.gid"})
        log.info("custom field exists: %s", name)
        if subtype == "enum":
            have = {o["name"] for o in full.get("enum_options", [])}
            for opt in options:
                if opt not in have:
                    log.info("  + enum option %r on %s", opt, name)
                    client.post(f"/custom_fields/{full['gid']}/enum_options",
                                {"name": opt, "color": ENUM_COLORS[len(have) % len(ENUM_COLORS)]})
            full = client.get(f"/custom_fields/{full['gid']}",
                              {"opt_fields": "name,resource_subtype,enum_options.name,enum_options.gid"})
        return full

    payload: dict = {"workspace": workspace, "name": name,
                     "resource_subtype": subtype, "type": subtype}
    if subtype == "enum":
        payload["enum_options"] = [
            {"name": o, "color": ENUM_COLORS[i % len(ENUM_COLORS)]}
            for i, o in enumerate(options)
        ]
    if subtype == "number":
        payload["precision"] = 0
    log.info("creating custom field: %s (%s)", name, subtype)
    created = client.post("/custom_fields", payload)
    return client.get(f"/custom_fields/{created['gid']}",
                      {"opt_fields": "name,resource_subtype,enum_options.name,enum_options.gid"}) \
        if not client.dry_run else {"gid": created["gid"], "name": name, "enum_options": []}


def ensure_project(client: AsanaClient, workspace: str, team: str | None,
                   name: str, field_gids: list[str]) -> str:
    matches = list(client.paginate("/projects",
                                   {"workspace": workspace, "archived": "false",
                                    "opt_fields": "name"}))
    proj = find_by_name(matches, name)
    if proj:
        log.info("project exists: %s", name)
        gid = proj["gid"]
    else:
        payload = {"workspace": workspace, "name": name,
                   "default_view": "list", "privacy_setting": "public_to_workspace"}
        if team:
            payload["team"] = team
        log.info("creating project: %s", name)
        gid = client.post("/projects", payload)["gid"]

    attached = {s["custom_field"]["gid"]
                for s in client.get(f"/projects/{gid}/custom_field_settings",
                                    {"opt_fields": "custom_field.gid"}) or []} \
        if not client.dry_run else set()
    for fg in field_gids:
        if fg not in attached:
            client.post(f"/projects/{gid}/addCustomFieldSetting",
                        {"custom_field": fg, "is_important": True,
                         "insert_after": None})
    return gid


def ensure_portfolio(client: AsanaClient, workspace: str, name: str,
                     owner_gid: str, project_gids: list[str]) -> str:
    matches = list(client.paginate("/portfolios",
                                   {"workspace": workspace, "owner": owner_gid,
                                    "opt_fields": "name"}))
    pf = find_by_name(matches, name)
    if pf:
        log.info("portfolio exists: %s", name)
        gid = pf["gid"]
    else:
        log.info("creating portfolio: %s", name)
        gid = client.post("/portfolios", {"workspace": workspace, "name": name,
                                          "owner": owner_gid})["gid"]

    current = {i["gid"] for i in
               (client.get(f"/portfolios/{gid}/items", {"opt_fields": "gid"}) or [])} \
        if not client.dry_run else set()
    for pg in project_gids:
        if pg not in current:
            client.post(f"/portfolios/{gid}/addItem", {"item": pg})
    return gid


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", type=Path, default=Path("config.yaml"))
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)-7s %(message)s")

    load_dotenv()
    cfg = yaml.safe_load(args.config.read_text(encoding="utf-8"))
    token = os.getenv("ASANA_TOKEN", "")
    workspace = os.getenv("ASANA_WORKSPACE_GID") or cfg.get("workspace_gid")
    team = os.getenv("ASANA_TEAM_GID") or cfg.get("team_gid")
    if not workspace:
        raise SystemExit("Set ASANA_WORKSPACE_GID in .env (see README step 2).")

    client = AsanaClient(token, dry_run=args.dry_run)
    me = client.me()
    log.info("authenticated as %s", me.get("name"))

    # Discover the real function and priority values from the workbooks so the
    # enum options match the data instead of a hand-maintained list.
    records, _ = parse_folder(Path(cfg["input_folder"]))
    functions = sorted({r.function for r in records})
    priorities = sorted({r.priority_no for r in records})
    CUSTOM_FIELDS["Function"] = ("enum", functions)
    CUSTOM_FIELDS["Priority (RISE)"] = ("enum", [f"P{p}" for p in priorities])
    log.info("functions: %s", ", ".join(functions))

    existing = list(client.paginate(f"/workspaces/{workspace}/custom_fields",
                                     {"opt_fields": "name"}))
    fields: dict[str, dict] = {}
    for name, (subtype, options) in CUSTOM_FIELDS.items():
        fields[name] = ensure_custom_field(client, workspace, name, subtype,
                                           options, existing)

    field_gids = [f["gid"] for f in fields.values()]
    prefix = cfg.get("project_prefix", "RISE Q3 2026 — ")
    projects = {fn: ensure_project(client, workspace, team, f"{prefix}{fn}", field_gids)
                for fn in functions}

    portfolio = ensure_portfolio(client, workspace, cfg.get("portfolio_name", "RISE Q3 2026"),
                                 me["gid"], list(projects.values()))

    state = {
        "workspace_gid": workspace,
        "portfolio_gid": portfolio,
        "projects": projects,
        "custom_fields": {
            name: {
                "gid": f["gid"],
                "enum_options": {o["name"]: o["gid"] for o in f.get("enum_options", [])},
            } for name, f in fields.items()
        },
    }
    out = Path(cfg.get("state_folder", "state")) / "asana_ids.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    if args.dry_run:
        log.info("[dry-run] would write %s", out)
    else:
        out.write_text(json.dumps(state, indent=2), encoding="utf-8")
        log.info("wrote %s", out)
    log.info("done — %d reads, %d writes", client.reads, client.writes)


if __name__ == "__main__":
    try:
        main()
    except AsanaError as exc:
        raise SystemExit(f"Asana error: {exc}")
