# RISE → Asana sync

Drop the team Excel workbooks in `inbox/`, and every morning Asana reflects
what's in them. Runs entirely on your machine; nothing leaves it except the
Asana API calls.

```
inbox/*.xlsx  →  parse_rise.py  →  build/records.json  →  asana_sync.py  →  Asana
                (normalize +                            (idempotent
                 validate)                               upsert)
```

---

## What it produces in Asana

One **project per function** (six of them), grouped in a **portfolio** called
`RISE Q3 2026`. Inside each project, tasks are grouped into sections by
priority.

Two kinds of task, distinguished by the `Record Type` field:

| | Milestone | KPI |
|---|---|---|
| Comes from | text in a week column (W1–W13) | a row in the KPI grid |
| Task name | the milestone text | `KPI · <name>` |
| Due date | the Sunday ending that week | none |
| Start date | the Monday of that week (spans merged cells correctly) | none |
| Notes carry | goal, thresholds, full cell text, source reference | goal, thresholds, every weekly reading |

Eight custom fields are created: `RISE Source ID`, `Record Type`, `Function`,
`Priority`, `RAG Status`, `KPI Type`, `Sprint Week`, `Last Synced`.

`RISE Source ID` is the load-bearing one — a hash of file, sheet, KPI and cell
content. It's how a second run recognises a task it already made, instead of
creating a duplicate. Don't edit or clear it.

---

## Setup (once, ~15 minutes)

### 1. Install

```bash
cd rise-asana-sync
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Get your Asana credentials

Create a personal access token in Asana under your profile settings →
*Apps* → *Developer apps* → *Personal access tokens*. Copy it immediately;
Asana shows it once.

```bash
cp .env.example .env
```

Put the token in `.env` as `ASANA_TOKEN`, then let the script find your
workspace and team IDs:

```bash
python whoami.py
```

Copy the workspace gid into `ASANA_WORKSPACE_GID`. If your Asana requires
projects to belong to a team, copy a team gid into `ASANA_TEAM_GID` too.

`.env` is gitignored. Don't paste the token into `config.yaml`.

### 3. Create the Asana structure

```bash
python setup_asana.py --dry-run     # review what it will create
python setup_asana.py               # create it
```

This writes `state/asana_ids.json`. Safe to re-run — it reuses anything that
already exists by name rather than duplicating it. Re-run it whenever a team
adds a new function or a new priority number, so the enum options keep up.

### 4. First sync

```bash
cp /path/to/your/*.xlsx inbox/
python run_sync.py --parse-only     # check the parse, no network
python run_sync.py --dry-run        # check the plan, no writes
python run_sync.py                  # go
```

Start with one team to build confidence:

```bash
python run_sync.py --only-function "Careers"
```

### 5. Schedule it

- **Windows** — `scheduler/windows_task_scheduler.md`
- **macOS** — `scheduler/com.rise.asanasync.plist`
- **cron** — `scheduler/crontab.txt`

All three are set for 07:30 on weekdays. On Windows, tick *Run task as soon as
possible after a scheduled start is missed* so a closed laptop doesn't skip a
day.

---

## Daily operation

Teams keep editing their own workbooks. The only rule is that the current
version has to land in `inbox/` before the sync runs — a synced folder
(OneDrive/SharePoint/Dropbox) pointed at `input_folder` is the least
disruptive way to do that, since nobody has to change how they work.

After each run:

- `build/last_run.json` — counts, timings, and any errors
- `build/issues.json` — data quality findings per file
- `logs/sync.log` — rotating detail log

### Guarantees

- **Idempotent.** Two runs over unchanged files produce zero writes. Verified
  in `tests/test_sync_offline.py`.
- **Non-destructive.** Nothing is ever deleted. A row removed from Excel moves
  its task to a *No longer in source file* section (configurable via
  `stale_policy`).
- **Targeted updates.** Only fields that actually differ get written, so task
  history in Asana stays readable.

### Excel is the source of truth

Editing a synced task's name, dates or fields in Asana works until the next
run overwrites it. Comments, attachments, assignees and subtasks are never
touched, so that's where to put discussion. To change a milestone, change the
spreadsheet.

---

## Configuration

`config.yaml`:

| Key | Purpose |
|---|---|
| `input_folder` | where workbooks are read from |
| `stale_policy` | `flag` · `complete` · `ignore` for rows removed from Excel |
| `milestones_as_asana_milestones` | render milestones as diamonds on the timeline |
| `fail_on_parse_error` | abort before touching Asana if parsing found errors |
| `project_prefix` / `portfolio_name` | naming in Asana |

---

## Known data issues in the current workbooks

The parser handles all of these, but four are worth fixing at source:

| Severity | File | Issue |
|---|---|---|
| Error | `Careers` | `#REF!` in Summary cells D13:F13 — the rollup formula is broken |
| Error | `AADA_EAA` | Tab `Priority 4` is a copy of DTCM's Priority 2, banner and all. Its content syncs under AADA_EAA and will double-count. |
| Warning | `AADA_EAA`, `DTCM_GDI` | Function lead still reads `[Function lead]` |
| Warning | `DTCM_GDI` P2, `AADA_EAA` P4, `2_0_LA` P2 | 14–29 rows of meeting notes pasted below the grid; excluded from sync |

Structural variation the parser absorbs without complaint: an optional
`Owner(s)` column that shifts the week columns right, header rows at row 5 or
6, split priorities (`Priority 1a` / `1b`), milestones in cells merged across
2–4 weeks, and week cells containing readings (`4.3`, `NA`) or status words
(`On Track`, `Completed`) rather than tasks.

---

## Adapting it

- **Different quarter** — change `SPRINT_W1_MONDAY` in `parse_rise.py`. Week
  dates are derived, not hardcoded per sheet.
- **A team renames a column** — add the new label to `COLUMN_SYNONYMS`.
  Matching is case- and whitespace-insensitive, so cosmetic edits already work.
- **New status word appears in a week cell** — add it to
  `NON_MILESTONE_TOKENS` so it doesn't become a phantom task.
- **Assignees** — the sheets hold names like "Dian" and "Joyce / Nisha", not
  emails, so nothing is auto-assigned. Add a `people.yaml` mapping names to
  Asana user gids and set `assignee` in `AsanaSync._desired` if you want that.

## Testing

```bash
python tests/test_sync_offline.py
```

Runs the full parse against `inbox/`, then exercises create / no-op / update /
stale against an in-memory fake Asana. No network, no token, no side effects.
