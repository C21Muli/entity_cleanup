# Entity cleanup — duplicates and Innovex-email entities

A two-step alternative to the interactive resolver in the duplicates_check repo, for when decisions are made on a copy (e.g. a dump) and the changes are applied by someone else as SQL. **Entities are never deleted** — they are deactivated (`status = 'INACTIVE'`), after their data has been moved to the entity that replaces them.

```
1. Plan      →   plan_cleanup.py          (read-only: writes an editable cleanup_plan.xlsx)
2. Generate  →   generate_cleanup_sql.py  (read-only: turns the edited plan into cleanup_auth.sql + cleanup_main.sql)
```

---

## Setup

Requires Python 3.12+ and [uv](https://github.com/astral-sh/uv).

```bash
uv venv
uv pip install --python .venv/bin/python -r requirements.txt
cp .env.example .env    # then fill in DSN (masterdata) and AUTH_DSN (auth)
```

Both scripts only **read** from the databases (to plan, validate and read the schema). All changes are made by running the generated SQL.

| File | Purpose |
|---|---|
| `plan_cleanup.py` | Step 1 — writes the editable plan |
| `generate_cleanup_sql.py` | Step 2 — validates the edited plan, writes the SQL |
| `entity_checks.py` | Shared read-only lookups (FKs to entities, reference counts, underwriters, users) |
| `matching.py` | Duplicate name matching, copied unchanged from `resolve_duplicates.py` in duplicates_check |

Logs (`plan_cleanup.log`, `generate_cleanup_sql.log`), the plan and the generated SQL are written next to the scripts and are gitignored, since they contain production entity data.

---

## Step 1 — plan

```bash
.venv/bin/python plan_cleanup.py                    # → cleanup_plan.xlsx
.venv/bin/python plan_cleanup.py --threshold 92     # stricter duplicate matching
```

The **Decisions** sheet has one row per entity in:
- every duplicate group across all entities (same matching as `resolve_duplicates.py`, but `@innovexsolutions.co.ke` emails never count as a shared email), and
- every Innovex-email entity that has no duplicate.

Each row shows underwriters, auth users, appointments, branches and every other reference, plus a suggestion pre-filled into the two yellow columns you edit:

| Column | Values |
|---|---|
| `action` | `KEEP`, `DEACTIVATE` or `SKIP` (= no decision, treated as KEEP). `REMOVE` from older plans is read as `DEACTIVATE` |
| `merge_into` | client key (or entity id) that receives a `DEACTIVATE`d entity's appointments, users, mappings, branches and other references before it is deactivated. Blank = deactivate in place, nothing moves |

Suggestions:
- **Duplicate group with a real email** — the most-used real-email entity is kept; the rest are merged into it, then deactivated.
- **Duplicate group where every member has an Innovex email** — the data is consolidated onto the most-used member, which is then deactivated too.
- **Innovex email, no duplicate** — deactivated in place (its data stays on it).

So every Innovex-email entity is suggested for deactivation. The **Instructions** sheet explains how to express "not duplicates", "keep a different one" and partial merges. **Underwriters**, **Users** and **References** sheets give the detail behind the counts.

## Step 2 — generate SQL

```bash
.venv/bin/python generate_cleanup_sql.py                  # reads cleanup_plan.xlsx
.venv/bin/python generate_cleanup_sql.py my_plan.xlsx     # or any other plan file
```

The plan is validated first (unknown actions, unknown or self `merge_into`, loops, a target without a client key); on any error no SQL is written. `merge_into` chains are followed to the final entity; a target that is itself deactivated in place still receives the data. Entities already `DELETED` are left as they are. Use `--status` if the application expects a different value than `INACTIVE`.

Two scripts are written — run the auth one first:

```bash
psql "<AUTH_DSN>" -v ON_ERROR_STOP=1 -f cleanup_auth.sql   # moves entity_users to the kept entity
psql "<DSN>"      -v ON_ERROR_STOP=1 -f cleanup_main.sql   # moves every reference, logs, deactivates
```

For each merge, `cleanup_main.sql` re-points every foreign key to `entities` (by `id` or `client_key`). Rows the kept entity already has — the same insurer mapping, entity type, user, or a clash on a unique index — are dropped *after* their dependants (`mapping_branches`, `mapping_services`) are moved to the kept entity's row, and an ACTIVE mapping is never lost to an inactive one. Each merged entity is logged to `removed_entities` (`resolved_in_auth_db = true`, since the auth script handles users), then set to `INACTIVE`. Entities deactivated in place keep all their data and users.

Safety:
- Each script is one transaction — any error rolls back everything.
- After every merge, `pg_temp.assert_unreferenced()` reads the FKs of the database it is running on and aborts if anything still points at the merged entity — so a table the move missed (e.g. one added after generation) stops the run instead of leaving data on an inactive entity.
- The SQL does not depend on the data at generation time, so it can be generated from a dump and run on the live databases — but re-plan if the live data has moved on significantly.
