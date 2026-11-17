#!/usr/bin/env python3
"""
Step 2 of the entity cleanup: turns an edited plan into SQL (read-only).

Reads the Decisions sheet written by plan_cleanup.py (only entity_id, action and
merge_into are used), validates it against the masterdata / auth databases, and
writes one SQL script per database:

    cleanup_auth.sql  — auth DB:  moves entity_users of each DEACTIVATEd entity
                        that has a merge_into to that target (drops rows the
                        target already has)
    cleanup_main.sql  — main DB:  for each DEACTIVATEd entity with a merge_into,
                        moves every reference (all FKs to entities, by id or
                        client_key) to the target and logs it to removed_entities;
                        then sets every DEACTIVATEd entity's status to INACTIVE
                        (--status). Entities are never deleted.

DEACTIVATE without merge_into deactivates the entity in place: nothing is moved,
its appointments, users etc. stay on it.

How rows are moved
------------------
  - Plain references (appointments, branches, offers, ...) are re-pointed.
  - Rows that would duplicate one the target already has — entity_to_entity
    (same insurer / provider), entity_entity_types (same type), entity_users
    (same user), and any table with a unique index on the FK column — are dropped,
    after their own dependants (e.g. mapping_branches, mapping_services) are
    re-pointed to the target's matching row. An ACTIVE entity_to_entity mapping
    is never lost to a non-ACTIVE one: the target's row is re-activated.
  - Child entities (parent_entity_id) are re-parented.

Safety
------
  - Each script runs as a single transaction: any error rolls back everything.
  - After each merge, a check discovers the FKs to entities in the database it
    runs on and aborts if anything still references the merged entity, so a
    table the move missed (e.g. one added after the SQL was generated) stops the
    run instead of leaving data on an inactive entity.
  - Only entities whose status is not DELETED are set to INACTIVE.
  - REMOVE in an older plan is accepted and treated as DEACTIVATE.

Run the auth script first, then the main one:
    psql "<AUTH_DSN>" -v ON_ERROR_STOP=1 -f cleanup_auth.sql
    psql "<DSN>"      -v ON_ERROR_STOP=1 -f cleanup_main.sql

Usage:
    python generate_cleanup_sql.py               # reads cleanup_plan.xlsx
    python generate_cleanup_sql.py my_plan.csv --main-out main.sql --auth-out auth.sql

Connections are read from .env (DSN, AUTH_DSN) or --dsn / --auth-dsn; they are
only read from, to validate the plan and to read the schema.
"""

import argparse
import csv
import logging
import os
import sys
from pathlib import Path
from datetime import datetime

from dotenv import load_dotenv
from openpyxl import load_workbook
from tabulate import tabulate

from entity_checks import (
    _safe_dsn, connect, count_references, fetch, load_entity_fks,
)

load_dotenv()

# Logs and default input / output files live next to this script.
HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(HERE / "generate_cleanup_sql.log"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)

ACTIONS = ("KEEP", "DEACTIVATE", "SKIP")
ALIASES = {"REMOVE": "DEACTIVATE"}  # plans made before the switch to deactivation

# Rows that mean the same thing when they share these columns, even without a
# unique index: a duplicate is dropped instead of moved.
LOGICAL_KEYS = {
    ("entity_to_entity", "service_provider_id"): ["insurance_id"],
    ("entity_to_entity", "insurance_id"):        ["service_provider_id"],
    ("entity_entity_types", "entity_id"):        ["entity_type_id"],
    ("entity_users", "entity_id"):               ["user_id"],
}
# Tables whose "status" must not go from ACTIVE to inactive when duplicates merge.
STATUS_TABLES = frozenset({"entity_to_entity"})

ASSERT_FUNCTION = """\
-- Aborts if anything still references the entity after its data was moved.
-- FKs are read from this database at run time, so tables that did not exist
-- when the SQL was generated are covered too.
CREATE FUNCTION pg_temp.assert_unreferenced(eid uuid) RETURNS void
LANGUAGE plpgsql AS $fn$
DECLARE
    ck  text;
    fk  record;
    hit boolean;
BEGIN
    SELECT client_key INTO ck FROM entities WHERE id = eid;
    FOR fk IN
        SELECT cn.nspname AS sch, cl.relname AS tbl, a.attname AS col, ra.attname AS refcol
        FROM pg_constraint con
        JOIN pg_class cl      ON cl.oid  = con.conrelid
        JOIN pg_namespace cn  ON cn.oid  = cl.relnamespace
        JOIN pg_class rcl     ON rcl.oid = con.confrelid
        JOIN pg_namespace rn  ON rn.oid  = rcl.relnamespace
        CROSS JOIN LATERAL unnest(con.conkey, con.confkey) AS k(src, ref)
        JOIN pg_attribute a  ON a.attrelid  = con.conrelid  AND a.attnum  = k.src
        JOIN pg_attribute ra ON ra.attrelid = con.confrelid AND ra.attnum = k.ref
        WHERE con.contype = 'f' AND rcl.relname = 'entities'
          AND rn.nspname = current_schema()
          AND array_length(con.conkey, 1) = 1
          AND NOT (cl.relname = 'entity_entity_types' AND a.attname = 'entity_id')
    LOOP
        IF fk.refcol = 'client_key' THEN
            CONTINUE WHEN ck IS NULL;
            EXECUTE format('SELECT EXISTS (SELECT 1 FROM %I.%I WHERE %I = $1)',
                           fk.sch, fk.tbl, fk.col) INTO hit USING ck;
        ELSE
            EXECUTE format('SELECT EXISTS (SELECT 1 FROM %I.%I WHERE %I = $1)',
                           fk.sch, fk.tbl, fk.col) INTO hit USING eid;
        END IF;
        IF hit THEN
            RAISE EXCEPTION 'Entity % is still referenced by %.% after the move — aborting, nothing applied',
                            eid, fk.tbl, fk.col;
        END IF;
    END LOOP;
END
$fn$;
"""

REMOVED_ENTITIES_DDL = """\
CREATE TABLE IF NOT EXISTS removed_entities (
    id                  uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    removed_entity_id   uuid NOT NULL,
    removed_client_key  varchar(50),
    entity_name         varchar(255),
    entity_type         varchar(255),
    kept_entity_id      uuid NOT NULL,
    kept_client_key     varchar(50),
    reason              text,
    removed_at          timestamptz NOT NULL DEFAULT now(),
    resolved_in_auth_db boolean NOT NULL DEFAULT false,
    notes               text
);
"""


# ---------------------------------------------------------------------------
# SQL text helpers
# ---------------------------------------------------------------------------

def q(ident: str) -> str:
    return '"' + ident.replace('"', '""') + '"'


def lit(value) -> str:
    return "NULL" if value is None else "'" + str(value).replace("'", "''") + "'"


def one_line(s) -> str:
    return " ".join(str(s or "").split())


# ---------------------------------------------------------------------------
# Schema lookups (whichever DB the statements are for)
# ---------------------------------------------------------------------------

class Schema:
    def __init__(self, cur):
        self.cur = cur
        self._unique: dict[str, list[list[str]]] = {}
        self._children: dict[str, list[dict]] = {}

    def unique_keys(self, table: str) -> list[list[str]]:
        """Column lists of unique, non-partial, plain-column indexes (not the PK)."""
        if table not in self._unique:
            rows = fetch(self.cur, """
                SELECT array_agg(a.attname ORDER BY k.ord) AS cols
                FROM pg_index ix
                JOIN pg_class t ON t.oid = ix.indrelid
                JOIN pg_namespace n ON n.oid = t.relnamespace
                CROSS JOIN LATERAL unnest(ix.indkey) WITH ORDINALITY AS k(attnum, ord)
                JOIN pg_attribute a ON a.attrelid = t.oid AND a.attnum = k.attnum
                WHERE t.relname = %s AND n.nspname = current_schema()
                  AND ix.indisunique AND NOT ix.indisprimary
                  AND ix.indpred IS NULL AND ix.indexprs IS NULL
                GROUP BY ix.indexrelid
            """, (table,))
            self._unique[table] = [r["cols"] for r in rows]
        return self._unique[table]

    def children(self, table: str) -> list[dict]:
        """Single-column FKs pointing at `table`: child table, child column, referenced column."""
        if table not in self._children:
            self._children[table] = fetch(self.cur, """
                SELECT cl.relname AS table_name, a.attname AS column_name,
                       ra.attname AS ref_column
                FROM pg_constraint con
                JOIN pg_class cl  ON cl.oid  = con.conrelid
                JOIN pg_class rcl ON rcl.oid = con.confrelid
                JOIN pg_namespace n ON n.oid = rcl.relnamespace
                CROSS JOIN LATERAL unnest(con.conkey, con.confkey) AS k(src, ref)
                JOIN pg_attribute a  ON a.attrelid  = con.conrelid  AND a.attnum  = k.src
                JOIN pg_attribute ra ON ra.attrelid = con.confrelid AND ra.attnum = k.ref
                WHERE con.contype = 'f' AND rcl.relname = %s
                  AND n.nspname = current_schema()
                  AND array_length(con.conkey, 1) = 1
                ORDER BY 1, 2
            """, (table,))
        return self._children[table]

    def dedupe_keys(self, table: str, column: str) -> list[list[str]]:
        """Other-column lists that identify a duplicate row once `column` is re-pointed."""
        keys = [[c for c in cols if c != column]
                for cols in self.unique_keys(table) if column in cols]
        logical = LOGICAL_KEYS.get((table, column))
        if logical and logical not in keys:
            keys.append(logical)
        return keys


def move_statements(schema: Schema, table: str, column: str, old, new) -> list[str]:
    """
    SQL that re-points table.column from `old` to `new`. Rows that would duplicate
    one `new` already has are dropped — after their dependants are moved onto
    that row.
    """
    t, c = q(table), q(column)
    out = []
    for key in schema.dedupe_keys(table, column):
        match = " AND ".join([f"k.{c} = {lit(new)}", f"r.{c} = {lit(old)}"]
                             + [f"k.{q(x)} = r.{q(x)}" for x in key])
        for ch in schema.children(table):
            ct, cc, ref = q(ch["table_name"]), q(ch["column_name"]), q(ch["ref_column"])
            for ukey in schema.dedupe_keys(ch["table_name"], ch["column_name"]):
                same = "".join(f" AND c2.{q(x)} = ch.{q(x)}" for x in ukey)
                out.append(
                    f"DELETE FROM {ct} ch USING {t} r, {t} k WHERE {match} "
                    f"AND ch.{cc} = r.{ref} AND EXISTS "
                    f"(SELECT 1 FROM {ct} c2 WHERE c2.{cc} = k.{ref}{same});")
            out.append(
                f"UPDATE {ct} ch SET {cc} = k.{ref} FROM {t} r, {t} k "
                f"WHERE {match} AND ch.{cc} = r.{ref};")
        if table in STATUS_TABLES:
            out.append(
                f"UPDATE {t} k SET status = 'ACTIVE' FROM {t} r WHERE {match} "
                f"AND r.status = 'ACTIVE' AND k.status IS DISTINCT FROM 'ACTIVE';")
        same = "".join(f" AND k.{q(x)} = r.{q(x)}" for x in key)
        out.append(
            f"DELETE FROM {t} r WHERE r.{c} = {lit(old)} AND EXISTS "
            f"(SELECT 1 FROM {t} k WHERE k.{c} = {lit(new)}{same});")
    out.append(f"UPDATE {t} SET {c} = {lit(new)} WHERE {c} = {lit(old)};")
    return out


# ---------------------------------------------------------------------------
# Plan loading and validation
# ---------------------------------------------------------------------------

def load_plan(path: str) -> list[dict]:
    if path.lower().endswith(".xlsx"):
        wb = load_workbook(path, read_only=True, data_only=True)
        if "Decisions" not in wb.sheetnames:
            sys.exit(f"{path}: no 'Decisions' sheet")
        it = wb["Decisions"].iter_rows(values_only=True)
        header = [str(h or "").strip() for h in next(it)]
        rows = [dict(zip(header, r)) for r in it]
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    for col in ("entity_id", "action", "merge_into"):
        if rows and col not in rows[0]:
            sys.exit(f"{path}: missing column '{col}'")
    plan = []
    for n, r in enumerate(rows, 2):
        eid = one_line(r.get("entity_id"))
        if not eid:
            continue
        plan.append({
            "line":       n,
            "group":      one_line(r.get("group")),
            "entity_id":  eid,
            "action":     ALIASES.get(a := one_line(r.get("action")).upper(), a),
            "merge_into": one_line(r.get("merge_into")),
        })
    logger.info(f"Loaded {len(plan)} decision row(s) from {path}")
    return plan


def resolve_plan(plan: list[dict], entities: dict[str, dict], errors: list[str],
                 warnings: list[str]) -> tuple[list[dict], list[dict]]:
    """
    Return (merges, in_place). merges are {row, removed, target} with merge_into
    chains followed; in_place are {row, removed} deactivated without moving data.
    """
    by_ck  = {e["client_key"]: e for e in entities.values() if e["client_key"]}
    by_ckl = {k.lower(): e for k, e in by_ck.items()}

    def find(ref: str):
        return entities.get(ref) or by_ck.get(ref) or by_ckl.get(ref.lower())

    seen, deactivates = set(), {}
    for r in plan:
        where = f"row {r['line']} ({r['group']})"
        if r["entity_id"] in seen:
            errors.append(f"{where}: entity {r['entity_id']} appears more than once")
            continue
        seen.add(r["entity_id"])
        if r["action"] == "":
            warnings.append(f"{where}: blank action — treated as SKIP")
            continue
        if r["action"] not in ACTIONS:
            errors.append(f"{where}: unknown action '{r['action']}' (use {', '.join(ACTIONS)})")
            continue
        if r["action"] != "DEACTIVATE":
            if r["merge_into"]:
                warnings.append(f"{where}: merge_into ignored — action is {r['action']}")
            continue
        if r["entity_id"] not in entities:
            errors.append(f"{where}: entity {r['entity_id']} no longer exists in the database")
            continue
        target = None
        if r["merge_into"]:
            target = find(r["merge_into"])
            if not target:
                errors.append(f"{where}: merge_into '{r['merge_into']}' matches no entity")
                continue
            if target["id"] == r["entity_id"]:
                errors.append(f"{where}: merge_into is the entity itself")
                continue
        deactivates[r["entity_id"]] = (r, target["id"] if target else None)

    merges, in_place = [], []
    for eid, (r, tid) in deactivates.items():
        where = f"row {r['line']} ({r['group']})"
        removed = entities[eid]
        if removed["status"] == "DELETED":
            warnings.append(f"{where}: {removed['name']} has status DELETED — left as DELETED")
        if tid is None:
            in_place.append({"row": r, "removed": removed})
            continue
        # A target that is itself merged elsewhere passes the data on; one that is
        # deactivated in place is the end of the chain (data lands on an inactive entity).
        chain = [eid]
        while tid in deactivates and deactivates[tid][1] is not None:
            if tid in chain:
                errors.append(f"{where}: merge_into chain loops back on itself")
                break
            chain.append(tid)
            tid = deactivates[tid][1]
        else:
            target = entities[tid]
            if len(chain) > 1:
                warnings.append(f"{where}: {removed['name']} merges via a chain into "
                                f"{target['name']} ({target['client_key'] or tid})")
            if target["status"] == "DELETED":
                warnings.append(f"{where}: merge target {target['name']} has status DELETED")
            if removed["client_key"] and not target["client_key"]:
                errors.append(f"{where}: merge target {target['name']} has no client_key, "
                              f"so client_key references cannot be moved")
                continue
            merges.append({"row": r, "removed": removed, "target": target,
                           "target_deactivated": tid in deactivates})
    return merges, in_place


# ---------------------------------------------------------------------------
# SQL generation
# ---------------------------------------------------------------------------

def header(title: str, plan_path: str, dsn: str, lines: list[str]) -> list[str]:
    return [
        f"-- Entity cleanup — {title}",
        f"-- Generated {datetime.now():%Y-%m-%d %H:%M} from {plan_path} by generate_cleanup_sql.py",
        f"-- Schema read from: {_safe_dsn(dsn)}",
        *[f"-- {line}" for line in lines],
        "-- Runs as one transaction: any error rolls everything back.",
        "",
        "BEGIN;",
        "",
    ]


def current_refs(refs: dict[str, int]) -> str:
    return ", ".join(f"{k}={v}" for k, v in sorted(refs.items())) or "nothing"


def deactivate_sql(entity: dict, status: str) -> str:
    return (f"UPDATE entities SET status = {lit(status)} WHERE id = {lit(entity['id'])} "
            f"AND status IS DISTINCT FROM 'DELETED';")


def build_main_sql(schema: Schema, fks: list[dict], merges: list[dict], in_place: list[dict],
                   refs: dict, plan_path: str, dsn: str, auth_file: str, status: str) -> str:
    out = header("MAIN (masterdata) database", plan_path, dsn, [
        f"{len(merges)} merge(s) then deactivation, {len(in_place)} deactivation(s) in place.",
        f"Deactivated entities get status = {status!r}; nothing is deleted.",
        f"Run AFTER {auth_file} (auth DB), e.g.:",
        f'  psql "<DSN>" -v ON_ERROR_STOP=1 -f <this file>',
    ])
    out += [REMOVED_ENTITIES_DDL, ASSERT_FUNCTION]

    total = len(merges) + len(in_place)
    for n, m in enumerate(merges, 1):
        r, k, row = m["removed"], m["target"], m["row"]
        out += [
            "",
            f"-- [{n}/{total}] {row['group']}  MERGE {one_line(r['name'])} ({r['client_key'] or '—'})"
            f"  →  {one_line(k['name'])} ({k['client_key'] or '—'}), then DEACTIVATE",
            f"--   deactivated id {r['id']}  |  kept id {k['id']}"
            + ("  (kept entity is deactivated too)" if m["target_deactivated"] else ""),
            f"--   currently referenced by: {current_refs(refs.get(r['id'], {}))}",
        ]
        for fk in fks:
            table, col = fk["table_name"], fk["column_name"]
            by_key = fk["ref_column"] == "client_key"
            old, new = (r["client_key"], k["client_key"]) if by_key else (r["id"], k["id"])
            if old is None:
                continue
            if table == "entities":  # self-reference: never make the kept entity its own parent
                out.append(f"UPDATE entities SET {q(col)} = NULL "
                           f"WHERE id = {lit(k['id'])} AND {q(col)} = {lit(old)};")
                out.append(f"UPDATE entities SET {q(col)} = {lit(new)} WHERE {q(col)} = {lit(old)};")
                continue
            out += move_statements(schema, table, col, old, new)
        reason = (f"Deactivated — merged into {k['name']} (client_key={k['client_key']}) "
                  f"— plan {row['group']}")
        notes  = f"entity_users moved by {auth_file}"
        out += [
            "INSERT INTO removed_entities (removed_entity_id, removed_client_key, entity_name, "
            "entity_type, kept_entity_id, kept_client_key, reason, resolved_in_auth_db, notes) VALUES ("
            f"{lit(r['id'])}, {lit(r['client_key'])}, {lit(r['name'])}, {lit(r['entity_type'])}, "
            f"{lit(k['id'])}, {lit(k['client_key'])}, {lit(reason)}, true, {lit(notes)});",
            f"SELECT pg_temp.assert_unreferenced({lit(r['id'])});",
            deactivate_sql(r, status),
        ]

    for n, d in enumerate(in_place, len(merges) + 1):
        r, row = d["removed"], d["row"]
        out += [
            "",
            f"-- [{n}/{total}] {row['group']}  DEACTIVATE in place {one_line(r['name'])} "
            f"({r['client_key'] or '—'}) — no merge target, its data stays on it",
            f"--   id {r['id']}  |  referenced by: {current_refs(refs.get(r['id'], {}))}",
            deactivate_sql(r, status),
        ]

    out += ["", "COMMIT;", ""]
    return "\n".join(out)


def build_auth_sql(schema: Schema | None, merges: list[dict], plan_path: str, dsn: str,
                   main_file: str) -> str:
    out = header("AUTH database", plan_path, dsn or "(not connected — default schema assumed)", [
        f"{len(merges)} merge(s). Entities deactivated in place keep their users.",
        f"Run BEFORE {main_file} (main DB), e.g.:",
        f'  psql "<AUTH_DSN>" -v ON_ERROR_STOP=1 -f <this file>',
    ])
    for n, m in enumerate(merges, 1):
        r, k, row = m["removed"], m["target"], m["row"]
        out += [
            "",
            f"-- [{n}/{len(merges)}] {row['group']}  {one_line(r['name'])} ({r['client_key'] or '—'})"
            f"  →  {one_line(k['name'])} ({k['client_key'] or '—'})",
        ]
        if schema:
            out += move_statements(schema, "entity_users", "entity_id", r["id"], k["id"])
        else:
            out += [
                f"DELETE FROM entity_users r WHERE r.entity_id = {lit(r['id'])} AND EXISTS "
                f"(SELECT 1 FROM entity_users k WHERE k.entity_id = {lit(k['id'])} "
                f"AND k.user_id = r.user_id);",
                f"UPDATE entity_users SET entity_id = {lit(k['id'])} WHERE entity_id = {lit(r['id'])};",
            ]
    out += ["", "COMMIT;", ""]
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("plan", nargs="?", default=str(HERE / "cleanup_plan.xlsx"),
                        help="Edited plan from plan_cleanup.py, .xlsx or .csv "
                             "(default: cleanup_plan.xlsx)")
    parser.add_argument("--dsn", default=os.getenv("DSN"),
                        help="Masterdata PostgreSQL DSN (or set DSN in .env)")
    parser.add_argument("--auth-dsn", default=os.getenv("AUTH_DSN"),
                        help="Auth PostgreSQL DSN (or set AUTH_DSN in .env)")
    parser.add_argument("--status", default="INACTIVE",
                        help="Status given to deactivated entities (default: INACTIVE)")
    parser.add_argument("--main-out", default=str(HERE / "cleanup_main.sql"), metavar="PATH",
                        help="SQL for the main DB (default: cleanup_main.sql)")
    parser.add_argument("--auth-out", default=str(HERE / "cleanup_auth.sql"), metavar="PATH",
                        help="SQL for the auth DB (default: cleanup_auth.sql)")
    args = parser.parse_args()

    if not args.dsn:
        parser.error("No DSN provided. Set DSN in .env or pass --dsn.")

    plan = load_plan(args.plan)

    conn = connect(args.dsn, "masterdata")
    cur  = conn.cursor()
    entities = {e["id"]: e for e in fetch(cur, """
        SELECT e.id::text AS id, e.name, e.client_key, e.status, et.name AS entity_type
        FROM entities e
        LEFT JOIN entity_type et ON et.id = e.entity_type_id
    """)}

    errors, warnings = [], []
    merges, in_place = resolve_plan(plan, entities, errors, warnings)

    fks  = load_entity_fks(cur)
    refs = count_references(cur, fks, [m["removed"] for m in merges]
                            + [d["removed"] for d in in_place])

    auth_schema, auth_conn = None, None
    if args.auth_dsn:
        auth_conn   = connect(args.auth_dsn, "auth")
        auth_schema = Schema(auth_conn.cursor())
    else:
        warnings.append("AUTH_DSN not set — auth SQL assumes the default entity_users schema")

    for w in warnings:
        logger.warning(w)
    if errors:
        for e in errors:
            logger.error(e)
        print(f"\n  {len(errors)} error(s) in {args.plan} — no SQL written. Fix the plan and re-run.")
        sys.exit(1)

    main_sql = build_main_sql(Schema(cur), fks, merges, in_place, refs,
                              os.path.basename(args.plan), args.dsn,
                              os.path.basename(args.auth_out), args.status)
    auth_sql = build_auth_sql(auth_schema, merges, os.path.basename(args.plan), args.auth_dsn,
                              os.path.basename(args.main_out))
    cur.close()
    conn.close()
    if auth_conn:
        auth_conn.close()

    with open(args.main_out, "w") as f:
        f.write(main_sql)
    with open(args.auth_out, "w") as f:
        f.write(auth_sql)

    actions = {a: sum(1 for r in plan if r["action"] == a) for a in ACTIONS}
    print()
    print(tabulate([
        ["KEEP", actions["KEEP"]],
        ["SKIP / blank", len(plan) - actions["KEEP"] - actions["DEACTIVATE"]],
        ["DEACTIVATE after merge", len(merges)],
        ["  of which into an entity also deactivated",
         sum(1 for m in merges if m["target_deactivated"])],
        ["DEACTIVATE in place", len(in_place)],
    ], headers=["Decision", "Entities"], tablefmt="rounded_outline"))
    print(f"\n  Warnings : {len(warnings)}  (see generate_cleanup_sql.log)")
    print(f"  Auth SQL : {args.auth_out}   ← run first")
    print(f"  Main SQL : {args.main_out}")


if __name__ == "__main__":
    main()
