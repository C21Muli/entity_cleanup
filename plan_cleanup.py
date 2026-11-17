#!/usr/bin/env python3
"""
Step 1 of the entity cleanup: builds an editable decision plan (read-only).

Lists, one row per entity:
  - every duplicate group across all entities (same fuzzy matching as
    resolve_duplicates.py, see matching.py; Innovex placeholder emails are not used as evidence), and
  - every entity with an Innovex email (@innovexsolutions.co.ke) that has no duplicate.

Each row carries what depends on the entity (underwriters, auth users, appointments,
branches, every other FK) and a suggested decision, pre-filled into the two
editable columns:

    action      KEEP | DEACTIVATE | SKIP
    merge_into  client key (or entity id) that receives a DEACTIVATEd entity's
                appointments, users, mappings, branches, ... before it is set
                INACTIVE; blank = deactivate in place (its data stays on it)

Entities are never deleted. Suggestions: in a duplicate group the most-used
real-email entity is kept and the rest are merged into it, then deactivated.
Every entity with an Innovex email is deactivated: in a group where all members
use one, the data is consolidated onto the most-used member, which is then
deactivated too; one with no duplicate is deactivated in place.

Review / edit the Decisions sheet, then run:
    python generate_cleanup_sql.py

Usage:
    python plan_cleanup.py
    python plan_cleanup.py --threshold 92
    python plan_cleanup.py --output cleanup_plan.csv      # Decisions sheet only

Connections are read from .env:
    DSN       — masterdata (main) database
    AUTH_DSN  — auth database (users are not counted if not set)
Both can be overridden with --dsn / --auth-dsn.
"""

import argparse
import csv
import logging
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from tabulate import tabulate

from entity_checks import (
    INNOVEX_DOMAIN, NOT_USAGE, connect, count_references, is_innovex,
    load_entity_fks, load_underwriters, load_users,
)
from matching import find_duplicate_groups, load_entities, similarity

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
        logging.FileHandler(HERE / "plan_cleanup.log"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)

ACTIONS = ("KEEP", "DEACTIVATE", "SKIP")
EDITABLE = ("action", "merge_into")

DECISION_COLUMNS = [
    "group", "group_type", "action", "merge_into",
    "suggested_action", "suggested_merge_into", "reason",
    "name", "email", "innovex_email", "client_key", "entity_id", "status",
    "entity_type", "match_to_keep", "underwriters_active", "users",
    "appointments", "branches", "other_references", "underwriters",
]
UNDERWRITER_COLUMNS = [
    "group", "entity_id", "name", "client_key", "insurance_name",
    "insurance_client_key", "mapping_status", "can_book_valuation", "markup", "vatable",
]
USER_COLUMNS = [
    "group", "entity_id", "name", "client_key", "user_id", "username", "user_email",
    "first_name", "last_name", "user_status", "entity_user_status",
]
REFERENCE_COLUMNS = ["group", "entity_id", "name", "client_key", "reference", "count"]

INSTRUCTIONS = [
    "Entity cleanup plan — edit the Decisions sheet, then run generate_cleanup_sql.py on this file.",
    "",
    "Only the yellow columns are read back: action and merge_into (rows are matched by entity_id).",
    "",
    "Entities are never deleted — DEACTIVATE sets the entity's status to INACTIVE.",
    "",
    "action = KEEP        leave the entity untouched.",
    "action = DEACTIVATE  with merge_into: first move its appointments, underwriter mappings,",
    "                     branches, users and every other reference to the entity named in",
    "                     merge_into (client key or entity id), then deactivate it.",
    "                     without merge_into: deactivate it in place — nothing is moved.",
    "action = SKIP        no decision yet — treated like KEEP.",
    "",
    "merge_into may name any entity (it does not have to be in the same group). If that entity",
    "is itself DEACTIVATEd with a merge_into, the data follows the chain to the final entity;",
    "if it is deactivated in place, the data lands on it and it is still deactivated.",
    "",
    "Not duplicates?          set every row of the group to KEEP (or DEACTIVATE in place).",
    "Keep a different one?    set it to KEEP, the others to DEACTIVATE with merge_into = its client key.",
    "Only some are duplicates DEACTIVATE those with merge_into, KEEP the rest.",
    "",
    "suggested_action / suggested_merge_into / reason show what the planner proposed.",
    "Usage counts are as of when the plan was generated.",
]

_HEADER_FONT = Font(bold=True)
_EDIT_FILL   = PatternFill("solid", fgColor="FFF2CC")
_BAND_FILL   = PatternFill("solid", fgColor="F2F2F2")


# ---------------------------------------------------------------------------
# Planning
# ---------------------------------------------------------------------------

def usage_of(e: dict, refs: dict, users: dict) -> dict:
    er = refs.get(e["id"], {})
    return {
        "appointments": er.get("appointments.service_provider_id", 0)
                        + er.get("appointments.client_key", 0),
        "branches":     er.get("entity_branches.entity_id", 0),
        "users":        len(users.get(e["id"], [])),
        "references":   sum(n for ref, n in er.items() if ref not in NOT_USAGE),
    }


def rank_key(e: dict, usage: dict, uw_active: int):
    """Higher is a better entity to keep."""
    u = usage[e["id"]]
    return (e["status"] != "DELETED", u["appointments"], uw_active, u["branches"],
            u["users"], u["references"])


def plan_group(members: list[dict], usage: dict, uw_active: dict, domain: str) -> dict:
    """Return {"keep": entity, "all_innovex": bool} for one duplicate group."""
    real = [e for e in members if not is_innovex(e["email"], domain)]
    pool = real or members
    keep = max(pool, key=lambda e: rank_key(e, usage, uw_active.get(e["id"], 0)))
    return {"keep": keep, "all_innovex": not real}


def ref_of(e: dict) -> str:
    return e["client_key"] or e["id"]


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def write_xlsx(path: str, decisions: list[dict], sheets: list[tuple[str, list[str], list[dict]]]):
    wb = Workbook()
    ins = wb.active
    ins.title = "Instructions"
    for line in INSTRUCTIONS:
        ins.append([line])
    ins.column_dimensions["A"].width = 110

    ws = wb.create_sheet("Decisions")
    ws.append(DECISION_COLUMNS)
    for cell in ws[1]:
        cell.font = _HEADER_FONT
    edit_idx = [DECISION_COLUMNS.index(c) for c in EDITABLE]
    band, prev_group = False, None
    for r in decisions:
        if r["group"] != prev_group:
            band, prev_group = not band, r["group"]
        ws.append([r.get(c) for c in DECISION_COLUMNS])
        for i, cell in enumerate(ws[ws.max_row]):
            if i in edit_idx:
                cell.fill = _EDIT_FILL
            elif band:
                cell.fill = _BAND_FILL
    for i in edit_idx:
        ws.cell(row=1, column=i + 1).fill = _EDIT_FILL
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = ws.dimensions
    widths = {"group": 7, "group_type": 18, "action": 9, "merge_into": 16,
              "suggested_action": 10, "suggested_merge_into": 16, "reason": 60,
              "name": 40, "email": 36, "client_key": 16, "entity_id": 38, "underwriters": 60}
    for i, col in enumerate(DECISION_COLUMNS):
        ws.column_dimensions[ws.cell(row=1, column=i + 1).column_letter].width = widths.get(col, 12)

    col = ws.cell(row=1, column=DECISION_COLUMNS.index("action") + 1).column_letter
    dv = DataValidation(type="list", formula1=f'"{",".join(ACTIONS)}"', allow_blank=False)
    dv.add(f"{col}2:{col}{max(ws.max_row, 2)}")
    ws.add_data_validation(dv)

    for title, columns, rows in sheets:
        s = wb.create_sheet(title)
        s.append(columns)
        for cell in s[1]:
            cell.font = _HEADER_FONT
        for r in rows:
            s.append([r.get(c) for c in columns])
        s.freeze_panes = "A2"

    wb.active = wb.sheetnames.index("Decisions")
    wb.save(path)


def write_csv(path: str, decisions: list[dict]):
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=DECISION_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(decisions)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dsn", default=os.getenv("DSN"),
                        help="Masterdata PostgreSQL DSN (or set DSN in .env)")
    parser.add_argument("--auth-dsn", default=os.getenv("AUTH_DSN"),
                        help="Auth PostgreSQL DSN (or set AUTH_DSN in .env)")
    parser.add_argument("--domain", default=INNOVEX_DOMAIN,
                        help=f"Placeholder email domain (default: {INNOVEX_DOMAIN})")
    parser.add_argument("--threshold", type=int, default=85,
                        help="Fuzzy name similarity %% for duplicates (default: 85)")
    parser.add_argument("--output", default=str(HERE / "cleanup_plan.xlsx"), metavar="PATH",
                        help="Plan file, .xlsx or .csv (default: cleanup_plan.xlsx)")
    args = parser.parse_args()

    if not args.dsn:
        parser.error("No DSN provided. Set DSN in .env or pass --dsn.")

    conn = connect(args.dsn, "masterdata")
    cur  = conn.cursor()
    entities = load_entities(cur)
    by_id    = {e["id"]: e for e in entities}
    logger.info(f"Loaded {len(entities)} entities")

    # Placeholder emails must not count as a shared email when grouping.
    scan = [{**e, "email": None if is_innovex(e["email"], args.domain) else e["email"]}
            for e in entities]
    groups = [[by_id[m["id"]] for m in g] for g in find_duplicate_groups(scan, args.threshold)]
    grouped_ids = {e["id"] for g in groups for e in g}
    standalone  = [e for e in entities
                   if is_innovex(e["email"], args.domain) and e["id"] not in grouped_ids]
    logger.info(f"{len(groups)} duplicate group(s), "
                f"{len(standalone)} Innovex-email entity(ies) without a duplicate")

    involved = [e for g in groups for e in g] + standalone
    ids      = [e["id"] for e in involved]
    refs     = count_references(cur, load_entity_fks(cur), involved)
    uw       = load_underwriters(cur, ids)
    cur.close()
    conn.close()

    if args.auth_dsn:
        auth_conn = connect(args.auth_dsn, "auth")
        auth_cur  = auth_conn.cursor()
        users     = load_users(auth_cur, ids)
        auth_cur.close()
        auth_conn.close()
    else:
        logger.warning("AUTH_DSN not set — auth users are not counted")
        users = {}

    usage     = {e["id"]: usage_of(e, refs, users) for e in involved}
    uw_active = {eid: sum(m["mapping_status"] == "ACTIVE" for m in maps) for eid, maps in uw.items()}

    def row(e: dict, group: str, group_type: str, action: str, merge_into: str | None,
            reason: str, match: float | None) -> dict:
        u = usage[e["id"]]
        return {
            "group": group, "group_type": group_type,
            "action": action, "merge_into": merge_into,
            "suggested_action": action, "suggested_merge_into": merge_into,
            "reason": reason,
            "name": e["name"], "email": e["email"],
            "innovex_email": "YES" if is_innovex(e["email"], args.domain) else "",
            "client_key": e["client_key"], "entity_id": e["id"], "status": e["status"],
            "entity_type": e["entity_type"],
            "match_to_keep": None if match is None else round(match),
            "underwriters_active": uw_active.get(e["id"], 0),
            "users": u["users"], "appointments": u["appointments"], "branches": u["branches"],
            "other_references": u["references"] - u["appointments"] - u["branches"]
                                - len(uw.get(e["id"], [])),
            "underwriters": ", ".join(m["insurance_name"] for m in uw.get(e["id"], [])
                                      if m["mapping_status"] == "ACTIVE"),
        }

    # Groups containing an Innovex email first, then the rest, each by kept name.
    planned = [(g, plan_group(g, usage, uw_active, args.domain)) for g in groups]
    planned.sort(key=lambda gp: (not any(is_innovex(e["email"], args.domain) for e in gp[0]),
                                 (gp[1]["keep"]["name"] or "").lower()))

    decisions = []
    for n, (members, p) in enumerate(planned, 1):
        gid  = f"D{n:03d}"
        keep = p["keep"]
        gtype = ("DUPLICATE+INNOVEX" if any(is_innovex(e["email"], args.domain) for e in members)
                 else "DUPLICATE")
        if p["all_innovex"]:
            decisions.append(row(keep, gid, gtype, "DEACTIVATE", None,
                                 "Every member uses an Innovex email — data consolidated here "
                                 "(most used), then deactivated", None))
        else:
            decisions.append(row(keep, gid, gtype, "KEEP", None,
                                 "Suggested keep: real email, most used", None))
        for e in sorted(members, key=lambda e: -similarity(e["name"], keep["name"])):
            if e["id"] == keep["id"]:
                continue
            score = similarity(e["name"], keep["name"])
            decisions.append(row(e, gid, gtype, "DEACTIVATE", ref_of(keep),
                                 f"Duplicate of {keep['name']} ({round(score)}%) — "
                                 f"data moves to {ref_of(keep)}, then deactivated", score))

    for n, e in enumerate(sorted(standalone, key=lambda e: (e["name"] or "").lower()), 1):
        in_use = usage[e["id"]]["references"] or usage[e["id"]]["users"]
        reason = ("Innovex email, no duplicate — deactivated in place"
                  + (" (still in use: its data stays on it)" if in_use else " (nothing references it)"))
        decisions.append(row(e, f"I{n:03d}", "INNOVEX_ONLY", "DEACTIVATE", None, reason, None))

    gid_of = {r["entity_id"]: r["group"] for r in decisions}
    detail = lambda rows_by_id: [  # noqa: E731
        {"group": gid_of[eid], "entity_id": eid, "name": by_id[eid]["name"],
         "client_key": by_id[eid]["client_key"], **r}
        for eid in gid_of for r in rows_by_id(eid)
    ]

    if args.output.lower().endswith(".xlsx"):
        write_xlsx(args.output, decisions, [
            ("Underwriters", UNDERWRITER_COLUMNS, detail(lambda eid: uw.get(eid, []))),
            ("Users",        USER_COLUMNS,        detail(lambda eid: users.get(eid, []))),
            ("References",   REFERENCE_COLUMNS,   detail(lambda eid: [
                {"reference": ref, "count": c} for ref, c in sorted(refs.get(eid, {}).items())])),
        ])
    else:
        write_csv(args.output, decisions)
    logger.info(f"Wrote {len(decisions)} row(s) to {args.output}")

    counts = {(t, a): 0 for t in ("DUPLICATE+INNOVEX", "DUPLICATE", "INNOVEX_ONLY") for a in ACTIONS}
    for d in decisions:
        counts[(d["group_type"], d["action"])] += 1
    print()
    print(tabulate(
        [[t, len({d["group"] for d in decisions if d["group_type"] == t}),
          counts[(t, "KEEP")], counts[(t, "DEACTIVATE")]]
         for t in ("DUPLICATE+INNOVEX", "DUPLICATE", "INNOVEX_ONLY")],
        headers=["Group type", "Groups", "Suggested KEEP", "Suggested DEACTIVATE"],
        tablefmt="rounded_outline",
    ))
    print(f"\n  Deactivate after merge (merge_into set) : "
          f"{sum(1 for d in decisions if d['action'] == 'DEACTIVATE' and d['merge_into'])}")
    print(f"  Deactivate in place (no merge_into)     : "
          f"{sum(1 for d in decisions if d['action'] == 'DEACTIVATE' and not d['merge_into'])}")
    print(f"  Plan file                               : {args.output}")
    print(f"\n  Review the Decisions sheet, then: "
          f".venv/bin/python generate_cleanup_sql.py {args.output}")


if __name__ == "__main__":
    main()
