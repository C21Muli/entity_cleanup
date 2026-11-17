"""
Shared read-only lookups for plan_cleanup.py and generate_cleanup_sql.py.

Covers the Innovex placeholder-email rule, the foreign keys that point at
entities (discovered from the schema, so new tables are picked up), per-entity
reference counts, underwriter mappings and auth-DB users.
"""

import logging
import re
import sys
from collections import defaultdict

import psycopg2

logger = logging.getLogger(__name__)

INNOVEX_DOMAIN = "innovexsolutions.co.ke"

# FK references that do not mean the entity is in use (its secondary types): the
# planner leaves them out of usage counts and the post-merge check ignores them.
NOT_USAGE = frozenset({"entity_entity_types.entity_id"})


def is_innovex(email: str | None, domain: str = INNOVEX_DOMAIN) -> bool:
    return (email or "").strip().lower().endswith(f"@{domain.lower()}")


# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------

def _safe_dsn(dsn: str) -> str:
    return re.sub(r"(password\s*=\s*)\S+", r"\1***", dsn, flags=re.IGNORECASE)


def connect(dsn: str, label: str):
    """Read-only connection — neither script ever writes to a database."""
    try:
        conn = psycopg2.connect(dsn)
        conn.set_session(readonly=True, autocommit=True)
        logger.info(f"Connected to {label}: {_safe_dsn(dsn)}")
        return conn
    except psycopg2.OperationalError as e:
        logger.error(f"Cannot connect to {label}: {e}")
        sys.exit(1)


def fetch(cur, sql: str, params=None) -> list[dict]:
    cur.execute(sql, params)
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


# ---------------------------------------------------------------------------
# Lookups — main DB
# ---------------------------------------------------------------------------

def load_entity_fks(cur) -> list[dict]:
    """Every single-column FK pointing at entities, with the column it references."""
    return fetch(cur, """
        SELECT cl.relname AS table_name, a.attname AS column_name,
               ra.attname AS ref_column
        FROM pg_constraint con
        JOIN pg_class cl  ON cl.oid  = con.conrelid
        JOIN pg_class rcl ON rcl.oid = con.confrelid
        JOIN pg_namespace n ON n.oid = rcl.relnamespace
        CROSS JOIN LATERAL unnest(con.conkey, con.confkey) AS k(src, ref)
        JOIN pg_attribute a  ON a.attrelid  = con.conrelid  AND a.attnum  = k.src
        JOIN pg_attribute ra ON ra.attrelid = con.confrelid AND ra.attnum = k.ref
        WHERE con.contype = 'f' AND rcl.relname = 'entities'
          AND n.nspname = current_schema()
          AND array_length(con.conkey, 1) = 1
        ORDER BY 1, 2
    """)


def count_references(cur, fks: list[dict], entities: list[dict]) -> dict[str, dict[str, int]]:
    """{entity_id: {"table.column": count}} for every FK, keyed by id or client_key."""
    ids  = [e["id"] for e in entities]
    keys = [e["client_key"] for e in entities if e["client_key"]]
    key_to_id = {e["client_key"]: e["id"] for e in entities if e["client_key"]}

    refs: dict[str, dict[str, int]] = defaultdict(dict)
    for fk in fks:
        ref = f"{fk['table_name']}.{fk['column_name']}"
        by_key = fk["ref_column"] == "client_key"
        cur.execute(
            f'SELECT "{fk["column_name"]}"::text, COUNT(*) FROM "{fk["table_name"]}" '
            f'WHERE "{fk["column_name"]}"::text = ANY(%s) GROUP BY 1',
            (keys if by_key else ids,),
        )
        for value, n in cur.fetchall():
            eid = key_to_id.get(value) if by_key else value
            if eid:
                refs[eid][ref] = n
    return refs


def load_underwriters(cur, ids: list[str]) -> dict[str, list[dict]]:
    """Insurers each entity is mapped to as a service provider (all statuses)."""
    rows = fetch(cur, """
        SELECT ete.service_provider_id::text AS eid,
               ins.name AS insurance_name, ins.client_key AS insurance_client_key,
               ete.status AS mapping_status, ete.can_book_valuation,
               ete.markup, ete.vatable
        FROM entity_to_entity ete
        JOIN entities ins ON ins.id = ete.insurance_id
        WHERE ete.service_provider_id::text = ANY(%s)
        ORDER BY ete.status, ins.name
    """, (ids,))
    result = defaultdict(list)
    for r in rows:
        result[r["eid"]].append(r)
    return result


# ---------------------------------------------------------------------------
# Lookups — auth DB
# ---------------------------------------------------------------------------

def load_users(auth_cur, ids: list[str]) -> dict[str, list[dict]]:
    rows = fetch(auth_cur, """
        SELECT eu.entity_id::text AS eid, eu.user_id::text AS user_id,
               eu.status::text AS entity_user_status,
               u.username, u.email AS user_email, u.first_name, u.last_name,
               u.status AS user_status
        FROM entity_users eu
        LEFT JOIN users u ON u.id = eu.user_id
        WHERE eu.entity_id::text = ANY(%s)
        ORDER BY u.email
    """, (ids,))
    result = defaultdict(list)
    for r in rows:
        result[r["eid"]].append(r)
    return result
