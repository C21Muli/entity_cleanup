"""
Fuzzy duplicate matching for entity names.

Copied unchanged from resolve_duplicates.py in the duplicates_check repo, so the
cleanup groups duplicates exactly the way that resolver does.

Names are compared with max(ratio, token_sort_ratio) on the normalised name
(& -> and, punctuation stripped, lowercase) and on the name with legal suffixes
(LTD, LIMITED, ...) removed. A shared real business email scores 100.
"""

import re
from itertools import combinations

from rapidfuzz import fuzz
from tqdm import tqdm

from entity_checks import fetch

_AMPERSAND_RE = re.compile(r"\s*&\s*")
_PUNCT_RE     = re.compile(r"[^\w\s]")
_SPACE_RE     = re.compile(r"\s+")
_SUFFIX_RE    = re.compile(
    r"\s*\b(ltd|limited|llp|llc|co|company|inc|plc|pvt|pty|group|associates?)\b\.?\s*$",
    re.IGNORECASE,
)

# Emails assigned to on-the-fly entity creation — not a real business email,
# so a shared address here does NOT indicate a duplicate.
_GARBAGE_EMAILS = frozenset({
    "development@innovexsolutions.co.ke",
})


def _normalize(s: str) -> str:
    s = _AMPERSAND_RE.sub(" and ", s)
    s = _PUNCT_RE.sub(" ", s)
    return _SPACE_RE.sub(" ", s).strip().lower()


def _strip_suffix(s: str) -> str:
    return _SUFFIX_RE.sub("", s).strip()


def similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    an, bn   = _normalize(a),     _normalize(b)
    as_, bs_ = _strip_suffix(an), _strip_suffix(bn)
    return max(
        fuzz.ratio(an, bn),
        fuzz.token_sort_ratio(an, bn),
        fuzz.ratio(as_, bs_),
        fuzz.token_sort_ratio(as_, bs_),
    )


def _real_email(e: dict) -> str:
    """Return lowercased email if it's a real business address, else empty string."""
    addr = (e.get("email") or "").strip().lower()
    return addr if addr and addr not in _GARBAGE_EMAILS else ""


def email_match(a: dict, b: dict) -> bool:
    ea, eb = _real_email(a), _real_email(b)
    return bool(ea and ea == eb)


def entity_similarity(a: dict, b: dict) -> float:
    """Name similarity, boosted to 100 when both entities share a real email."""
    if email_match(a, b):
        return 100.0
    return similarity(a["name"], b["name"])


def load_entities(cur) -> list[dict]:
    return fetch(cur, """
        SELECT e.id, e.name, e.email, e.phone_number, e.status,
               e.client_key, e.is_onboarded, e.v_one_id, e.physical_address,
               et.name AS entity_type
        FROM entities e
        LEFT JOIN entity_type et ON et.id = e.entity_type_id
        WHERE e.name IS NOT NULL AND e.name != ''
        ORDER BY e.name
    """)


def find_duplicate_groups(entities: list[dict], threshold: int) -> list[list[dict]]:
    """
    Complete-linkage clustering: an entity joins a group only if it meets the
    threshold with every existing member. This prevents chain-linking where
    A≈B and B≈C at 60% pulls unrelated A and C into the same group.
    Pairs are processed strongest-first so high-confidence pairs form group cores.
    """
    n = len(entities)
    if n < 2:
        return []

    # Collect all pairs above threshold with their scores.
    sim_pairs: dict[tuple[int, int], float] = {}
    total = n * (n - 1) // 2
    with tqdm(total=total, desc="Scanning for duplicates", unit="pair") as pbar:
        for i, j in combinations(range(n), 2):
            s = entity_similarity(entities[i], entities[j])
            if s >= threshold:
                sim_pairs[(i, j)] = s  # i < j always (combinations guarantees this)
            pbar.update(1)

    # Complete-linkage clustering — process strongest pairs first.
    group_id: dict[int, int] = {}
    members:  dict[int, set[int]] = {}
    next_gid  = 0

    def all_connected(idx: int, grp: set[int]) -> bool:
        return all((min(idx, m), max(idx, m)) in sim_pairs for m in grp)

    for (i, j) in sorted(sim_pairs, key=lambda p: -sim_pairs[p]):
        gi, gj = group_id.get(i), group_id.get(j)

        if gi is None and gj is None:
            members[next_gid] = {i, j}
            group_id[i] = group_id[j] = next_gid
            next_gid += 1

        elif gi is None:
            if all_connected(i, members[gj]):
                members[gj].add(i)
                group_id[i] = gj

        elif gj is None:
            if all_connected(j, members[gi]):
                members[gi].add(j)
                group_id[j] = gi

        elif gi != gj:
            # Merge only when every cross-group pair meets threshold.
            if all((min(a, b), max(a, b)) in sim_pairs
                   for a in members[gi] for b in members[gj]):
                keep, drop = (gi, gj) if len(members[gi]) >= len(members[gj]) else (gj, gi)
                for m in members[drop]:
                    group_id[m] = keep
                members[keep].update(members.pop(drop))

    return [
        [entities[i] for i in sorted(grp)]
        for grp in members.values()
        if len(grp) >= 2
    ]
