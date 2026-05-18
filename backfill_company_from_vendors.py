"""Back-fill component_price_master.company from the vendor tables.

Only fills rows where company IS NULL or blank — never overwrites an
existing value. Reports per-source counts so you can see exactly which
vendor pricelist contributed which matches.

Rules (highest signal first):
  1. Exact part-code match against solenoid / dembla / cair / L&T / gas-train
  2. "Control Valve NN NB"        -> DEMBLA   (sole control-valve vendor)
  3. "Shut Off Ball Valve NN"     -> CAIR
  4. "L1RF*", "L2FF*", "L3FBT*",   -> L&T     (cat-no prefixes from lt_ball_valve_master)
     "L3RBT*", "L6FBT*", etc.
  5. Any "Butterfly Valve" + NB   -> L&T
  6. "Flexible Hose"              -> BENGAL INDUSTRIES
  7. "Gas Train"                  -> IAPL
  8. "Gas Regulator" / "Slam Shut"-> MADAS
  9. Solenoid part-code prefixes  -> MADAS
"""
from __future__ import annotations

import re
import sqlite3

DB = "vlph.db"


def _norm(s: str) -> str:
    """Compact key: uppercase, drop non-alphanumerics. So 'CONTROL VALVE 50 NB'
    matches 'Control Valve 050NB' matches 'CONTROL VALVE 050 NB'."""
    return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())


def main() -> None:
    conn = sqlite3.connect(DB)

    # ── 1. Build a normalized lookup from every vendor table → vendor name.
    lookup: dict[str, str] = {}

    def _add(key: str, vendor: str):
        k = _norm(key)
        if not k:
            return
        # Don't overwrite a more specific earlier mapping
        lookup.setdefault(k, vendor)

    # Dembla: valve_type + " " + nb (e.g. "Control Valve 50")
    for vt, nb in conn.execute("SELECT valve_type, nb FROM dembla_valve_master"):
        _add(f"{vt} {nb} NB", "DEMBLA")
        _add(f"{vt} {int(nb):03d}NB", "DEMBLA")

    # Cair: valve_type + " " + nb
    for vt, nb in conn.execute("SELECT valve_type, nb FROM cair_motorized_valve_master"):
        _add(f"{vt} {nb} NB", "CAIR")
        _add(f"{vt} {int(nb):03d}NB", "CAIR")

    # L&T ball valves: cat_no is the literal item key in component_price_master
    for (cat,) in conn.execute("SELECT cat_no FROM lt_ball_valve_master"):
        _add(cat, "L&T")

    # L&T butterfly: there's no make column, but they're all L&T.
    for model, nb in conn.execute("SELECT model, nb FROM lt_butterfly_valve_master"):
        _add(model, "L&T")
        _add(f"Butterfly Valve {nb} NB", "L&T")

    # Flexible hose: make column present
    for size_nb, length, make in conn.execute("SELECT size_nb, length_mm, make FROM flexible_hose_master"):
        _add(f"FLEXIBLE HOSE-{size_nb}*{length}MM", make or "BENGAL INDUSTRIES")
        _add(f"FLEXIBLE HOSE-{size_nb}*{length}MM (OIL)",   make or "BENGAL INDUSTRIES")
        _add(f"FLEXIBLE HOSE-{size_nb}*{length}MM (AIR)",   make or "BENGAL INDUSTRIES")

    # Gas regulator: all entries are MADAS
    for (pc,) in conn.execute("SELECT part_code FROM gas_regulator_master"):
        _add(pc, "MADAS")

    # Solenoid valves: part_code (e.g. EVAP04V-008) is the strongest signal
    for (pc,) in conn.execute("SELECT part_code FROM solenoidvalve_component_master"):
        _add(pc, "MADAS")

    # Gas train: type column has 'IAPL Make Gas Train' style
    for typ, part_code in conn.execute("SELECT type, part_code FROM gas_train_master"):
        make = "IAPL" if "iapl" in str(typ).lower() else "IAPL"
        _add(part_code, make)

    print(f"Built {len(lookup)} vendor identifiers")

    # ── 2. Walk component_price_master where company is blank,
    #       apply exact-match first, then pattern-based fallbacks.
    PATTERN_RULES = [
        # The order matters — first match wins.
        (re.compile(r"^control\s*valve\b",                re.I), "DEMBLA"),
        (re.compile(r"^shut\s*off\s*ball\s*valve\b",      re.I), "CAIR"),
        (re.compile(r"^motorized\s*valve\b.*\bnb\b",      re.I), "CAIR"),
        (re.compile(r"^butterfly\s*valve\b.*(\d+(\.\d+)?\s*(\"|nb)|nb\s*\d+)", re.I), "L&T"),
        (re.compile(r"^l[123][rf][fbsrw]",                re.I), "L&T"),
        (re.compile(r"^flexible\s*hose\b",                re.I), "BENGAL INDUSTRIES"),
        (re.compile(r"^gas\s*train\b",                    re.I), "IAPL"),
        (re.compile(r"^gas\s*regulator\b",                re.I), "MADAS"),
        (re.compile(r"^gas\s*filter\b",                   re.I), "MADAS"),
        (re.compile(r"^slam\s*shut\b",                    re.I), "MADAS"),
        (re.compile(r"^pneumatic\s*valve\b",              re.I), "MADAS"),
        (re.compile(r"^safety\s*relief\b",                re.I), "MADAS"),
        (re.compile(r"^gas\s*pressure\s*switch\b",        re.I), "MADAS"),
        (re.compile(r"^solenoid\s*valve\b",               re.I), "MADAS"),
        (re.compile(r"^(evap|evp|evo|evpc|evpcf|cm\d|cx\d|cn-\d|ec\d|ecs\d)", re.I), "MADAS"),
        (re.compile(r"^thermocouple\b",                   re.I), "TEMPSENS"),
        (re.compile(r"^rotary\s*joint\b",                 re.I), "ROTOFLOW"),
        (re.compile(r"^orifice\s*plate\b",                re.I), "ENCON"),
        (re.compile(r"^compensator\b",                    re.I), "ENCON"),
    ]

    # ── 2a. De-dup before backfill — UNIQUE(item, company) treats NULL=NULL as
    # distinct in SQLite, so two NULL-company rows for the same item can coexist
    # today. Updating them both to a vendor would violate the constraint, so
    # collapse exact duplicates (same item, same price) down to one row.
    dups = conn.execute("""
        SELECT item, price, COUNT(*) AS n FROM component_price_master
         WHERE company IS NULL OR TRIM(company)=''
      GROUP BY item, price HAVING n > 1
    """).fetchall()
    dup_removed = 0
    for item, price, _n in dups:
        rowids = [r[0] for r in conn.execute(
            "SELECT rowid FROM component_price_master "
            "WHERE item=? AND (company IS NULL OR TRIM(company)='') "
            "  AND (price=? OR (price IS NULL AND ? IS NULL)) ORDER BY rowid",
            (item, price, price)
        ).fetchall()]
        for rid in rowids[1:]:
            conn.execute("DELETE FROM component_price_master WHERE rowid=?", (rid,))
            dup_removed += 1
    conn.commit()
    print(f"De-duped {dup_removed} exact-duplicate NULL-company rows")

    candidates = conn.execute(
        "SELECT rowid, item FROM component_price_master "
        "WHERE company IS NULL OR TRIM(company)=''"
    ).fetchall()

    updates: list[tuple[int, str, str, str]] = []  # rowid, item, vendor, source
    for rowid, item in candidates:
        # Exact normalized match wins
        v = lookup.get(_norm(item))
        src = "exact"
        if not v:
            for pat, vendor in PATTERN_RULES:
                if pat.search(item):
                    v, src = vendor, "pattern"
                    break
        if v:
            updates.append((rowid, item, v, src))

    print(f"Candidates without company: {len(candidates)}")
    print(f"Resolvable now:             {len(updates)}")

    # ── 3. Apply updates. If a row with the same (item, target_company)
    # already exists, prefer to delete THIS blank-company row (it's a stale
    # duplicate) instead of updating, since the UNIQUE constraint would refuse.
    applied = 0
    skipped_dup = 0
    for rowid, item, vendor, src in updates:
        clash = conn.execute(
            "SELECT 1 FROM component_price_master WHERE item=? AND company=? AND rowid<>?",
            (item, vendor, rowid),
        ).fetchone()
        if clash:
            conn.execute("DELETE FROM component_price_master WHERE rowid=?", (rowid,))
            skipped_dup += 1
            continue
        conn.execute(
            "UPDATE component_price_master SET company=? WHERE rowid=? "
            "AND (company IS NULL OR TRIM(company)='')",
            (vendor, rowid),
        )
        applied += 1
    conn.commit()
    print(f"Updated:    {applied}")
    print(f"Removed-as-dup (vendor row already existed): {skipped_dup}")

    # ── 4. Breakdown per vendor
    print()
    print("=== Assigned ===")
    by_vendor: dict[str, int] = {}
    by_src:    dict[str, int] = {}
    for _, _, v, s in updates:
        by_vendor[v] = by_vendor.get(v, 0) + 1
        by_src[s]    = by_src.get(s, 0) + 1
    for v, n in sorted(by_vendor.items(), key=lambda x: -x[1]):
        print(f"  {v:24}  {n}")
    print(f"  (exact-match: {by_src.get('exact', 0)}, pattern: {by_src.get('pattern', 0)})")

    # ── 5. Final coverage snapshot
    total = conn.execute("SELECT COUNT(*) FROM component_price_master").fetchone()[0]
    with_co = conn.execute(
        "SELECT COUNT(*) FROM component_price_master "
        "WHERE company IS NOT NULL AND TRIM(company) <> ''"
    ).fetchone()[0]
    print()
    print(f"Coverage now: {with_co} / {total}  ({with_co/total*100:.1f}%)")
    conn.close()


if __name__ == "__main__":
    main()
