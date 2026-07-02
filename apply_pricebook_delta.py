"""One-off: apply the (NEW Excel − OLD Excel) delta to live vlph.db.

Strategy:
  1. Parse the OLD pricebook into a temp DB.
  2. Parse the NEW pricebook into a temp DB.
  3. For each master table, compute the diff: rows that changed or
     are new in the NEW DB compared to the OLD.
  4. Apply ONLY those changed rows to the live vlph.db.
     Manual rate edits in live that touch items NOT in the delta are
     preserved.

Run from project root after dropping the new file in as
'Pricelist_WorkBook_updated.xlsx'. Prints a per-table summary.
"""
from __future__ import annotations

import os
import sqlite3
import tempfile

from bom.pricelist_parser import parse_all

OLD = "uploads/Pricelist WorkBook 28-08-2025.xlsx"
NEW = "uploads/Pricelist WorkBook 18-05-2026.xlsx"
LIVE_DB = "vlph.db"


def _init_db_schema(path: str) -> None:
    """The parsers expect the master tables to exist already.
    Copy the live DB schema (and data) so parse_all has a target."""
    import shutil
    shutil.copy(LIVE_DB, path)


def _table_rows(conn: sqlite3.Connection, table: str) -> list[tuple]:
    cur = conn.execute(f"SELECT * FROM {table}")
    cols = [d[0] for d in cur.description]
    return cols, cur.fetchall()


def _diff_table(old_conn, new_conn, table: str, key_cols: list[str]):
    """Return (changed_rows, added_rows) using key_cols as identity.
    Each row is a dict[col -> value]."""
    cols_o, rows_o = _table_rows(old_conn, table)
    cols_n, rows_n = _table_rows(new_conn, table)
    if cols_o != cols_n:
        # Schema drift — bail out, full reset is the only safe option.
        return None, None, f"schema differs ({cols_o} vs {cols_n})"
    key_idx = [cols_o.index(k) for k in key_cols]

    def keyof(r): return tuple(r[i] for i in key_idx)

    map_o = {keyof(r): r for r in rows_o}
    map_n = {keyof(r): r for r in rows_n}
    changed = []
    added = []
    for k, r in map_n.items():
        if k not in map_o:
            added.append(dict(zip(cols_n, r)))
        elif map_o[k] != r:
            changed.append((dict(zip(cols_o, map_o[k])), dict(zip(cols_n, r))))
    return changed, added, None


def main() -> None:
    if not os.path.exists(NEW):
        raise SystemExit(f"missing: {NEW}")
    if not os.path.exists(OLD):
        raise SystemExit(f"missing: {OLD}")

    # Build temp DBs by re-parsing each Excel. Use NamedTemporaryFile only
    # to claim a unique path, then close it immediately so Windows lets us
    # reopen/delete the file later.
    f_old = tempfile.NamedTemporaryFile(suffix="_old.db", delete=False)
    f_new = tempfile.NamedTemporaryFile(suffix="_new.db", delete=False)
    tmp_old = f_old.name; f_old.close()
    tmp_new = f_new.name; f_new.close()
    _init_db_schema(tmp_old)
    _init_db_schema(tmp_new)

    print(f"Parsing OLD ({OLD}) -> {tmp_old}")
    with sqlite3.connect(tmp_old) as c:
        parse_all(OLD, c)
    print(f"Parsing NEW ({NEW}) -> {tmp_new}")
    with sqlite3.connect(tmp_new) as c:
        parse_all(NEW, c)

    # Tables to diff and their natural keys (taken from PRAGMA table_info).
    table_keys = {
        "component_price_master":     ["item"],
        "burner_pricelist_master":    ["section", "burner_size", "component"],
        "blower_pricelist_master":    ["section", "model", "hp"],
        "hpu_master":                 ["unit_kw", "variant", "item"],
        "horizontal_master":          ["model", "particular"],
        "vertical_master":            ["model", "particular"],
        "recuperator_master":         ["type", "model"],
        "gail_gas_burner_master":     ["section", "burner_size"],
        "rad_heat_master":            ["section", "item"],
        "rad_heat_tata_master":       ["section", "item"],
    }

    old_conn = sqlite3.connect(tmp_old)
    new_conn = sqlite3.connect(tmp_new)
    live_conn = sqlite3.connect(LIVE_DB)

    summary = {}
    for table, key_cols in table_keys.items():
        try:
            changed, added, err = _diff_table(old_conn, new_conn, table, key_cols)
        except sqlite3.OperationalError as e:
            summary[table] = f"(skipped: {e})"
            continue
        if err:
            summary[table] = f"(skipped: {err})"
            continue
        if not changed and not added:
            summary[table] = "no changes"
            continue

        # Apply changes to live DB.
        for old_row, new_row in changed:
            sets = ", ".join(f"{c}=?" for c in new_row if c not in key_cols)
            where = " AND ".join(f"{k}=?" for k in key_cols)
            vals = [new_row[c] for c in new_row if c not in key_cols] + [new_row[k] for k in key_cols]
            live_conn.execute(f"UPDATE {table} SET {sets} WHERE {where}", vals)

        for r in added:
            cols = list(r.keys())
            placeholders = ", ".join("?" for _ in cols)
            live_conn.execute(
                f"INSERT OR IGNORE INTO {table} ({', '.join(cols)}) VALUES ({placeholders})",
                [r[c] for c in cols],
            )
        summary[table] = f"{len(changed)} updated, {len(added)} inserted"

    live_conn.commit()
    live_conn.close()
    old_conn.close()
    new_conn.close()

    print()
    print("===== APPLIED =====")
    for t, s in summary.items():
        print(f"  {t:30s} {s}")

    # Best-effort cleanup — Windows occasionally holds SQLite handles a
    # tick longer than Python expects.
    import gc, time
    gc.collect(); time.sleep(0.2)
    for p in (tmp_old, tmp_new):
        try:
            os.remove(p)
        except OSError:
            pass


if __name__ == "__main__":
    main()
