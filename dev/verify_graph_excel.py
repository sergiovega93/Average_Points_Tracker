"""
Smoke test for the Microsoft Graph Excel integration.

Run from the repo root:
    py -3 dev\\verify_graph_excel.py

Checks (read-only, no PATCH):
    1) MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET present.
    2) Graph token can be acquired.
    3) SharePoint site + default drive resolve.
    4) Both workbooks (patcher + tracker) are reachable by their drive-relative paths.
    5) Tables tbl_of and tbl_avgpoints exist; columns are listed.
    6) Reads first 3 rows of each table.

Optional:
    Set CONFIRM_WRITE=1 to also try a no-op cell PATCH against tbl_avgpoints
    (writes the SAME value the cell already has). Use only for live debugging.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import config  # noqa: E402
from services.graph_excel import (  # noqa: E402
    get_table_columns,
    get_token,
    read_table,
    resolve_drive_id,
    resolve_item_id,
    resolve_site_id,
    update_table_cell,
    workbook,
)


def _section(title: str) -> None:
    print("\n--- " + title + " ---")


def main() -> int:
    ok = True

    _section("1) Microsoft Graph credentials")
    missing = [
        n
        for n in ("MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET")
        if not getattr(config, n, "")
    ]
    if missing:
        print(f"  MISSING: {', '.join(missing)}")
        return 2
    print("  MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET present")

    _section("2) Token")
    try:
        tok = get_token()
        print(f"  token len={len(tok)}")
    except Exception as exc:
        print(f"  FAIL get_token: {exc}")
        return 2

    _section("3) Site + drive")
    site_id = resolve_site_id(tok)
    drive_id = resolve_drive_id(tok, site_id)
    print(f"  site_id ends with ...{site_id[-12:]}")
    print(f"  drive_id ends with ...{drive_id[-12:]}")

    for label, sp_path, table in (
        (
            "Patcher",
            config.EXCEL_PATCHER_SP_PATH,
            config.PATCHER_TABLE_NAME,
        ),
        (
            "Tracker",
            config.EXCEL_TRACKER_SP_PATH,
            config.ENRICHER_TABLE_NAME,
        ),
    ):
        _section(f"4) {label}: file + table")
        try:
            item_id = resolve_item_id(tok, drive_id, sp_path)
            print(f"  file ok: {sp_path}  (item ...{item_id[-8:]})")
            with workbook(sp_path, persist=False) as wb:
                cols = get_table_columns(wb, table)
                print(f"  table {table!r} has {len(cols)} columns:")
                for i, c in enumerate(cols):
                    print(f"    [{i:02d}] {c!r}")
                df = read_table(wb, table)
                print(f"  rows on table: {len(df)}")
        except Exception as exc:
            print(f"  FAIL {label}: {exc}")
            ok = False

    if os.getenv("CONFIRM_WRITE") == "1":
        _section("5) Optional no-op PATCH (single cell)")
        try:
            with workbook(config.EXCEL_TRACKER_SP_PATH, persist=True) as wb:
                df = read_table(wb, config.ENRICHER_TABLE_NAME)
                if df.empty or config.ORIGINATION_FEE_COLUMN not in df.columns:
                    print("  skip: tracker table empty or missing Origination Fee column")
                else:
                    current = df.at[0, config.ORIGINATION_FEE_COLUMN]
                    print(f"  rewriting row 0 Origination Fee -> {current!r}")
                    update_table_cell(
                        wb,
                        config.ENRICHER_TABLE_NAME,
                        0,
                        config.ORIGINATION_FEE_COLUMN,
                        current,
                    )
                    print("  no-op single-cell PATCH ok")
        except Exception as exc:
            print(f"  FAIL no-op single-cell PATCH: {exc}")
            ok = False

        _section("6) Optional no-op PATCH (single column rewrite)")
        try:
            from services.graph_excel import update_table_column

            with workbook(config.EXCEL_TRACKER_SP_PATH, persist=True) as wb:
                df = read_table(wb, config.ENRICHER_TABLE_NAME)
                if df.empty or config.ORIGINATION_FEE_COLUMN not in df.columns:
                    print("  skip: tracker table empty")
                else:
                    values = df[config.ORIGINATION_FEE_COLUMN].tolist()
                    print(
                        f"  rewriting full {config.ORIGINATION_FEE_COLUMN!r} column "
                        f"({len(values)} cells) with the SAME values"
                    )
                    update_table_column(
                        wb,
                        config.ENRICHER_TABLE_NAME,
                        config.ORIGINATION_FEE_COLUMN,
                        values,
                    )
                    print("  no-op column PATCH ok")
                    print(
                        "  ** verify in Excel that 'Points %', 'Weighted Average Points', "
                        "and other user-managed columns are unchanged. **"
                    )
        except Exception as exc:
            print(f"  FAIL no-op column PATCH: {exc}")
            ok = False
    else:
        print("\n(set CONFIRM_WRITE=1 to also test no-op PATCHes on the tracker)")

    print("\nALL OK" if ok else "\nFAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
