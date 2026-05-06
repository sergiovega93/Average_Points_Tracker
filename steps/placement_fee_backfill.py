"""
Backfill Origination Fee on the Master Tracker (table tbl_avgpoints) from the
Placement Fee Patching workbook (table tbl_of), via Microsoft Graph Excel API.
Only fills rows where Origination Fee is still empty/zero after MA enrichment.
"""
from __future__ import annotations

from datetime import datetime

import pandas as pd

import config
from services.graph_excel import (
    lookup_row_index_by_key,
    read_table,
    update_table_cell,
    workbook,
)

from .placement_patcher import get_table_dataframe, to_number


def _fee_is_missing_or_zero(val) -> bool:
    if val is None or val == "":
        return True
    try:
        return abs(float(val)) <= config.NONZERO_TOL
    except (TypeError, ValueError):
        return True


def load_patcher_fees() -> dict[int, float]:
    """Map of API Loan ID -> Origination Fee from the patcher SharePoint table."""
    df = get_table_dataframe(config.EXCEL_PATCHER_SP_PATH, config.PATCHER_TABLE_NAME)
    out: dict[int, float] = {}
    for _, row in df.iterrows():
        lid = to_number(row.get(config.PATCHER_COL_LOAN_ID))
        fee = to_number(row.get(config.PATCHER_COL_ORIG_FEE))
        if lid is None or fee is None or fee <= 0:
            continue
        out[int(lid)] = round(float(fee), 2)
    return out


def preview_backfill_on_dataframe(
    df: pd.DataFrame, fees_by_loan: dict[int, float] | None = None
) -> list[dict]:
    """
    Given the Master Tracker DataFrame already merged with MA fields (in memory),
    return rows where Origination Fee is missing/zero AND patcher has a fee.
    """
    if df.empty:
        return []
    if fees_by_loan is None:
        fees_by_loan = load_patcher_fees()
    if (
        config.API_LOAN_ID_COLUMN not in df.columns
        or config.ORIGINATION_FEE_COLUMN not in df.columns
    ):
        return []
    out: list[dict] = []
    for i, raw_id in enumerate(df[config.API_LOAN_ID_COLUMN].tolist()):
        if raw_id is None or raw_id == "":
            continue
        try:
            lid = int(str(raw_id).strip().replace(",", ""))
        except (TypeError, ValueError):
            continue
        current = df.at[i, config.ORIGINATION_FEE_COLUMN] if i in df.index else None
        if not _fee_is_missing_or_zero(current):
            continue
        desired = fees_by_loan.get(lid)
        if desired is None:
            continue
        out.append(
            {
                "loan_id": lid,
                "table_row_index_zero_based": i,
                "fee_would_set_from_patcher": desired,
            }
        )
    return out


def run_backfill(*, dry_run: bool = False, only_loan_ids: set[int] | None = None) -> dict:
    """
    Open the Master Tracker workbook via Graph, read tbl_avgpoints, fill missing
    Origination Fee cells from the patcher table, write back via Graph.
    """
    fees_by_loan = load_patcher_fees()

    filled = 0
    skipped_has_value = 0
    skipped_no_patcher_row = 0
    rows_on_table = 0

    with workbook(config.EXCEL_TRACKER_SP_PATH, persist=not dry_run) as wb:
        df = read_table(wb, config.ENRICHER_TABLE_NAME)
        rows_on_table = len(df)
        if df.empty or config.API_LOAN_ID_COLUMN not in df.columns:
            print(
                f"  [backfill] tbl {config.ENRICHER_TABLE_NAME!r} empty or missing "
                f"{config.API_LOAN_ID_COLUMN!r} column; nothing to do."
            )
        elif config.ORIGINATION_FEE_COLUMN not in df.columns:
            raise ValueError(
                f"Column {config.ORIGINATION_FEE_COLUMN!r} not found in table "
                f"{config.ENRICHER_TABLE_NAME!r}. Run enricher once so the column exists."
            )
        else:
            for i, raw_id in enumerate(df[config.API_LOAN_ID_COLUMN].tolist()):
                if raw_id is None or raw_id == "":
                    continue
                try:
                    lid = int(str(raw_id).strip().replace(",", ""))
                except (TypeError, ValueError):
                    continue
                if only_loan_ids is not None and lid not in only_loan_ids:
                    continue
                current = (
                    df.at[i, config.ORIGINATION_FEE_COLUMN] if i in df.index else None
                )
                if not _fee_is_missing_or_zero(current):
                    skipped_has_value += 1
                    continue
                desired = fees_by_loan.get(lid)
                if desired is None:
                    skipped_no_patcher_row += 1
                    continue
                if dry_run:
                    print(
                        f"  [dry-run] table row {i} loan {lid}: would set fee -> {desired}"
                    )
                else:
                    update_table_cell(
                        wb,
                        config.ENRICHER_TABLE_NAME,
                        i,
                        config.ORIGINATION_FEE_COLUMN,
                        desired,
                    )
                filled += 1

    summary = {
        "filled_from_patcher": filled,
        "skipped_row_already_has_fee": skipped_has_value,
        "skipped_no_fee_in_patcher_table": skipped_no_patcher_row,
        "patcher_table_loans": len(fees_by_loan),
        "rows_on_master_table": rows_on_table,
    }

    if dry_run:
        print(f"  [dry-run] backfill summary: {summary}")
    else:
        print(f"  Backfill saved. summary: {summary}")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    with open(config.LOGS_DIR / f"backfill_summary_{stamp}.txt", "w", encoding="utf-8") as f:
        for k, v in summary.items():
            f.write(f"{k}={v}\n")

    return summary


def find_master_row_index_for_loan(df: pd.DataFrame, loan_id: int) -> int | None:
    """0-based table-row index for a loan_id, used by single-loan merge helpers."""
    return lookup_row_index_by_key(df, config.API_LOAN_ID_COLUMN, loan_id)
