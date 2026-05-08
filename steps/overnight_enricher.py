"""
Overnight / legacy-style Master Tracker enricher: fixed Creation Date window,
unconditional enrichment column overwrites except hybrid Origination Fee rule.
"""
from __future__ import annotations

import logging
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import requests
from tqdm import tqdm

import config
from services.graph_excel import (
    GraphExcelError,
    add_table_columns_if_missing,
    get_table_columns,
    jsonify_value,
    lookup_row_index_by_key,
    read_table,
    update_table_column,
    workbook,
)
from steps.points_enricher import (
    ENRICHMENT_ORDER,
    _api_auth_token,
    _fee_is_missing_or_zero,
    _realign_to_current_table,
    fetch_ma_enrichment_row,
)

log = logging.getLogger(__name__)

# Zapier / Master Tracker: e.g. 12/02/25 02:48PM — evita UserWarning de inferencia lenta
CREATION_DATE_Z_COL = "Creation Date Z"
ZAPIER_DATETIME_FORMAT = "%m/%d/%y %I:%M%p"
ALWAYS_OVERWRITE_COLUMNS = [c for c in ENRICHMENT_ORDER if c != config.ORIGINATION_FEE_COLUMN]


def _select_loan_ids(df: pd.DataFrame, cutoff: datetime) -> tuple[list[int], int]:
    """
    Returns (loan_ids_in_window, count_out_of_window).
    Includes rows where Creation Date Z is present and >= cutoff.
    Rows with no valid API Loan ID are skipped silently.
    """
    FILTER_COL = CREATION_DATE_Z_COL

    if FILTER_COL in df.columns:
        cd = pd.to_datetime(
            df[FILTER_COL],
            format=ZAPIER_DATETIME_FORMAT,
            errors="coerce",
            utc=False,
        )
        if hasattr(cd, "dt") and cd.dt.tz is not None:
            cd = cd.dt.tz_convert(None)
    else:
        cd = pd.Series([pd.NaT] * len(df))

    recent_mask = cd.notna() & (cd >= cutoff)
    out_of_window = int((~recent_mask).sum())

    ids: list[int] = []
    for raw in df.loc[recent_mask, config.API_LOAN_ID_COLUMN]:
        if raw is None or raw == "":
            continue
        try:
            ids.append(int(str(raw).strip().replace(",", "")))
        except (TypeError, ValueError):
            continue
    return ids, out_of_window


def _recent_mask_series(df: pd.DataFrame, cutoff: datetime) -> pd.Series:
    FILTER_COL = CREATION_DATE_Z_COL
    if FILTER_COL in df.columns:
        cd = pd.to_datetime(
            df[FILTER_COL],
            format=ZAPIER_DATETIME_FORMAT,
            errors="coerce",
            utc=False,
        )
        if hasattr(cd, "dt") and cd.dt.tz is not None:
            cd = cd.dt.tz_convert(None)
    else:
        cd = pd.Series([pd.NaT] * len(df), index=df.index)
    return cd.notna() & (cd >= cutoff)


def _parse_loan_id_cell(raw) -> int | None:
    if raw is None or raw == "" or (isinstance(raw, float) and pd.isna(raw)):
        return None
    try:
        return int(str(raw).strip().replace(",", ""))
    except (TypeError, ValueError):
        return None


def run_overnight_enricher(
    *,
    dry_run: bool = False,
    lookback_days: int = 180,
) -> dict:
    """
    Enriches Master Tracker from MA API. Legacy-style: overwrites all enrichment
    columns for every loan whose Creation Date is within the lookback window (present and >= cutoff).
    Exception: Origination Fee is only overwritten when MA returns a non-zero, non-null value,
    preserving any fee that was backfilled from Placement Fee Patching (NOT_SUPPORTED_OLD_CORE case).
    """
    config.require_ma_credentials()
    config.require_graph_credentials()

    cutoff = datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(days=max(1, lookback_days))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log.info(
        "Overnight enricher: lookback_days=%s cutoff=%s dry_run=%s",
        lookback_days,
        cutoff.date(),
        dry_run,
    )

    zero_summary = {
        "rows_on_sheet": 0,
        "loans_in_window": 0,
        "enriched": 0,
        "errors": 0,
        "skipped_out_of_window": 0,
        "origination_fee_overwritten": 0,
        "origination_fee_preserved": 0,
    }

    csv_rows: list[dict] = []
    n_enriched = 0
    n_errors = 0
    n_fee_overwritten = 0
    n_fee_preserved = 0

    with workbook(config.EXCEL_TRACKER_SP_PATH, persist=False) as wb:
        existing_cols = get_table_columns(wb, config.ENRICHER_TABLE_NAME)
        merge_cols = ALWAYS_OVERWRITE_COLUMNS + [config.ORIGINATION_FEE_COLUMN]
        missing_cols = [c for c in merge_cols if c not in existing_cols]
        if missing_cols:
            add_table_columns_if_missing(wb, config.ENRICHER_TABLE_NAME, missing_cols)

        df = read_table(wb, config.ENRICHER_TABLE_NAME)
        rows_on_sheet = len(df)

        if df.empty or config.API_LOAN_ID_COLUMN not in df.columns:
            log.info(
                "Overnight enricher: table %r empty or missing %r; nothing to do.",
                config.ENRICHER_TABLE_NAME,
                config.API_LOAN_ID_COLUMN,
            )
            zero_summary["rows_on_sheet"] = rows_on_sheet
            config.LOGS_DIR.mkdir(parents=True, exist_ok=True)
            csv_path_early = config.LOGS_DIR / f"overnight_enricher_{stamp}.csv"
            pd.DataFrame(
                columns=[
                    "loan_id",
                    "action",
                    "origination_fee_from_ma",
                    "origination_fee_written",
                    "error_detail",
                ]
            ).to_csv(csv_path_early, index=False, encoding="utf-8-sig")
            print(f"  Overnight enricher audit CSV -> {csv_path_early}")
            return zero_summary

        loan_ids, out_of_window = _select_loan_ids(df, cutoff)
        recent_mask = _recent_mask_series(df, cutoff)

        print(
            f"  Overnight enricher: rows_on_sheet={rows_on_sheet} "
            f"loans_in_window={len(loan_ids)} skipped_out_of_window={out_of_window}"
        )

        api_col = config.API_LOAN_ID_COLUMN
        for i in df.index:
            if not bool(recent_mask.loc[i]):
                lid = _parse_loan_id_cell(df.at[i, api_col]) if api_col in df.columns else None
                csv_rows.append(
                    {
                        "loan_id": lid if lid is not None else "",
                        "action": "OUT_OF_WINDOW",
                        "origination_fee_from_ma": "",
                        "origination_fee_written": "",
                        "error_detail": "",
                    }
                )

        for i in df.index:
            if not bool(recent_mask.loc[i]):
                continue
            lid = _parse_loan_id_cell(df.at[i, api_col]) if api_col in df.columns else None
            if lid is None:
                csv_rows.append(
                    {
                        "loan_id": "",
                        "action": "SKIPPED_NO_API_ID",
                        "origination_fee_from_ma": "",
                        "origination_fee_written": "",
                        "error_detail": "",
                    }
                )

    enriched_rows: list[dict] = []
    session = requests.Session()
    t0 = time.time()
    for loan_id in tqdm(loan_ids, desc="  MA API (overnight)", unit="loan"):
        try:
            row = fetch_ma_enrichment_row(int(loan_id), session=session)
            enriched_rows.append(row)
        except Exception as e:
            enriched_rows.append(
                {config.API_LOAN_ID_COLUMN: int(loan_id), "Error": str(e)}
            )
        time.sleep(config.SLEEP_BETWEEN_CALLS)
    print(f"  Overnight enricher: fetched in {time.time() - t0:.1f}s — merging...")

    df_merged = df.copy()
    fee_col = config.ORIGINATION_FEE_COLUMN

    for enriched in enriched_rows:
        try:
            lid = int(enriched.get(config.API_LOAN_ID_COLUMN))
        except (TypeError, ValueError):
            continue
        idx = lookup_row_index_by_key(df_merged, config.API_LOAN_ID_COLUMN, lid)
        if idx is None:
            continue

        is_error = bool(enriched.get("Error"))
        if is_error:
            n_errors += 1
        else:
            n_enriched += 1

        for col in ALWAYS_OVERWRITE_COLUMNS:
            if col not in df_merged.columns:
                continue
            val = enriched.get(col)
            df_merged.at[idx, col] = "" if val in (None, {}, []) else val

        if fee_col in df_merged.columns:
            fee_ma = enriched.get(fee_col)
            if not _fee_is_missing_or_zero(fee_ma):
                df_merged.at[idx, fee_col] = fee_ma
                n_fee_overwritten += 1
            else:
                n_fee_preserved += 1

        fee_written = df_merged.at[idx, fee_col] if fee_col in df_merged.columns else ""
        fee_from_ma = enriched.get(fee_col)
        csv_rows.append(
            {
                "loan_id": lid,
                "action": "ERROR" if is_error else "ENRICHED",
                "origination_fee_from_ma": "" if fee_from_ma in (None, {}, []) else fee_from_ma,
                "origination_fee_written": fee_written,
                "error_detail": (enriched.get("Error", "") if is_error else ""),
            }
        )

    if not dry_run:
        with workbook(config.EXCEL_TRACKER_SP_PATH, persist=True) as wb:
            write_cols = ALWAYS_OVERWRITE_COLUMNS + [fee_col]
            df_aligned = _realign_to_current_table(wb, df_merged)
            print(
                f"  Overnight enricher: pushing {len(write_cols)} column(s) to Graph "
                f"({len(df_merged)} rows)..."
            )
            for col_name in write_cols:
                if col_name not in df_merged.columns:
                    continue
                for attempt in range(3):
                    try:
                        update_table_column(
                            wb,
                            config.ENRICHER_TABLE_NAME,
                            col_name,
                            df_aligned[col_name].tolist(),
                        )
                        break
                    except GraphExcelError as exc:
                        msg = str(exc)
                        is_size_error = (
                            "doesn't match the size" in msg
                            or "doesn't match the dimensions" in msg
                            or "InvalidArgument" in msg
                            and "rows or columns" in msg
                        )
                        if not is_size_error or attempt == 2:
                            raise
                        print(
                            f"  [overnight] {col_name}: size mismatch on attempt "
                            f"{attempt + 1}/3; retrying..."
                        )
                        df_aligned = _realign_to_current_table(wb, df_merged)

    csv_path = config.LOGS_DIR / f"overnight_enricher_{stamp}.csv"
    config.LOGS_DIR.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(csv_rows).to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"  Overnight enricher audit CSV -> {csv_path}")

    return {
        "rows_on_sheet": rows_on_sheet,
        "loans_in_window": len(loan_ids),
        "enriched": n_enriched,
        "errors": n_errors,
        "skipped_out_of_window": out_of_window,
        "origination_fee_overwritten": n_fee_overwritten,
        "origination_fee_preserved": n_fee_preserved,
    }
