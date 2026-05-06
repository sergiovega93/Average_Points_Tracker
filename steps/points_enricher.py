"""
Master Tracker enricher — MA loans/get, merge into the SharePoint workbook
(table tbl_avgpoints on sheet 'MA API ID') via Microsoft Graph Excel API.
Preserves Origination Fee when MA returns empty/zero.
"""
from __future__ import annotations

import hashlib
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
    update_table_cell,
    update_table_column,
    workbook,
)


def _api_auth_token() -> str:
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d-%H")
    raw = f"{config.MA_LENDER_ID}-{config.MA_API_KEY}-loans/get-{ts}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def _headers() -> dict:
    return {
        "ACCOUNT-ID": config.MA_LENDER_ID,
        "API-AUTH": _api_auth_token(),
        "Content-Type": "application/json",
    }


FIELD_MAP = {
    "result.property.address.street": "Street",
    "result.property.address.city": "City",
    "result.property.address.prov": "State",
    "result.property.address.zip": "Zip Code",
    "result.mortgage.interest": "Interest Rate",
    "result.mortgage.funding_date": "Funding Date",
    "result.mortgage.term": "Term Length",
    "result.mortgage.total": "Total Loan Amount",
    "result.mortgage.fees.placement_fee": "Origination Fee",
    "result.repair_cost": "Rehab Budget",
    "result.property.purchase_price": "Purchase Price",
    "result.creation_date": "Creation Date",
    "result.mortgage.type": "Loan Type",
    "result.status": "Status Pipeline",
    "result.custom_fields.macf_ltc_uw_2.response": "LTC UW",
    "result.custom_fields.macf_ltv_uw.response": "LTV UW",
}

ENRICHMENT_ORDER = list(FIELD_MAP.values()) + [
    "Sales Rep",
    "Office Rep",
    "Terms Sent Date",
    "Terms Accepted Date",
    "Title Pending Date",
    "Docs Pending Date",
    "Review Date",
    "Ready to Send Date",
    "Ready to Fund Date",
    "Funded Date",
    "Days: Terms Sent Date → Terms Accepted Date",
    "Days: Terms Accepted Date → Title Pending Date",
    "Days: Title Pending Date → Docs Pending Date",
    "Days: Docs Pending Date → Review Date",
    "Days: Review Date → Ready to Send Date",
    "Days: Ready to Send Date → Ready to Fund Date",
    "Days: Ready to Fund Date → Funded Date",
]


def deep_get(dictionary, keys, default=None):
    for key in keys.split("."):
        if "[" in key and "]" in key:
            base, idx_part = key.split("[", 1)
            index = int(idx_part.replace("]", ""))
            dictionary = dictionary.get(base, [])
            dictionary = dictionary[index] if len(dictionary) > index else {}
        else:
            dictionary = dictionary.get(key, {}) if isinstance(dictionary, dict) else {}
        if dictionary in (None, {}, []):
            return default
    return dictionary


def get_role_user(data, role_name: str) -> str:
    for role in data.get("result", {}).get("roles", []):
        if role.get("name") == role_name and role.get("users"):
            return role["users"][0].get("name", "") or ""
    return ""


def extract_status_dates(data) -> dict:
    milestone_map = {
        "Terms Sent": "Terms Sent Date",
        "Terms Accepted": "Terms Accepted Date",
        "Title Pending": "Title Pending Date",
        "Docs Pending": "Docs Pending Date",
        "Review": "Review Date",
        "Ready to Send": "Ready to Send Date",
        "Ready to Fund": "Ready to Fund Date",
        "funded loan": "Funded Date",
    }
    result = {v: "" for v in milestone_map.values()}
    timestamps: dict = {}

    for note in data.get("result", {}).get("notes", []):
        text = note.get("note", "")
        ts = note.get("date", 0)
        date_str = (
            datetime.fromtimestamp(ts, timezone.utc).strftime("%Y-%m-%d") if ts else ""
        )

        if text.strip() == "funded loan":
            result["Funded Date"] = date_str
            timestamps["Funded Date"] = ts
        elif text.strip() == "Auto-imported from API":
            result["Terms Sent Date"] = date_str
            timestamps["Terms Sent Date"] = ts
        elif text.startswith("changed status to "):
            full = text.replace("changed status to ", "").strip()
            for base, col in milestone_map.items():
                if base in full:
                    result[col] = date_str
                    timestamps[col] = ts
                    break

    pairs = [
        ("Terms Sent Date", "Terms Accepted Date"),
        ("Terms Accepted Date", "Title Pending Date"),
        ("Title Pending Date", "Docs Pending Date"),
        ("Docs Pending Date", "Review Date"),
        ("Review Date", "Ready to Send Date"),
        ("Ready to Send Date", "Ready to Fund Date"),
        ("Ready to Fund Date", "Funded Date"),
    ]
    for start, end in pairs:
        key = f"Days: {start} → {end}"
        if start in timestamps and end in timestamps:
            result[key] = (timestamps[end] - timestamps[start]) // 86400
        else:
            result[key] = ""

    return result


def _fee_is_missing_or_zero(val) -> bool:
    if val is None or val == "":
        return True
    try:
        return abs(float(val)) <= config.NONZERO_TOL
    except (TypeError, ValueError):
        return True


def _utc_now_naive() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def _parse_iso_utc_to_naive(s: str) -> datetime | None:
    s = (s or "").strip()
    if not s:
        return None
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        return None
    if dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    return dt


def _enricher_cutoff_naive(*, lookback_days_override: int | None = None) -> datetime:
    now_n = _utc_now_naive()
    if lookback_days_override is not None:
        return now_n - timedelta(days=max(1, int(lookback_days_override)))
    mode = (config.ENRICHER_LOOKBACK_MODE or "days").strip().lower()
    if mode == "since_last_run":
        p = Path(config.ENRICHER_LAST_RUN_FILE)
        if p.is_file():
            try:
                last = _parse_iso_utc_to_naive(p.read_text(encoding="utf-8"))
                if last is not None:
                    overlap = timedelta(hours=config.ENRICHER_LOOKBACK_OVERLAP_HOURS)
                    return last - overlap
            except OSError:
                pass
        return now_n - timedelta(days=config.ENRICHER_LOOKBACK_DAYS)
    return now_n - timedelta(days=config.ENRICHER_LOOKBACK_DAYS)


def _write_last_enricher_run_marker() -> None:
    p = Path(config.ENRICHER_LAST_RUN_FILE)
    p.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).isoformat()
    p.write_text(stamp + "\n", encoding="utf-8")


def fetch_ma_enrichment_row(loan_id: int, *, session: requests.Session | None = None) -> dict:
    """Single loans/get call mapped through FIELD_MAP + role + status dates helpers."""
    s = session or requests.Session()
    r = s.post(
        f"{config.MA_API_BASE_V1}/loans/get",
        headers=_headers(),
        json={"loan_id": int(loan_id)},
        timeout=config.TIMEOUT_S,
    )
    try:
        data = r.json()
    except ValueError:
        return {config.API_LOAN_ID_COLUMN: int(loan_id), "Error": "loans_get non-json"}
    if not isinstance(data, dict) or not data.get("result"):
        return {
            config.API_LOAN_ID_COLUMN: int(loan_id),
            "Error": "loans_get returned no result",
        }
    row: dict = {config.API_LOAN_ID_COLUMN: int(loan_id)}
    for key, label in FIELD_MAP.items():
        val = deep_get(data, key)
        if "date" in key and isinstance(val, int):
            val = datetime.fromtimestamp(val, tz=timezone.utc).strftime("%Y-%m-%d")
        row[label] = "" if val in (None, {}, []) else val
    row["Sales Rep"] = get_role_user(data, "Sales")
    row["Office Rep"] = get_role_user(data, "Office Rep")
    row.update(extract_status_dates(data))
    return row


def _filter_loan_ids_by_creation_date(
    df: pd.DataFrame, cutoff_naive: datetime
) -> list[int]:
    """Pick rows whose Creation Date is missing OR >= cutoff (naive UTC)."""
    if config.API_LOAN_ID_COLUMN not in df.columns:
        return []
    if "Creation Date" in df.columns:
        cd = pd.to_datetime(df["Creation Date"], errors="coerce", utc=False)
        if hasattr(cd, "dt") and cd.dt.tz is not None:
            cd = cd.dt.tz_convert(None)
        mask = cd.isna() | (cd >= cutoff_naive)
        ids_raw = df.loc[mask, config.API_LOAN_ID_COLUMN].tolist()
    else:
        ids_raw = df[config.API_LOAN_ID_COLUMN].tolist()
    out: list[int] = []
    for raw in ids_raw:
        if raw is None or raw == "":
            continue
        try:
            out.append(int(str(raw).strip().replace(",", "")))
        except (TypeError, ValueError):
            continue
    return out


def _merge_enriched_into_df(
    df: pd.DataFrame, enriched_rows: list[dict]
) -> tuple[pd.DataFrame, dict]:
    """
    Apply enrichment to the in-memory DataFrame. Preserves Origination Fee
    when MA returned empty/zero but the existing cell has a non-zero value.
    Returns (df, per_loan_audit_rows_for_dryrun).
    """
    df = df.astype(object).copy()
    audit_rows: list[dict] = []
    n_orig_ma = n_orig_pres = n_orig_empty = n_orig_api_err = 0
    preserved = 0

    for enriched in enriched_rows:
        try:
            lid = int(enriched.get(config.API_LOAN_ID_COLUMN))
        except (TypeError, ValueError):
            continue
        idx = lookup_row_index_by_key(df, config.API_LOAN_ID_COLUMN, lid)
        if idx is None:
            continue
        api_err = enriched.get("Error")
        fee_ma_raw = enriched.get(config.ORIGINATION_FEE_COLUMN)
        existing_orig_cell = (
            df.at[idx, config.ORIGINATION_FEE_COLUMN]
            if config.ORIGINATION_FEE_COLUMN in df.columns
            else None
        )
        orig_source: str | None = None

        for col_name in ENRICHMENT_ORDER:
            if col_name not in df.columns:
                continue
            val = enriched.get(col_name)
            if col_name == config.ORIGINATION_FEE_COLUMN:
                if _fee_is_missing_or_zero(val):
                    if not _fee_is_missing_or_zero(existing_orig_cell):
                        val = existing_orig_cell
                        preserved += 1
                        orig_source = "preserved_master"
                    else:
                        orig_source = "empty_after_ma"
                else:
                    orig_source = "ma"
            df.at[idx, col_name] = "" if val in (None, {}, []) else val

        if api_err:
            orig_source = orig_source or "api_row_error"
        if not orig_source:
            orig_source = "unknown"

        if orig_source == "ma":
            n_orig_ma += 1
        elif orig_source == "preserved_master":
            n_orig_pres += 1
        elif orig_source == "empty_after_ma":
            n_orig_empty += 1
        elif orig_source == "api_row_error":
            n_orig_api_err += 1

        audit_rows.append(
            {
                "loan_id": lid,
                "origination_fee_from_ma_api": fee_ma_raw,
                "origination_fee_master_cell_before_merge": existing_orig_cell,
                "origination_fee_source_after_step2": orig_source,
                "api_row_error": api_err or "",
            }
        )

    summary = {
        "origination_fee_preserved": preserved,
        "orig_from_ma": n_orig_ma,
        "orig_preserved_master": n_orig_pres,
        "orig_empty_after_ma": n_orig_empty,
        "orig_api_row_error": n_orig_api_err,
    }
    return df, {"summary": summary, "audit_rows": audit_rows}


def _realign_to_current_table(wb, df_merged: pd.DataFrame) -> pd.DataFrame:
    """
    Re-read tbl_avgpoints fresh from Graph and project df_merged's enrichment
    values onto whatever rows currently exist in SharePoint, matched by
    API Loan ID. Any rows present in the live table but not in df_merged
    (e.g. added by a teammate during processing) keep their current SharePoint
    values for the enricher-managed columns. Returns a DataFrame whose row
    count exactly matches the current table size (so update_table_column
    PATCHes will not error on dimension mismatch).
    """
    df_now = read_table(wb, config.ENRICHER_TABLE_NAME)
    api_col = config.API_LOAN_ID_COLUMN
    if df_now.empty or api_col not in df_now.columns:
        return df_now

    lookup: dict[int, pd.Series] = {}
    if api_col in df_merged.columns:
        for _, r in df_merged.iterrows():
            lid_raw = r.get(api_col)
            if pd.isna(lid_raw):
                continue
            try:
                lookup[int(lid_raw)] = r
            except (TypeError, ValueError):
                continue

    enrich_cols = [
        c for c in ENRICHMENT_ORDER
        if c in df_merged.columns and c in df_now.columns
    ]
    if not enrich_cols or not lookup:
        return df_now

    aligned = df_now.copy()
    for idx, fresh_row in df_now.iterrows():
        lid_raw = fresh_row.get(api_col)
        if pd.isna(lid_raw):
            continue
        try:
            lid = int(lid_raw)
        except (TypeError, ValueError):
            continue
        merged_row = lookup.get(lid)
        if merged_row is None:
            continue
        for c in enrich_cols:
            aligned.at[idx, c] = merged_row.get(c)
    return aligned


def run_enricher(
    *,
    dry_run: bool = False,
    lookback_days_override: int | None = None,
) -> dict:
    """
    Enrich the Master Tracker table from MA. With dry_run=True, no PATCH is sent.
    Returns a summary dict.
    """
    cutoff = _enricher_cutoff_naive(lookback_days_override=lookback_days_override)
    mode = (config.ENRICHER_LOOKBACK_MODE or "days").strip().lower()
    if lookback_days_override is not None:
        print(
            f"  Enricher window: **override** lookback_days={lookback_days_override} only "
            f"(cutoff_utc_naive={cutoff:%Y-%m-%d %H:%M})"
        )
    else:
        print(
            f"  Enricher window: mode={mode!r} cutoff_utc_naive={cutoff:%Y-%m-%d %H:%M} "
            f"lookback_days={config.ENRICHER_LOOKBACK_DAYS} "
            f"overlap_h={config.ENRICHER_LOOKBACK_OVERLAP_HOURS}"
        )
    if lookback_days_override is None and mode == "since_last_run":
        p = Path(config.ENRICHER_LAST_RUN_FILE)
        print(f"  Last-run marker file: {p} (exists={p.is_file()})")

    config.require_graph_credentials()

    merge_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    summary: dict = {}

    with workbook(config.EXCEL_TRACKER_SP_PATH, persist=not dry_run) as wb:
        existing_cols = get_table_columns(wb, config.ENRICHER_TABLE_NAME)
        missing_cols = [c for c in ENRICHMENT_ORDER if c not in existing_cols]
        if missing_cols:
            print(f"  Adding missing enrichment columns: {missing_cols}")
            add_table_columns_if_missing(wb, config.ENRICHER_TABLE_NAME, missing_cols)

        df = read_table(wb, config.ENRICHER_TABLE_NAME)
        rows_on_table = len(df)
        if df.empty or config.API_LOAN_ID_COLUMN not in df.columns:
            print(
                f"  [enrich] table {config.ENRICHER_TABLE_NAME!r} empty or missing "
                f"{config.API_LOAN_ID_COLUMN!r}; nothing to do."
            )
            return {
                "rows_on_sheet": rows_on_table,
                "origination_fee_preserved": 0,
                "enrich_loan_count": 0,
                "orig_from_ma": 0,
                "orig_preserved_master": 0,
                "orig_empty_after_ma": 0,
                "orig_api_row_error": 0,
            }

        loan_ids = _filter_loan_ids_by_creation_date(df, cutoff)
        print(f"  Enriching {len(loan_ids)} loan(s)...")

        enriched_rows: list[dict] = []
        t0 = time.time()
        session = requests.Session()
        for loan_id in tqdm(loan_ids, desc="  MA API", unit="loan"):
            try:
                row = fetch_ma_enrichment_row(int(loan_id), session=session)
                enriched_rows.append(row)
            except Exception as e:
                enriched_rows.append(
                    {config.API_LOAN_ID_COLUMN: int(loan_id), "Error": str(e)}
                )
            time.sleep(config.SLEEP_BETWEEN_CALLS)
        print(f"  Fetched in {time.time() - t0:.1f}s — merging in memory...")

        df_merged, merge_meta = _merge_enriched_into_df(df, enriched_rows)

        summary = {
            "rows_on_sheet": rows_on_table,
            "enrich_loan_count": len(loan_ids),
            **merge_meta["summary"],
        }

        if dry_run:
            from .placement_fee_backfill import preview_backfill_on_dataframe

            bf_prev = preview_backfill_on_dataframe(df_merged)
            summary["backfill_preview_rows"] = len(bf_prev)
            audit_rows = merge_meta["audit_rows"]
            if audit_rows:
                p2 = config.LOGS_DIR / f"dryrun_step2_by_loan_{merge_stamp}.csv"
                pd.DataFrame(audit_rows).to_csv(p2, index=False, encoding="utf-8-sig")
                print(f"  [dry-run] per-loan enrich -> {p2}")
            if bf_prev:
                p3 = config.LOGS_DIR / f"dryrun_step3_backfill_preview_{merge_stamp}.csv"
                pd.DataFrame(bf_prev).to_csv(p3, index=False, encoding="utf-8-sig")
                print(f"  [dry-run] backfill preview (post-merge in memory) -> {p3}")
            print(
                f"  [dry-run] would push enrichment; preserved_fee_cells="
                f"{summary['origination_fee_preserved']}; backfill rows that would get "
                f"Excel fee: {len(bf_prev)}"
            )
        else:
            print("  Re-reading table to detect concurrent edits before writing...")
            df_aligned = _realign_to_current_table(wb, df_merged)
            size_diff = len(df_aligned) - len(df_merged)
            if size_diff != 0:
                kind = "added" if size_diff > 0 else "removed"
                print(
                    f"  [enrich] table size changed during processing: "
                    f"{len(df_merged)} -> {len(df_aligned)} rows "
                    f"({abs(size_diff)} {kind} by other editors; "
                    f"new rows kept as-is, will be enriched on next run)."
                )

            touched_cols = [c for c in ENRICHMENT_ORDER if c in df_aligned.columns]
            print(
                f"  Pushing {len(touched_cols)} column(s) to Graph "
                f"(leaves user-managed columns + formulas untouched)..."
            )
            for col_name in touched_cols:
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
                            f"  [enrich] {col_name}: size mismatch on attempt "
                            f"{attempt + 1}/3; re-reading table & realigning..."
                        )
                        df_aligned = _realign_to_current_table(wb, df_merged)
            _write_last_enricher_run_marker()
            print(
                f"  Workbook saved (column-by-column Graph PATCH). "
                f"origination_fee_preserved={summary['origination_fee_preserved']}; "
                f"last-run marker updated."
            )

    return summary


def merge_single_loan_into_master(
    loan_id: int,
    *,
    dry_run: bool = False,
    update_last_run_marker: bool = False,
) -> dict:
    """
    Webhook / single-loan path: fetch one MA loan and PATCH only that table row.
    """
    config.require_graph_credentials()
    enriched = fetch_ma_enrichment_row(int(loan_id))
    if enriched.get("Error"):
        return {"error": enriched["Error"], "loan_id": int(loan_id)}

    with workbook(config.EXCEL_TRACKER_SP_PATH, persist=not dry_run) as wb:
        existing_cols = get_table_columns(wb, config.ENRICHER_TABLE_NAME)
        missing_cols = [c for c in ENRICHMENT_ORDER if c not in existing_cols]
        if missing_cols:
            print(f"  Adding missing enrichment columns: {missing_cols}")
            add_table_columns_if_missing(wb, config.ENRICHER_TABLE_NAME, missing_cols)

        df = read_table(wb, config.ENRICHER_TABLE_NAME)
        idx = lookup_row_index_by_key(df, config.API_LOAN_ID_COLUMN, loan_id)
        if idx is None:
            return {"error": "master_row_not_found", "loan_id": int(loan_id)}

        df_merged, merge_meta = _merge_enriched_into_df(df, [enriched])
        preserved = merge_meta["summary"]["origination_fee_preserved"]

        result = {
            "loan_id": int(loan_id),
            "table_row_index": idx,
            "origination_fee_preserved": preserved,
        }

        if dry_run:
            result["dry_run"] = True
            return result

        touched_cols = [c for c in ENRICHMENT_ORDER if c in df_merged.columns]
        for col_name in touched_cols:
            update_table_cell(
                wb,
                config.ENRICHER_TABLE_NAME,
                idx,
                col_name,
                jsonify_value(df_merged.at[idx, col_name]),
            )
        if update_last_run_marker:
            _write_last_enricher_run_marker()

    return result
