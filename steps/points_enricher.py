"""
Master Tracker enricher — MA loans/get, merge into workbook.
Preserves Origination Fee when MA returns empty/zero (see config.ORIGINATION_FEE_COLUMN).
"""
from __future__ import annotations

import hashlib
import time
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests
from openpyxl import load_workbook
from tqdm import tqdm

import config


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


def run_enricher(*, dry_run: bool = False) -> dict:
    """
    Enrich Master Tracker from MA. Does not save workbook if dry_run.
    Returns a small summary dict (rows written, preserved fees count).
    """
    cutoff = datetime.now(timezone.utc) - timedelta(days=config.ENRICHER_LOOKBACK_DAYS)
    print(f"  Cutoff (UTC): {cutoff.strftime('%Y-%m-%d %H:%M')} lookback_days={config.ENRICHER_LOOKBACK_DAYS}")

    df = pd.read_excel(
        config.EXCEL_TRACKER_PATH,
        sheet_name=config.ENRICHER_SHEET_NAME,
        engine="openpyxl",
    )
    if "Creation Date" in df.columns:
        df["Creation Date"] = pd.to_datetime(df["Creation Date"], errors="coerce")
        mask = df["Creation Date"].isna() | (df["Creation Date"] >= cutoff.replace(tzinfo=None))
        loan_ids = df.loc[mask, config.API_LOAN_ID_COLUMN].dropna().astype(int)
    else:
        loan_ids = df[config.API_LOAN_ID_COLUMN].dropna().astype(int)

    print(f"  Enriching {len(loan_ids)} loan(s)...")
    enriched_rows = []
    t0 = time.time()
    session = requests.Session()

    for loan_id in tqdm(loan_ids, desc="  MA API", unit="loan"):
        try:
            r = session.post(
                "https://app.mortgageautomator.com/api/loans/get",
                headers=_headers(),
                json={"loan_id": int(loan_id)},
                timeout=config.TIMEOUT_S,
            )
            data = r.json()
            row = {config.API_LOAN_ID_COLUMN: int(loan_id)}
            for key, label in FIELD_MAP.items():
                val = deep_get(data, key)
                if "date" in key and isinstance(val, int):
                    val = datetime.fromtimestamp(
                        val, tz=timezone.utc
                    ).strftime("%Y-%m-%d")
                row[label] = "" if val in (None, {}, []) else val
            row["Sales Rep"] = get_role_user(data, "Sales")
            row["Office Rep"] = get_role_user(data, "Office Rep")
            row.update(extract_status_dates(data))
            enriched_rows.append(row)
        except Exception as e:
            enriched_rows.append({config.API_LOAN_ID_COLUMN: int(loan_id), "Error": str(e)})

        time.sleep(config.SLEEP_BETWEEN_CALLS)

    print(f"  Fetched in {time.time() - t0:.1f}s — merging into sheet...")
    df_enriched = pd.DataFrame(enriched_rows)
    wb = load_workbook(config.EXCEL_TRACKER_PATH)
    ws = wb[config.ENRICHER_SHEET_NAME]

    header_row = [cell.value for cell in ws[1]]
    static_cols = 3
    while (
        static_cols < len(header_row)
        and header_row[static_cols]
        and header_row[static_cols] not in ENRICHMENT_ORDER
    ):
        static_cols += 1
    col_start = static_cols + 1
    api_id_list = [cell.value for cell in ws["B"][1:]]

    preserved = 0
    for offset, col_name in enumerate(ENRICHMENT_ORDER):
        ws.cell(row=1, column=col_start + offset).value = col_name

    for idx, api_id in enumerate(api_id_list, start=2):
        if api_id is None or api_id == "":
            continue
        try:
            lid = int(api_id)
        except (TypeError, ValueError):
            continue
        match = df_enriched[df_enriched[config.API_LOAN_ID_COLUMN] == lid]
        if match.empty:
            continue
        d = match.iloc[0].to_dict()
        for offset, col_name in enumerate(ENRICHMENT_ORDER):
            val = d.get(col_name)
            if col_name == config.ORIGINATION_FEE_COLUMN:
                if _fee_is_missing_or_zero(val):
                    existing = ws.cell(row=idx, column=col_start + offset).value
                    if not _fee_is_missing_or_zero(existing):
                        val = existing
                        preserved += 1
            ws.cell(row=idx, column=col_start + offset).value = (
                "" if val in (None, {}, []) else val
            )

    summary = {"rows_on_sheet": len(api_id_list), "origination_fee_preserved": preserved}
    if dry_run:
        print(f"  [dry-run] would save workbook; preserved_fee_cells={preserved}")
        wb.close()
    else:
        wb.save(config.EXCEL_TRACKER_PATH)
        print(f"  Workbook saved. origination_fee_preserved={preserved}")
    return summary
