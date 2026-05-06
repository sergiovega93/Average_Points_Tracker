"""
Run placement + enrich + targeted backfill for a single API loan id (webhook / CLI).
All workbook I/O goes through Microsoft Graph (SharePoint Excel API).
"""
from __future__ import annotations

import logging

import requests

import config

from .placement_fee_backfill import run_backfill
from .placement_patcher import get_table_dataframe, patch_one, to_number
from .points_enricher import merge_single_loan_into_master

log = logging.getLogger(__name__)


def _patcher_fee_and_street(loan_id: int) -> tuple[float, str] | None:
    df = get_table_dataframe(config.EXCEL_PATCHER_SP_PATH, config.PATCHER_TABLE_NAME)
    for _, row in df.iterrows():
        lid = to_number(row.get(config.PATCHER_COL_LOAN_ID))
        if lid is None or int(lid) != int(loan_id):
            continue
        fee = to_number(row.get(config.PATCHER_COL_ORIG_FEE))
        street = row.get(config.PATCHER_COL_STREET)
        if fee is None or fee <= 0:
            return None
        st = "" if street is None else str(street)
        return round(float(fee), 2), st
    return None


def run_pipeline_for_loan_id(loan_id: int, *, dry_run: bool = False) -> dict:
    """
    1) Patch MA from patcher row if present.
    2) Merge MA fields into the single Master row for this loan_id.
    3) Backfill Origination Fee from patcher table for this loan_id only if still empty.
    """
    lid = int(loan_id)
    session = requests.Session()
    out: dict = {"loan_id": lid}

    pr = _patcher_fee_and_street(lid)
    if pr:
        fee, street = pr
        out["placement"] = patch_one(session, lid, fee, street, dry_run=dry_run)
        log.info("single_loan placement: %s", out["placement"].get("action"))
    else:
        out["placement"] = None
        log.info("single_loan: no patcher row for loan_id=%s (skip placement)", lid)

    out["enrich"] = merge_single_loan_into_master(lid, dry_run=dry_run)
    log.info("single_loan enrich: %s", out["enrich"])

    out["backfill"] = run_backfill(dry_run=dry_run, only_loan_ids={lid})
    log.info("single_loan backfill: %s", out["backfill"])

    return out
