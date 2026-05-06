"""
Aggregate CSV for --dry-run full pipeline (placement + enrich + virtual backfill preview).
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

import config


def write_dry_run_aggregate_csv(
    stamp: str,
    placement_results: list[dict],
    enrich_summary: dict,
) -> Path:
    """
    Two-column CSV: metric, value (+ notes row for caveats).
    """
    pr = placement_results or []
    n = len(pr)
    skipped = sum(1 for r in pr if r.get("action") == "SKIPPED_ALREADY_SET")
    would = sum(
        1
        for r in pr
        if r.get("action") == "DRY_RUN_WOULD_ATTEMPT_PATCH"
        or r.get("action") in ("WOULD_PATCH", "WOULD_REVIEW")  # legacy
    )
    get_err = sum(1 for r in pr if r.get("action") == "ERROR_LOANS_GET")
    updated = sum(1 for r in pr if r.get("action") == "UPDATED")
    err = sum(1 for r in pr if r.get("action") == "ERROR")
    oc = sum(1 for r in pr if r.get("action") == "NOT_SUPPORTED_OLD_CORE")

    rows = [
        ("placement_loans_evaluated_from_excel", n, ""),
        (
            "placement_skipped_already_set_fee_in_ma",
            skipped,
            "GET showed non-zero placement_fee; no POST in dry-run.",
        ),
        (
            "placement_dry_run_would_attempt_patch",
            would,
            "MA fee was zero/missing; real run needs POST to know UPDATED vs OLD_CORE vs ERROR.",
        ),
        (
            "placement_loans_get_errors",
            get_err,
            "loans/get failed or returned no result; fix credentials before trusting dry-run.",
        ),
        ("placement_updated_live_run_only", updated, ""),
        ("placement_error_live_run_only", err, ""),
        ("placement_old_core_live_run_only", oc, ""),
        ("", "", ""),
        (
            "enrich_loans_in_api_scope",
            enrich_summary.get("enrich_loan_count", ""),
            "",
        ),
        (
            "enrich_origination_fee_source_ma",
            enrich_summary.get("orig_from_ma", ""),
            "After virtual step2 merge.",
        ),
        (
            "enrich_origination_fee_preserved_existing_master",
            enrich_summary.get("orig_preserved_master", ""),
            "MA empty/zero; kept existing Master cell.",
        ),
        (
            "enrich_origination_fee_still_empty_after_ma",
            enrich_summary.get("orig_empty_after_ma", ""),
            "Would need step3 backfill if patcher has fee.",
        ),
        (
            "backfill_preview_would_fill_from_patcher",
            int(enrich_summary.get("backfill_preview_rows", 0) or 0),
            "Simulated on in-memory sheet after step2 merge (dry-run accurate).",
        ),
    ]

    out = Path(config.LOGS_DIR) / f"dryrun_aggregate_{stamp}.csv"
    pd.DataFrame(rows, columns=["metric", "value", "notes"]).to_csv(
        out, index=False, encoding="utf-8-sig"
    )
    return out


def dry_run_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")
