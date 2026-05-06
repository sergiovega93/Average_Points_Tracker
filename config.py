"""
Central configuration. Loads environment from DOTENV_PATH or project .env.
All workbook I/O now goes through Microsoft Graph (SharePoint Excel API);
the legacy EXCEL_PATCHER_PATH / EXCEL_TRACKER_PATH local-file vars are gone.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent

_dotenv_path = os.getenv("DOTENV_PATH", "").strip()
if _dotenv_path:
    load_dotenv(_dotenv_path, override=False)
else:
    load_dotenv(BASE_DIR / ".env", override=False)

LOGS_DIR = BASE_DIR / "logs"
LOGS_DIR.mkdir(parents=True, exist_ok=True)


# --- Mortgage Automator ---
# Accept the names this repo defined originally, AND the names the rest of mysite
# already uses (MA_ACCOUNT_ID / MA_BASE_URL) so we share one ~/.secrets/lmc.env.
MA_LENDER_ID = (
    os.getenv("MA_LENDER_ID") or os.getenv("MA_ACCOUNT_ID") or ""
).strip()
MA_API_KEY = os.getenv("MA_API_KEY", "").strip()

_ma_base_default = "https://app.mortgageautomator.com/api"
_ma_base_env = (
    os.getenv("MA_API_BASE_V1") or os.getenv("MA_BASE_URL") or _ma_base_default
).rstrip("/")
MA_API_BASE_V1 = _ma_base_env
MA_API_BASE_V2 = os.getenv("MA_API_BASE_V2", f"{_ma_base_env}/v2").rstrip("/")

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "").strip()


# --- Microsoft Graph (SharePoint Excel) ---
MS_TENANT_ID = os.getenv("MS_TENANT_ID", "").strip()
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", "").strip()
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "").strip()

SP_HOSTNAME = os.getenv("SP_HOSTNAME", "loanmountaincapital.sharepoint.com").strip()
SP_SITE_RELATIVE_PATH = os.getenv(
    "SP_SITE_RELATIVE_PATH", "sites/LoanMountainCapital-General"
).strip()

EXCEL_PATCHER_SP_PATH = os.getenv(
    "EXCEL_PATCHER_SP_PATH",
    "General/Project Raptor 3/Underwritting/Placement Fee Patching.xlsx",
).strip()
EXCEL_TRACKER_SP_PATH = os.getenv(
    "EXCEL_TRACKER_SP_PATH",
    "General/Project Raptor 3/Sales/Master Tracker Points Average.xlsx",
).strip()


# --- Patcher (table on the Placement workbook) ---
PATCHER_TABLE_NAME = os.getenv("PATCHER_TABLE_NAME", "tbl_of").strip()
PATCHER_COL_LOAN_ID = os.getenv("PATCHER_COL_LOAN_ID", "API Loan ID").strip()
PATCHER_COL_ORIG_FEE = os.getenv("PATCHER_COL_ORIG_FEE", "Origination Fee").strip()
PATCHER_COL_STREET = os.getenv("PATCHER_COL_STREET", "Address").strip()
SKIP_IF_EXISTING_NONZERO = (
    os.getenv("SKIP_IF_EXISTING_NONZERO", "true").lower() in ("1", "true", "yes")
)
NONZERO_TOL = float(os.getenv("NONZERO_TOL", "0.005"))
VERIFY_TOL = float(os.getenv("VERIFY_TOL", "0.01"))
VERIFY_SETTLE_WAIT = float(os.getenv("VERIFY_SETTLE_WAIT", "1.2"))
RETRIES_PER_ENDPOINT = int(os.getenv("RETRIES_PER_ENDPOINT", "1"))
SLEEP_BETWEEN_CALLS = float(os.getenv("SLEEP_BETWEEN_CALLS", "0.4"))
TIMEOUT_S = int(os.getenv("TIMEOUT_S", "30"))


# --- Enricher (table on the Master Tracker workbook) ---
ENRICHER_SHEET_NAME = os.getenv("ENRICHER_SHEET_NAME", "MA API ID").strip()
ENRICHER_TABLE_NAME = os.getenv("ENRICHER_TABLE_NAME", "tbl_avgpoints").strip()

# days: include rows with Creation Date >= now - N days (UTC).
# since_last_run: use logs/last_enricher_run.iso (minus overlap); if missing, fall back to days.
ENRICHER_LOOKBACK_MODE = os.getenv("ENRICHER_LOOKBACK_MODE", "days").strip().lower()
ENRICHER_LOOKBACK_DAYS = int(os.getenv("ENRICHER_LOOKBACK_DAYS", "14"))
ENRICHER_LOOKBACK_OVERLAP_HOURS = int(os.getenv("ENRICHER_LOOKBACK_OVERLAP_HOURS", "24"))
_last_run_default = LOGS_DIR / "last_enricher_run.iso"
ENRICHER_LAST_RUN_FILE = os.getenv(
    "ENRICHER_LAST_RUN_FILE", str(_last_run_default)
).strip()

ORIGINATION_FEE_COLUMN = os.getenv(
    "ORIGINATION_FEE_COLUMN", "Origination Fee"
).strip()
API_LOAN_ID_COLUMN = os.getenv("API_LOAN_ID_COLUMN", "API Loan ID").strip()


# --- Webhook / job queue (used by flask_webhook.py + loan_automator_queue.py) ---
LOAN_AUTOMATOR_WEBHOOK_SECRET = os.getenv("LOAN_AUTOMATOR_WEBHOOK_SECRET", "").strip()
QUEUE_DB_PATH = os.getenv(
    "QUEUE_DB_PATH",
    str(LOGS_DIR / "loan_automator_queue.sqlite"),
).strip()


def require_ma_credentials() -> None:
    if not MA_LENDER_ID or not MA_API_KEY:
        print(
            "Missing MA_LENDER_ID or MA_API_KEY. Set them in .env or DOTENV_PATH file.",
            file=sys.stderr,
        )
        raise SystemExit(2)


def require_graph_credentials() -> None:
    missing = [
        n for n in ("MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET")
        if not globals().get(n)
    ]
    if missing:
        print(
            f"Missing Microsoft Graph credentials: {', '.join(missing)}. "
            "Set them in .env or DOTENV_PATH file.",
            file=sys.stderr,
        )
        raise SystemExit(2)
