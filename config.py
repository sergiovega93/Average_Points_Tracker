"""
Central configuration. Loads environment from DOTENV_PATH or project .env.
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

def _env_path(name: str, default: Path) -> str:
    raw = os.getenv(name, "").strip()
    return str(default) if not raw else raw


EXCEL_PATCHER_PATH = _env_path(
    "EXCEL_PATCHER_PATH", BASE_DIR / "data" / "Placement_Fee_Patching.xlsx"
)
EXCEL_TRACKER_PATH = _env_path(
    "EXCEL_TRACKER_PATH", BASE_DIR / "data" / "Master_Tracker_Points_Average.xlsx"
)

MA_LENDER_ID = os.getenv("MA_LENDER_ID", "").strip()
MA_API_KEY = os.getenv("MA_API_KEY", "").strip()

MA_API_BASE_V1 = os.getenv(
    "MA_API_BASE_V1", "https://app.mortgageautomator.com/api"
).rstrip("/")
MA_API_BASE_V2 = os.getenv(
    "MA_API_BASE_V2", "https://app.mortgageautomator.com/api/v2"
).rstrip("/")

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "").strip()

# --- Patcher ---
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
WRITE_XLSX_REPORT = os.getenv("WRITE_XLSX_REPORT", "false").lower() in (
    "1",
    "true",
    "yes",
)

# --- Enricher ---
ENRICHER_SHEET_NAME = os.getenv("ENRICHER_SHEET_NAME", "MA API ID").strip()
ENRICHER_LOOKBACK_DAYS = int(os.getenv("ENRICHER_LOOKBACK_DAYS", "180"))
# Column title in Master Tracker that maps from MA placement_fee
ORIGINATION_FEE_COLUMN = os.getenv(
    "ORIGINATION_FEE_COLUMN", "Origination Fee"
).strip()
API_LOAN_ID_COLUMN = os.getenv("API_LOAN_ID_COLUMN", "API Loan ID").strip()


def require_ma_credentials() -> None:
    if not MA_LENDER_ID or not MA_API_KEY:
        print(
            "Missing MA_LENDER_ID or MA_API_KEY. Set them in .env or DOTENV_PATH file.",
            file=sys.stderr,
        )
        raise SystemExit(2)
