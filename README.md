# Loan Automator

Automates Placement Fee patching against Mortgage Automator, enriches the Master Tracker from the MA API, then backfills **Origination Fee** from the patcher workbook when MA still has no fee (for example old-core loans).

## Flow

1. **Placement patcher** — reads the patcher Excel table, `POST loans/update` when the fee in MA is missing/zero (unless skipped by policy).
2. **Points enricher** — refreshes MA-driven columns on the `MA API ID` sheet. If MA returns an empty/zero **Origination Fee**, the cell is **not overwritten** when the workbook already has a non-zero value.
3. **Backfill** — for rows where **Origination Fee** is still empty/zero, writes the amount from the patcher table (same `API Loan ID`).

## Layout

```
loan-automator/
  orchestrator.py          # entry point
  config.py
  steps/
    placement_patcher.py
    points_enricher.py
    placement_fee_backfill.py
  data/                    # place workbooks here (ignored by git)
  logs/                    # run logs + CSV (ignored by git)
```

## Local setup

```powershell
cd loan-automator
py -3 -m pip install -r requirements.txt
copy .env.example .env
# Edit .env: MA_LENDER_ID, MA_API_KEY, EXCEL_* paths
```

Copy your workbooks into `data/` or point `EXCEL_PATCHER_PATH` / `EXCEL_TRACKER_PATH` to absolute paths.

### Smoke test (synthetic workbooks)

```powershell
py -3 dev\create_sample_workbooks.py
$env:MA_LENDER_ID="10134"; $env:MA_API_KEY="your_key"
py -3 orchestrator.py --verify
py -3 orchestrator.py --step backfill --dry-run
```

Replace sample files with real workbooks before using MA steps.

### Optional: shared secrets file (same idea as WSGI on PythonAnywhere)

```powershell
set DOTENV_PATH=C:\path\to\lmc.env
py -3 orchestrator.py --verify
```

## CLI

```text
py -3 orchestrator.py --verify              # paths, tables, credentials
py -3 orchestrator.py --verify --verify-api # also POST loans/get for VERIFY_LOAN_ID
py -3 orchestrator.py --dry-run             # full pipeline, no writes to MA or workbooks
py -3 orchestrator.py --step placement      # single step
py -3 orchestrator.py --step enrich
py -3 orchestrator.py --step backfill     # no MA calls; only needs Excel paths
```

`--dry-run` still performs **GET** calls to MA where the step needs them, so credentials must be valid.

## PythonAnywhere

- Clone the repo, `pip install --user -r requirements.txt`.
- Upload workbooks (SFTP) under e.g. `/home/<user>/loan-automator/data/`.
- Either export variables in `~/.bashrc` or set `DOTENV_PATH` to your secrets file and call `load_dotenv` from `orchestrator.py` (already supported via `python-dotenv` and `DOTENV_PATH`).
- Scheduled task example:

```bash
/home/<user>/.local/bin/python /home/<user>/loan-automator/orchestrator.py
```

If the scheduled task does not load `~/.bashrc`, use a wrapper script that exports variables or sets `DOTENV_PATH`.

## Security

- Never commit `.env` or real `.xlsx` data.
- Rotate any API key that has appeared in old scripts or chat logs.
