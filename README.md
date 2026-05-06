# Loan Automator

Automates Placement Fee patching against Mortgage Automator, enriches the Master Tracker from the MA API, then backfills **Origination Fee** from the patcher workbook when MA still has no fee (for example old-core loans).

All workbook reads/writes go through **Microsoft Graph (SharePoint Excel API)**: the patcher and tracker live in SharePoint and are edited in place via Graph workbook sessions. There is no local-file fallback.

## Flow

1. **Placement patcher** — reads the patcher SharePoint table (`tbl_of`), `POST loans/update` when the fee in MA is missing/zero (unless skipped by policy).
2. **Points enricher** — refreshes MA-driven columns on the tracker table (`tbl_avgpoints`, sheet `MA API ID`). If MA returns an empty/zero **Origination Fee**, the cell is **not overwritten** when the workbook already has a non-zero value.
3. **Backfill** — for rows where **Origination Fee** is still empty/zero, writes the amount from the patcher table (matched by `API Loan ID`).

## Layout

```
loan-automator/
  orchestrator.py              # entry point (CLI + scheduled task)
  flask_webhook.py             # register_routes(app) for PythonAnywhere / Zapier
  loan_automator_queue.py      # SQLite queue + background worker thread
  config.py
  services/
    graph_excel.py             # Microsoft Graph Excel helper (token, sessions, table I/O)
  steps/
    placement_patcher.py
    points_enricher.py
    placement_fee_backfill.py
    single_loan.py             # one loan_id pipeline (webhook / --loan-id)
    dry_run_report.py
  dev/
    verify_graph_excel.py      # smoke test for Graph + SharePoint paths
  logs/                        # run logs + CSV + queue sqlite (ignored by git)
```

## Local setup

```powershell
cd loan-automator
py -3 -m pip install -r requirements.txt
copy .env.example .env
# Edit .env: MA_LENDER_ID, MA_API_KEY, MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET
```

The default SharePoint paths in `config.py` already point to:

- Patcher: `General/Project Raptor 3/Underwritting/Placement Fee Patching.xlsx` (table `tbl_of`)
- Tracker: `General/Project Raptor 3/Sales/Master Tracker Points Average.xlsx` (table `tbl_avgpoints`, sheet `MA API ID`)

Override `EXCEL_PATCHER_SP_PATH` / `EXCEL_TRACKER_SP_PATH` in `.env` if those files move.

### Smoke test

```powershell
py -3 dev\verify_graph_excel.py
```

Validates Graph credentials, resolves site + drive, opens both workbooks, lists table columns, prints the first 3 rows. Set `CONFIRM_WRITE=1` to also try a no-op cell PATCH (writes the cell's current value back).

### Enricher window (avoid re-fetching 180 days every run)

- Default is **`ENRICHER_LOOKBACK_MODE=days`** with **`ENRICHER_LOOKBACK_DAYS=14`** (override in `.env`, e.g. `2` for quick tests).
- For frequent runs / webhooks, set **`ENRICHER_LOOKBACK_MODE=since_last_run`**: after each successful enricher save, a timestamp is written to `logs/last_enricher_run.iso`; the next run enriches rows whose **Creation Date** is on or after that time minus **`ENRICHER_LOOKBACK_OVERLAP_HOURS`** (default 24). If the marker file is missing, it falls back to **`ENRICHER_LOOKBACK_DAYS`**.

### Full dry-run (`py -3 orchestrator.py --dry-run`)

Writes under `logs/`:

- `placement_patch_*_dryrun.csv` — per loan from the patcher workbook (`DRY_RUN_WOULD_ATTEMPT_PATCH` vs `SKIPPED_ALREADY_SET` vs `ERROR_LOANS_GET`).
- `dryrun_step2_by_loan_*.csv` — each merged Master row: origination fee from MA, existing cell, and source after virtual step 2 (`ma` / `preserved_master` / `empty_after_ma`).
- `dryrun_step3_backfill_preview_*.csv` — loans that would receive Origination Fee from the patcher table after step 2 (computed on the **in-memory** merged table; step 3 disk PATCH is skipped in full dry-run so this stays consistent).
- `dryrun_aggregate_*.csv` — rolled-up counts + notes (including that **UPDATED vs OLD_CORE** only appear after a real `loans/update`).

### Optional: shared secrets file (same idea as WSGI on PythonAnywhere)

```powershell
set DOTENV_PATH=C:\path\to\lmc.env
py -3 orchestrator.py --verify
```

## CLI

```text
py -3 orchestrator.py --verify              # MA creds, Graph creds, SharePoint reachability, tables exist
py -3 orchestrator.py --verify --verify-api # also POST loans/get for VERIFY_LOAN_ID
py -3 orchestrator.py --dry-run             # full pipeline, no writes to MA or workbooks (Graph PATCH is skipped)
py -3 orchestrator.py --step placement      # single step
py -3 orchestrator.py --step enrich
py -3 orchestrator.py --step backfill       # no MA calls; only needs Graph
py -3 orchestrator.py --loan-id 1202091     # single loan: patch + enrich row + backfill that id
py -3 orchestrator.py --loan-id 1202091 --dry-run
```

`--dry-run` still performs **GET** calls to MA where the step needs them, so credentials must be valid.

### Zapier / Flask webhook (PythonAnywhere)

On PA this repo is cloned at **`/home/sergiovegadev93/loan-automator`** (sibling of `mysite`). The Flask app `mysite` registers the route via a shim in `mysite/app/blueprints/loan_automator.py` that loads `flask_webhook.py` from the sibling repo with `importlib`.

1. Set the secrets in **`/home/sergiovegadev93/.secrets/lmc.env`** (loaded by WSGI):
   - `MA_LENDER_ID`, `MA_API_KEY`
   - `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`
   - `LOAN_AUTOMATOR_WEBHOOK_SECRET` (also set the same value in Zapier).
2. **Zapier Custom Request (POST)** to `https://sergiovegadev93.pythonanywhere.com/webhook/loan-automator` with:
   - **Headers:** `Content-Type: application/json`, **`X-Webhook-Secret`**: same as `LOAN_AUTOMATOR_WEBHOOK_SECRET`.
   - **Body (JSON):** `{"loan_id": 1202091}` or **`{"api_loan_id": 1202091}`** (both supported, aligned with other zaps).

Optional: JSON field `"secret"` instead of header. Also supported: **`X-Loan-Automator-Secret`**, **`Authorization: Bearer …`**, **`?token=…`**.

3. The handler returns **202** with `job_id` and starts a **daemon thread** worker that processes the SQLite queue (**`QUEUE_DB_PATH`**, default under `logs/`). One worker per web process; on PythonAnywhere that is usually one process — jobs are still serialized in SQLite, so two webhooks never edit the workbook at the same time.

Local testing without a secret: set **`LOAN_AUTOMATOR_WEBHOOK_ALLOW_INSECURE=1`** (never in production).

### Live pilot (writes MA + Master — use with care)

Process only the **last 5** loans from the Placement table and enrich Master rows from roughly the **last 2 days** (UTC):

```powershell
cd C:\Users\serch\loan-automator
py -3 orchestrator.py --placement-last-n 5 --enrich-lookback-days 2
```

Omit both flags to run the **full** worklist and the enricher window from `.env`. For a dry-run with the same scope:

```powershell
py -3 orchestrator.py --dry-run --placement-last-n 5 --enrich-lookback-days 2
```

## PythonAnywhere

- Clone the repo to `/home/sergiovegadev93/loan-automator`, `pip install --user -r requirements.txt`.
- Put secrets in `/home/sergiovegadev93/.secrets/lmc.env`; make sure the WSGI loads it (or export `DOTENV_PATH=/home/sergiovegadev93/.secrets/lmc.env`).
- Reload the web app after every secret or code change. Auto-deploy on push is wired through `mysite`'s `/webhooks/git-pull-loan-automator` (same `GITHUB_WEBHOOK_SECRET`).
- Scheduled task example:

```bash
DOTENV_PATH=/home/sergiovegadev93/.secrets/lmc.env \
  /home/sergiovegadev93/.local/bin/python /home/sergiovegadev93/loan-automator/orchestrator.py
```

## Microsoft Graph permissions

The Azure AD app referenced by `MS_CLIENT_ID` must have **application** permission `Files.ReadWrite.All` or `Sites.ReadWrite.All` (with admin consent granted) so the client-credentials flow can edit cells via the Excel workbook session API. The same app is used by `mysite/sharepoint_draw_upload`, so this is usually already provisioned.

## Security

- Never commit `.env` or real `.xlsx` data.
- Rotate any API key that has appeared in old scripts or chat logs.
