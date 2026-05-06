"""
Microsoft Graph Excel helper for loan-automator.

All workbook reads/writes go through Graph (no openpyxl, no local file copy).
Designed for the two SharePoint workbooks declared in config.py:
    EXCEL_PATCHER_SP_PATH   (drive-relative, contains table tbl_of)
    EXCEL_TRACKER_SP_PATH   (drive-relative, contains sheet 'MA API ID' with table tbl_avgpoints)

Auth uses the same client-credentials flow as mysite/sharepoint_draw_upload_webhook.py
(MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET).
"""
from __future__ import annotations

import logging
import threading
import time
from contextlib import contextmanager
from typing import Any, Iterator
from urllib.parse import quote

import pandas as pd
import requests

import config

log = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL_TMPL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
DEFAULT_TIMEOUT = (15, 60)

_token_cache: dict[str, Any] = {"token": None, "exp": 0.0}
_id_cache: dict[str, str] = {}
_id_lock = threading.Lock()


class GraphExcelError(RuntimeError):
    """Wrap any non-200 from Graph with a small body excerpt."""


def _require_ms_credentials() -> None:
    missing = [
        n for n in ("MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET")
        if not getattr(config, n, "")
    ]
    if missing:
        raise GraphExcelError(
            f"Missing Microsoft Graph credentials: {', '.join(missing)} "
            "(set in .env / DOTENV_PATH)."
        )


def get_token() -> str:
    """Cached app-only access token. Refreshes ~5 min before expiry."""
    now = time.time()
    if _token_cache["token"] and _token_cache["exp"] - 300 > now:
        return _token_cache["token"]
    _require_ms_credentials()
    url = TOKEN_URL_TMPL.format(tenant=config.MS_TENANT_ID)
    data = {
        "client_id": config.MS_CLIENT_ID,
        "client_secret": config.MS_CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }
    r = requests.post(url, data=data, timeout=DEFAULT_TIMEOUT)
    if r.status_code != 200:
        raise GraphExcelError(f"Graph token HTTP {r.status_code}: {r.text[:600]}")
    js = r.json()
    tok = js.get("access_token")
    if not tok:
        raise GraphExcelError(f"Graph token missing access_token: {str(js)[:600]}")
    _token_cache["token"] = tok
    _token_cache["exp"] = now + int(js.get("expires_in", 3600))
    return tok


def _h(token: str, *, session_id: str | None = None, content_type: bool = False) -> dict:
    h = {"Authorization": f"Bearer {token}"}
    if content_type:
        h["Content-Type"] = "application/json"
    if session_id:
        h["Workbook-Session-Id"] = session_id
    return h


def _do(method: str, url: str, **kw) -> dict:
    """Execute a Graph call, raise GraphExcelError on non-2xx, return parsed JSON or {}."""
    kw.setdefault("timeout", DEFAULT_TIMEOUT)
    r = requests.request(method, url, **kw)
    if not (200 <= r.status_code < 300):
        raise GraphExcelError(
            f"Graph {method} {url.split('graph.microsoft.com')[-1]} "
            f"HTTP {r.status_code}: {(r.text or '')[:600]}"
        )
    if not r.content:
        return {}
    try:
        return r.json()
    except ValueError:
        return {}


def resolve_site_id(token: str) -> str:
    key = f"site::{config.SP_HOSTNAME}::{config.SP_SITE_RELATIVE_PATH}"
    with _id_lock:
        if key in _id_cache:
            return _id_cache[key]
    url = f"{GRAPH_BASE}/sites/{config.SP_HOSTNAME}:/{config.SP_SITE_RELATIVE_PATH}"
    js = _do("GET", url, headers=_h(token))
    site_id = js.get("id")
    if not site_id:
        raise GraphExcelError(f"site lookup missing id: {str(js)[:400]}")
    with _id_lock:
        _id_cache[key] = site_id
    return site_id


def resolve_drive_id(token: str, site_id: str) -> str:
    key = f"drive::{site_id}"
    with _id_lock:
        if key in _id_cache:
            return _id_cache[key]
    url = f"{GRAPH_BASE}/sites/{site_id}/drive"
    js = _do("GET", url, headers=_h(token))
    drive_id = js.get("id")
    if not drive_id:
        raise GraphExcelError(f"default drive lookup missing id: {str(js)[:400]}")
    with _id_lock:
        _id_cache[key] = drive_id
    return drive_id


def _quote_path(path: str) -> str:
    """Quote a drive-relative path for Graph (path-style addressing)."""
    return quote(path.lstrip("/"), safe="/")


def resolve_item_id(token: str, drive_id: str, sp_path: str) -> str:
    key = f"item::{drive_id}::{sp_path}"
    with _id_lock:
        if key in _id_cache:
            return _id_cache[key]
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{_quote_path(sp_path)}"
    js = _do("GET", url, headers=_h(token))
    item_id = js.get("id")
    if not item_id:
        raise GraphExcelError(f"item lookup '{sp_path}' missing id: {str(js)[:400]}")
    with _id_lock:
        _id_cache[key] = item_id
    return item_id


def open_workbook_session(token: str, drive_id: str, item_id: str, *, persist: bool) -> str:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/createSession"
    js = _do(
        "POST",
        url,
        headers=_h(token, content_type=True),
        json={"persistChanges": bool(persist)},
    )
    sid = js.get("id")
    if not sid:
        raise GraphExcelError(f"createSession missing id: {str(js)[:400]}")
    return sid


def close_workbook_session(token: str, drive_id: str, item_id: str, session_id: str) -> None:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/closeSession"
    try:
        _do("POST", url, headers=_h(token, session_id=session_id, content_type=True))
    except GraphExcelError as exc:
        log.warning("closeSession ignored: %s", exc)


class WorkbookHandle:
    """Bundle of (token, drive_id, item_id, session_id) for a single workbook."""

    def __init__(self, token: str, drive_id: str, item_id: str, session_id: str, sp_path: str):
        self.token = token
        self.drive_id = drive_id
        self.item_id = item_id
        self.session_id = session_id
        self.sp_path = sp_path

    @property
    def workbook_url(self) -> str:
        return f"{GRAPH_BASE}/drives/{self.drive_id}/items/{self.item_id}/workbook"

    def headers(self, *, content_type: bool = False) -> dict:
        return _h(self.token, session_id=self.session_id, content_type=content_type)


@contextmanager
def workbook(sp_path: str, *, persist: bool) -> Iterator[WorkbookHandle]:
    """Open a Graph workbook session, yield a handle, always close on exit."""
    token = get_token()
    site_id = resolve_site_id(token)
    drive_id = resolve_drive_id(token, site_id)
    item_id = resolve_item_id(token, drive_id, sp_path)
    sid = open_workbook_session(token, drive_id, item_id, persist=persist)
    log.info(
        "Graph workbook session opened: %s (persist=%s)", sp_path, persist
    )
    try:
        yield WorkbookHandle(token, drive_id, item_id, sid, sp_path)
    finally:
        close_workbook_session(token, drive_id, item_id, sid)
        log.info("Graph workbook session closed: %s", sp_path)


def read_table(handle: WorkbookHandle, table_name: str) -> pd.DataFrame:
    """
    Read an entire Excel Table by name. Returns a DataFrame with header row as columns.
    Uses /tables/{name}/range to get headers + values in one request.
    """
    url = f"{handle.workbook_url}/tables('{quote(table_name)}')/range"
    js = _do("GET", url, headers=handle.headers())
    values = js.get("values") or []
    if not values:
        return pd.DataFrame()
    headers = [("" if h is None else str(h).strip()) for h in values[0]]
    df = pd.DataFrame(values[1:], columns=headers)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def get_table_address(handle: WorkbookHandle, table_name: str) -> str:
    """Return the absolute address ('Sheet1!A1:C42') of the table's full range."""
    url = f"{handle.workbook_url}/tables('{quote(table_name)}')/range"
    js = _do("GET", url, headers=handle.headers())
    addr = js.get("address") or ""
    if not addr:
        raise GraphExcelError(f"table {table_name!r} address missing")
    return addr


def get_table_columns(handle: WorkbookHandle, table_name: str) -> list[str]:
    url = f"{handle.workbook_url}/tables('{quote(table_name)}')/columns?$select=name"
    js = _do("GET", url, headers=handle.headers())
    return [c.get("name", "") for c in (js.get("value") or [])]


def add_table_columns_if_missing(
    handle: WorkbookHandle, table_name: str, missing: list[str]
) -> None:
    """Append columns to the right of an existing table. No-op for empty list."""
    if not missing:
        return
    url = f"{handle.workbook_url}/tables('{quote(table_name)}')/columns/add"
    for name in missing:
        body = {"name": name, "values": None}
        _do("POST", url, headers=handle.headers(content_type=True), json=body)
        log.info("Added column %r to table %r", name, table_name)


def update_table_row(
    handle: WorkbookHandle,
    table_name: str,
    row_index_zero_based: int,
    ordered_values: list[Any],
) -> None:
    """
    PATCH a single row of a table by 0-based data-row index (excludes header).
    `ordered_values` must match the table's current column order.
    """
    url = (
        f"{handle.workbook_url}/tables('{quote(table_name)}')"
        f"/rows/itemAt(index={int(row_index_zero_based)})"
    )
    body = {"values": [list(ordered_values)]}
    _do("PATCH", url, headers=handle.headers(content_type=True), json=body)


def update_table_cell(
    handle: WorkbookHandle,
    table_name: str,
    row_index_zero_based: int,
    column_name: str,
    value: Any,
) -> None:
    """
    PATCH a single cell at the intersection of a table row and a named column.
    Uses range-of-table-column-data-body addressing.
    """
    url = (
        f"{handle.workbook_url}/tables('{quote(table_name)}')"
        f"/columns('{quote(column_name)}')/dataBodyRange"
    )
    js = _do("GET", url, headers=handle.headers())
    address = js.get("address")
    if not address:
        raise GraphExcelError(
            f"dataBodyRange address missing for {table_name}.{column_name}"
        )
    sheet_name, col_range = address.split("!", 1)
    sheet_name = sheet_name.strip("'")
    cell_addr = _shift_first_row(col_range, row_index_zero_based)
    rng_url = (
        f"{handle.workbook_url}/worksheets('{quote(sheet_name)}')"
        f"/range(address='{quote(cell_addr)}')"
    )
    _do(
        "PATCH",
        rng_url,
        headers=handle.headers(content_type=True),
        json={"values": [[value]]},
    )


def _shift_first_row(col_range: str, row_offset: int) -> str:
    """
    Given a single-column range like 'B2:B42' and a 0-based offset, return the
    single-cell address at that offset, e.g. row_offset=0 -> 'B2', row_offset=5 -> 'B7'.
    """
    if ":" not in col_range:
        return col_range
    start, _ = col_range.split(":", 1)
    col_letters = "".join(ch for ch in start if ch.isalpha())
    row_digits = "".join(ch for ch in start if ch.isdigit())
    if not col_letters or not row_digits:
        raise GraphExcelError(f"cannot parse column range {col_range!r}")
    return f"{col_letters}{int(row_digits) + int(row_offset)}"


def lookup_row_index_by_key(
    df: pd.DataFrame, key_col: str, key_value: Any
) -> int | None:
    """0-based data-row index in the DataFrame returned by read_table."""
    if key_col not in df.columns or df.empty:
        return None
    try:
        target = int(str(key_value).strip().replace(",", ""))
    except (TypeError, ValueError):
        return None
    for i, raw in enumerate(df[key_col].tolist()):
        if raw is None or raw == "":
            continue
        try:
            if int(str(raw).strip().replace(",", "")) == target:
                return i
        except (TypeError, ValueError):
            continue
    return None


def update_worksheet_range(
    handle: WorkbookHandle,
    sheet_name: str,
    address: str,
    values_2d: list[list[Any]],
) -> None:
    """PATCH a 2D range on a worksheet (e.g. 'A2:Z500')."""
    url = (
        f"{handle.workbook_url}/worksheets('{quote(sheet_name)}')"
        f"/range(address='{quote(address)}')"
    )
    _do(
        "PATCH",
        url,
        headers=handle.headers(content_type=True),
        json={"values": values_2d},
    )


def get_table_data_body_address(handle: WorkbookHandle, table_name: str) -> str:
    """Address (e.g. \"'MA API ID'!A2:Z500\") of the table's body, excluding header."""
    url = f"{handle.workbook_url}/tables('{quote(table_name)}')/dataBodyRange"
    js = _do("GET", url, headers=handle.headers())
    addr = js.get("address") or ""
    if not addr:
        raise GraphExcelError(f"dataBodyRange address missing for {table_name!r}")
    return addr


def get_table_column_data_body_address(
    handle: WorkbookHandle, table_name: str, column_name: str
) -> str:
    """Address of a single named column's data body (excludes header)."""
    url = (
        f"{handle.workbook_url}/tables('{quote(table_name)}')"
        f"/columns('{quote(column_name)}')/dataBodyRange"
    )
    js = _do("GET", url, headers=handle.headers())
    addr = js.get("address") or ""
    if not addr:
        raise GraphExcelError(
            f"dataBodyRange address missing for {table_name}.{column_name}"
        )
    return addr


def update_table_column(
    handle: WorkbookHandle,
    table_name: str,
    column_name: str,
    values_1d: list[Any],
) -> None:
    """
    PATCH all data cells of a named column with a 1D list of values.
    Leaves all other columns (and any formulas they contain) untouched.
    """
    address = get_table_column_data_body_address(handle, table_name, column_name)
    sheet, rng = split_address(address)
    values_2d = [[jsonify_value(v)] for v in values_1d]
    url = (
        f"{handle.workbook_url}/worksheets('{quote(sheet)}')"
        f"/range(address='{quote(rng)}')"
    )
    _do(
        "PATCH",
        url,
        headers=handle.headers(content_type=True),
        json={"values": values_2d},
    )


def split_address(address: str) -> tuple[str, str]:
    """Split a Graph address ('Sheet1!A1:C5' or \"'My Sheet'!A1:C5\") into (sheet, range)."""
    if "!" not in address:
        return "", address
    sheet, rng = address.split("!", 1)
    return sheet.strip().strip("'"), rng


def jsonify_value(v: Any) -> Any:
    """Coerce a Python/pandas value into a Graph-safe primitive (None, str, int, float, bool)."""
    import math as _math

    if v is None:
        return None
    if isinstance(v, float):
        if _math.isnan(v) or _math.isinf(v):
            return None
        return v
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v
    if isinstance(v, str):
        return v
    try:
        import pandas as _pd  # noqa: F401

        if v is _pd.NaT or (hasattr(v, "isna") and bool(v.isna())):
            return None
    except Exception:
        pass
    try:
        import numpy as _np

        if isinstance(v, _np.integer):
            return int(v)
        if isinstance(v, _np.floating):
            f = float(v)
            return None if (_math.isnan(f) or _math.isinf(f)) else f
        if isinstance(v, _np.bool_):
            return bool(v)
    except Exception:
        pass
    return str(v)


def df_to_jsonable_2d(df) -> list[list[Any]]:
    """Convert a pandas DataFrame's body (no headers) into a Graph-safe 2D list."""
    out: list[list[Any]] = []
    for _, row in df.iterrows():
        out.append([jsonify_value(v) for v in row.tolist()])
    return out
