"""
Certified Stone Inventory Application
-------------------------------------
Internal tool for certificate lookup, inventory management,
and media reference for certified stones.

Revision notes (vs. previous version)
  * Upserts now MERGE instead of overwrite: a failed/partial re-fetch can no
    longer blank existing grades; new non-empty values (e.g. a new video URL)
    now win over old ones; rows without a real certificate number are skipped;
    duplicates inside a batch are collapsed.
  * SQLite schema now stores the three media columns and auto-migrates.
  * Sheets are written with RAW input (no formula injection / no number
    coercion); Excel exports are sanitised; reads never numericise.
  * Header handling is name-based; unknown/extra sheet columns are preserved.
  * Outbound fetches go through `safe_get` (host allow-list for lab PDFs,
    private-address blocking for video URLs, redirect re-validation, size caps).
  * Drive files are private unless `[drive] public_link = true` in secrets.
  * Optional password gate via `[auth] password` in secrets.
  * Edit tab saves only changed cells and detects conflicting edits.
  * Add tab keeps results in session_state (row-click no longer wipes them);
    save status is honest (no false "saved" on quota fallback).
  * Parsers consolidated into one code path; lab detection / shape matching fixed.
  * Certificate ZIP and preview downloads are parallel / cached.
  * Dead Playwright code and Streamlit-compat fallbacks removed
    (requires a Streamlit version with st.dialog and dataframe selection).
"""
import streamlit as st
import pandas as pd
import requests
import re
import io
import hmac
import ipaddress
import logging
import socket
import sqlite3
import os
import zipfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from datetime import datetime, date
from typing import Optional, Dict, Any, List, Tuple
from urllib.parse import urlparse, urljoin
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Optional: pdfplumber for better PDF text extraction
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

def _safe_name(s: Any) -> str:
    return re.sub(r"[^\w\-]+", "_", str(s))


# Streamlit renamed use_container_width=True -> width="stretch" (around 1.50).
def _stretch_kwargs() -> Dict[str, Any]:
    try:
        ver = tuple(int(x) for x in st.__version__.split(".")[:2])
    except Exception:
        ver = (0, 0)
    return {"width": "stretch"} if ver >= (1, 50) else {"use_container_width": True}


_STRETCH = _stretch_kwargs()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("stone_inventory")

st.set_page_config(
    page_title="Certified Stone Inventory",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------
# Optional password gate  ([auth] password = "..." in secrets)
# For real SSO use st.login(); this is a minimal guard for internal tools.
# -------------------------------------------------
def _require_login() -> None:
    try:
        expected = st.secrets.get("auth", {}).get("password")
    except Exception:
        expected = None
    if not expected:
        return
    if st.session_state.get("_authed"):
        return
    st.markdown("## Certified Stone Inventory")
    pw = st.text_input("Password", type="password", key="_login_pw")
    if pw:
        if hmac.compare_digest(str(pw), str(expected)):
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()


_require_login()

# -------------------------------------------------
# Constants
# -------------------------------------------------
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "inventory.db")

INVENTORY_COLUMNS = [
    "Certificate Number", "DESCRIPTION", "SHAPE AND CUT", "CARAT WEIGHT",
    "COLOR GRADE", "CLARITY GRADE", "MEASUREMENTS", "CUT", "POLISH",
    "SYMMETRY", "FLUORESCENCE", "Process",
    "FGI Item Number", "Date Shipped from India", "For stock or Customer",
    "Production Order #, if for stock", "Current Location/Status",
    "Price Per Carat", "Price For stone", "CERTIFICATE LINK", "Comment",
    "Video URL", "Video Backup Drive Link",
    "Certificate PDF Backup Drive Link",
    "Lab", "Status", "Notes", "date_added", "last_updated"
]

EDITABLE_COLUMNS = [
    "FGI Item Number", "Date Shipped from India", "For stock or Customer",
    "Production Order #, if for stock", "Current Location/Status",
    "Price Per Carat", "Price For stone", "Comment",
    "Video URL", "Video Backup Drive Link",
    "Certificate PDF Backup Drive Link",
]

# Fields that, when present, mean a row holds real certificate data
CORE_FIELDS = ["SHAPE AND CUT", "CARAT WEIGHT", "COLOR GRADE", "CLARITY GRADE", "MEASUREMENTS"]

# Mapping of sheet/app column name -> SQLite column name
DB_COLUMN_MAP = {
    "Certificate Number": "certificate_number",
    "DESCRIPTION": "description",
    "SHAPE AND CUT": "shape_and_cut",
    "CARAT WEIGHT": "carat_weight",
    "COLOR GRADE": "color_grade",
    "CLARITY GRADE": "clarity_grade",
    "MEASUREMENTS": "measurements",
    "CUT": "cut_grade",
    "POLISH": "polish",
    "SYMMETRY": "symmetry",
    "FLUORESCENCE": "fluorescence",
    "Process": "process",
    "FGI Item Number": "fgi_item_number",
    "Date Shipped from India": "date_shipped_from_india",
    "For stock or Customer": "for_stock_or_customer",
    "Production Order #, if for stock": "production_order",
    "Current Location/Status": "current_location_status",
    "Price Per Carat": "price_per_carat",
    "Price For stone": "price_for_stone",
    "CERTIFICATE LINK": "certificate_link",
    "Comment": "comment",
    "Video URL": "video_url",
    "Video Backup Drive Link": "video_backup_drive_link",
    "Certificate PDF Backup Drive Link": "certificate_pdf_backup_drive_link",
    "Lab": "lab",
    "Status": "status",
    "Notes": "notes",
    "date_added": "date_added",
    "last_updated": "last_updated",
}
DB_COLUMNS = [DB_COLUMN_MAP[c] for c in INVENTORY_COLUMNS]

# Hosts we are willing to fetch lab certificates from (suffix match)
CERT_ALLOWED_HOST_SUFFIXES = ("igi.org", "gia.edu", "amazonaws.com")

MAX_PDF_BYTES = 25 * 1024 * 1024
MAX_VIDEO_BYTES = 300 * 1024 * 1024
CARAT_TOLERANCE = 0.12          # fractional carat search window (e.g. "1.5ct")

FAILURE_STATUSES = {"Not Found", "Parse Error", "Link Ready", "Error", "Unknown Format"}


def _is_failure_status(status: str) -> bool:
    s = (status or "").strip()
    return s in FAILURE_STATUSES or s.startswith("HTTP ")


def _valid_cert(cert: Any) -> bool:
    c = str(cert or "").strip()
    return bool(c) and c.lower() not in ("unknown", "none", "nan")


def _blank(v: Any) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


# -------------------------------------------------
# Safe outbound HTTP
# -------------------------------------------------
_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
)


def _build_session() -> requests.Session:
    s = requests.Session()
    retry = Retry(
        total=2, backoff_factor=0.6,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET", "HEAD"), raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=10, pool_maxsize=10)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    s.headers.update({"User-Agent": _USER_AGENT})
    return s


_HTTP = _build_session()


def _host_allowed(host: str, suffixes: Optional[Tuple[str, ...]]) -> bool:
    host = (host or "").lower().rstrip(".")
    return any(host == s or host.endswith("." + s) for s in (suffixes or ()))


def _is_public_host(host: str) -> bool:
    """Reject loopback / private / link-local / reserved targets (SSRF guard).
    Note: does not defend against DNS rebinding between check and connect."""
    try:
        infos = socket.getaddrinfo(host, None)
    except socket.gaierror:
        return False
    for info in infos:
        ip = ipaddress.ip_address(info[4][0])
        if (ip.is_private or ip.is_loopback or ip.is_link_local
                or ip.is_reserved or ip.is_multicast or ip.is_unspecified):
            return False
    return True


class UnsafeURL(Exception):
    pass


def safe_get(
    url: str,
    *,
    allowed_host_suffixes: Optional[Tuple[str, ...]] = None,
    stream: bool = False,
    timeout: int = 30,
    headers: Optional[Dict[str, str]] = None,
    max_redirects: int = 5,
) -> requests.Response:
    """GET with scheme check, optional host allow-list, public-address check and
    per-hop redirect validation. Raises UnsafeURL if any hop is not permitted."""
    current = str(url).strip()
    for _ in range(max_redirects + 1):
        p = urlparse(current)
        if p.scheme != "https" and not (p.scheme == "http" and allowed_host_suffixes is None):
            raise UnsafeURL(f"Blocked URL scheme: {p.scheme or 'none'}")
        if not p.hostname:
            raise UnsafeURL("Invalid URL")
        if allowed_host_suffixes is not None and not _host_allowed(p.hostname, allowed_host_suffixes):
            raise UnsafeURL(f"Host not allowed: {p.hostname}")
        if not _is_public_host(p.hostname):
            raise UnsafeURL(f"Blocked non-public address: {p.hostname}")
        resp = _HTTP.get(current, timeout=timeout, headers=headers, stream=stream, allow_redirects=False)
        if resp.is_redirect or resp.status_code in (301, 302, 303, 307, 308):
            nxt = resp.headers.get("Location")
            resp.close()
            if not nxt:
                raise UnsafeURL("Redirect without Location")
            current = urljoin(current, nxt)
            continue
        return resp
    raise UnsafeURL("Too many redirects")


def read_capped(resp: requests.Response, max_bytes: int) -> bytes:
    """Read a streamed response body, aborting if it exceeds max_bytes."""
    buf = bytearray()
    for chunk in resp.iter_content(chunk_size=1024 * 256):
        if not chunk:
            continue
        buf.extend(chunk)
        if len(buf) > max_bytes:
            resp.close()
            raise ValueError(f"File exceeds {max_bytes // (1024 * 1024)} MB limit")
    return bytes(buf)


# -------------------------------------------------
# Excel export helpers (formula-injection safe)
# -------------------------------------------------
def _xl_safe(v: Any) -> Any:
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + v
    return v


@st.cache_data(show_spinner=False, max_entries=8)
def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Inventory") -> bytes:
    safe = df.copy()
    for c in safe.columns:
        if not pd.api.types.is_numeric_dtype(safe[c]) and not pd.api.types.is_bool_dtype(safe[c]):
            safe[c] = safe[c].map(_xl_safe)      # works for object and pandas>=3 string dtypes
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        safe.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()

# -------------------------------------------------
# Database layer: Google Sheets (preferred) + SQLite fallback
# -------------------------------------------------
def _has_gsheets_secrets() -> bool:
    try:
        return bool(
            "gcp_service_account" in st.secrets
            and "sheets" in st.secrets
            and st.secrets["sheets"].get("spreadsheet_id")
        )
    except Exception:
        return False


def _has_gdrive_folder() -> bool:
    try:
        return bool(st.secrets.get("drive", {}).get("folder_id"))
    except Exception:
        return False


def _drive_public_link_enabled() -> bool:
    """Drive files are private unless [drive] public_link = true."""
    try:
        return bool(st.secrets.get("drive", {}).get("public_link", False))
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def _get_google_creds():
    from google.oauth2.service_account import Credentials
    # NOTE: "drive.file" is narrower than "drive" and is enough for files this
    # app creates in a folder shared with the service account. Switch back to
    # the full "drive" scope only if your folder setup requires it.
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
    ]
    return Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scopes
    )


@st.cache_resource(show_spinner=False)
def _get_gspread_client():
    import gspread
    return gspread.authorize(_get_google_creds())


@st.cache_resource(show_spinner=False)
def _get_drive_service():
    from googleapiclient.discovery import build
    return build("drive", "v3", credentials=_get_google_creds(), cache_discovery=False)


@st.cache_resource(show_spinner=False)
def _open_worksheet():
    """Cached worksheet handle – avoids re-opening on every rerun."""
    client = _get_gspread_client()
    sid = st.secrets["sheets"]["spreadsheet_id"]
    ws_name = st.secrets["sheets"].get("worksheet_name", "Inventory")
    sh = client.open_by_key(sid)
    try:
        ws = sh.worksheet(ws_name)
    except Exception:
        ws = sh.add_worksheet(title=ws_name, rows=2000, cols=len(INVENTORY_COLUMNS))
        ws.append_row(INVENTORY_COLUMNS, value_input_option="RAW")
    return ws


def use_gsheets() -> bool:
    return _has_gsheets_secrets()


# ---------------- SQLite ----------------
@contextmanager
def db():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA journal_mode=WAL")
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    """Create the SQLite table and migrate older databases (adds missing columns)."""
    cols_sql = ",\n            ".join(
        f"{c} TEXT" + (" PRIMARY KEY" if c == "certificate_number" else "")
        for c in DB_COLUMNS
    )
    with db() as conn:
        conn.execute(f"CREATE TABLE IF NOT EXISTS inventory (\n            {cols_sql}\n        )")
        have = {r["name"] for r in conn.execute("PRAGMA table_info(inventory)")}
        for c in DB_COLUMNS:
            if c not in have:
                conn.execute(f"ALTER TABLE inventory ADD COLUMN {c} TEXT")
                logger.info("Migrated SQLite: added column %s", c)


# ---------------- Row merge logic ----------------
def _s(v: Any) -> str:
    return "" if _blank(v) else str(v).strip()


def _normalize_row(row: Dict[str, Any]) -> Dict[str, str]:
    return {c: _s(row.get(c)) for c in INVENTORY_COLUMNS}


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def merge_row(new: Dict[str, Any], old: Optional[Dict[str, Any]]) -> Optional[Dict[str, str]]:
    """
    Combine a freshly fetched/edited row with the stored one.

    * New row, nothing stored        -> insert as is.
    * Stored row exists              -> non-empty new values win; empty new
                                        values never blank stored ones.
    * New row is a failure status
      (Not Found / Link Ready / ...) and stored row already has certificate
      data                           -> only user-editable fields may change.
    Returns None when nothing would change.
    """
    now = _now()
    new = _normalize_row(new)
    if old is None:
        out = dict(new)
        out["date_added"] = out["date_added"] or now
        out["last_updated"] = now
        return out

    old = _normalize_row(old)
    old_has_data = any(old[c] for c in CORE_FIELDS)

    if _is_failure_status(new["Status"]) and old_has_data:
        out = dict(old)
        changed = False
        for col in EDITABLE_COLUMNS:
            if new[col] and new[col] != old[col]:
                out[col] = new[col]
                changed = True
        if not changed:
            return None
        out["last_updated"] = now
        return out

    out = dict(old)
    changed = False
    for col in INVENTORY_COLUMNS:
        if col in ("date_added", "last_updated"):
            continue
        if new[col] and new[col] != old[col]:
            out[col] = new[col]
            changed = True
    if not changed:
        return None
    out["date_added"] = old["date_added"] or new["date_added"] or now
    out["last_updated"] = now
    return out


def _collapse_batch(rows: List[Dict[str, Any]]) -> Tuple[List[Dict[str, str]], int]:
    """Drop rows with no real certificate number and merge duplicates inside the batch."""
    by_cert: Dict[str, Dict[str, str]] = {}
    skipped = 0
    for r in rows or []:
        n = _normalize_row(r)
        cert = n["Certificate Number"]
        if not _valid_cert(cert):
            skipped += 1
            continue
        if cert in by_cert:
            for col, val in n.items():
                if val and col not in ("date_added", "last_updated"):
                    by_cert[cert][col] = val
        else:
            by_cert[cert] = n
    return list(by_cert.values()), skipped


# ---------------- Google Sheets helpers ----------------
def _gspread_error_detail(e: Exception) -> str:
    """Pull status + body from gspread APIError when available."""
    parts = [str(e)]
    resp = getattr(e, "response", None)
    if resp is not None:
        try:
            parts.append(f"HTTP {getattr(resp, 'status_code', '?')}")
            text = getattr(resp, "text", None) or getattr(resp, "content", b"")
            if isinstance(text, bytes):
                text = text.decode("utf-8", errors="ignore")
            if text:
                parts.append(str(text)[:400])
        except Exception:
            logger.debug("Could not parse gspread error body", exc_info=True)
    return " | ".join(parts)[:600]


def _sheets_api_call(fn, retries: int = 3):
    """Run a Sheets API call with limited retry on quota / transient errors."""
    for attempt in range(retries):
        try:
            return fn()
        except Exception as e:
            msg = _gspread_error_detail(e)
            retryable = any(
                x in msg
                for x in ("429", "Quota", "RATE_LIMIT", "503", "500", "Backend Error", "Timeout", "timed out")
            )
            if not retryable or attempt >= retries - 1:
                raise
            time.sleep(3 + 2 ** attempt)


def _ws_update(ws, range_name: str, values_2d: List[List[Any]], value_input_option: str = "RAW"):
    """gspread v5/v6 compatible single-range update."""
    return _sheets_api_call(lambda: ws.update(
        values=values_2d, range_name=range_name, value_input_option=value_input_option,
    ))


def _ensure_sheet_headers(ws) -> List[str]:
    """
    Make sure every app column exists in row 1 WITHOUT reordering or overwriting
    what is already there (missing columns are appended on the right).
    Returns the header row as it now stands.
    """
    current = _sheets_api_call(lambda: ws.row_values(1))
    if not any((h or "").strip() for h in current):
        _ws_update(ws, "A1", [list(INVENTORY_COLUMNS)])
        return list(INVENTORY_COLUMNS)
    have = {(h or "").strip() for h in current}
    missing = [h for h in INVENTORY_COLUMNS if h not in have]
    if not missing:
        return list(current)
    new_header = list(current) + missing
    try:
        if getattr(ws, "col_count", 0) < len(new_header):
            _sheets_api_call(lambda: ws.resize(cols=len(new_header)))
    except Exception:
        logger.warning("Could not widen worksheet grid", exc_info=True)
    _ws_update(ws, "A1", [new_header])
    return new_header


def _header_index(header: List[str]) -> Dict[str, int]:
    idx: Dict[str, int] = {}
    for i, h in enumerate(header):
        idx.setdefault((h or "").strip(), i)   # first occurrence wins
    return idx


def _cell(row: List[str], idx: Dict[str, int], name: str) -> str:
    i = idx.get(name)
    return str(row[i]).strip() if i is not None and i < len(row) else ""


def _sheet_frame() -> pd.DataFrame:
    """Read the worksheet as strings, mapped by header NAME (never numericised)."""
    ws = _open_worksheet()
    try:
        header = _ensure_sheet_headers(ws)
    except Exception:
        logger.warning("Header check failed", exc_info=True)
        header = _sheets_api_call(lambda: ws.row_values(1))
    values = _sheets_api_call(lambda: ws.get_all_values())
    idx = _header_index(header)
    rows = []
    for raw in values[1:]:
        if not any(str(c).strip() for c in raw):
            continue
        rows.append({c: _cell(raw, idx, c) for c in INVENTORY_COLUMNS})
    df = pd.DataFrame(rows, columns=INVENTORY_COLUMNS)
    if not df.empty:
        df = df[df["Certificate Number"].astype(str).str.strip() != ""]
        df = df.drop_duplicates(subset="Certificate Number", keep="first")
    return df.reset_index(drop=True)


def _upsert_rows_gsheets(rows: List[Dict[str, str]]) -> Dict[str, int]:
    """Merge-upsert using one batch_update + one append_rows."""
    ws = _open_worksheet()
    header = _ensure_sheet_headers(ws)
    existing = _sheets_api_call(lambda: ws.get_all_values())
    idx = _header_index(header)

    index_by_cert: Dict[str, int] = {}
    for i, raw in enumerate(existing[1:], start=2):
        c = _cell(raw, idx, "Certificate Number")
        if c:
            index_by_cert.setdefault(c, i)

    batch_data, to_append = [], []
    counts = {"inserted": 0, "updated": 0, "unchanged": 0}

    for new in rows:
        cert = new["Certificate Number"]
        ridx = index_by_cert.get(cert)
        if ridx is not None:
            raw = existing[ridx - 1]
            old = {c: _cell(raw, idx, c) for c in INVENTORY_COLUMNS}
            merged = merge_row(new, old)
            if merged is None:
                counts["unchanged"] += 1
                continue
            values = list(raw) + [""] * max(0, len(header) - len(raw))   # keep unknown columns intact
            for c in INVENTORY_COLUMNS:
                if c in idx:
                    values[idx[c]] = merged[c]
            batch_data.append({"range": f"A{ridx}", "values": [values]})
            counts["updated"] += 1
        else:
            merged = merge_row(new, None)
            values = [""] * len(header)
            for c in INVENTORY_COLUMNS:
                if c in idx:
                    values[idx[c]] = merged[c]
            to_append.append(values)
            counts["inserted"] += 1

    if batch_data:
        _sheets_api_call(lambda: ws.batch_update(batch_data, value_input_option="RAW"))
    if to_append:
        _sheets_api_call(lambda: ws.append_rows(to_append, value_input_option="RAW", table_range="A1"))
    return counts


def _upsert_rows_sqlite(rows: List[Dict[str, str]]) -> Dict[str, int]:
    counts = {"inserted": 0, "updated": 0, "unchanged": 0}
    col_list = ", ".join(DB_COLUMNS)
    placeholders = ", ".join(f":{c}" for c in DB_COLUMNS)
    with db() as conn:
        for new in rows:
            cert = new["Certificate Number"]
            r = conn.execute("SELECT * FROM inventory WHERE certificate_number = ?", (cert,)).fetchone()
            old = None
            if r is not None:
                old = {c: r[DB_COLUMN_MAP[c]] for c in INVENTORY_COLUMNS}
            merged = merge_row(new, old)
            if merged is None:
                counts["unchanged"] += 1
                continue
            conn.execute(
                f"INSERT OR REPLACE INTO inventory ({col_list}) VALUES ({placeholders})",
                {DB_COLUMN_MAP[c]: merged[c] for c in INVENTORY_COLUMNS},
            )
            counts["updated" if old is not None else "inserted"] += 1
    return counts


def upsert_rows(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Insert or merge rows. Returns a result dict:
      inserted / updated / unchanged / skipped (no valid certificate number),
      warning (str | None)  – set when data was NOT written to the primary store,
      local_only (bool)     – True when only the local SQLite copy was written.
    """
    result: Dict[str, Any] = {
        "inserted": 0, "updated": 0, "unchanged": 0, "skipped": 0,
        "warning": None, "local_only": False,
    }
    valid, result["skipped"] = _collapse_batch(rows)
    if not valid:
        return result

    if not use_gsheets():
        result.update(_upsert_rows_sqlite(valid))
        clear_inventory_cache()
        return result

    try:
        result.update(_upsert_rows_gsheets(valid))
        clear_inventory_cache()
        return result
    except Exception as e:
        detail = _gspread_error_detail(e)
        logger.exception("Google Sheets upsert failed")
        sqlite_ok = False
        try:
            _upsert_rows_sqlite(valid)
            sqlite_ok = True
        except Exception:
            logger.exception("SQLite fallback also failed")
        hint = " A local copy was kept, but it will NOT sync back to the sheet automatically." if sqlite_ok else ""
        low = detail.lower()
        if "429" in detail or "quota" in low:
            msg = "Google Sheets write quota exceeded. Wait 1–2 minutes, then save a smaller batch." + hint
            if sqlite_ok:
                result.update(warning=msg, local_only=True, inserted=0, updated=0)
                return result
            raise RuntimeError(msg) from e
        if "403" in detail or "permission" in low:
            raise RuntimeError(
                "Google Sheets permission error. Share the spreadsheet with the service account as Editor." + hint
            ) from e
        if "404" in detail or "not found" in low:
            raise RuntimeError(
                "Google Sheet not found. Check spreadsheet_id / worksheet_name in secrets." + hint
            ) from e
        raise RuntimeError(f"Google Sheets save failed: {detail}" + hint) from e


def report_upsert(res: Dict[str, Any]) -> None:
    """Render an honest status line for an upsert result."""
    if res.get("warning"):
        st.warning(res["warning"])
        return
    n_in, n_up = res.get("inserted", 0), res.get("updated", 0)
    parts = []
    if n_in:
        parts.append(f"{n_in} added")
    if n_up:
        parts.append(f"{n_up} updated")
    if res.get("unchanged"):
        parts.append(f"{res['unchanged']} unchanged")
    if parts:
        st.success("Saved: " + ", ".join(parts) + ".")
    if res.get("skipped"):
        st.warning(
            f"{res['skipped']} row(s) not saved because no certificate number could be determined."
        )


# ---------------- Reads ----------------
@st.cache_data(ttl=90, show_spinner=False)
def _load_gsheets_inventory() -> pd.DataFrame:
    """Cached read from Google Sheets (90s) to stay under API quota."""
    return _sheet_frame()


@st.cache_data(ttl=30, show_spinner=False)
def _load_sqlite_inventory() -> pd.DataFrame:
    with db() as conn:
        raw = pd.read_sql_query("SELECT * FROM inventory ORDER BY date_added DESC", conn)
    inv = {v: k for k, v in DB_COLUMN_MAP.items()}
    df = raw.rename(columns=inv)
    for c in INVENTORY_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    return df[INVENTORY_COLUMNS].fillna("").astype(str)


def _normalize_lab_column(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure Lab is filled from certificate number when blank."""
    if df is None or df.empty:
        return df
    df = df.copy()
    df["Lab"] = [infer_lab(cert, lab) for cert, lab in zip(df["Certificate Number"], df["Lab"])]
    return df


def load_all_inventory() -> pd.DataFrame:
    if use_gsheets():
        try:
            return _normalize_lab_column(_load_gsheets_inventory())
        except Exception as e:
            logger.exception("Inventory load failed")
            msg = str(e)
            if "429" in msg or "Quota" in msg:
                st.warning("Data service is temporarily busy. Please try again shortly.")
            else:
                st.error("Unable to load inventory data.")
            return pd.DataFrame(columns=INVENTORY_COLUMNS)
    df = _load_sqlite_inventory()
    return _normalize_lab_column(df) if not df.empty else pd.DataFrame(columns=INVENTORY_COLUMNS)


def clear_inventory_cache():
    """Call after any write so the next read is fresh."""
    for fn in (_load_gsheets_inventory, _load_sqlite_inventory):
        try:
            fn.clear()
        except Exception:
            logger.debug("cache clear failed", exc_info=True)


def count_added_today(df: Optional[pd.DataFrame] = None) -> int:
    df = load_all_inventory() if df is None else df
    if df.empty or "date_added" not in df.columns:
        return 0
    return int(df["date_added"].astype(str).str.startswith(date.today().isoformat()).sum())


# ---------------- Edit tab: diff-based save with conflict detection ----------------
Changes = Dict[str, Dict[str, Tuple[str, str]]]   # cert -> {column: (old, new)}


def compute_edit_changes(original: pd.DataFrame, edited: pd.DataFrame) -> Changes:
    """Cells whose editable value differs between the table as loaded and as edited."""
    if original.empty or edited.empty:
        return {}
    orig = {
        _s(r["Certificate Number"]): r
        for r in original.to_dict("records") if _s(r.get("Certificate Number"))
    }
    changes: Changes = {}
    for r in edited.to_dict("records"):
        cert = _s(r.get("Certificate Number"))
        o = orig.get(cert)
        if not cert or o is None:
            continue
        for col in EDITABLE_COLUMNS:
            if col not in r:
                continue
            old, new = _s(o.get(col)), _s(r.get(col))
            if old != new:
                changes.setdefault(cert, {})[col] = (old, new)
    return changes


def update_editable_fields(changes: Changes) -> Dict[str, Any]:
    """
    Write only the changed editable cells. A cell is skipped as a CONFLICT when
    someone else changed it since this table was loaded (current value is
    neither the old nor the new value).
    Returns {"updated": n_cells, "conflicts": [(cert, column), ...]}.
    """
    out: Dict[str, Any] = {"updated": 0, "conflicts": []}
    if not changes:
        return out
    now = _now()

    if use_gsheets():
        import gspread
        ws = _open_worksheet()
        header = _ensure_sheet_headers(ws)
        existing = _sheets_api_call(lambda: ws.get_all_values())
        idx = _header_index(header)
        rows_by_cert: Dict[str, int] = {}
        for i, raw in enumerate(existing[1:], start=2):
            c = _cell(raw, idx, "Certificate Number")
            if c:
                rows_by_cert.setdefault(c, i)
        cells, touched = [], set()
        for cert, cols in changes.items():
            ridx = rows_by_cert.get(cert)
            if ridx is None:
                out["conflicts"].append((cert, "(row no longer exists)"))
                continue
            raw = existing[ridx - 1]
            for col, (old, new) in cols.items():
                if col not in idx:
                    continue
                current = _cell(raw, idx, col)
                if current != old and current != new:
                    out["conflicts"].append((cert, col))
                    continue
                cells.append(gspread.Cell(ridx, idx[col] + 1, new))
                touched.add(ridx)
                out["updated"] += 1
        for ridx in touched:
            cells.append(gspread.Cell(ridx, idx["last_updated"] + 1, now))
        if cells:
            _sheets_api_call(lambda: ws.update_cells(cells, value_input_option="RAW"))
    else:
        with db() as conn:
            for cert, cols in changes.items():
                r = conn.execute("SELECT * FROM inventory WHERE certificate_number = ?", (cert,)).fetchone()
                if r is None:
                    out["conflicts"].append((cert, "(row no longer exists)"))
                    continue
                sets, params = [], {}
                for col, (old, new) in cols.items():
                    dbc = DB_COLUMN_MAP[col]            # whitelisted via EDITABLE_COLUMNS
                    current = _s(r[dbc])
                    if current != old and current != new:
                        out["conflicts"].append((cert, col))
                        continue
                    sets.append(f"{dbc} = :{dbc}")
                    params[dbc] = new
                    out["updated"] += 1
                if sets:
                    params["_cert"], params["_now"] = cert, now
                    conn.execute(
                        f"UPDATE inventory SET {', '.join(sets)}, last_updated = :_now "
                        f"WHERE certificate_number = :_cert", params,
                    )
    clear_inventory_cache()
    return out


# -------------------------------------------------
# Inventory helpers (unchanged)
# -------------------------------------------------
def _parse_carat(val) -> float:
    try:
        m = re.search(r"[\d.]+", str(val))
        return float(m.group()) if m else 0.0
    except Exception:
        return 0.0


def _parse_money(val) -> float:
    try:
        s = re.sub(r"[^\d.]", "", str(val))
        return float(s) if s else 0.0
    except Exception:
        return 0.0


def inventory_subtotals(df: pd.DataFrame) -> Dict[str, Any]:
    """Compute summary stats for a filtered inventory dataframe."""
    if df is None or df.empty:
        return {
            "count": 0,
            "total_carat": 0.0,
            "avg_carat": 0.0,
            "total_price": 0.0,
            "by_shape": {},
            "by_color": {},
            "by_clarity": {},
            "with_video": 0,
        }
    carats = df["CARAT WEIGHT"].apply(_parse_carat) if "CARAT WEIGHT" in df.columns else pd.Series([0.0] * len(df))
    prices = df["Price For stone"].apply(_parse_money) if "Price For stone" in df.columns else pd.Series([0.0] * len(df))
    with_video = 0
    if "Video URL" in df.columns:
        with_video = int((df["Video URL"].fillna("").astype(str).str.strip() != "").sum())

    by_shape, by_color, by_clarity = {}, {}, {}
    if "SHAPE AND CUT" in df.columns:
        by_shape = df["SHAPE AND CUT"].fillna("").astype(str).str.strip().replace("", "Unknown").value_counts().head(8).to_dict()
    if "COLOR GRADE" in df.columns:
        by_color = df["COLOR GRADE"].fillna("").astype(str).str.strip().replace("", "Unknown").value_counts().to_dict()
    if "CLARITY GRADE" in df.columns:
        by_clarity = df["CLARITY GRADE"].fillna("").astype(str).str.strip().replace("", "Unknown").value_counts().to_dict()

    total_ct = float(carats.sum())
    return {
        "count": len(df),
        "total_carat": total_ct,
        "avg_carat": float(carats.mean()) if len(df) else 0.0,
        "total_price": float(prices.sum()),
        "by_shape": by_shape,
        "by_color": by_color,
        "by_clarity": by_clarity,
        "with_video": with_video,
    }


# -------------------------------------------------
# Certificate download / ZIP / preview
# -------------------------------------------------
def resolve_certificate_pdf_url(cert: str, cert_link: str = "") -> Optional[str]:
    """Build best-effort public PDF URL for a certificate."""
    cert = str(cert or "").strip()
    link = str(cert_link or "").strip()
    if link and (".pdf" in link.lower() or "pdf.igi.org" in link.lower() or "pdf.gia.edu" in link.lower()):
        return link
    if is_igi(cert):
        number_only = re.sub(r"^LG", "", cert, flags=re.I)
        return f"https://pdf.igi.org/FDR{number_only}.pdf"
    if link.startswith("http"):
        return link
    return None


def download_certificate_pdf(cert: str, cert_link: str = "") -> Dict[str, Any]:
    """Download certificate PDF bytes. Handles IGI FDR URLs and GIA token pages.
    Only hosts in CERT_ALLOWED_HOST_SUFFIXES are contacted (including redirects)."""
    url = resolve_certificate_pdf_url(cert, cert_link)
    if not url:
        return {"success": False, "error": "No PDF URL available for this certificate"}
    hdrs = {"Accept": "application/pdf,*/*"}
    try:
        resp = safe_get(url, allowed_host_suffixes=CERT_ALLOWED_HOST_SUFFIXES,
                        stream=True, timeout=40, headers=hdrs)
        if resp.status_code == 404:
            resp.close()
            return {"success": False, "error": "PDF not found (404) on lab server"}
        if resp.status_code != 200:
            code = resp.status_code
            resp.close()
            return {"success": False, "error": f"HTTP {code}"}

        ct = (resp.headers.get("Content-Type") or "").lower()
        content = read_capped(resp, MAX_PDF_BYTES)

        # GIA token pages return HTML that redirects to an S3 PDF
        if "pdf" not in ct and content[:4] != b"%PDF":
            html = content.decode("utf-8", errors="ignore")
            m = re.search(r'https://[^"\']+\.pdf[^"\']*', html)
            if not m:
                m = re.search(r'window\.location\s*=\s*\(\s*["\'](https://[^"\']+)["\']', html)
            if not m:
                return {"success": False, "error": "Response is not a PDF"}
            pdf_url = m.group(1) if m.lastindex else m.group(0)
            r2 = safe_get(pdf_url, allowed_host_suffixes=CERT_ALLOWED_HOST_SUFFIXES,
                          stream=True, timeout=40, headers=hdrs)
            if r2.status_code != 200:
                r2.close()
                return {"success": False, "error": "Could not download redirected PDF"}
            content = read_capped(r2, MAX_PDF_BYTES)

        if content[:4] != b"%PDF":
            return {"success": False, "error": "Downloaded file is not a valid PDF"}
        return {"success": True, "pdf_bytes": content, "source_url": url}
    except UnsafeURL as e:
        return {"success": False, "error": f"Blocked: {e}"}
    except Exception as e:
        logger.warning("Certificate download failed for %s: %s", cert, e)
        return {"success": False, "error": str(e)[:200]}


@st.cache_data(ttl=900, show_spinner=False, max_entries=64)
def _download_pdf_cached(cert: str, link: str) -> bytes:
    dl = download_certificate_pdf(cert, link)
    if not dl.get("success"):
        raise RuntimeError(dl.get("error", "download failed"))   # failures are not cached
    return dl["pdf_bytes"]


def get_certificate_pdf(cert: str, cert_link: str = "") -> Dict[str, Any]:
    """Cached wrapper (15 min) used by the preview and Drive backup."""
    try:
        return {"success": True, "pdf_bytes": _download_pdf_cached(str(cert), str(cert_link or ""))}
    except RuntimeError as e:
        return {"success": False, "error": str(e)}


def build_certificates_zip(
    df: pd.DataFrame,
    max_files: int = 80,
    workers: int = 5,
) -> Tuple[Optional[bytes], List[Dict[str, str]]]:
    """
    Download certificate PDFs for rows in df (in parallel) and pack into a ZIP.
    Returns (zip_bytes or None, status log rows).
    """
    if df is None or df.empty:
        return None, []

    jobs = []
    for _, row in df.head(max_files).iterrows():
        cert = _s(row.get("Certificate Number"))
        if cert:
            jobs.append((cert, _s(row.get("CERTIFICATE LINK"))))

    with ThreadPoolExecutor(max_workers=workers) as pool:
        downloads = list(pool.map(lambda j: download_certificate_pdf(*j), jobs))

    logs: List[Dict[str, str]] = []
    buf = io.BytesIO()
    ok_count = 0
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for (cert, _link), dl in zip(jobs, downloads):
            if dl.get("success") and dl.get("pdf_bytes"):
                zf.writestr(f"{_safe_name(cert)}_certificate.pdf", dl["pdf_bytes"])
                ok_count += 1
                logs.append({"Certificate": cert, "Status": "OK"})
            else:
                logs.append({
                    "Certificate": cert, "Status": "Failed",
                    "Note": str(dl.get("error") or "download failed")[:120],
                })
    if ok_count == 0:
        return None, logs
    return buf.getvalue(), logs


def pdf_to_preview_images(pdf_bytes: bytes, max_pages: int = 2, scale: float = 1.5) -> List[Any]:
    """Render PDF pages to PIL images for in-app preview."""
    images = []
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(pdf_bytes)
        for i in range(min(len(doc), max_pages)):
            images.append(doc[i].render(scale=scale).to_pil())
        return images
    except Exception:
        logger.debug("pypdfium2 render failed", exc_info=True)
    # Fallback: text-only card via pdfplumber
    try:
        from PIL import Image, ImageDraw
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages[:max_pages]:
                txt = (page.extract_text() or "")[:1200]
                img = Image.new("RGB", (900, 1200), "white")
                draw = ImageDraw.Draw(img)
                y = 20
                for line in txt.splitlines()[:50]:
                    draw.text((20, y), line[:90], fill="black")
                    y += 18
                images.append(img)
        return images
    except Exception:
        logger.debug("pdfplumber fallback render failed", exc_info=True)
        return []


# -------------------------------------------------
# Google Drive
# -------------------------------------------------
def upload_fileobj_to_drive(filename: str, fileobj, mime: str = "application/pdf") -> Dict[str, Any]:
    """Upload a seekable file object to the configured Drive folder (private by default)."""
    if not _has_gdrive_folder():
        return {"success": False, "error": "Drive folder_id not configured in secrets"}
    try:
        from googleapiclient.http import MediaIoBaseUpload
        folder_id = str(st.secrets["drive"]["folder_id"]).strip()
        if not folder_id:
            return {"success": False, "error": "drive.folder_id is empty in secrets"}
        service = _get_drive_service()
        media = MediaIoBaseUpload(fileobj, mimetype=mime, resumable=True)
        f = (
            service.files()
            .create(body={"name": filename, "parents": [folder_id]}, media_body=media,
                    fields="id,webViewLink", supportsAllDrives=True)
            .execute()
        )
        if _drive_public_link_enabled():
            try:
                service.permissions().create(
                    fileId=f["id"], body={"type": "anyone", "role": "reader"}, supportsAllDrives=True,
                ).execute()
            except Exception:
                logger.warning("Could not set public link permission", exc_info=True)
        link = f.get("webViewLink") or f"https://drive.google.com/file/d/{f['id']}/view"
        return {"success": True, "drive_link": link, "file_id": f["id"], "filename": filename}
    except Exception as e:
        logger.exception("Drive upload failed")
        err = str(e)
        if "File not found" in err or "404" in err:
            err += " — Check folder_id and that the folder is shared with the service account as Editor."
        if "googleapiclient" in err or "No module" in err:
            err += " — Add google-api-python-client to requirements.txt and reboot the app."
        return {"success": False, "error": err[:400]}


def upload_bytes_to_drive(filename: str, data: bytes, mime: str = "application/pdf") -> Dict[str, Any]:
    return upload_fileobj_to_drive(filename, io.BytesIO(data), mime)


def backup_certificate_pdf_to_drive(cert: str, cert_link: str = "") -> Dict[str, Any]:
    """Download lab PDF and store a permanent copy on Google Drive."""
    try:
        dl = get_certificate_pdf(cert, cert_link)
        if not dl.get("success"):
            return {"success": False, "error": f"Download failed: {dl.get('error', 'unknown')}"}
        filename = f"{_safe_name(cert)}_certificate.pdf"
        return upload_bytes_to_drive(filename, dl["pdf_bytes"], "application/pdf")
    except Exception as e:
        logger.exception("Backup failed")
        return {"success": False, "error": f"Backup error: {str(e)[:300]}"}


# -------------------------------------------------
# Video links
# -------------------------------------------------
def _read_head(resp: requests.Response, n: int) -> bytes:
    buf = bytearray()
    for chunk in resp.iter_content(chunk_size=65536):
        buf.extend(chunk)
        if len(buf) >= n:
            break
    return bytes(buf[:n])


def probe_video_url(url: str) -> Dict[str, Any]:
    """Check (without downloading the body) whether a URL is a direct video, or
    find an mp4 inside an HTML page."""
    try:
        r = safe_get(url, stream=True, timeout=20)
        try:
            ct = (r.headers.get("Content-Type") or "").lower()
            size = r.headers.get("Content-Length")
            final = r.url
            ok = r.status_code == 200
            is_video = ("video/" in ct) or final.lower().endswith((".mp4", ".webm", ".mov"))
            found_mp4 = None
            if ok and "html" in ct:
                text = _read_head(r, 1_000_000).decode("utf-8", errors="ignore")
                mp4s = re.findall(r'https?://[^\s"\'<>]+\.mp4[^\s"\'<>]*', text, flags=re.I)
                if mp4s:
                    found_mp4 = mp4s[0].replace("&amp;", "&")
                    is_video = True
                    final = found_mp4
        finally:
            r.close()
        return {
            "ok": ok, "status": r.status_code, "content_type": ct,
            "size": int(size) if size and str(size).isdigit() else None,
            "is_direct_video": bool(is_video and ok),
            "final_url": final, "extracted_mp4": found_mp4,
        }
    except Exception as e:
        return {
            "ok": False, "status": 0, "content_type": "", "size": None,
            "is_direct_video": False, "final_url": url, "error": str(e),
        }


def download_and_upload_video_to_drive(cert: str, url: str) -> Dict[str, Any]:
    """Stream a direct video file (size-capped, via a temp file) to the Drive folder."""
    if not _has_gdrive_folder():
        return {"success": False, "error": "Drive folder_id not configured in secrets"}

    probe = probe_video_url(url)
    if probe.get("error"):
        return {"success": False, "error": f"Could not reach URL: {probe['error'][:150]}"}
    if not probe.get("is_direct_video"):
        return {
            "success": False,
            "error": f"Not a direct video file (type={probe.get('content_type') or 'unknown'}). "
                     "Vision360/DNA HTML cannot be converted to MP4 automatically.",
            "link_saved_only": True,
        }
    if probe.get("size") and probe["size"] > MAX_VIDEO_BYTES:
        return {"success": False, "error": f"Video larger than {MAX_VIDEO_BYTES // (1024 * 1024)} MB limit"}

    import tempfile
    download_url = probe.get("extracted_mp4") or probe.get("final_url") or url
    try:
        r = safe_get(download_url, stream=True, timeout=60)
        if r.status_code != 200:
            code = r.status_code
            r.close()
            return {"success": False, "error": f"Download HTTP {code}"}
        ct = (r.headers.get("Content-Type") or "").lower()
        ext = ".webm" if "webm" in ct else (".mov" if ("quicktime" in ct or "mov" in ct) else ".mp4")
        total = 0
        with tempfile.SpooledTemporaryFile(max_size=32 * 1024 * 1024) as tmp:
            for chunk in r.iter_content(chunk_size=1024 * 512):
                total += len(chunk)
                if total > MAX_VIDEO_BYTES:
                    r.close()
                    return {"success": False, "error": f"Video exceeds {MAX_VIDEO_BYTES // (1024 * 1024)} MB limit"}
                tmp.write(chunk)
            tmp.seek(0)
            return upload_fileobj_to_drive(f"{cert}_video{ext}", tmp, ct if "video" in ct else "video/mp4")
    except UnsafeURL as e:
        return {"success": False, "error": f"Blocked: {e}"}
    except Exception as e:
        logger.exception("Video backup failed")
        return {"success": False, "error": str(e)[:200]}


def save_video_urls(pairs: List[tuple], try_drive_backup: bool = True) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    pairs: list of (certificate_number, video_url)
    Updates inventory Video URL (new URL replaces the old); optionally backs up
    direct videos to Drive. Returns (per-item results, upsert result).
    """
    results: List[Dict[str, Any]] = []
    existing = load_all_inventory()
    existing_map: Dict[str, Dict[str, Any]] = {}
    if not existing.empty:
        for r in existing.to_dict("records"):
            existing_map[_s(r["Certificate Number"])] = r

    to_upsert = []
    for cert, url in pairs:
        cert, url = str(cert).strip(), str(url).strip()
        if not _valid_cert(cert) or not url:
            continue
        if not url.lower().startswith(("http://", "https://")):
            results.append({"Certificate Number": cert, "Video URL": url, "backup": "rejected (not an http/https URL)"})
            continue
        row = dict(existing_map.get(cert) or empty_inventory_row(cert, infer_lab(cert)))
        if cert not in existing_map:
            row["Status"] = "Media only"
        row["Certificate Number"] = cert
        row["Lab"] = infer_lab(cert, row.get("Lab") or "")
        row["Video URL"] = url
        info = {"Certificate Number": cert, "Video URL": url, "backup": "skipped"}

        if try_drive_backup and _has_gdrive_folder():
            up = download_and_upload_video_to_drive(cert, url)
            if up.get("success"):
                row["Video Backup Drive Link"] = up["drive_link"]
                info.update(backup="uploaded", **{"Video Backup Drive Link": up["drive_link"]})
            elif up.get("link_saved_only"):
                info.update(backup="link only (not direct video)", note=up.get("error", ""))
            else:
                info.update(backup="failed", note=up.get("error", ""))
        else:
            info["backup"] = "link saved (Drive not configured or disabled)"
        to_upsert.append(row)
        results.append(info)

    res = upsert_rows(to_upsert) if to_upsert else {"inserted": 0, "updated": 0, "unchanged": 0, "skipped": 0, "warning": None}
    return results, res


# -------------------------------------------------
# Certificate preview (dialog)
# -------------------------------------------------
def render_certificate_preview(cert: str, cert_link: str = ""):
    """Overlay: certificate image first, then compact specs and actions."""
    cert = str(cert or "").strip()
    if not cert:
        return
    existing = load_all_inventory()
    link = cert_link
    backup_link = ""
    match_row = None
    if not existing.empty:
        match = existing[existing["Certificate Number"].astype(str) == cert]
        if not match.empty:
            match_row = match.iloc[0].to_dict()
            link = link or str(match_row.get("CERTIFICATE LINK") or "")
            backup_link = str(match_row.get("Certificate PDF Backup Drive Link") or "").strip()

    def _v(key, default="—"):
        if not match_row:
            return default
        val = match_row.get(key)
        return str(val) if val not in (None, "") else default

    with st.spinner("Loading certificate…"):
        dl = get_certificate_pdf(cert, link)
    pdf_bytes = dl.get("pdf_bytes") if dl.get("success") else None
    if pdf_bytes:
        images = pdf_to_preview_images(pdf_bytes, max_pages=1, scale=1.6)
        if images:
            st.image(images[0], **_STRETCH)
        st.download_button(
            "Download PDF", data=pdf_bytes, file_name=f"{cert}_certificate.pdf",
            mime="application/pdf", key=f"prev_dl_{cert}",
        )
    else:
        st.warning(dl.get("error", "Certificate file is not available."))

    st.caption(
        f"{_v('Lab')} · {_v('SHAPE AND CUT')} · {_v('CARAT WEIGHT')} · "
        f"{_v('COLOR GRADE')} / {_v('CLARITY GRADE')} · {_v('MEASUREMENTS')} · {_v('Process')}"
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        do_backup = st.button("Save PDF copy", key=f"prev_bak_{cert}", disabled=not _has_gdrive_folder())
    with c2:
        if link:
            st.link_button("Open report", link, **_STRETCH)
    with c3:
        if backup_link:
            st.link_button("Saved copy", backup_link, **_STRETCH)

    if do_backup and _has_gdrive_folder():
        with st.spinner("Saving certificate…"):
            res = backup_certificate_pdf_to_drive(cert, link)
        if res.get("success"):
            row = dict(match_row or empty_inventory_row(cert, infer_lab(cert)))
            if not match_row:
                row["Status"] = "Link only"
            row["Certificate Number"] = cert
            row["Certificate PDF Backup Drive Link"] = res["drive_link"]
            if link:
                row["CERTIFICATE LINK"] = link
            report_upsert(upsert_rows([row]))
        else:
            st.error(res.get("error", "Unable to save certificate copy."))


@st.dialog("Certificate preview", width="large")
def certificate_preview_dialog(cert: str, cert_link: str = ""):
    render_certificate_preview(cert, cert_link)


def open_certificate_preview(cert: str, cert_link: str = ""):
    cert = str(cert or "").strip()
    if cert:
        certificate_preview_dialog(cert, cert_link)


def maybe_open_preview(state_key: str, sel_rows, frame: pd.DataFrame) -> None:
    """Open the preview dialog once per newly selected row (not on every rerun)."""
    if sel_rows and sel_rows[0] < len(frame):
        crow = frame.iloc[sel_rows[0]]
        cert = _s(crow.get("Certificate Number"))
        link = _s(crow.get("CERTIFICATE LINK"))
        if st.session_state.get(state_key) != cert:
            st.session_state[state_key] = cert
            open_certificate_preview(cert, link)
    else:
        st.session_state[state_key] = None


def render_inventory_subtotals(df: pd.DataFrame, title: str = "Subtotals", expanded: bool = False):
    """Compact summary in an expander so the table stays primary."""
    s = inventory_subtotals(df)
    price_txt = f"{s['total_price']:,.0f}" if s["total_price"] else "—"
    label = (
        f"Σ {title}:  {s['count']} pcs · {s['total_carat']:.2f} ct · "
        f"avg {s['avg_carat']:.2f} · price {price_txt} · video {s['with_video']}"
    )
    with st.expander(label, expanded=expanded):
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Pieces", s["count"])
        c2.metric("Total ct", f"{s['total_carat']:.2f}")
        c3.metric("Avg ct", f"{s['avg_carat']:.2f}")
        c4.metric("Price Σ", price_txt)
        c5.metric("Video", s["with_video"])
        if s["count"] == 0:
            return
        b1, b2, b3 = st.columns(3)
        with b1:
            if s["by_shape"]:
                st.dataframe(
                    pd.DataFrame({"Shape": list(s["by_shape"].keys()), "#": list(s["by_shape"].values())}),
                    hide_index=True, **_STRETCH, height=160,
                )
        with b2:
            if s["by_color"]:
                st.dataframe(
                    pd.DataFrame({"Color": list(s["by_color"].keys()), "#": list(s["by_color"].values())}),
                    hide_index=True, **_STRETCH, height=160,
                )
        with b3:
            if s["by_clarity"]:
                st.dataframe(
                    pd.DataFrame({"Clarity": list(s["by_clarity"].keys()), "#": list(s["by_clarity"].values())}),
                    hide_index=True, **_STRETCH, height=160,
                )


init_db()

# -------------------------------------------------
# Helper functions (unchanged)
# -------------------------------------------------

def clean_report_number(raw: str) -> str:
    """
    Clean and normalize report number.
    Handles formats like:
      LG833638417
      lg 833638417
      GIA# 1553362669
      GIA 6555540885
      1553362669
    """
    cleaned = raw.strip().upper()
    # Remove common prefixes and symbols (GIA#, GIA , IGI, etc.)
    cleaned = re.sub(r'^(GIA|IGI)[\s#:.\-]*', '', cleaned)
    cleaned = re.sub(r'[^A-Za-z0-9]', '', cleaned)
    return cleaned


def is_gia_token_url(text: str) -> bool:
    """Detect special GIA PDF token links."""
    text = text.strip()
    return (
        "pdf.gia.edu" in text.lower()
        or "report-check-objects" in text.lower()
        or (text.startswith("http") and "ReportNumber=" in text)
    )


def is_igi(report_no: str) -> bool:
    """Detect IGI report (starts with LG)."""
    return str(report_no).strip().upper().startswith("LG")


def is_gia(report_no: str) -> bool:
    """
    Detect GIA-style report.
    Accepts pure numeric strings of reasonable length (typical GIA reports are 9–12 digits).
    """
    s = str(report_no).strip()
    return s.isdigit() and 8 <= len(s) <= 13


def infer_lab(cert: str, current_lab: str = "") -> str:
    """Fill blank/unknown Lab from certificate number pattern."""
    lab = (current_lab or "").strip()
    if lab in ("IGI", "GIA"):
        return lab
    c = str(cert or "").strip()
    if is_igi(c):
        return "IGI"
    if is_gia(c):
        return "GIA"
    return lab or "Unknown"



def extract_text_from_pdf(content: bytes) -> str:
    """
    Robust multi-method text extraction from IGI PDF bytes.
    Tries pdfplumber first, then falls back to other strategies.
    """
    text = ""

    # Method 1: pdfplumber (best quality)
    if HAS_PDFPLUMBER:
        try:
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text() or ""
                    text += page_text + "\n"
            if len(text.strip()) > 200:
                return text
        except Exception:
            pass

    # Method 2: Try pypdf if available
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(content))
        for page in reader.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
        if len(text.strip()) > 200:
            return text
    except Exception:
        pass

    # Method 3: Last-resort raw decode (captures a lot of the text stream)
    try:
        raw = content.decode("latin-1", errors="ignore")
        # Keep only printable-ish characters and normalize whitespace
        raw = re.sub(r'[^\x20-\x7E\n\r\t]', ' ', raw)
        text = re.sub(r'[ \t]+', ' ', raw)
        return text
    except Exception:
        return text


def empty_inventory_row(cert: str = "", lab: str = "") -> Dict[str, Any]:
    """Full inventory schema matching the Excel template."""
    return {
        "Certificate Number": cert,
        "DESCRIPTION": "LABORATORY GROWN DIAMOND" if lab == "IGI" else ("NATURAL DIAMOND" if lab == "GIA" else ""),
        "SHAPE AND CUT": None,
        "CARAT WEIGHT": None,
        "COLOR GRADE": None,
        "CLARITY GRADE": None,
        "MEASUREMENTS": None,
        "CUT": None,
        "POLISH": None,
        "SYMMETRY": None,
        "FLUORESCENCE": None,
        "Process": None,
        "FGI Item Number": "",
        "Date Shipped from India": "",
        "For stock or Customer": "",
        "Production Order #, if for stock": "",
        "Current Location/Status": "",
        "Price Per Carat": "",
        "Price For stone": "",
        "CERTIFICATE LINK": "",
        "Comment": "",
        "Video URL": "",
        "Video Backup Drive Link": "",
        "Certificate PDF Backup Drive Link": "",
        "Lab": lab,
        "Status": "Error",
        "Notes": "",
    }


# -------------------------------------------------
# PDF parsing (single shared path for IGI / GIA / uploaded PDFs)
# -------------------------------------------------
_SEP = r"[\s.:]*"     # GIA pads labels with dot leaders: "Carat Weight ....... 0.50 carat"
_CLARITY = r"((?:FL|IF|VVS\s*[12]|VS\s*[12]|SI\s*[12]|I\s*[123]))"


def normalize_pdf_text(text: str) -> str:
    """Upper-case and normalise punctuation so one set of regexes fits all labs."""
    t = (text or "").upper()
    t = t.replace("×", "X").replace("–", "-").replace("—", "-")
    t = re.sub(r"(?<=[A-Z])-(?=[A-Z])", " ", t)      # CUT-CORNERED -> CUT CORNERED
    return t


def detect_lab_from_pdf_text(text_upper: str) -> str:
    """Decide the lab by report-number shape first, then by whichever name appears first."""
    if re.search(r"\bLG\d{8,12}\b", text_upper):
        return "IGI"
    igi = re.search(r"INTERNATIONAL GEMOLOGICAL INSTITUTE|\bIGI\b", text_upper)
    gia = re.search(r"GEMOLOGICAL INSTITUTE OF AMERICA|\bGIA\b", text_upper)
    if igi and gia:
        return "IGI" if igi.start() < gia.start() else "GIA"
    if igi:
        return "IGI"
    if gia:
        return "GIA"
    return "Unknown"


def extract_report_number_from_pdf_text(text_upper: str, lab: str = "") -> str:
    m = re.search(r"\b(LG\d{8,12})\b", text_upper)
    if m:
        return m.group(1)
    m = re.search(rf"(?:IGI\s+)?REPORT\s+NUMBER{_SEP}(LG?\d{{8,12}})", text_upper)
    if m:
        val = m.group(1)
        if not val.startswith("LG") and lab == "IGI":
            return f"LG{val}" if val.isdigit() else val
        return val
    m = re.search(rf"GIA\s+REPORT\s+(?:NUMBER|#)?{_SEP}(\d{{8,13}})", text_upper)
    if m:
        return m.group(1)
    m = re.search(rf"REPORT\s+NUMBER{_SEP}(\d{{8,13}})", text_upper)
    if m:
        return m.group(1)
    return ""


def apply_grading_fields_from_text(result: Dict[str, Any], text: str) -> Dict[str, Any]:
    """Fill shape, measurements, carat, color, clarity, cut quality, process from PDF text."""
    text_upper = normalize_pdf_text(text)

    shape = parse_shape_and_cutting_style(text_upper)
    if shape:
        result["SHAPE AND CUT"] = shape

    for pat in [
        rf"MEASUREMENTS?{_SEP}([\d\.\-\sX]+?\s*MM)",
        r"([\d\.]+\s*-\s*[\d\.]+\s*X\s*[\d\.]+\s*MM)",
        r"([\d\.]+\s*X\s*[\d\.]+\s*X\s*[\d\.]+\s*MM)",
    ]:
        m = re.search(pat, text_upper)
        if m:
            size = re.sub(r"\s+", " ", m.group(1).strip().replace("X", "×"))
            size = size.replace("MM", "mm") if "MM" in size else size + " mm"
            result["MEASUREMENTS"] = size
            break

    for pat in [
        rf"CARAT\s+WEIGHT{_SEP}(\d+\.\d{{1,3}})\s*CARATS?",     # labelled value first
        r"(\d+\.\d{2})\s*CARATS?",
        r"(\d+\.\d{2})\s*CT\b",
    ]:
        m = re.search(pat, text_upper)
        if m:
            result["CARAT WEIGHT"] = f"{m.group(1)} ct"
            break

    color = parse_color_grade(text_upper)
    if color:
        result["COLOR GRADE"] = color

    for pat in [rf"\bCLARITY\s+GRADE{_SEP}{_CLARITY}", rf"\bCLARITY{_SEP}{_CLARITY}"]:
        m = re.search(pat, text_upper)
        if m:
            result["CLARITY GRADE"] = re.sub(r"\s+", "", m.group(1)).upper()
            break

    quality = r"(IDEAL|EXCELLENT|VERY GOOD|GOOD|FAIR|POOR)"
    for field, label in (("CUT", r"CUT\s+GRADE"), ("POLISH", r"POLISH"), ("SYMMETRY", r"SYMMETRY")):
        m = re.search(rf"\b{label}{_SEP}{quality}", text_upper)
        if m:
            result[field] = m.group(1).title()
    m = re.search(rf"\bFLUORESCENCE{_SEP}(NONE|FAINT|MEDIUM|STRONG|VERY STRONG)", text_upper)
    if m:
        result["FLUORESCENCE"] = m.group(1).title()

    if re.search(r"CHEMICAL\s+VAPOR\s+DEPOSITION|\bCVD\b", text_upper):
        result["Process"] = "CVD"
    elif re.search(r"HIGH\s+PRESSURE\s+HIGH\s+TEMPERATURE|\bHPHT\b", text_upper):
        result["Process"] = "HPHT"

    if "LABORATORY GROWN" in text_upper or "LAB GROWN" in text_upper:
        result["DESCRIPTION"] = "LABORATORY GROWN DIAMOND"
    elif "NATURAL" in text_upper and "DIAMOND" in text_upper:
        result["DESCRIPTION"] = "NATURAL DIAMOND"

    return result


def finalize_result(
    result: Dict[str, Any], source: str, success_at: int = 3,
    partial_at: int = 1, success_label: str = "Success",
) -> Dict[str, Any]:
    """Set Status / Notes from how many core fields were extracted."""
    filled = sum(1 for k in CORE_FIELDS if result.get(k))
    if filled >= success_at:
        result["Status"] = success_label
        result["Notes"] = f"{source} • {filled}/5 core fields"
    elif filled >= partial_at:
        result["Status"] = "Partial"
        result["Notes"] = f"{source} • partial {filled}/5 fields"
    else:
        result["Status"] = "Parse Error"
        result["Notes"] = f"{source} • key fields not found"
    return result


def parse_certificate_pdf_bytes(pdf_bytes: bytes, filename: str = "") -> Dict[str, Any]:
    """Parse an uploaded IGI/GIA certificate PDF into an inventory row."""
    if not pdf_bytes or pdf_bytes[:4] != b"%PDF":
        row = empty_inventory_row("", "Unknown")
        row.update(Status="Error", Notes="Not a valid PDF file")
        return row

    text = extract_text_from_pdf(pdf_bytes)
    if not text or len(text.strip()) < 40:
        row = empty_inventory_row("", "Unknown")
        row.update(Status="Parse Error", Notes="Could not extract text from PDF")
        return row

    text_upper = normalize_pdf_text(text)
    lab = detect_lab_from_pdf_text(text_upper)
    cert = extract_report_number_from_pdf_text(text_upper, lab)
    if not cert and filename:
        m = re.search(r"(LG\d{8,12}|\d{8,13})", filename.upper())
        if m:
            cert = m.group(1)

    lab = lab if lab != "Unknown" else infer_lab(cert or "")
    result = empty_inventory_row(cert or "Unknown", lab)
    result["Lab"] = lab
    if cert:
        if lab == "IGI" and cert.upper().startswith("LG"):
            result["CERTIFICATE LINK"] = f"https://pdf.igi.org/FDR{re.sub(r'^LG', '', cert, flags=re.I)}.pdf"
        elif lab == "GIA" and cert.isdigit():
            result["CERTIFICATE LINK"] = f"https://www.gia.edu/report-check?reportno={cert}"

    apply_grading_fields_from_text(result, text_upper)
    finalize_result(result, "Uploaded PDF")
    if not cert:
        result["Notes"] += " · report number not found (not saved)"
    return result


KNOWN_SHAPES = [
    "CUT CORNERED RECTANGULAR MODIFIED BRILLIANT",
    "CUT CORNERED SQUARE MODIFIED BRILLIANT",
    "SQUARE CUSHION MODIFIED BRILLIANT",
    "RECTANGULAR CUSHION MODIFIED BRILLIANT",
    "RECTANGULAR MODIFIED BRILLIANT",
    "SQUARE MODIFIED BRILLIANT",
    "CUSHION MODIFIED BRILLIANT",
    "PEAR MODIFIED BRILLIANT",
    "OVAL MODIFIED BRILLIANT",
    "HEART MODIFIED BRILLIANT",
    "MARQUISE MODIFIED BRILLIANT",
    "SQUARE EMERALD CUT",
    "ROUND BRILLIANT",
    "PRINCESS CUT",
    "CUSHION BRILLIANT",
    "OVAL BRILLIANT",
    "PEAR BRILLIANT",
    "MARQUISE BRILLIANT",
    "EMERALD CUT",
    "RADIANT CUT",
    "HEART BRILLIANT",
    "ASSCHER CUT",
    "TRIANGLE BRILLIANT",
    "TRILLIANT",
    "BAGUETTE",
    "TAPERED BAGUETTE",
]
_SHAPE_WORDS = {
    "CUT", "CORNERED", "RECTANGULAR", "SQUARE", "MODIFIED", "BRILLIANT",
    "ROUND", "PRINCESS", "CUSHION", "OVAL", "PEAR", "MARQUISE", "EMERALD",
    "RADIANT", "HEART", "ASSCHER", "TRIANGLE", "TRILLIANT", "BAGUETTE",
    "TAPERED", "STEP", "MIXED",
}


def parse_shape_and_cutting_style(text_upper: str) -> Optional[str]:
    """
    Extract the full shape name from lab PDF text. Searches the text that follows
    the "Shape and Cutting Style" label (multi-line safe, ignores % noise); the
    longest matching known phrase wins. Only when the label is absent does it
    fall back to scanning the whole document.
    """
    if not text_upper:
        return None

    m = re.search(
        r"SHAPE\s+AND\s+CUTTING\s+STYLE[\s.:]*(.{5,120}?)(?:"
        r"MEASUREMENTS?|GRADING\s+RESULTS|CARAT\s+WEIGHT|COLOR\s+GRADE|CLARITY\s+GRADE|"
        r"PROPORTIONS|ADDITIONAL\s+GRADING|$)",
        text_upper, flags=re.DOTALL,
    )
    if not m:
        hits = [p for p in KNOWN_SHAPES if p in text_upper]
        return max(hits, key=len).title() if hits else None

    blob = re.sub(r"\s+", " ", re.sub(r"[\n\r]+", " ", m.group(1))).strip()
    blob = re.sub(r"\s+", " ", re.sub(r"\b\d+\.?\d*\s*%?", " ", blob)).strip()

    hits = [p for p in KNOWN_SHAPES if p in blob]
    if hits:
        return max(hits, key=len).title()

    tokens: List[str] = []
    for tok in re.findall(r"[A-Z]+", blob):
        if tok in _SHAPE_WORDS:
            tokens.append(tok)
        elif tokens and tok in ("MM", "CARAT", "CARATS", "GRADING", "RESULTS"):
            break
        elif len(tokens) >= 2:
            break
    if len(tokens) >= 2:
        shape = " ".join(tokens)
        if 4 < len(shape) < 80:
            return shape.title()
    return None


# -------------------------------------------------
# Lab fetchers
# -------------------------------------------------
def fetch_igi_report(report_no: str) -> Dict[str, Any]:
    """Fetch a live IGI Laboratory Grown Diamond report from the official PDF."""
    number_only = re.sub(r"^LG", "", report_no, flags=re.IGNORECASE)
    cert = report_no if report_no.upper().startswith("LG") else f"LG{number_only}"
    pdf_url = f"https://pdf.igi.org/FDR{number_only}.pdf"

    result = empty_inventory_row(cert, "IGI")
    result["CERTIFICATE LINK"] = pdf_url
    try:
        resp = safe_get(pdf_url, allowed_host_suffixes=CERT_ALLOWED_HOST_SUFFIXES,
                        stream=True, timeout=25, headers={"Accept": "application/pdf,*/*"})
        if resp.status_code == 404:
            resp.close()
            result.update(
                Status="Not Found",
                Notes="Certificate PDF not available from laboratory at this time",
            )
            result["CERTIFICATE LINK"] = f"https://www.igi.org/Verify-Your-Report/?r={cert}"
            return result
        if resp.status_code != 200:
            result.update(Status=f"HTTP {resp.status_code}", Notes="Could not download PDF")
            resp.close()
            return result

        content = read_capped(resp, MAX_PDF_BYTES)
        text = extract_text_from_pdf(content) if content[:4] == b"%PDF" else ""
        if not text or len(text.strip()) < 100:
            result.update(Status="Parse Error", Notes="Could not extract readable text from PDF")
            return result

        apply_grading_fields_from_text(result, text)
        if not result.get("Process"):
            result["Process"] = "Unknown"
        finalize_result(result, "Fetched live")
    except UnsafeURL as e:
        result.update(Status="Error", Notes=f"Blocked: {e}")
    except Exception as e:
        logger.warning("IGI fetch failed for %s: %s", cert, e)
        result.update(Status="Error", Notes=str(e)[:150])
    return result


def fetch_gia_from_token(token_url: str) -> Dict[str, Any]:
    """
    Download a GIA report PDF using a special token link
    (https://pdf.gia.edu/?ReportNumber=HASH) and extract grading data.
    """
    result = empty_inventory_row("Unknown", "GIA")
    result["Process"] = "—"
    result["CERTIFICATE LINK"] = token_url
    hdrs = {"Accept": "application/pdf,text/html,*/*"}
    try:
        resp = safe_get(token_url, allowed_host_suffixes=CERT_ALLOWED_HOST_SUFFIXES,
                        stream=True, timeout=20, headers=hdrs)
        ct = (resp.headers.get("Content-Type") or "").lower()
        body = read_capped(resp, MAX_PDF_BYTES)

        pdf_url = None
        if body[:4] == b"%PDF" or "pdf" in ct:
            pdf_bytes = body
        else:
            html = body.decode("utf-8", errors="ignore")
            m = re.search(r'https://[^"\']+\.pdf[^"\']*', html) or \
                re.search(r'window\.location\s*=\s*\(\s*["\'](https://[^"\']+)["\']', html)
            if not m:
                result["Notes"] = "Could not find PDF URL in token page"
                return result
            pdf_url = m.group(1) if m.lastindex else m.group(0)
            r2 = safe_get(pdf_url, allowed_host_suffixes=CERT_ALLOWED_HOST_SUFFIXES,
                          stream=True, timeout=30, headers=hdrs)
            if r2.status_code != 200:
                result["Notes"] = f"PDF download failed (HTTP {r2.status_code})"
                r2.close()
                return result
            pdf_bytes = read_capped(r2, MAX_PDF_BYTES)

        text = extract_text_from_pdf(pdf_bytes)
        if not text or len(text.strip()) < 80:
            result.update(Status="Parse Error", Notes="Downloaded PDF but could not extract text")
            return result

        text_upper = normalize_pdf_text(text)
        num = re.search(r"/(\d{8,13})\.pdf", pdf_url or token_url)
        if num:
            result["Certificate Number"] = num.group(1)
        else:
            m = re.search(rf"GIA\s+REPORT\s+NUMBER{_SEP}(\d{{8,13}})", text_upper)
            if m:
                result["Certificate Number"] = m.group(1)

        apply_grading_fields_from_text(result, text_upper)
        if not result.get("Process"):
            result["Process"] = "—"
        finalize_result(result, "Extracted from GIA PDF", success_at=4, partial_at=2,
                        success_label="Success (Token PDF)")
        if not _valid_cert(result["Certificate Number"]):
            result["Notes"] += " · report number not found (not saved)"
    except UnsafeURL as e:
        result.update(Status="Error", Notes=f"Blocked: {e}")
    except Exception as e:
        logger.warning("GIA token fetch failed: %s", e)
        result.update(Status="Error", Notes=str(e)[:120])
    return result


def fetch_gia_info(report_no: str) -> Dict[str, Any]:
    """
    GIA lookup by number. GIA blocks scripted access to report data, so this
    returns a clean report-check link only; use a token link or PDF upload to
    capture grading fields.
    """
    result = empty_inventory_row(report_no, "GIA")
    result["Process"] = "—"
    result["DESCRIPTION"] = "NATURAL DIAMOND"
    result["CERTIFICATE LINK"] = f"https://www.gia.edu/report-check?reportno={report_no}"
    result["Status"] = "Link Ready"
    result["Notes"] = "Open laboratory report check link"
    return result


def parse_color_grade(text_upper: str) -> Optional[str]:
    """
    D–Z colorless scale OR fancy color descriptions, e.g.:
      FANCY INTENSE YELLOW, FANCY VIVID PINK, FANCY LIGHT BLUE
    """
    if not text_upper:
        return None

    # Fancy / colored grades (check before single-letter D–Z)
    fancy = re.search(
        r"COLOR\s+GRADE[\s.:]*"
        r"((?:FANCY\s+)?"
        r"(?:VIVID|INTENSE|DEEP|DARK|LIGHT|VERY\s+LIGHT|FAINT)?\s*"
        r"(?:YELLOW|ORANGE|PINK|BLUE|GREEN|PURPLE|VIOLET|BROWN|GRAY|GREY|"
        r"CHAMPAGNE|COGNAC|BLACK|WHITE|RED|GRAYISH|GREYISH|BROWNISH|GREENISH|"
        r"YELLOWISH|ORANGEY|PINKISH|BLUISH)?"
        r"(?:\s+(?:YELLOW|ORANGE|PINK|BLUE|GREEN|PURPLE|VIOLET|BROWN|GRAY|GREY|"
        r"CHAMPAGNE|COGNAC|BLACK|WHITE|RED))*"
        r")",
        text_upper,
    )
    if fancy:
        val = re.sub(r"\s+", " ", fancy.group(1)).strip()
        # Reject empty / noise
        if val and not re.fullmatch(r"[D-Z]", val):
            # Must contain FANCY or a color word
            if "FANCY" in val or re.search(
                r"YELLOW|ORANGE|PINK|BLUE|GREEN|PURPLE|BROWN|GRAY|GREY|BLACK|WHITE|RED|CHAMPAGNE|COGNAC",
                val,
            ):
                return val.title()

    # Explicit "FANCY …" anywhere near color section
    m2 = re.search(
        r"\b(FANCY\s+(?:VIVID|INTENSE|DEEP|DARK|LIGHT|VERY\s+LIGHT|FAINT)?\s*"
        r"(?:YELLOW|ORANGE|PINK|BLUE|GREEN|PURPLE|VIOLET|BROWN|GRAY|GREY|"
        r"CHAMPAGNE|COGNAC|BLACK|WHITE|RED|"
        r"(?:GRAYISH|GREYISH|BROWNISH|GREENISH|YELLOWISH|ORANGEY|PINKISH|BLUISH)\s+\w+)"
        r")\b",
        text_upper,
    )
    if m2:
        return re.sub(r"\s+", " ", m2.group(1)).strip().title()

    # Standard D–Z
    m = re.search(r"COLOR\s+GRADE[\s.:]*([D-Z])\b", text_upper)
    if m:
        return m.group(1)
    return None


# -------------------------------------------------
# Inventory search (moved to module level so it can be tested)
# -------------------------------------------------
def smart_filter(df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Filter with partial keywords, cert fragments, shape/color/clarity/carat."""
    if df.empty or not query or not query.strip():
        return df

    q = query.strip().upper()
    mask = pd.Series([True] * len(df), index=df.index)
    search_cols = [
        c for c in [
            "Certificate Number", "SHAPE AND CUT", "CARAT WEIGHT", "COLOR GRADE",
            "CLARITY GRADE", "Process", "FGI Item Number", "Current Location/Status",
            "For stock or Customer", "Comment", "DESCRIPTION", "Lab",
            "Production Order #, if for stock",
        ] if c in df.columns
    ]

    # --- Structured filters (AND) ---
    carat_m = re.search(r"(\d+\.?\d*)\s*(?:CT|CARAT|CTS|CARATS)\b", q)
    if carat_m and "CARAT WEIGHT" in df.columns:
        target = float(carat_m.group(1))
        is_whole = abs(target - round(target)) < 1e-9

        def _near_carat(val):
            try:
                num = float(re.search(r"[\d.]+", str(val).upper()).group())
                if is_whole:
                    return target <= num < target + 1.0
                return abs(num - target) <= CARAT_TOLERANCE
            except Exception:
                return False

        mask &= df["CARAT WEIGHT"].apply(_near_carat)

    shapes = [
        "ROUND", "OVAL", "PEAR", "MARQUISE", "PRINCESS", "CUSHION",
        "EMERALD", "RADIANT", "HEART", "ASSCHER", "RECTANGULAR", "SQUARE",
        "BRILLIANT", "CUT CORNERED",
    ]
    for shape in shapes:
        if re.search(rf"\b{shape}\b", q) and "SHAPE AND CUT" in df.columns:
            mask &= df["SHAPE AND CUT"].astype(str).str.upper().str.contains(shape, na=False)

    color_m2 = re.search(r"(?:COLOR|COLOUR)\s*([D-Z])\b|\b([D-Z])\s*(?:COLOR|COLOUR)\b", q)
    if color_m2 and "COLOR GRADE" in df.columns:
        letter = color_m2.group(1) or color_m2.group(2)
        mask &= df["COLOR GRADE"].astype(str).str.upper() == letter

    clar_m = re.search(
        r"\b(FL|IF|VVS\s*1|VVS\s*2|VVS1|VVS2|VS\s*1|VS\s*2|VS1|VS2|"
        r"SI\s*1|SI\s*2|SI1|SI2|I\s*1|I\s*2|I\s*3|I1|I2|I3)\b",
        q,
    )
    if clar_m and "CLARITY GRADE" in df.columns:
        clar = re.sub(r"\s+", "", clar_m.group(1)).upper()
        mask &= df["CLARITY GRADE"].astype(str).str.upper().str.replace(" ", "", regex=False) == clar

    if re.search(r"\bCVD\b", q) and "Process" in df.columns:
        mask &= df["Process"].astype(str).str.upper().str.contains("CVD", na=False)
    if re.search(r"\bHPHT\b", q) and "Process" in df.columns:
        mask &= df["Process"].astype(str).str.upper().str.contains("HPHT", na=False)

    # --- Partial keyword / cert fragment tokens (AND across tokens) ---
    tokens = re.findall(r"[A-Z0-9]{2,}", q)
    skip = {
        "CT", "CARAT", "CARATS", "CTS", "SHOW", "ME", "IN", "THE", "WITH", "AND",
        "COLOR", "COLOUR", "CLARITY", "SHAPE", "CUT", "GRADE", "CVD", "HPHT",
    }
    skip |= set(shapes)
    for tok in tokens:
        if tok in skip:
            continue
        if re.match(r"^\d+\.?\d*(CT|CARAT|CTS|CARATS)$", tok):
            continue
        if re.match(r"^(FL|IF|VVS[12]?|VS[12]?|SI[12]?|I[123]?)$", tok):
            continue
        # Digit-only or short cert-like → prefer certificate number partial match
        if re.match(r"^\d{3,}$", tok) or (tok.startswith("LG") and len(tok) >= 4):
            if "Certificate Number" in df.columns:
                mask &= df["Certificate Number"].astype(str).str.upper().str.contains(
                    re.escape(tok), na=False
                )
            continue
        # General partial: token must appear in any searchable column
        if search_cols:
            sub = pd.Series([False] * len(df), index=df.index)
            for c in search_cols:
                sub |= df[c].astype(str).str.upper().str.contains(re.escape(tok), na=False)
            mask &= sub

    # If query is a single short fragment with no structured hit, still partial-search all
    if len(tokens) == 1 and tokens[0] not in skip and search_cols:
        tok = tokens[0]
        if not (
            carat_m or color_m2 or clar_m
            or any(re.search(rf"\b{s}\b", q) for s in shapes)
        ):
            sub = pd.Series([False] * len(df), index=df.index)
            for c in search_cols:
                sub |= df[c].astype(str).str.upper().str.contains(re.escape(tok), na=False)
            mask = sub

    return df[mask]


# -------------------------------------------------
# Batch lookup (runs in worker threads – no Streamlit calls in here)
# -------------------------------------------------
def fetch_item(item_type: str, value: str) -> Dict[str, Any]:
    if item_type == "token":
        return fetch_gia_from_token(value)
    if is_igi(value):
        return fetch_igi_report(value)
    if is_gia(value):
        return fetch_gia_info(value)
    data = empty_inventory_row(value, "Unknown")
    data["Status"] = "Unknown Format"
    data["Notes"] = "Could not determine if IGI or GIA"
    return data


def _clean_url(v: Any) -> str:
    s = _s(v)
    return s if s.lower().startswith(("http://", "https://")) else ""


def _clear_cert_input() -> None:
    st.session_state["cert_input"] = ""
    st.session_state.pop("batch_results", None)
    st.session_state["batch_preview_cert"] = None


# -------------------------------------------------
# Streamlit UI
# -------------------------------------------------
st.markdown("## Certified Stone Inventory")
st.caption("FGI · Certificate lookup and inventory")

with st.sidebar:
    st.markdown("**Overview**")
    _stats_df = load_all_inventory()
    st.metric("Total certificates", len(_stats_df))
    st.metric("Added today", count_added_today(_stats_df))
    st.caption("Storage: Google Sheets" if use_gsheets() else "Storage: local SQLite (not shared)")

tab_add, tab_all, tab_edit, tab_video = st.tabs([
    "Add certificates",
    "Inventory",
    "Edit records",
    "Media links",
])

# =========================================================
# TAB 1 – Add new certificates
# =========================================================
with tab_add:
    st.markdown("##### Fetch by certificate number")
    report_text = st.text_area(
        "Certificate numbers",
        height=140,
        placeholder="Enter one certificate number per line (IGI or GIA)",
        key="cert_input",
    )
    col1, col2 = st.columns([1, 1])
    with col1:
        fetch_btn = st.button("Fetch & save", type="primary", **_STRETCH)
    with col2:
        st.button("Clear", on_click=_clear_cert_input, **_STRETCH)

    st.divider()
    st.markdown("##### Upload certificate PDF")
    st.caption("Use when live fetch is blocked (e.g. some IGI/GIA reports). Upload the PDF from your browser download.")
    uploaded_files = st.file_uploader(
        "Certificate PDF files", type=["pdf"], accept_multiple_files=True, key="cert_pdf_uploader",
    )
    upload_btn = st.button("Parse & save PDFs", type="primary", key="parse_pdfs_btn", **_STRETCH)

    if upload_btn and uploaded_files:
        results_upload = []
        progress = st.progress(0)
        for i, uf in enumerate(uploaded_files):
            results_upload.append(parse_certificate_pdf_bytes(uf.read(), filename=uf.name))
            progress.progress((i + 1) / len(uploaded_files))
        progress.empty()
        try:
            report_upsert(upsert_rows(results_upload))
        except RuntimeError as e:
            st.error(str(e))
        st.session_state["upload_results"] = pd.DataFrame(results_upload)
    elif upload_btn and not uploaded_files:
        st.warning("Please choose one or more PDF files first.")

    _up = st.session_state.get("upload_results")
    if _up is not None and not _up.empty:
        show_cols = [c for c in [
            "Certificate Number", "Lab", "SHAPE AND CUT", "CARAT WEIGHT", "COLOR GRADE",
            "CLARITY GRADE", "MEASUREMENTS", "Process", "Status", "Notes",
        ] if c in _up.columns]
        st.dataframe(_up[show_cols], hide_index=True, **_STRETCH)

    if fetch_btn and report_text.strip():
        raw_lines = [line.strip() for line in report_text.strip().splitlines() if line.strip()]
        seen, items = set(), []
        for line in raw_lines:
            if is_gia_token_url(line):
                if line not in seen:
                    seen.add(line)
                    items.append(("token", line))
            else:
                cleaned = clean_report_number(line)
                if cleaned and cleaned not in seen:
                    seen.add(cleaned)
                    items.append(("number", cleaned))

        progress = st.progress(0)
        status_text = st.empty()
        results: List[Optional[Dict[str, Any]]] = [None] * len(items)
        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = {pool.submit(fetch_item, t, v): i for i, (t, v) in enumerate(items)}
            for n_done, fut in enumerate(as_completed(futures), start=1):
                i = futures[fut]
                try:
                    results[i] = fut.result()
                except Exception as e:           # never let one lookup kill the batch
                    logger.exception("Lookup failed")
                    row = empty_inventory_row(items[i][1], "Unknown")
                    row.update(Status="Error", Notes=str(e)[:150])
                    results[i] = row
                status_text.text(f"Retrieved {n_done}/{len(items)}")
                progress.progress(n_done / len(items))
        progress.empty()
        status_text.empty()

        # Save everything that has a usable certificate number (including "Not Found",
        # so grades can be filled later). Merge rules protect existing data.
        to_save = [r for r in results if r.get("Status") not in ("Unknown Format", "Error")]
        try:
            report_upsert(upsert_rows(to_save))
        except RuntimeError as e:
            st.error(str(e))

        df_new = pd.DataFrame(results)
        df_new["Saved"] = [
            r.get("Status") not in ("Unknown Format", "Error") and _valid_cert(r.get("Certificate Number"))
            for r in results
        ]
        st.session_state["batch_results"] = df_new[[c for c in INVENTORY_COLUMNS + ["Saved"] if c in df_new.columns]]
        st.session_state["batch_preview_cert"] = None

    _batch = st.session_state.get("batch_results")
    if _batch is not None and not _batch.empty:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Processed", len(_batch))
        m2.metric("IGI", int((_batch["Lab"] == "IGI").sum()))
        m3.metric("GIA", int((_batch["Lab"] == "GIA").sum()))
        m4.metric("Added today", count_added_today())
        ev_new = st.dataframe(
            _batch, hide_index=True, on_select="rerun", selection_mode="single-row",
            key="batch_table_select", **_STRETCH,
        )
        maybe_open_preview("batch_preview_cert", ev_new.selection.rows if ev_new else [], _batch)

        export = _batch.drop(columns=["Lab", "Status", "Notes", "date_added", "last_updated", "Saved"], errors="ignore")
        st.download_button(
            "Export Excel",
            data=df_to_xlsx_bytes(export, "Batch"),
            file_name=f"Batch_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# =========================================================
# TAB 2 – All Inventory
# =========================================================
with tab_all:
    df_all = load_all_inventory()

    t1, t2, t3 = st.columns([3, 1, 1])
    with t1:
        search_q = st.text_input(
            "Search", value="",
            placeholder="Search by certificate, shape, color, clarity, carat…",
            label_visibility="collapsed", key="inventory_search",
        )
    with t2:
        filter_lab = st.selectbox("Lab", ["All", "IGI", "GIA", "Unknown"], key="lab_filter", label_visibility="collapsed")
    with t3:
        filter_today = st.checkbox("Today only", value=False, key="today_filter")

    view = df_all.copy()
    if filter_lab != "All" and not view.empty:
        view = view[view["Lab"] == filter_lab]
    if filter_today and not view.empty:
        view = view[view["date_added"].astype(str).str.startswith(date.today().isoformat())]
    if search_q.strip() and not view.empty:
        view = smart_filter(view, search_q)

    list_cols = [c for c in INVENTORY_COLUMNS if c in view.columns]
    display_df = view[list_cols].copy() if list_cols else view.copy()
    col_config = {}
    for url_col in ("CERTIFICATE LINK", "Video URL", "Video Backup Drive Link", "Certificate PDF Backup Drive Link"):
        if url_col in display_df.columns:
            display_df[url_col] = display_df[url_col].map(_clean_url)
            col_config[url_col] = st.column_config.TextColumn(url_col, width="medium")

    render_inventory_subtotals(view, title=f"{len(view)}/{len(df_all)} shown")
    if view.empty:
        st.info("No matching records.")
    else:
        event = st.dataframe(
            display_df, hide_index=True, height=600,
            on_select="rerun", selection_mode="single-row",
            key="inventory_table_select", column_config=col_config or None, **_STRETCH,
        )
        maybe_open_preview("inv_preview_cert", event.selection.rows if event else [], view)

        e1, e2 = st.columns(2)
        with e1:
            export = view.drop(columns=[c for c in ["Lab", "Status", "Notes"] if c in view.columns], errors="ignore")
            st.download_button(
                "Export Excel",
                data=df_to_xlsx_bytes(export, "Inventory"),
                file_name="Filtered_Inventory.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                **_STRETCH,
            )
        with e2:
            n_view = len(view)
            max_zip = 80
            zip_key = "|".join(view["Certificate Number"].astype(str).head(max_zip))
            st.caption(f"Certificate PDFs for current filter ({min(n_view, max_zip)} max per run)")
            if st.button("Prepare certificate ZIP", key="prep_cert_zip", disabled=n_view == 0, **_STRETCH):
                with st.spinner(f"Downloading up to {min(n_view, max_zip)} certificate PDF(s)…"):
                    zip_bytes, zip_logs = build_certificates_zip(view, max_files=max_zip)
                st.session_state.update(
                    cert_zip_bytes=zip_bytes, cert_zip_logs=zip_logs, cert_zip_key=zip_key,
                    cert_zip_name=f"Certificates_{date.today().isoformat()}.zip",
                )

            zip_ready = st.session_state.get("cert_zip_key") == zip_key
            logs = st.session_state.get("cert_zip_logs") or []
            if zip_ready and st.session_state.get("cert_zip_bytes"):
                st.download_button(
                    "Download ZIP", data=st.session_state["cert_zip_bytes"],
                    file_name=st.session_state.get("cert_zip_name") or "Certificates.zip",
                    mime="application/zip", key="dl_cert_zip", **_STRETCH,
                )
                ok_n = sum(1 for x in logs if x.get("Status") == "OK")
                st.caption(f"ZIP ready: {ok_n}/{len(logs)} PDF(s) included")
            elif zip_ready and logs:
                st.warning("No PDFs could be downloaded for this filter (lab block or missing files).")
            elif st.session_state.get("cert_zip_key"):
                st.caption("The filter changed since the ZIP was prepared — prepare it again.")
            if zip_ready and logs:
                with st.expander("Download log"):
                    st.dataframe(pd.DataFrame(logs), hide_index=True, **_STRETCH)

# =========================================================
# TAB 3 – Edit editable fields
# =========================================================
with tab_edit:
    _flash = st.session_state.pop("edit_flash", None)
    if _flash:
        getattr(st, _flash[0])(_flash[1])

    df_all = load_all_inventory()
    if df_all.empty:
        st.info("No inventory records available.")
    else:
        edit_cols = [
            "Certificate Number", "SHAPE AND CUT", "CARAT WEIGHT", "COLOR GRADE", "CLARITY GRADE",
        ] + EDITABLE_COLUMNS
        edit_cols = [c for c in edit_cols if c in df_all.columns]
        df_edit = df_all[edit_cols].reset_index(drop=True)

        edit_q = st.text_input(
            "Filter rows", value="", key="edit_filter",
            placeholder="Filter before editing (save pending edits first — changing the filter resets them)",
        )
        df_view = smart_filter(df_edit, edit_q).reset_index(drop=True) if edit_q.strip() else df_edit

        b1, b2, b3 = st.columns([1, 3, 1])
        with b1:
            save_clicked = st.button("Save", type="primary", **_STRETCH)
        with b2:
            pick_e = st.selectbox(
                "Preview", df_view["Certificate Number"].astype(str).tolist(),
                key="edit_preview_pick", label_visibility="collapsed",
            )
        with b3:
            preview_clicked = st.button("Preview", key="edit_preview_btn", **_STRETCH)

        editor_key = f"inventory_editor_{st.session_state.get('edit_gen', 0)}_{abs(hash(edit_q))}"
        edited = st.data_editor(
            df_view, hide_index=True, num_rows="fixed", height=560, key=editor_key,
            column_config={
                "Certificate Number": st.column_config.TextColumn(disabled=True, width="small"),
                "SHAPE AND CUT": st.column_config.TextColumn(disabled=True, width="medium"),
                "CARAT WEIGHT": st.column_config.TextColumn(disabled=True, width="small"),
                "COLOR GRADE": st.column_config.TextColumn(disabled=True, width="small"),
                "CLARITY GRADE": st.column_config.TextColumn(disabled=True, width="small"),
            },
            **_STRETCH,
        )

        changes = compute_edit_changes(df_view, edited)
        n_cells = sum(len(v) for v in changes.values())
        if n_cells:
            st.caption(f"{n_cells} unsaved cell change(s) in {len(changes)} row(s).")

        if save_clicked:
            if not changes:
                st.info("No changes to save.")
            else:
                try:
                    res = update_editable_fields(changes)
                except Exception as e:
                    logger.exception("Edit save failed")
                    st.error(f"Save failed: {str(e)[:300]}")
                else:
                    conflicts = res["conflicts"]
                    if conflicts:
                        sample = ", ".join(f"{c} · {col}" for c, col in conflicts[:5])
                        st.session_state["edit_flash"] = (
                            "warning",
                            f"Saved {res['updated']} cell(s). Skipped {len(conflicts)} cell(s) that someone "
                            f"else changed since you loaded the table ({sample}"
                            f"{'…' if len(conflicts) > 5 else ''}). Review and save again if needed.",
                        )
                    else:
                        st.session_state["edit_flash"] = ("success", f"Saved {res['updated']} cell(s).")
                    st.session_state["edit_gen"] = st.session_state.get("edit_gen", 0) + 1
                    st.rerun()

        if preview_clicked:
            open_certificate_preview(pick_e, "")

# =========================================================
# TAB 4 – Video links & Drive backup
# =========================================================
with tab_video:
    video_text = st.text_area(
        "Certificate number and media URL", value="", height=180,
        placeholder="One pair per line: certificate number, then URL (tab or space separated)",
        key="video_bulk_input",
    )
    try_backup = st.checkbox("Upload direct video files to linked storage", value=False)

    if st.button("Save media links", type="primary", key="save_videos_btn"):
        pairs = []
        for line in video_text.strip().splitlines():
            parts = re.split(r"\s+", line.strip(), maxsplit=1)
            if len(parts) == 2:
                pairs.append((parts[0], parts[1]))

        if not pairs:
            st.error("No valid certificate / URL pairs found.")
        else:
            with st.spinner("Saving…"):
                try:
                    results, upsert_res = save_video_urls(pairs, try_drive_backup=try_backup)
                except RuntimeError as e:
                    results, upsert_res = [], {"warning": str(e)}
            report_upsert(upsert_res)
            if results:
                st.dataframe(pd.DataFrame(results), hide_index=True, **_STRETCH)

st.caption("Certified Stone Inventory · FGI")
