"""
Certified Stone Inventory Application
-------------------------------------
Internal tool for certificate lookup, inventory management,
and media reference for certified stones.
"""
import streamlit as st
import pandas as pd
import requests
import re
import io
import sqlite3
import os
from datetime import datetime, date
from typing import Optional, Dict, Any, List
import time

# Optional: pdfplumber for better PDF text extraction
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

# Optional: Playwright for GIA scraping
try:
    from playwright.sync_api import sync_playwright
    HAS_PLAYWRIGHT = True
except ImportError:
    HAS_PLAYWRIGHT = False

st.set_page_config(
    page_title="Certified Stone Inventory",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------
# Database layer: Google Sheets (preferred) + SQLite fallback
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

SHEET_HEADERS = INVENTORY_COLUMNS  # same order written to Google Sheet


def _has_gsheets_secrets() -> bool:
    try:
        return (
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


@st.cache_resource(show_spinner=False)
def _get_google_creds():
    from google.oauth2.service_account import Credentials
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
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
        ws = sh.add_worksheet(title=ws_name, rows=2000, cols=len(SHEET_HEADERS))
        ws.append_row(SHEET_HEADERS)
    return ws


def use_gsheets() -> bool:
    return _has_gsheets_secrets()


def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    """Initialize SQLite fallback table. Does not hit Google Sheets on every run."""
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS inventory (
            certificate_number TEXT PRIMARY KEY,
            description TEXT,
            shape_and_cut TEXT,
            carat_weight TEXT,
            color_grade TEXT,
            clarity_grade TEXT,
            measurements TEXT,
            cut_grade TEXT,
            polish TEXT,
            symmetry TEXT,
            fluorescence TEXT,
            process TEXT,
            fgi_item_number TEXT,
            date_shipped_from_india TEXT,
            for_stock_or_customer TEXT,
            production_order TEXT,
            current_location_status TEXT,
            price_per_carat TEXT,
            price_for_stone TEXT,
            certificate_link TEXT,
            comment TEXT,
            lab TEXT,
            status TEXT,
            notes TEXT,
            date_added TEXT,
            last_updated TEXT
        )
    """)
    conn.commit()
    conn.close()


def row_to_db_dict(row: Dict[str, Any]) -> Dict[str, Any]:
    now = datetime.now().isoformat(timespec="seconds")
    return {
        "certificate_number": row.get("Certificate Number") or "",
        "description": row.get("DESCRIPTION") or "",
        "shape_and_cut": row.get("SHAPE AND CUT") or "",
        "carat_weight": row.get("CARAT WEIGHT") or "",
        "color_grade": row.get("COLOR GRADE") or "",
        "clarity_grade": row.get("CLARITY GRADE") or "",
        "measurements": row.get("MEASUREMENTS") or "",
        "cut_grade": row.get("CUT") or "",
        "polish": row.get("POLISH") or "",
        "symmetry": row.get("SYMMETRY") or "",
        "fluorescence": row.get("FLUORESCENCE") or "",
        "process": row.get("Process") or "",
        "fgi_item_number": row.get("FGI Item Number") or "",
        "date_shipped_from_india": row.get("Date Shipped from India") or "",
        "for_stock_or_customer": row.get("For stock or Customer") or "",
        "production_order": row.get("Production Order #, if for stock") or "",
        "current_location_status": row.get("Current Location/Status") or "",
        "price_per_carat": row.get("Price Per Carat") or "",
        "price_for_stone": row.get("Price For stone") or "",
        "certificate_link": row.get("CERTIFICATE LINK") or "",
        "comment": row.get("Comment") or "",
        "video_url": row.get("Video URL") or "",
        "video_backup_drive_link": row.get("Video Backup Drive Link") or "",
        "certificate_pdf_backup_drive_link": row.get("Certificate PDF Backup Drive Link") or "",
        "lab": row.get("Lab") or "",
        "status": row.get("Status") or "",
        "notes": row.get("Notes") or "",
        "date_added": row.get("date_added") or now,
        "last_updated": now,
    }


def db_dict_to_row(d: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "Certificate Number": d.get("certificate_number"),
        "DESCRIPTION": d.get("description"),
        "SHAPE AND CUT": d.get("shape_and_cut"),
        "CARAT WEIGHT": d.get("carat_weight"),
        "COLOR GRADE": d.get("color_grade"),
        "CLARITY GRADE": d.get("clarity_grade"),
        "MEASUREMENTS": d.get("measurements"),
        "CUT": d.get("cut_grade"),
        "POLISH": d.get("polish"),
        "SYMMETRY": d.get("symmetry"),
        "FLUORESCENCE": d.get("fluorescence"),
        "Process": d.get("process"),
        "FGI Item Number": d.get("fgi_item_number"),
        "Date Shipped from India": d.get("date_shipped_from_india"),
        "For stock or Customer": d.get("for_stock_or_customer"),
        "Production Order #, if for stock": d.get("production_order"),
        "Current Location/Status": d.get("current_location_status"),
        "Price Per Carat": d.get("price_per_carat"),
        "Price For stone": d.get("price_for_stone"),
        "CERTIFICATE LINK": d.get("certificate_link"),
        "Comment": d.get("comment"),
        "Video URL": d.get("video_url"),
        "Video Backup Drive Link": d.get("video_backup_drive_link"),
        "Certificate PDF Backup Drive Link": d.get("certificate_pdf_backup_drive_link"),
        "Lab": d.get("lab"),
        "Status": d.get("status"),
        "Notes": d.get("notes"),
        "date_added": d.get("date_added"),
        "last_updated": d.get("last_updated"),
    }


def _row_to_sheet_list(row: Dict[str, Any]) -> List[str]:
    """Ordered values matching SHEET_HEADERS."""
    d = row_to_db_dict(row)
    mapping = {
        "Certificate Number": d["certificate_number"],
        "DESCRIPTION": d["description"],
        "SHAPE AND CUT": d["shape_and_cut"],
        "CARAT WEIGHT": d["carat_weight"],
        "COLOR GRADE": d["color_grade"],
        "CLARITY GRADE": d["clarity_grade"],
        "MEASUREMENTS": d["measurements"],
        "CUT": d["cut_grade"],
        "POLISH": d["polish"],
        "SYMMETRY": d["symmetry"],
        "FLUORESCENCE": d["fluorescence"],
        "Process": d["process"],
        "FGI Item Number": d["fgi_item_number"],
        "Date Shipped from India": d["date_shipped_from_india"],
        "For stock or Customer": d["for_stock_or_customer"],
        "Production Order #, if for stock": d["production_order"],
        "Current Location/Status": d["current_location_status"],
        "Price Per Carat": d["price_per_carat"],
        "Price For stone": d["price_for_stone"],
        "CERTIFICATE LINK": d["certificate_link"],
        "Comment": d["comment"],
        "Video URL": d["video_url"],
        "Video Backup Drive Link": d["video_backup_drive_link"],
        "Certificate PDF Backup Drive Link": d["certificate_pdf_backup_drive_link"],
        "Lab": d["lab"],
        "Status": d["status"],
        "Notes": d["notes"],
        "date_added": d["date_added"],
        "last_updated": d["last_updated"],
    }
    return [str(mapping.get(h, "") or "") for h in SHEET_HEADERS]


def _ensure_sheet_headers(ws) -> None:
    """Rewrite row 1 if it has blanks/duplicates or is missing new columns."""
    current = ws.row_values(1)
    needs_fix = (
        not current
        or any((h or "").strip() == "" for h in current)
        or len(set(current)) != len(current)
        or list(current[: len(SHEET_HEADERS)]) != list(SHEET_HEADERS)
    )
    if needs_fix:
        ws.update("A1", [SHEET_HEADERS], value_input_option="USER_ENTERED")


def _sheet_records() -> List[Dict[str, str]]:
    ws = _open_worksheet()
    try:
        _ensure_sheet_headers(ws)
    except Exception:
        pass

    # Prefer expected_headers (gspread ≥5) so duplicate blanks don't crash
    try:
        records = ws.get_all_records(expected_headers=SHEET_HEADERS)
        return records or []
    except TypeError:
        pass
    except Exception:
        pass

    # Manual parse fallback
    values = ws.get_all_values()
    if len(values) < 2:
        return []
    header = SHEET_HEADERS
    rows = []
    for raw in values[1:]:
        if not any(str(c).strip() for c in raw):
            continue
        padded = list(raw) + [""] * max(0, len(header) - len(raw))
        rows.append({header[i]: padded[i] for i in range(len(header))})
    return rows


def upsert_rows(rows: List[Dict[str, Any]]):
    """Insert or update rows. Preserves existing editable fields on update."""
    if use_gsheets():
        _upsert_rows_gsheets(rows)
        clear_inventory_cache()
    else:
        _upsert_rows_sqlite(rows)


def _upsert_rows_gsheets(rows: List[Dict[str, Any]]):
    ws = _open_worksheet()
    existing = ws.get_all_values()
    if not existing:
        ws.append_row(SHEET_HEADERS)
        existing = [SHEET_HEADERS]

    header = existing[0]
    # Map certificate number -> row index (1-based in Sheets)
    cert_col = 0
    try:
        cert_col = header.index("Certificate Number")
    except ValueError:
        cert_col = 0

    index_by_cert = {}
    for i, row in enumerate(existing[1:], start=2):
        if len(row) > cert_col and row[cert_col]:
            index_by_cert[row[cert_col].strip()] = i

    # Column indices for editable fields (to preserve)
    def col_idx(name):
        try:
            return header.index(name)
        except ValueError:
            return None

    editable_idx = {name: col_idx(name) for name in EDITABLE_COLUMNS}
    date_added_idx = col_idx("date_added")

    for row in rows:
        cert = (row.get("Certificate Number") or "").strip()
        if not cert:
            continue
        values = _row_to_sheet_list(row)

        if cert in index_by_cert:
            ridx = index_by_cert[cert]
            old = existing[ridx - 1] if ridx - 1 < len(existing) else []
            # Preserve non-empty editable fields
            for name, cidx in editable_idx.items():
                if cidx is not None and cidx < len(old) and old[cidx]:
                    # Find position in SHEET_HEADERS
                    try:
                        pos = SHEET_HEADERS.index(name)
                        values[pos] = old[cidx]
                    except ValueError:
                        pass
            # Preserve original date_added
            if date_added_idx is not None and date_added_idx < len(old) and old[date_added_idx]:
                try:
                    pos = SHEET_HEADERS.index("date_added")
                    values[pos] = old[date_added_idx]
                except ValueError:
                    pass
            # Update row
            end_col = chr(ord('A') + len(values) - 1) if len(values) <= 26 else "Z"
            # Use update with range for full row
            ws.update(f"A{ridx}", [values], value_input_option="USER_ENTERED")
        else:
            ws.append_row(values, value_input_option="USER_ENTERED")


def _upsert_rows_sqlite(rows: List[Dict[str, Any]]):
    conn = get_conn()
    cur = conn.cursor()
    for row in rows:
        cert = (row.get("Certificate Number") or "").strip()
        if not cert:
            continue
        existing = cur.execute(
            "SELECT fgi_item_number, date_shipped_from_india, for_stock_or_customer, "
            "production_order, current_location_status, price_per_carat, price_for_stone, "
            "comment, date_added FROM inventory WHERE certificate_number = ?",
            (cert,)
        ).fetchone()

        d = row_to_db_dict(row)
        if existing:
            (fgi, shipped, stock, po, loc, ppc, pfs, comment, date_added) = existing
            if fgi:
                d["fgi_item_number"] = fgi
            if shipped:
                d["date_shipped_from_india"] = shipped
            if stock:
                d["for_stock_or_customer"] = stock
            if po:
                d["production_order"] = po
            if loc:
                d["current_location_status"] = loc
            if ppc:
                d["price_per_carat"] = ppc
            if pfs:
                d["price_for_stone"] = pfs
            if comment:
                d["comment"] = comment
            d["date_added"] = date_added

        cur.execute("""
            INSERT INTO inventory (
                certificate_number, description, shape_and_cut, carat_weight, color_grade,
                clarity_grade, measurements, cut_grade, polish, symmetry, fluorescence,
                process, fgi_item_number, date_shipped_from_india, for_stock_or_customer,
                production_order, current_location_status, price_per_carat, price_for_stone,
                certificate_link, comment, lab, status, notes, date_added, last_updated
            ) VALUES (
                :certificate_number, :description, :shape_and_cut, :carat_weight, :color_grade,
                :clarity_grade, :measurements, :cut_grade, :polish, :symmetry, :fluorescence,
                :process, :fgi_item_number, :date_shipped_from_india, :for_stock_or_customer,
                :production_order, :current_location_status, :price_per_carat, :price_for_stone,
                :certificate_link, :comment, :lab, :status, :notes, :date_added, :last_updated
            )
            ON CONFLICT(certificate_number) DO UPDATE SET
                description=excluded.description,
                shape_and_cut=excluded.shape_and_cut,
                carat_weight=excluded.carat_weight,
                color_grade=excluded.color_grade,
                clarity_grade=excluded.clarity_grade,
                measurements=excluded.measurements,
                cut_grade=excluded.cut_grade,
                polish=excluded.polish,
                symmetry=excluded.symmetry,
                fluorescence=excluded.fluorescence,
                process=excluded.process,
                fgi_item_number=excluded.fgi_item_number,
                date_shipped_from_india=excluded.date_shipped_from_india,
                for_stock_or_customer=excluded.for_stock_or_customer,
                production_order=excluded.production_order,
                current_location_status=excluded.current_location_status,
                price_per_carat=excluded.price_per_carat,
                price_for_stone=excluded.price_for_stone,
                certificate_link=excluded.certificate_link,
                comment=excluded.comment,
                lab=excluded.lab,
                status=excluded.status,
                notes=excluded.notes,
                last_updated=excluded.last_updated
        """, d)
    conn.commit()
    conn.close()


@st.cache_data(ttl=90, show_spinner=False)
def _load_gsheets_inventory() -> pd.DataFrame:
    """Cached read from Google Sheets (90s) to stay under API quota."""
    records = _sheet_records()
    if not records:
        return pd.DataFrame(columns=INVENTORY_COLUMNS)
    df = pd.DataFrame(records)
    for c in INVENTORY_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    return df[INVENTORY_COLUMNS]


def _normalize_lab_column(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure Lab is filled from certificate number when blank."""
    if df is None or df.empty:
        return df
    if "Lab" not in df.columns:
        df["Lab"] = ""
    if "Certificate Number" in df.columns:
        df["Lab"] = [
            infer_lab(cert, lab)
            for cert, lab in zip(df["Certificate Number"], df["Lab"])
        ]
    return df


def load_all_inventory() -> pd.DataFrame:
    if use_gsheets():
        try:
            return _normalize_lab_column(_load_gsheets_inventory())
        except Exception as e:
            # Don't spam errors on quota – show once-friendly message
            msg = str(e)
            if "429" in msg or "Quota" in msg:
                st.warning("Data service is temporarily busy. Please try again shortly.")
            else:
                st.error("Unable to load inventory data.")
            return pd.DataFrame(columns=INVENTORY_COLUMNS)

    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM inventory ORDER BY date_added DESC", conn)
    conn.close()
    if df.empty:
        return pd.DataFrame(columns=INVENTORY_COLUMNS)
    rows = [db_dict_to_row(r) for r in df.to_dict(orient="records")]
    return _normalize_lab_column(pd.DataFrame(rows))


def clear_inventory_cache():
    """Call after any write so the next read is fresh."""
    try:
        _load_gsheets_inventory.clear()
    except Exception:
        pass


def update_editable_fields(df_edit: pd.DataFrame):
    """Save only the editable columns for existing certificates."""
    if use_gsheets():
        _update_editable_gsheets(df_edit)
        clear_inventory_cache()
    else:
        _update_editable_sqlite(df_edit)


def _update_editable_gsheets(df_edit: pd.DataFrame):
    """Batch-update editable fields (far fewer API calls than update_cell)."""
    import gspread
    ws = _open_worksheet()
    existing = ws.get_all_values()
    if len(existing) < 2:
        return
    header = existing[0]
    try:
        cert_col = header.index("Certificate Number")
    except ValueError:
        return

    index_by_cert = {}
    for i, row in enumerate(existing[1:], start=2):
        if len(row) > cert_col and row[cert_col]:
            index_by_cert[row[cert_col].strip()] = i

    now = datetime.now().isoformat(timespec="seconds")
    cells = []
    for _, row in df_edit.iterrows():
        cert = str(row.get("Certificate Number", "")).strip()
        if not cert or cert not in index_by_cert:
            continue
        ridx = index_by_cert[cert]
        for name in EDITABLE_COLUMNS:
            if name not in header:
                continue
            cidx = header.index(name) + 1
            val = str(row.get(name) or "")
            cells.append(gspread.Cell(ridx, cidx, val))
        if "last_updated" in header:
            cells.append(gspread.Cell(ridx, header.index("last_updated") + 1, now))

    if cells:
        ws.update_cells(cells, value_input_option="USER_ENTERED")


def _update_editable_sqlite(df_edit: pd.DataFrame):
    conn = get_conn()
    cur = conn.cursor()
    now = datetime.now().isoformat(timespec="seconds")
    for _, row in df_edit.iterrows():
        cert = str(row.get("Certificate Number", "")).strip()
        if not cert:
            continue
        cur.execute("""
            UPDATE inventory SET
                fgi_item_number = ?,
                date_shipped_from_india = ?,
                for_stock_or_customer = ?,
                production_order = ?,
                current_location_status = ?,
                price_per_carat = ?,
                price_for_stone = ?,
                comment = ?,
                last_updated = ?
            WHERE certificate_number = ?
        """, (
            str(row.get("FGI Item Number") or ""),
            str(row.get("Date Shipped from India") or ""),
            str(row.get("For stock or Customer") or ""),
            str(row.get("Production Order #, if for stock") or ""),
            str(row.get("Current Location/Status") or ""),
            str(row.get("Price Per Carat") or ""),
            str(row.get("Price For stone") or ""),
            str(row.get("Comment") or ""),
            now,
            cert,
        ))
    conn.commit()
    conn.close()


def count_added_today() -> int:
    today = date.today().isoformat()
    df = load_all_inventory()
    if df.empty or "date_added" not in df.columns:
        return 0
    return int(df["date_added"].astype(str).str.startswith(today).sum())


def total_count() -> int:
    df = load_all_inventory()
    return len(df)


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


def render_certificate_preview(cert: str, cert_link: str = "", show_backup_btn: bool = True):
    """Show certificate PDF preview + optional Drive backup (used after row click)."""
    cert = str(cert or "").strip()
    if not cert:
        return
    st.markdown(f"### Certificate {cert}")
    existing = load_all_inventory()
    link = cert_link
    backup_link = ""
    match_row = None
    if not existing.empty and "Certificate Number" in existing.columns:
        match = existing[existing["Certificate Number"].astype(str) == cert]
        if not match.empty:
            match_row = match.iloc[0].to_dict()
            if not link:
                link = str(match_row.get("CERTIFICATE LINK") or "")
            backup_link = str(match_row.get("Certificate PDF Backup Drive Link") or "").strip()

    c1, c2, c3 = st.columns(3)
    with c1:
        do_backup = st.button(
            "Save PDF copy",
            key=f"prev_bak_{cert}",
            disabled=not _has_gdrive_folder(),
        )
    with c2:
        st.button("Refresh", key=f"prev_rel_{cert}")
    with c3:
        if backup_link:
            st.markdown(f"[Open saved copy]({backup_link})")

    pdf_bytes = None
    if do_backup and _has_gdrive_folder():
        with st.spinner("Saving certificate…"):
            res = backup_certificate_pdf_to_drive(cert, link)
        if res.get("success"):
            row = match_row or empty_inventory_row(cert, infer_lab(cert))
            row["Certificate Number"] = cert
            row["Certificate PDF Backup Drive Link"] = res["drive_link"]
            if link:
                row["CERTIFICATE LINK"] = link
            upsert_rows([row])
            st.success("Certificate copy saved.")
            pdf_bytes = res.get("pdf_bytes")
        else:
            st.error(res.get("error", "Unable to save certificate copy."))

    if pdf_bytes is None:
        with st.spinner("Loading certificate…"):
            dl = download_certificate_pdf(cert, link)
        if not dl.get("success"):
            st.warning(dl.get("error", "Certificate file is not available."))
            return
        pdf_bytes = dl["pdf_bytes"]

    images = pdf_to_preview_images(pdf_bytes, max_pages=2, scale=1.5)
    if images:
        for i, img in enumerate(images):
            st.image(img, caption=f"Page {i + 1}", use_container_width=True)
    st.download_button(
        "Download PDF",
        data=pdf_bytes,
        file_name=f"{cert}_certificate.pdf",
        mime="application/pdf",
        key=f"prev_dl_{cert}",
    )


# Modal popup when Streamlit supports st.dialog
try:
    @st.dialog("Certificate preview", width="large")
    def certificate_preview_dialog(cert: str, cert_link: str = ""):
        render_certificate_preview(cert, cert_link)
    HAS_DIALOG = True
except Exception:
    HAS_DIALOG = False


def open_certificate_preview(cert: str, cert_link: str = ""):
    """Open preview as dialog if available, else store in session for inline panel."""
    cert = str(cert or "").strip()
    if not cert:
        return
    if HAS_DIALOG:
        certificate_preview_dialog(cert, cert_link)
    else:
        st.session_state["inline_preview_cert"] = cert
        st.session_state["inline_preview_link"] = cert_link or ""


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
                    hide_index=True, use_container_width=True, height=160,
                )
        with b2:
            if s["by_color"]:
                st.dataframe(
                    pd.DataFrame({"Color": list(s["by_color"].keys()), "#": list(s["by_color"].values())}),
                    hide_index=True, use_container_width=True, height=160,
                )
        with b3:
            if s["by_clarity"]:
                st.dataframe(
                    pd.DataFrame({"Clarity": list(s["by_clarity"].keys()), "#": list(s["by_clarity"].values())}),
                    hide_index=True, use_container_width=True, height=160,
                )


init_db()

# -------------------------------------------------
# Helper functions
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


def probe_video_url(url: str) -> Dict[str, Any]:
    """Check if URL is a direct downloadable video, or find an mp4 inside an HTML page."""
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        r = requests.get(url, headers=headers, timeout=20, allow_redirects=True)
        ct = (r.headers.get("Content-Type") or "").lower()
        size = r.headers.get("Content-Length")
        final = r.url
        is_video = ("video/" in ct) or final.lower().endswith((".mp4", ".webm", ".mov"))

        # If HTML page, try to extract a direct .mp4 URL from the source
        found_mp4 = None
        if r.status_code == 200 and "html" in ct:
            mp4s = re.findall(r'https?://[^\s"\'<>]+\.mp4[^\s"\'<>]*', r.text, flags=re.I)
            if mp4s:
                found_mp4 = mp4s[0].replace("&amp;", "&")
                is_video = True
                final = found_mp4

        return {
            "ok": r.status_code == 200,
            "status": r.status_code,
            "content_type": ct,
            "size": int(size) if size and str(size).isdigit() else None,
            "is_direct_video": bool(is_video and r.status_code == 200),
            "final_url": final,
            "extracted_mp4": found_mp4,
        }
    except Exception as e:
        return {
            "ok": False, "status": 0, "content_type": "", "size": None,
            "is_direct_video": False, "final_url": url, "error": str(e),
        }


def download_and_upload_video_to_drive(cert: str, url: str) -> Dict[str, Any]:
    """
    Download a direct video file and upload to Google Drive folder.
    Returns dict with drive_link or error.
    """
    if not _has_gdrive_folder():
        return {"success": False, "error": "Drive folder_id not configured in secrets"}

    probe = probe_video_url(url)
    if not probe.get("is_direct_video"):
        return {
            "success": False,
            "error": f"Not a direct video file (type={probe.get('content_type') or 'unknown'}). Vision360/DNA HTML cannot be converted to MP4 automatically.",
            "link_saved_only": True,
        }

    download_url = probe.get("extracted_mp4") or probe.get("final_url") or url
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        r = requests.get(download_url, headers=headers, timeout=180, allow_redirects=True)
        if r.status_code != 200:
            return {"success": False, "error": f"Download HTTP {r.status_code}"}

        # Pick extension
        ct = (r.headers.get("Content-Type") or "").lower()
        ext = ".mp4"
        if "webm" in ct:
            ext = ".webm"
        elif "quicktime" in ct or "mov" in ct:
            ext = ".mov"

        filename = f"{cert}_video{ext}"
        folder_id = st.secrets["drive"]["folder_id"]
        service = _get_drive_service()

        from googleapiclient.http import MediaIoBaseUpload
        media = MediaIoBaseUpload(io.BytesIO(r.content), mimetype=ct or "video/mp4", resumable=True)
        meta = {"name": filename, "parents": [folder_id]}
        f = service.files().create(body=meta, media_body=media, fields="id,webViewLink,webContentLink").execute()

        # Make readable by link (optional)
        try:
            service.permissions().create(
                fileId=f["id"],
                body={"type": "anyone", "role": "reader"},
            ).execute()
        except Exception:
            pass

        link = f.get("webViewLink") or f"https://drive.google.com/file/d/{f['id']}/view"
        return {"success": True, "drive_link": link, "file_id": f["id"], "filename": filename}
    except Exception as e:
        return {"success": False, "error": str(e)[:200]}


def save_video_urls(pairs: List[tuple], try_drive_backup: bool = True) -> List[Dict[str, Any]]:
    """
    pairs: list of (certificate_number, video_url)
    Updates inventory Video URL; optionally backs up direct videos to Drive.
    """
    results = []
    existing = load_all_inventory()
    existing_map = {}
    if not existing.empty and "Certificate Number" in existing.columns:
        for _, row in existing.iterrows():
            existing_map[str(row["Certificate Number"]).strip()] = row.to_dict()

    to_upsert = []
    for cert, url in pairs:
        cert = str(cert).strip()
        url = str(url).strip()
        if not cert or not url:
            continue
        row = existing_map.get(cert) or empty_inventory_row(cert, infer_lab(cert))
        row["Certificate Number"] = cert
        row["Lab"] = infer_lab(cert, row.get("Lab") or "")
        row["Video URL"] = url
        info = {"Certificate Number": cert, "Video URL": url, "backup": "skipped"}

        if try_drive_backup and _has_gdrive_folder():
            up = download_and_upload_video_to_drive(cert, url)
            if up.get("success"):
                row["Video Backup Drive Link"] = up["drive_link"]
                info["backup"] = "uploaded"
                info["Video Backup Drive Link"] = up["drive_link"]
            elif up.get("link_saved_only"):
                info["backup"] = "link only (not direct video)"
                info["note"] = up.get("error", "")
            else:
                info["backup"] = "failed"
                info["note"] = up.get("error", "")
        else:
            info["backup"] = "link saved (Drive not configured or disabled)"

        to_upsert.append(row)
        results.append(info)

    if to_upsert:
        upsert_rows(to_upsert)
    return results


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
    """Download certificate PDF bytes. Handles IGI FDR URLs and GIA token pages."""
    url = resolve_certificate_pdf_url(cert, cert_link)
    if not url:
        return {"success": False, "error": "No PDF URL available for this certificate"}

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/pdf,*/*",
    }
    try:
        resp = requests.get(url, headers=headers, timeout=40, allow_redirects=True)
        if resp.status_code == 404:
            return {"success": False, "error": "PDF not found (404) on lab server"}
        if resp.status_code != 200:
            return {"success": False, "error": f"HTTP {resp.status_code}"}

        content = resp.content
        ct = (resp.headers.get("Content-Type") or "").lower()

        # GIA token pages return HTML that redirects to S3 PDF
        if "pdf" not in ct and content[:4] != b"%PDF":
            html = content.decode("utf-8", errors="ignore")
            m = re.search(r'https://[^"\']+\.pdf[^"\']*', html)
            if not m:
                m = re.search(r'window\.location\s*=\s*\(\s*["\'](https://[^"\']+)["\']', html)
            if m:
                pdf_url = m.group(1) if m.lastindex else m.group(0)
                pdf_resp = requests.get(pdf_url, headers=headers, timeout=40)
                if pdf_resp.status_code != 200 or pdf_resp.content[:4] != b"%PDF":
                    return {"success": False, "error": "Could not download redirected PDF"}
                content = pdf_resp.content
            else:
                return {"success": False, "error": "Response is not a PDF"}

        if content[:4] != b"%PDF":
            return {"success": False, "error": "Downloaded file is not a valid PDF"}

        return {"success": True, "pdf_bytes": content, "source_url": url}
    except Exception as e:
        return {"success": False, "error": str(e)[:200]}


def pdf_to_preview_images(pdf_bytes: bytes, max_pages: int = 2, scale: float = 1.5) -> List[Any]:
    """Render PDF pages to PIL images for in-app preview."""
    images = []
    try:
        import pypdfium2 as pdfium
        from PIL import Image
        doc = pdfium.PdfDocument(pdf_bytes)
        n = min(len(doc), max_pages)
        for i in range(n):
            page = doc[i]
            bitmap = page.render(scale=scale)
            pil = bitmap.to_pil()
            images.append(pil)
        return images
    except Exception:
        pass
    # Fallback: try pdfplumber + blank placeholder message via None
    try:
        import pdfplumber
        from PIL import Image, ImageDraw, ImageFont
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for i, page in enumerate(pdf.pages[:max_pages]):
                # pdfplumber can't always rasterize; draw text extract as fallback card
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
        return []


def upload_bytes_to_drive(filename: str, data: bytes, mime: str = "application/pdf") -> Dict[str, Any]:
    """Upload raw bytes to configured Google Drive folder."""
    if not _has_gdrive_folder():
        return {"success": False, "error": "Drive folder_id not configured in secrets"}
    try:
        from googleapiclient.http import MediaIoBaseUpload
        folder_id = str(st.secrets["drive"]["folder_id"]).strip()
        if not folder_id:
            return {"success": False, "error": "drive.folder_id is empty in secrets"}
        service = _get_drive_service()
        media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
        meta = {"name": filename, "parents": [folder_id]}
        f = (
            service.files()
            .create(
                body=meta,
                media_body=media,
                fields="id,webViewLink",
                supportsAllDrives=True,
            )
            .execute()
        )
        try:
            service.permissions().create(
                fileId=f["id"],
                body={"type": "anyone", "role": "reader"},
                supportsAllDrives=True,
            ).execute()
        except Exception:
            pass
        link = f.get("webViewLink") or f"https://drive.google.com/file/d/{f['id']}/view"
        return {"success": True, "drive_link": link, "file_id": f["id"], "filename": filename}
    except Exception as e:
        err = str(e)
        # Common causes
        if "File not found" in err or "404" in err:
            err += " — Check folder_id and that the folder is shared with the service account as Editor."
        if "googleapiclient" in err or "No module" in err:
            err += " — Add google-api-python-client to requirements.txt and reboot the app."
        return {"success": False, "error": err[:400]}


def backup_certificate_pdf_to_drive(cert: str, cert_link: str = "") -> Dict[str, Any]:
    """Download lab PDF and store a permanent copy on Google Drive."""
    try:
        dl = download_certificate_pdf(cert, cert_link)
        if not dl.get("success"):
            return {"success": False, "error": f"Download failed: {dl.get('error', 'unknown')}"}
        safe = re.sub(r"[^\w\-]", "_", str(cert))
        filename = f"{safe}_certificate.pdf"
        up = upload_bytes_to_drive(filename, dl["pdf_bytes"], "application/pdf")
        if up.get("success"):
            up["pdf_bytes"] = dl["pdf_bytes"]
        return up
    except Exception as e:
        return {"success": False, "error": f"Backup error: {str(e)[:300]}"}


def fetch_igi_report(report_no: str) -> Dict[str, Any]:
    """
    Fetch live IGI Laboratory Grown Diamond report from official PDF.
    Returns full inventory-format row.
    """
    number_only = re.sub(r'^LG', '', report_no, flags=re.IGNORECASE)
    cert = report_no if report_no.upper().startswith("LG") else f"LG{number_only}"
    pdf_url = f"https://pdf.igi.org/FDR{number_only}.pdf"

    result = empty_inventory_row(cert, "IGI")
    result["CERTIFICATE LINK"] = pdf_url

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "application/pdf,*/*"
        }
        resp = requests.get(pdf_url, headers=headers, timeout=25)

        if resp.status_code == 404:
            result["Status"] = "Not Found"
            result["Notes"] = "Certificate PDF not available from laboratory at this time"
            result["CERTIFICATE LINK"] = f"https://www.igi.org/Verify-Your-Report/?r={cert}"
            return result
        if resp.status_code != 200:
            result["Status"] = f"HTTP {resp.status_code}"
            result["Notes"] = "Could not download PDF"
            return result

        text = extract_text_from_pdf(resp.content)
        if not text or len(text.strip()) < 100:
            result["Status"] = "Parse Error"
            result["Notes"] = "Could not extract readable text from PDF"
            return result

        text_upper = text.upper()
        text_upper = text_upper.replace("×", "X").replace("–", "-").replace("—", "-")

        # Shape
        shape_patterns = [
            r"SHAPE AND CUTTING STYLE\s*([A-Z0-9\s\-]+?)(?:\n|MEASUREMENTS|GRADING)",
            r"SHAPE AND CUTTING STYLE\s*([^\n]{5,60})",
            r"(ROUND BRILLIANT|PRINCESS CUT|CUSHION MODIFIED BRILLIANT|CUSHION BRILLIANT|"
            r"OVAL BRILLIANT|PEAR BRILLIANT|MARQUISE BRILLIANT|EMERALD CUT|"
            r"SQUARE EMERALD CUT|RADIANT CUT|HEART BRILLIANT|ASSCHER CUT|"
            r"CUT CORNERED RECTANGULAR|TRIANGLE BRILLIANT|TRILLIANT)",
        ]
        for pat in shape_patterns:
            m = re.search(pat, text_upper)
            if m:
                shape = re.sub(r'\s+', ' ', m.group(1)).strip()
                shape = re.split(r'\s{2,}|\d|MM|CARAT', shape)[0].strip()
                if 4 < len(shape) < 50:
                    result["SHAPE AND CUT"] = shape.title()
                    break

        # Measurements
        meas_patterns = [
            r"MEASUREMENTS?\s*([\d\.\-\sX]+?\s*MM)",
            r"([\d\.]+\s*[-–]\s*[\d\.]+\s*[Xx×]\s*[\d\.]+\s*MM)",
            r"([\d\.]+\s*[Xx×]\s*[\d\.]+\s*[Xx×]\s*[\d\.]+\s*MM)",
        ]
        for pat in meas_patterns:
            m = re.search(pat, text_upper)
            if m:
                size = m.group(1).strip().replace("X", "×").replace("x", "×")
                size = re.sub(r'\s+', ' ', size)
                if "MM" not in size.upper():
                    size += " mm"
                else:
                    size = size.replace("MM", "mm")
                result["MEASUREMENTS"] = size
                break

        # Carat
        for pat in [r"(\d+\.\d{2})\s*CARATS?", r"CARAT WEIGHT\s*(\d+\.\d{2})", r"(\d+\.\d{2})\s*CT\b"]:
            m = re.search(pat, text_upper)
            if m:
                result["CARAT WEIGHT"] = f"{m.group(1)} ct"
                break

        # Color
        for pat in [r"COLOR GRADE\s*([D-Z])\b", r"COLOR\s*GRADE\s*[:\s]*([D-Z])\b"]:
            m = re.search(pat, text_upper)
            if m:
                result["COLOR GRADE"] = m.group(1)
                break

        # Clarity
        for pat in [
            r"CLARITY GRADE\s*((?:FL|IF|VVS\s*[12]|VS\s*[12]|SI\s*[12]|I\s*[123]))",
            r"CLARITY\s*GRADE\s*[:\s]*((?:FL|IF|VVS\s*[12]|VS\s*[12]|SI\s*[12]|I\s*[123]))",
        ]:
            m = re.search(pat, text_upper)
            if m:
                result["CLARITY GRADE"] = re.sub(r'\s+', '', m.group(1)).upper()
                break

        # Cut / Polish / Symmetry / Fluorescence
        m = re.search(r"CUT GRADE\s*(IDEAL|EXCELLENT|VERY GOOD|GOOD|FAIR|POOR)", text_upper)
        if m:
            result["CUT"] = m.group(1).title()

        m = re.search(r"POLISH\s*(EXCELLENT|VERY GOOD|GOOD|FAIR|POOR)", text_upper)
        if m:
            result["POLISH"] = m.group(1).title()

        m = re.search(r"SYMMETRY\s*(EXCELLENT|VERY GOOD|GOOD|FAIR|POOR)", text_upper)
        if m:
            result["SYMMETRY"] = m.group(1).title()

        m = re.search(r"FLUORESCENCE\s*(NONE|FAINT|MEDIUM|STRONG|VERY STRONG)", text_upper)
        if m:
            result["FLUORESCENCE"] = m.group(1).title()

        # Process
        if re.search(r"CHEMICAL\s+VAPOR\s+DEPOSITION|\bCVD\b", text_upper):
            result["Process"] = "CVD"
        elif re.search(r"HIGH\s+PRESSURE\s+HIGH\s+TEMPERATURE|\bHPHT\b", text_upper):
            result["Process"] = "HPHT"
        else:
            result["Process"] = "Unknown"

        filled = sum(1 for k in ["SHAPE AND CUT", "MEASUREMENTS", "CARAT WEIGHT", "COLOR GRADE", "CLARITY GRADE"] if result[k])
        if filled >= 3:
            result["Status"] = "Success"
            result["Notes"] = f"Fetched live • {filled}/5 core fields"
        elif filled >= 1:
            result["Status"] = "Partial"
            result["Notes"] = f"Partial extraction • {filled}/5 fields"
        else:
            result["Status"] = "Parse Error"
            result["Notes"] = "Text extracted but key fields not found"

    except Exception as e:
        result["Status"] = "Error"
        result["Notes"] = str(e)[:150]

    return result


def fetch_gia_from_token(token_url: str) -> Dict[str, Any]:
    """
    Download a GIA report PDF using a special token link
    (https://pdf.gia.edu/?ReportNumber=HASH) and extract grading data.
    Returns full inventory-format row.
    """
    result = empty_inventory_row("Unknown", "GIA")
    result["Process"] = "—"
    result["CERTIFICATE LINK"] = token_url

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        }

        # Step 1: Get the token page (contains JS redirect to S3)
        resp = requests.get(token_url, headers=headers, timeout=20, allow_redirects=True)

        pdf_url = None
        content_type = resp.headers.get("Content-Type", "").lower()

        if "pdf" in content_type:
            # Direct PDF
            pdf_bytes = resp.content
            # Try to extract report number from URL if possible
            m = re.search(r"/(\d{8,13})\.pdf", token_url)
            if m:
                result["Certificate Number"] = m.group(1)
        else:
            # Look for the S3 redirect URL inside the HTML/JS
            html = resp.text
            m = re.search(r'https://[^"\']+\.pdf[^"\']*', html)
            if not m:
                # Sometimes it is window.location = ("...")
                m = re.search(r'window\.location\s*=\s*\(\s*["\'](https://[^"\']+)["\']', html)
            if m:
                pdf_url = m.group(1) if m.lastindex else m.group(0)
            else:
                result["Notes"] = "Could not find PDF URL in token page"
                return result

            # Extract report number from the S3 path if possible
            num_match = re.search(r"/(\d{8,13})\.pdf", pdf_url)
            if num_match:
                result["Certificate Number"] = num_match.group(1)

            # Step 2: Download the actual PDF
            pdf_resp = requests.get(pdf_url, headers=headers, timeout=30)
            if pdf_resp.status_code != 200:
                result["Notes"] = f"PDF download failed (HTTP {pdf_resp.status_code})"
                return result
            pdf_bytes = pdf_resp.content

        # Step 3: Extract text from PDF (reuse the same robust extractor)
        text = extract_text_from_pdf(pdf_bytes)
        if not text or len(text.strip()) < 80:
            result["Status"] = "Parse Error"
            result["Notes"] = "Downloaded PDF but could not extract text"
            return result

        text_upper = text.upper()

        # GIA reports use many dots between label and value, e.g.:
        # Shape and Cutting Style ............................................ Round Brilliant
        # Carat Weight ........................................................................ 0.50 carat
        # Color Grade ........................................................................................ G
        # Clarity Grade ......................................................................................SI1

        # ---------- Report Number ----------
        if result.get("Certificate Number") in ("Unknown", "", None):
            m = re.search(r"GIA REPORT NUMBER[\s.]+(\d{8,13})", text_upper)
            if m:
                result["Certificate Number"] = m.group(1)

        # ---------- Shape ----------
        m = re.search(r"SHAPE AND CUTTING STYLE[\s.]+([A-Z][A-Z\s\-]+?)(?:\n|MEASUREMENTS|$)", text_upper)
        if not m:
            m = re.search(r"(ROUND BRILLIANT|PRINCESS CUT|CUSHION MODIFIED BRILLIANT|CUSHION BRILLIANT|"
                          r"OVAL BRILLIANT|PEAR BRILLIANT|MARQUISE BRILLIANT|EMERALD CUT|"
                          r"SQUARE EMERALD CUT|RADIANT CUT|HEART BRILLIANT|ASSCHER CUT)", text_upper)
        if m:
            result["SHAPE AND CUT"] = re.sub(r'\s+', ' ', m.group(1)).strip().title()

        # ---------- Measurements ----------
        m = re.search(r"MEASUREMENTS[\s.]+([\d\.\s\-X]+)\s*MM", text_upper)
        if m:
            size = m.group(1).strip().replace("X", "×")
            size = re.sub(r'\s+', ' ', size) + " mm"
            result["MEASUREMENTS"] = size

        # ---------- Carat Weight ----------
        m = re.search(r"CARAT WEIGHT[\s.]+([\d\.]+)\s*CARAT", text_upper)
        if m:
            result["CARAT WEIGHT"] = f"{m.group(1)} ct"

        # ---------- Color ----------
        m = re.search(r"COLOR GRADE[\s.]+([D-Z])\b", text_upper)
        if m:
            result["COLOR GRADE"] = m.group(1)

        # ---------- Clarity ----------
        m = re.search(r"CLARITY GRADE[\s.]+((?:FL|IF|VVS\s*[12]|VS\s*[12]|SI\s*[12]|I\s*[123]))", text_upper)
        if m:
            result["CLARITY GRADE"] = re.sub(r'\s+', '', m.group(1)).upper()

        # Cut / Polish / Symmetry / Fluorescence (GIA)
        m = re.search(r"CUT GRADE[\s.]+(EXCELLENT|VERY GOOD|GOOD|FAIR|POOR)", text_upper)
        if m:
            result["CUT"] = m.group(1).title()
        m = re.search(r"POLISH[\s.]+(EXCELLENT|VERY GOOD|GOOD|FAIR|POOR)", text_upper)
        if m:
            result["POLISH"] = m.group(1).title()
        m = re.search(r"SYMMETRY[\s.]+(EXCELLENT|VERY GOOD|GOOD|FAIR|POOR)", text_upper)
        if m:
            result["SYMMETRY"] = m.group(1).title()
        m = re.search(r"FLUORESCENCE[\s.]+(NONE|FAINT|MEDIUM|STRONG|VERY STRONG)", text_upper)
        if m:
            result["FLUORESCENCE"] = m.group(1).title()

        result["Process"] = "—"
        result["DESCRIPTION"] = "NATURAL DIAMOND"

        filled = sum(1 for k in ["SHAPE AND CUT", "MEASUREMENTS", "CARAT WEIGHT", "COLOR GRADE", "CLARITY GRADE"] if result[k])
        if filled >= 4:
            result["Status"] = "Success (Token PDF)"
            result["Notes"] = f"Extracted from GIA PDF • {filled}/5 fields"
        elif filled >= 2:
            result["Status"] = "Partial"
            result["Notes"] = f"PDF downloaded • {filled}/5 fields found"
        else:
            result["Status"] = "Parse Error"
            result["Notes"] = "PDF downloaded but key fields not found"


    except Exception as e:
        result["Status"] = "Error"
        result["Notes"] = str(e)[:120]

    return result


def fetch_gia_info(report_no: str, try_playwright: bool = False) -> Dict[str, Any]:
    """
    GIA lookup.
    By default only returns a clean link.
    If try_playwright=True, attempts browser extraction (often fails due to GIA protection).
    """
    result = empty_inventory_row(report_no, "GIA")
    result["Process"] = "—"
    result["DESCRIPTION"] = "NATURAL DIAMOND"
    result["CERTIFICATE LINK"] = f"https://www.gia.edu/report-check?reportno={report_no}"
    result["Status"] = "Link Ready"
    result["Notes"] = "Open laboratory report check link"

    if not try_playwright or not HAS_PLAYWRIGHT:
        if try_playwright and not HAS_PLAYWRIGHT:
            result["Notes"] = "Open laboratory report check link"
        return result

    url = f"https://www.gia.edu/report-check?reportno={report_no}"

    try:
        with sync_playwright() as p:
            # Hardened launch args to avoid HTTP/2 and anti-bot issues
            browser = p.chromium.launch(
                headless=True,
                args=[
                    "--disable-http2",
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-web-security",
                    "--ignore-certificate-errors",
                ]
            )
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
                viewport={"width": 1280, "height": 900},
                locale="en-US",
                java_script_enabled=True,
            )
            # Hide webdriver flag
            page = context.new_page()
            page.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            """)

            # Use domcontentloaded instead of networkidle (more reliable)
            page.goto(url, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(3500)  # allow JS to render

            content = page.inner_text("body")
            browser.close()

        if not content or len(content) < 100:
            raise Exception("Empty page content")

        text = content.upper()

        # ---------- Shape ----------
        shape_match = re.search(
            r"(ROUND|PRINCESS|CUSHION|OVAL|PEAR|MARQUISE|EMERALD|RADIANT|HEART|ASSCHER)"
            r"(?:\s+(?:BRILLIANT|CUT|MODIFIED|MODIFIED BRILLIANT))?",
            text
        )
        if shape_match:
            result["Shape"] = shape_match.group(0).title().strip()

        # ---------- Measurements ----------
        meas = re.search(r"([\d\.]+\s*[xX×]\s*[\d\.]+\s*[xX×]\s*[\d\.]+\s*mm)", text, re.IGNORECASE)
        if not meas:
            meas = re.search(r"([\d\.]+\s*-\s*[\d\.]+\s*[xX×]\s*[\d\.]+\s*mm)", text, re.IGNORECASE)
        if meas:
            result["MM Size"] = meas.group(1).replace("X", "×").replace("x", "×")

        # ---------- Carat ----------
        carat = re.search(r"([\d\.]+)\s*CARAT", text)
        if carat:
            result["Weight"] = f"{carat.group(1)} ct"

        # ---------- Color ----------
        color = re.search(r"COLOR\s*(?:GRADE)?\s*[:\s]*([D-Z])\b", text)
        if color:
            result["Color"] = color.group(1)

        # ---------- Clarity ----------
        clarity = re.search(r"CLARITY\s*(?:GRADE)?\s*[:\s]*((?:FL|IF|VVS[12]|VS[12]|SI[12]|I[123]))", text)
        if clarity:
            result["Clarity"] = clarity.group(1)

        filled = sum(1 for k in ["Shape", "MM Size", "Weight", "Color", "Clarity"] if result[k])
        if filled >= 2:
            result["Status"] = "Success (Playwright)"
            result["Notes"] = f"Extracted via browser • {filled}/5 fields"
        else:
            result["Status"] = "Partial / Link"
            result["Notes"] = "Limited data extracted – see links below"

    except Exception as e:
        err_msg = str(e)
        # Shorten common long errors
        if "ERR_HTTP2" in err_msg or "PROTOCOL_ERROR" in err_msg:
            short = "HTTP/2 protocol error (common with GIA)"
        elif "Timeout" in err_msg:
            short = "Page load timeout"
        else:
            short = err_msg[:70]
        result["Status"] = "Link Ready"
        result["Notes"] = f"Browser failed ({short}) – use links below"

    return result


# -------------------------------------------------
# Streamlit UI
# -------------------------------------------------

st.title("Certified Stone Inventory")
st.caption("Centralized inventory · Certificate lookup · Media reference")

# Sidebar
with st.sidebar:
    st.subheader("Summary")
    _stats_df = load_all_inventory()
    st.metric("Total certificates", len(_stats_df))
    _today = date.today().isoformat()
    _added_today = 0
    if not _stats_df.empty and "date_added" in _stats_df.columns:
        _added_today = int(_stats_df["date_added"].astype(str).str.startswith(_today).sum())
    st.metric("Added today", _added_today)
    st.divider()
    use_playwright = False  # reserved for internal GIA automation
    if use_gsheets():
        st.caption("Connected data store")
    else:
        st.caption("Local data store")

# ---------- Tabs ----------
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
    report_text = st.text_area(
        "Certificate numbers",
        value="",
        height=160,
        placeholder="Enter one certificate number per line (IGI or GIA)",
    )
    col1, col2 = st.columns([1, 1])
    with col1:
        fetch_btn = st.button("Fetch & save", type="primary", use_container_width=True)
    with col2:
        clear_btn = st.button("Clear", use_container_width=True)

    if clear_btn:
        st.rerun()

    if fetch_btn and report_text.strip():
        raw_lines = [line.strip() for line in report_text.strip().splitlines() if line.strip()]
        seen = set()
        items_to_process = []
        for line in raw_lines:
            if is_gia_token_url(line):
                key = line
                if key not in seen:
                    seen.add(key)
                    items_to_process.append(("token", line))
            else:
                cleaned = clean_report_number(line)
                if cleaned and cleaned not in seen:
                    seen.add(cleaned)
                    items_to_process.append(("number", cleaned))

        progress = st.progress(0)
        status_text = st.empty()
        results = []

        for i, (item_type, value) in enumerate(items_to_process):
            display = value if len(value) < 60 else value[:50] + "..."
            status_text.text(f"Retrieving {display} ({i+1}/{len(items_to_process)})")
            progress.progress((i + 1) / len(items_to_process))

            if item_type == "token":
                data = fetch_gia_from_token(value)
            elif is_igi(value):
                data = fetch_igi_report(value)
            elif is_gia(value):
                data = fetch_gia_info(value, try_playwright=use_playwright)
            else:
                data = empty_inventory_row(value, "Unknown")
                data["Status"] = "Unknown Format"
                data["Notes"] = "Could not determine if IGI or GIA"
            results.append(data)
            time.sleep(0.4)

        progress.empty()
        status_text.empty()

        # Save rows to DB (including Not Found so user can fill grades later)
        to_save = [r for r in results if r.get("Status") not in ("Unknown Format", "Error")]
        if to_save:
            upsert_rows(to_save)
            st.success(f"{len(to_save)} record(s) saved to inventory.")

        df_new = pd.DataFrame(results)
        inv_cols = [c for c in INVENTORY_COLUMNS if c in df_new.columns]
        df_new = df_new[inv_cols]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Processed", len(df_new))
        m2.metric("IGI", len(df_new[df_new["Lab"] == "IGI"]) if "Lab" in df_new.columns else 0)
        m3.metric("GIA", len(df_new[df_new["Lab"] == "GIA"]) if "Lab" in df_new.columns else 0)
        m4.metric("Added today", count_added_today())
        try:
            ev_new = st.dataframe(
                df_new, use_container_width=True, hide_index=True,
                on_select="rerun", selection_mode="single-row", key="batch_table_select",
            )
            sel_n = ev_new.selection.rows if ev_new and ev_new.selection else []
            if sel_n and sel_n[0] < len(df_new):
                crow = df_new.iloc[sel_n[0]]
                open_certificate_preview(
                    str(crow.get("Certificate Number", "")),
                    str(crow.get("CERTIFICATE LINK", "") or ""),
                )
        except TypeError:
            st.dataframe(df_new, use_container_width=True, hide_index=True)

        # Download this batch
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            export = df_new.drop(columns=[c for c in ["Lab", "Status", "Notes", "date_added", "last_updated"] if c in df_new.columns], errors="ignore")
            export.to_excel(writer, index=False, sheet_name="Batch")
        st.download_button(
            "Export Excel",
            data=buffer.getvalue(),
            file_name=f"Batch_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# =========================================================
# TAB 2 – All Inventory
# =========================================================
with tab_all:
    df_all = load_all_inventory()

    # Compact toolbar: search + filters on one row
    t1, t2, t3 = st.columns([3, 1, 1])
    with t1:
        search_q = st.text_input(
            "Search",
            value="",
            placeholder="Search by certificate, shape, color, clarity, carat…",
            label_visibility="collapsed",
            key="inventory_search",
        )
    with t2:
        filter_lab = st.selectbox("Lab", ["All", "IGI", "GIA", "Unknown"], key="lab_filter", label_visibility="collapsed")
    with t3:
        filter_today = st.checkbox("Today only", value=False, key="today_filter")

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
                    return abs(num - target) <= 0.12
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

    view = df_all.copy()
    if filter_lab != "All" and not view.empty and "Lab" in view.columns:
        view = view[view["Lab"] == filter_lab]
    if filter_today and not view.empty and "date_added" in view.columns:
        today_str = date.today().isoformat()
        view = view[view["date_added"].astype(str).str.startswith(today_str)]
    if search_q.strip() and not view.empty:
        view = smart_filter(view, search_q)

    render_inventory_subtotals(view, title=f"{len(view)}/{len(df_all)} shown")
    if view.empty:
        st.info("No matching records.")
    else:
        try:
            event = st.dataframe(
                view,
                use_container_width=True,
                hide_index=True,
                height=520,
                on_select="rerun",
                selection_mode="single-row",
                key="inventory_table_select",
            )
            sel_rows = event.selection.rows if event and event.selection else []
            if sel_rows:
                ridx = sel_rows[0]
                if ridx < len(view):
                    crow = view.iloc[ridx]
                    open_certificate_preview(
                        str(crow.get("Certificate Number", "")),
                        str(crow.get("CERTIFICATE LINK", "") or ""),
                    )
        except TypeError:
            # Older Streamlit without on_select
            st.dataframe(view, use_container_width=True, hide_index=True, height=520)
            pick = st.selectbox(
                "Or pick certificate to preview",
                view["Certificate Number"].astype(str).tolist(),
                key="inv_preview_pick",
            )
            if st.button("Preview selected", key="inv_preview_btn"):
                link = ""
                m = view[view["Certificate Number"].astype(str) == pick]
                if not m.empty:
                    link = str(m.iloc[0].get("CERTIFICATE LINK") or "")
                open_certificate_preview(pick, link)

        # Inline fallback panel when dialog is unavailable
        if not HAS_DIALOG and st.session_state.get("inline_preview_cert"):
            with st.expander(f"Preview: {st.session_state['inline_preview_cert']}", expanded=True):
                render_certificate_preview(
                    st.session_state["inline_preview_cert"],
                    st.session_state.get("inline_preview_link", ""),
                )

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            export = view.drop(columns=[c for c in ["Lab", "Status", "Notes"] if c in view.columns], errors="ignore")
            export.to_excel(writer, index=False, sheet_name="Inventory")
        st.download_button(
            "Export Excel",
            data=buffer.getvalue(),
            file_name="Filtered_Inventory.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# =========================================================
# TAB 3 – Edit editable fields
# =========================================================
with tab_edit:
    df_all = load_all_inventory()
    if df_all.empty:
        st.info("No inventory records available.")
    else:
        # Single editable table only (no second grid)
        edit_cols = [
            "Certificate Number", "SHAPE AND CUT", "CARAT WEIGHT",
            "COLOR GRADE", "CLARITY GRADE",
        ] + EDITABLE_COLUMNS
        edit_cols = [c for c in edit_cols if c in df_all.columns]
        df_edit = df_all[edit_cols].copy()

        b1, b2, b3 = st.columns([1, 3, 1])
        with b1:
            save_clicked = st.button("Save", type="primary", use_container_width=True)
        with b2:
            pick_e = st.selectbox(
                "Preview",
                df_edit["Certificate Number"].astype(str).tolist(),
                key="edit_preview_pick",
                label_visibility="collapsed",
            )
        with b3:
            preview_clicked = st.button("Preview", use_container_width=True, key="edit_preview_btn")

        edited = st.data_editor(
            df_edit,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            height=560,
            key="inventory_editor",
            column_config={
                "Certificate Number": st.column_config.TextColumn(disabled=True, width="small"),
                "SHAPE AND CUT": st.column_config.TextColumn(disabled=True, width="medium"),
                "CARAT WEIGHT": st.column_config.TextColumn(disabled=True, width="small"),
                "COLOR GRADE": st.column_config.TextColumn(disabled=True, width="small"),
                "CLARITY GRADE": st.column_config.TextColumn(disabled=True, width="small"),
            },
        )

        if save_clicked:
            update_editable_fields(edited)
            st.success("Changes saved.")
            st.rerun()

        if preview_clicked:
            open_certificate_preview(pick_e, "")

        if not HAS_DIALOG and st.session_state.get("inline_preview_cert"):
            with st.expander(f"Certificate {st.session_state['inline_preview_cert']}", expanded=True):
                render_certificate_preview(
                    st.session_state["inline_preview_cert"],
                    st.session_state.get("inline_preview_link", ""),
                )

# =========================================================
# TAB 4 – Video links & Drive backup
# =========================================================
with tab_video:
    video_text = st.text_area(
        "Certificate number and media URL",
        value="",
        height=180,
        placeholder="One pair per line: certificate number, then URL (tab or space separated)",
        key="video_bulk_input",
    )
    try_backup = st.checkbox("Upload direct video files to linked storage", value=False)

    if st.button("Save media links", type="primary", key="save_videos_btn"):
        pairs = []
        for line in video_text.strip().splitlines():
            line = line.strip()
            if not line:
                continue
            parts = re.split(r"\s+", line, maxsplit=1)
            if len(parts) < 2:
                continue
            pairs.append((parts[0], parts[1]))

        if not pairs:
            st.error("No valid certificate / URL pairs found.")
        else:
            with st.spinner("Saving…"):
                results = save_video_urls(pairs, try_drive_backup=try_backup)
            st.success(f"{len(results)} media link(s) saved.")
            st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)

st.divider()
st.caption("Certified Stone Inventory · FGI")
