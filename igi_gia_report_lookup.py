"""
IGI + GIA Diamond Inventory Manager - Streamlit App
----------------------------------------------------
Features:
- Live fetch IGI / GIA reports
- Full inventory column format
- SQLite database – keeps growing as you add certificates
- Tabs: Add New | All Inventory | Edit Editable Fields
- Track how many certificates added today
- Edit FGI Item #, prices, location, etc. and save

How to run:
    pip install streamlit pandas openpyxl requests pdfplumber pypdf
    streamlit run igi_gia_report_lookup.py
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
    page_title="IGI + GIA Inventory Manager",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------
# Database (SQLite – persists across sessions)
# -------------------------------------------------
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "inventory.db")

INVENTORY_COLUMNS = [
    "Certificate Number", "DESCRIPTION", "SHAPE AND CUT", "CARAT WEIGHT",
    "COLOR GRADE", "CLARITY GRADE", "MEASUREMENTS", "CUT", "POLISH",
    "SYMMETRY", "FLUORESCENCE", "Process",
    "FGI Item Number", "Date Shipped from India", "For stock or Customer",
    "Production Order #, if for stock", "Current Location/Status",
    "Price Per Carat", "Price For stone", "CERTIFICATE LINK", "Comment",
    "Lab", "Status", "Notes", "date_added", "last_updated"
]

EDITABLE_COLUMNS = [
    "FGI Item Number", "Date Shipped from India", "For stock or Customer",
    "Production Order #, if for stock", "Current Location/Status",
    "Price Per Carat", "Price For stone", "Comment"
]


def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
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
        "Lab": d.get("lab"),
        "Status": d.get("status"),
        "Notes": d.get("notes"),
        "date_added": d.get("date_added"),
        "last_updated": d.get("last_updated"),
    }


def upsert_rows(rows: List[Dict[str, Any]]):
    """Insert or update rows. Preserves existing editable fields on update."""
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
            # Keep user-editable fields that already have values
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
            d["date_added"] = date_added  # keep original add date

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


def load_all_inventory() -> pd.DataFrame:
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM inventory ORDER BY date_added DESC", conn)
    conn.close()
    if df.empty:
        return pd.DataFrame(columns=INVENTORY_COLUMNS)
    rows = [db_dict_to_row(r) for r in df.to_dict(orient="records")]
    return pd.DataFrame(rows)


def update_editable_fields(df_edit: pd.DataFrame):
    """Save only the editable columns for existing certificates."""
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
    conn = get_conn()
    n = conn.execute(
        "SELECT COUNT(*) FROM inventory WHERE date_added LIKE ?",
        (f"{today}%",)
    ).fetchone()[0]
    conn.close()
    return n


def total_count() -> int:
    conn = get_conn()
    n = conn.execute("SELECT COUNT(*) FROM inventory").fetchone()[0]
    conn.close()
    return n


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
    return report_no.startswith("LG")


def is_gia(report_no: str) -> bool:
    """
    Detect GIA-style report.
    Accepts pure numeric strings of reasonable length (typical GIA reports are 9–12 digits).
    """
    return report_no.isdigit() and 8 <= len(report_no) <= 13


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
        "Lab": lab,
        "Status": "Error",
        "Notes": "",
    }


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
            result["Notes"] = "PDF not public yet – open Verify on igi.org and fill grades in Edit tab"
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
    result["Notes"] = "See clickable links below the table"

    if not try_playwright or not HAS_PLAYWRIGHT:
        if try_playwright and not HAS_PLAYWRIGHT:
            result["Notes"] = "Playwright not installed – see links below"
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

st.title("💎 IGI + GIA Inventory Manager")
st.caption("Live report lookup + persistent inventory database")

# Sidebar
with st.sidebar:
    st.header("📊 Inventory Stats")
    st.metric("Total certificates", total_count())
    st.metric("Added today", count_added_today())
    st.divider()
    st.header("Settings")
    use_playwright = st.toggle(
        "Try Playwright for GIA (experimental)",
        value=False,
        help="Usually blocked by GIA. Keep OFF."
    )
    st.markdown(
        """
        **Tips**
        - IGI: `LG831649004`
        - GIA number or `GIA# 123...`
        - GIA token: full `pdf.gia.edu` link
        - Editable fields are saved when you click **Save**
        """
    )
    st.divider()
    st.caption("SQLite database: inventory.db")

# ---------- Tabs ----------
tab_add, tab_all, tab_edit = st.tabs([
    "➕ Add New Certificates",
    "📋 All Inventory",
    "✏️ Edit Editable Fields"
])

# =========================================================
# TAB 1 – Add new certificates
# =========================================================
with tab_add:
    st.subheader("Fetch & add certificates")
    default_example = """LG831649004
LG833663017
LG834619611"""
    report_text = st.text_area(
        "Report numbers or GIA PDF token links (one per line)",
        value=default_example,
        height=160,
        help="Mix IGI numbers, GIA numbers, and special GIA PDF token links"
    )
    col1, col2 = st.columns([1, 1])
    with col1:
        fetch_btn = st.button("🔍 Fetch & Save to Inventory", type="primary", use_container_width=True)
    with col2:
        clear_btn = st.button("Clear input", use_container_width=True)

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

        st.info(f"Processing **{len(items_to_process)}** unique item(s)...")
        progress = st.progress(0)
        status_text = st.empty()
        results = []

        for i, (item_type, value) in enumerate(items_to_process):
            display = value if len(value) < 60 else value[:50] + "..."
            status_text.text(f"Looking up {display} ({i+1}/{len(items_to_process)})...")
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
            not_found = sum(1 for r in to_save if r.get("Status") == "Not Found")
            msg = f"Saved **{len(to_save)}** certificate(s) to inventory database."
            if not_found:
                msg += f" ({not_found} without public PDF – fill grades in Edit tab)"
            st.success(msg)

        df_new = pd.DataFrame(results)
        inv_cols = [c for c in INVENTORY_COLUMNS if c in df_new.columns]
        df_new = df_new[inv_cols]

        st.subheader("Just added / fetched (this session)")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Processed", len(df_new))
        m2.metric("IGI", len(df_new[df_new["Lab"] == "IGI"]) if "Lab" in df_new.columns else 0)
        m3.metric("GIA", len(df_new[df_new["Lab"] == "GIA"]) if "Lab" in df_new.columns else 0)
        m4.metric("Added today (total)", count_added_today())

        st.dataframe(df_new, use_container_width=True, hide_index=True)

        # Download this batch
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            export = df_new.drop(columns=[c for c in ["Lab", "Status", "Notes", "date_added", "last_updated"] if c in df_new.columns], errors="ignore")
            export.to_excel(writer, index=False, sheet_name="Batch")
        st.download_button(
            "📥 Download this batch (Excel)",
            data=buffer.getvalue(),
            file_name=f"Batch_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# =========================================================
# TAB 2 – All Inventory
# =========================================================
with tab_all:
    st.subheader("Full inventory database")
    df_all = load_all_inventory()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total in DB", len(df_all))
    c2.metric("Added today", count_added_today())
    if not df_all.empty and "Lab" in df_all.columns:
        c3.metric("IGI", len(df_all[df_all["Lab"] == "IGI"]))
        c4.metric("GIA", len(df_all[df_all["Lab"] == "GIA"]))

    # ---- Smart search bar ----
    st.markdown("**🔍 Smart search** — e.g. `2 CT Oval VS2` · `E color pear` · `CVD 3ct` · `LG831`")
    search_q = st.text_input(
        "Search inventory",
        value="",
        placeholder="Type shape, carat, color, clarity, process, certificate #…",
        label_visibility="collapsed",
        key="inventory_search",
    )

    f1, f2 = st.columns(2)
    with f1:
        filter_lab = st.selectbox("Filter by Lab", ["All", "IGI", "GIA", "Unknown"], key="lab_filter")
    with f2:
        filter_today = st.checkbox("Show only added today", value=False, key="today_filter")

    def smart_filter(df: pd.DataFrame, query: str) -> pd.DataFrame:
        """Filter inventory using natural-ish phrases like '2 CT Oval VS2' or 'emerald 2ct'."""
        if df.empty or not query or not query.strip():
            return df

        q = query.strip().upper()
        mask = pd.Series([True] * len(df), index=df.index)

        # Certificate number fragment
        cert_hits = re.findall(r"\bLG?\d{6,}\b", q)
        for c in cert_hits:
            col = df["Certificate Number"].astype(str).str.upper()
            mask &= col.str.contains(c, na=False)

        # Carat: "2 CT", "2CT", "2.5 carat", "2ct"
        carat_m = re.search(r"(\d+\.?\d*)\s*(?:CT|CARAT|CTS|CARATS)\b", q)
        if carat_m and "CARAT WEIGHT" in df.columns:
            target = float(carat_m.group(1))
            # Whole number (e.g. 2ct) → match 2.00–2.99; decimal → tight tolerance
            is_whole = abs(target - round(target)) < 1e-9

            def _near_carat(val):
                try:
                    s = str(val).upper()
                    num = float(re.search(r"[\d.]+", s).group())
                    if is_whole:
                        return target <= num < target + 1.0
                    return abs(num - target) <= 0.12
                except Exception:
                    return False

            mask &= df["CARAT WEIGHT"].apply(_near_carat)

        # Shape keywords
        shapes = [
            "ROUND", "OVAL", "PEAR", "MARQUISE", "PRINCESS", "CUSHION",
            "EMERALD", "RADIANT", "HEART", "ASSCHER", "RECTANGULAR", "SQUARE",
            "BRILLIANT", "CUT CORNERED",
        ]
        for shape in shapes:
            if re.search(rf"\b{shape}\b", q) and "SHAPE AND CUT" in df.columns:
                mask &= df["SHAPE AND CUT"].astype(str).str.upper().str.contains(shape, na=False)

        # Color: "E color", "color E", or explicit "COLOR E"
        color_m2 = re.search(r"(?:COLOR|COLOUR)\s*([D-Z])\b|\b([D-Z])\s*(?:COLOR|COLOUR)\b", q)
        if color_m2 and "COLOR GRADE" in df.columns:
            letter = color_m2.group(1) or color_m2.group(2)
            mask &= df["COLOR GRADE"].astype(str).str.upper() == letter

        # Clarity
        clar_m = re.search(
            r"\b(FL|IF|VVS\s*1|VVS\s*2|VVS1|VVS2|VS\s*1|VS\s*2|VS1|VS2|"
            r"SI\s*1|SI\s*2|SI1|SI2|I\s*1|I\s*2|I\s*3|I1|I2|I3)\b",
            q,
        )
        if clar_m and "CLARITY GRADE" in df.columns:
            clar = re.sub(r"\s+", "", clar_m.group(1)).upper()
            mask &= df["CLARITY GRADE"].astype(str).str.upper().str.replace(" ", "", regex=False) == clar

        # Process
        if re.search(r"\bCVD\b", q) and "Process" in df.columns:
            mask &= df["Process"].astype(str).str.upper().str.contains("CVD", na=False)
        if re.search(r"\bHPHT\b", q) and "Process" in df.columns:
            mask &= df["Process"].astype(str).str.upper().str.contains("HPHT", na=False)

        # Free-text tokens (location, FGI item, comments) — skip already-handled tokens
        tokens = re.findall(r"[A-Z0-9.]{2,}", q)
        skip = {
            "CT", "CARAT", "CARATS", "CTS", "SHOW", "ME", "IN", "THE", "WITH", "AND",
            "COLOR", "COLOUR", "CLARITY", "SHAPE", "CUT", "GRADE",
        }
        for tok in tokens:
            if tok in skip:
                continue
            # Skip pure numbers and carat-like tokens (2CT, 2.5CT)
            if re.match(r"^\d+\.?\d*$", tok):
                continue
            if re.match(r"^\d+\.?\d*(CT|CARAT|CTS|CARATS)$", tok):
                continue
            if tok in shapes or tok in ("CVD", "HPHT"):
                continue
            if re.match(r"^(FL|IF|VVS[12]?|VS[12]?|SI[12]?|I[123]?)$", tok):
                continue
            if len(tok) == 1 and tok.isalpha():
                continue
            cols = [
                c for c in [
                    "Certificate Number", "SHAPE AND CUT", "FGI Item Number",
                    "Current Location/Status", "For stock or Customer",
                    "Comment", "DESCRIPTION",
                ] if c in df.columns
            ]
            if cols:
                sub = pd.Series([False] * len(df), index=df.index)
                for c in cols:
                    sub |= df[c].astype(str).str.upper().str.contains(re.escape(tok), na=False)
                mask &= sub

        return df[mask]

    view = df_all.copy()
    if filter_lab != "All" and not view.empty and "Lab" in view.columns:
        view = view[view["Lab"] == filter_lab]
    if filter_today and not view.empty and "date_added" in view.columns:
        today_str = date.today().isoformat()
        view = view[view["date_added"].astype(str).str.startswith(today_str)]
    if search_q.strip() and not view.empty:
        view = smart_filter(view, search_q)

    st.caption(f"Showing **{len(view)}** of **{len(df_all)}** certificates")
    st.dataframe(view, use_container_width=True, hide_index=True)

    if not view.empty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            export = view.drop(columns=[c for c in ["Lab", "Status", "Notes"] if c in view.columns], errors="ignore")
            export.to_excel(writer, index=False, sheet_name="Inventory")
        st.download_button(
            "📥 Download filtered inventory (Excel)",
            data=buffer.getvalue(),
            file_name="Filtered_Inventory.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# =========================================================
# TAB 3 – Edit editable fields
# =========================================================
with tab_edit:
    st.subheader("Edit user fields (FGI Item #, prices, location, etc.)")
    st.caption("Change the yellow columns, then click **Save changes**.")

    df_all = load_all_inventory()
    if df_all.empty:
        st.info("No certificates in the database yet. Add some in the first tab.")
    else:
        # Show only key + editable columns for editing
        edit_cols = ["Certificate Number", "SHAPE AND CUT", "CARAT WEIGHT", "COLOR GRADE", "CLARITY GRADE"] + EDITABLE_COLUMNS
        edit_cols = [c for c in edit_cols if c in df_all.columns]
        df_edit = df_all[edit_cols].copy()

        edited = st.data_editor(
            df_edit,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="inventory_editor",
            column_config={
                "Certificate Number": st.column_config.TextColumn(disabled=True),
                "SHAPE AND CUT": st.column_config.TextColumn(disabled=True),
                "CARAT WEIGHT": st.column_config.TextColumn(disabled=True),
                "COLOR GRADE": st.column_config.TextColumn(disabled=True),
                "CLARITY GRADE": st.column_config.TextColumn(disabled=True),
            }
        )

        if st.button("💾 Save changes to database", type="primary"):
            update_editable_fields(edited)
            st.success("Editable fields saved.")
            st.rerun()

st.divider()
st.caption("Data sourced live from official IGI PDFs and GIA Report Check. Inventory stored in local SQLite (inventory.db).")
