"""
IGI + GIA Diamond Report Live Lookup - Streamlit App
----------------------------------------------------
Features:
- Paste multiple report numbers (one per line)
- Auto-detects IGI (LG...) vs GIA (numeric)
- Live fetches IGI reports from official PDF
- Robust multi-method PDF text extraction + many parsing fallbacks
- GIA: tries Playwright browser automation to extract data, falls back to link
- Displays results in interactive table
- Download as Excel / CSV
- Clean, professional UI

How to run:
    pip install streamlit pandas openpyxl requests pdfplumber pypdf playwright
    playwright install chromium
    streamlit run igi_gia_report_lookup.py
"""

import streamlit as st
import pandas as pd
import requests
import re
import io
from typing import Optional, Dict, Any
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
    page_title="IGI + GIA Report Lookup",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded"
)

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


def fetch_igi_report(report_no: str) -> Dict[str, Any]:
    """
    Fetch live IGI Laboratory Grown Diamond report from official PDF.
    Uses multiple extraction + parsing fallbacks for maximum reliability.
    """
    number_only = re.sub(r'^LG', '', report_no, flags=re.IGNORECASE)
    pdf_url = f"https://pdf.igi.org/FDR{number_only}.pdf"

    result = {
        "Report Number": report_no if report_no.upper().startswith("LG") else f"LG{number_only}",
        "Lab": "IGI",
        "Shape": None,
        "MM Size": None,
        "Weight": None,
        "Color": None,
        "Clarity": None,
        "Process": None,
        "Status": "Error",
        "Notes": ""
    }

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "application/pdf,*/*"
        }
        resp = requests.get(pdf_url, headers=headers, timeout=25)

        if resp.status_code == 404:
            result["Status"] = "Not Found"
            result["Notes"] = "PDF not available on IGI website"
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
        # Normalize common OCR / extraction noise
        text_upper = text_upper.replace("×", "X").replace("–", "-").replace("—", "-")

        # ---------- SHAPE ----------
        shape_patterns = [
            r"SHAPE AND CUTTING STYLE\s*([A-Z0-9\s\-]+?)(?:\n|MEASUREMENTS|GRADING)",
            r"SHAPE AND CUTTING STYLE\s*([^\n]{5,60})",
            r"(ROUND BRILLIANT|PRINCESS CUT|CUSHION MODIFIED BRILLIANT|CUSHION BRILLIANT|"
            r"OVAL BRILLIANT|PEAR BRILLIANT|MARQUISE BRILLIANT|EMERALD CUT|"
            r"SQUARE EMERALD CUT|RADIANT CUT|HEART BRILLIANT|ASSCHER CUT|"
            r"TRIANGLE BRILLIANT|TRILLIANT)",
        ]
        for pat in shape_patterns:
            m = re.search(pat, text_upper)
            if m:
                shape = m.group(1).strip()
                shape = re.sub(r'\s+', ' ', shape)
                # Remove trailing garbage
                shape = re.split(r'\s{2,}|\d|MM|CARAT', shape)[0].strip()
                if 4 < len(shape) < 50:
                    result["Shape"] = shape.title()
                    break

        # ---------- MEASUREMENTS (MM Size) ----------
        meas_patterns = [
            r"MEASUREMENTS?\s*([\d\.\-\sX]+?\s*MM)",
            r"([\d\.]+\s*[-–]\s*[\d\.]+\s*[Xx×]\s*[\d\.]+\s*MM)",
            r"([\d\.]+\s*[Xx×]\s*[\d\.]+\s*[Xx×]\s*[\d\.]+\s*MM)",
            r"([\d\.]+\s*[-–]\s*[\d\.]+\s*[Xx×]\s*[\d\.]+)",
        ]
        for pat in meas_patterns:
            m = re.search(pat, text_upper)
            if m:
                size = m.group(1).strip()
                size = size.replace("X", "×").replace("x", "×")
                size = re.sub(r'\s+', ' ', size)
                if "MM" not in size.upper():
                    size += " mm"
                else:
                    size = size.replace("MM", "mm")
                result["MM Size"] = size
                break

        # ---------- CARAT WEIGHT ----------
        weight_patterns = [
            r"(\d+\.\d{2})\s*CARATS?",
            r"CARAT WEIGHT\s*(\d+\.\d{2})",
            r"(\d+\.\d{2})\s*CT\b",
        ]
        for pat in weight_patterns:
            m = re.search(pat, text_upper)
            if m:
                result["Weight"] = f"{m.group(1)} ct"
                break

        # ---------- COLOR ----------
        color_patterns = [
            r"COLOR GRADE\s*([D-Z])\b",
            r"COLOR\s*GRADE\s*[:\s]*([D-Z])\b",
            r"\b([D-Z])\s+(?:VS|VVS|SI|IF|FL|IDEAL)",
            r"\bCOLOR\s+([D-Z])\b",
        ]
        for pat in color_patterns:
            m = re.search(pat, text_upper)
            if m:
                result["Color"] = m.group(1)
                break

        # ---------- CLARITY ----------
        clarity_patterns = [
            r"CLARITY GRADE\s*((?:FL|IF|VVS\s*[12]|VS\s*[12]|SI\s*[12]|I\s*[123]))",
            r"CLARITY\s*GRADE\s*[:\s]*((?:FL|IF|VVS\s*[12]|VS\s*[12]|SI\s*[12]|I\s*[123]))",
            r"\b(VVS\s*[12]|VS\s*[12]|SI\s*[12]|IF|FL)\b",
        ]
        for pat in clarity_patterns:
            m = re.search(pat, text_upper)
            if m:
                clar = m.group(1).replace(" ", "").upper()
                result["Clarity"] = clar
                break

        # ---------- PROCESS (CVD / HPHT) ----------
        if re.search(r"CHEMICAL\s+VAPOR\s+DEPOSITION|\bCVD\b", text_upper):
            result["Process"] = "CVD"
        elif re.search(r"HIGH\s+PRESSURE\s+HIGH\s+TEMPERATURE|\bHPHT\b", text_upper):
            result["Process"] = "HPHT"
        else:
            result["Process"] = "Unknown"

        # Final status
        filled = sum(1 for k in ["Shape", "MM Size", "Weight", "Color", "Clarity"] if result[k])
        if filled >= 3:
            result["Status"] = "Success"
            result["Notes"] = f"Fetched live • {filled}/5 core fields"
        elif filled >= 1:
            result["Status"] = "Partial"
            result["Notes"] = f"Partial extraction • {filled}/5 fields found"
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
    """
    result = {
        "Report Number": "Unknown",
        "Lab": "GIA",
        "Shape": None,
        "MM Size": None,
        "Weight": None,
        "Color": None,
        "Clarity": None,
        "Process": "—",
        "Status": "Error",
        "Notes": ""
    }

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
                result["Report Number"] = m.group(1)
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
                result["Report Number"] = num_match.group(1)

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

        # ---------- Shape ----------
        shape_match = re.search(
            r"(ROUND|PRINCESS|CUSHION|OVAL|PEAR|MARQUISE|EMERALD|RADIANT|HEART|ASSCHER)"
            r"(?:\s+(?:BRILLIANT|CUT|MODIFIED|MODIFIED BRILLIANT))?",
            text_upper
        )
        if shape_match:
            result["Shape"] = shape_match.group(0).title().strip()

        # ---------- Measurements ----------
        meas = re.search(r"([\d\.]+\s*[xX×]\s*[\d\.]+\s*[xX×]\s*[\d\.]+\s*mm)", text_upper, re.IGNORECASE)
        if not meas:
            meas = re.search(r"([\d\.]+\s*-\s*[\d\.]+\s*[xX×]\s*[\d\.]+\s*mm)", text_upper, re.IGNORECASE)
        if meas:
            result["MM Size"] = meas.group(1).replace("X", "×").replace("x", "×")

        # ---------- Carat ----------
        carat = re.search(r"([\d\.]+)\s*CARAT", text_upper)
        if carat:
            result["Weight"] = f"{carat.group(1)} ct"

        # ---------- Color ----------
        color = re.search(r"COLOR\s*(?:GRADE)?\s*[:\s]*([D-Z])\b", text_upper)
        if color:
            result["Color"] = color.group(1)

        # ---------- Clarity ----------
        clarity = re.search(r"CLARITY\s*(?:GRADE)?\s*[:\s]*((?:FL|IF|VVS[12]|VS[12]|SI[12]|I[123]))", text_upper)
        if clarity:
            result["Clarity"] = clarity.group(1)

        filled = sum(1 for k in ["Shape", "MM Size", "Weight", "Color", "Clarity"] if result[k])
        if filled >= 2:
            result["Status"] = "Success (Token PDF)"
            result["Notes"] = f"Extracted from GIA PDF token • {filled}/5 fields"
        else:
            result["Status"] = "Partial"
            result["Notes"] = f"PDF downloaded but only {filled}/5 fields found"

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
    result = {
        "Report Number": report_no,
        "Lab": "GIA",
        "Shape": None,
        "MM Size": None,
        "Weight": None,
        "Color": None,
        "Clarity": None,
        "Process": "—",
        "Status": "Link Ready",
        "Notes": "See clickable links below the table"
    }

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

st.title("💎 IGI + GIA Diamond Report Live Lookup")
st.markdown(
    """
    Paste **IGI report numbers**, **GIA report numbers**, or **GIA PDF token links** (one per line).  
    The app will auto-detect the type and extract data where possible.
    """
)

with st.sidebar:
    st.header("About")
    st.markdown(
        """
        **IGI Support**  
        Full live lookup from official PDF reports  
        (Shape, Measurements, Weight, Color, Clarity, Process)

        **GIA Support**  
        Official Report Check links (reliable & fast).  
        Optional experimental browser extraction (often blocked by GIA).
        """
    )

    use_playwright = st.toggle(
        "Try Playwright for GIA (experimental)",
        value=False,
        help="Attempts automatic extraction. Usually times out or gets blocked. Keep OFF for best experience."
    )

    st.markdown(
        """
        **Tips**
        - IGI numbers usually start with `LG`
        - GIA numbers: `GIA# 123...` or just the number
        - GIA PDF tokens: paste the full `https://pdf.gia.edu/?ReportNumber=...` link
        - You can mix everything in the same list
        """
    )
    st.divider()
    st.caption("Built for diamond professionals • Live data from official sources")

# Input area
default_example = """LG833638417
LG727519767
GIA# 6555540885
https://pdf.gia.edu/?ReportNumber=A0B42B94F6E0CCFEC856F5971E9D171A"""
report_text = st.text_area(
    "Report numbers or GIA PDF token links (one per line)",
    value=default_example,
    height=180,
    help="You can mix IGI numbers, GIA numbers, and special GIA PDF token links"
)

col1, col2, col3 = st.columns([1, 1, 2])
with col1:
    fetch_btn = st.button("🔍 Fetch Reports", type="primary", use_container_width=True)
with col2:
    clear_btn = st.button("Clear", use_container_width=True)

if clear_btn:
    st.rerun()

if fetch_btn and report_text.strip():
    raw_lines = [line.strip() for line in report_text.strip().splitlines() if line.strip()]

    # Keep original lines so we can detect token URLs
    seen = set()
    items_to_process = []  # list of (type, value)
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
            data = {
                "Report Number": value,
                "Lab": "Unknown",
                "Shape": None,
                "MM Size": None,
                "Weight": None,
                "Color": None,
                "Clarity": None,
                "Process": None,
                "Status": "Unknown Format",
                "Notes": "Could not determine if IGI or GIA"
            }
        results.append(data)
        time.sleep(0.5)

    progress.empty()
    status_text.empty()

    df = pd.DataFrame(results)

    # Reorder columns
    cols = ["Report Number", "Lab", "Shape", "MM Size", "Weight", "Color", "Clarity", "Process", "Status", "Notes"]
    df = df[[c for c in cols if c in df.columns]]

    st.success(f"Completed • {len(df)} reports processed")

    # Summary metrics
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total", len(df))
    m2.metric("IGI", len(df[df["Lab"] == "IGI"]))
    m3.metric("GIA", len(df[df["Lab"] == "GIA"]))
    m4.metric("Success / Link", len(df[df["Status"].isin(["Success", "Link Ready"])]))

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Notes": st.column_config.TextColumn("Notes", width="large")
        }
    )

    # Download buttons
    st.subheader("Download Results")
    c1, c2 = st.columns(2)

    with c1:
        # Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Reports")
        st.download_button(
            label="📥 Download Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name="IGI_GIA_Reports.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with c2:
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name="IGI_GIA_Reports.csv",
            mime="text/csv",
            use_container_width=True
        )

    # GIA links section – clickable links
    gia_rows = df[df["Lab"] == "GIA"]
    if not gia_rows.empty:
        st.subheader("🔗 GIA Official Report Check Links")
        st.caption("Click any link below to open the official GIA Report Check page in a new tab.")
        for _, row in gia_rows.iterrows():
            url = f"https://www.gia.edu/report-check?reportno={row['Report Number']}"
            st.markdown(f"**{row['Report Number']}** → [Open GIA Report Check]({url})")

else:
    st.markdown(
        """
        ### How it works
        1. Paste your report numbers above (IGI or GIA)
        2. Click **Fetch Reports**
        3. View the live results table
        4. Download Excel / CSV

        **Example IGI numbers** (from previous session):  
        `LG833638417`, `LG727519767`, `LG793621413` …

        **Example GIA number**:  
        `2141438171`
        """
    )

st.divider()
st.caption("Data sourced live from official IGI PDFs and GIA Report Check. Always verify critical results on the official websites.")
