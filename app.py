"""
Employee Payslip Generator
==========================
Streamlit app that turns a salary spreadsheet (.xlsx or .csv) into one
PDF payslip per employee, plus a single ZIP of all of them.

SECURITY NOTES (read before deploying with real payroll data)
-------------------------------------------------------------
Payslips contain PII and salary figures. Always deploy behind authentication:
  * Streamlit Community Cloud -> use a PRIVATE app with a viewer allow-list.
  * Hugging Face Spaces       -> use a PRIVATE or Protected Space.
  * Google Cloud Run          -> front it with IAM / Identity-Aware Proxy.

This app keeps data in memory only: it never writes the source file or the
generated PDFs to disk, and it does not log cell values. The optional
APP_PASSWORD gate (via st.secrets) is a SECONDARY layer, not a replacement
for platform-level single sign-on.
"""

from __future__ import annotations

import hmac
import io
import re
import zipfile
from datetime import date
from xml.sax.saxutils import escape as xml_escape

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    Image as RLImage,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ------------------------------------------------------------------
# Fonts
# ------------------------------------------------------------------
# reportlab's built-in Helvetica has no glyph for many currency symbols
# (e.g. the Naira sign). Register DejaVu (ships via packages.txt ->
# fonts-dejavu-core on Streamlit Cloud) so symbols render. Falls back to
# Helvetica if the font files are not present.
FONT_REGULAR, FONT_BOLD = "Helvetica", "Helvetica-Bold"


def _register_fonts():
    import os
    global FONT_REGULAR, FONT_BOLD
    reg_paths = ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                 "/usr/share/fonts/dejavu/DejaVuSans.ttf"]
    bold_paths = ["/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                  "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"]
    reg = next((p for p in reg_paths if os.path.exists(p)), None)
    bold = next((p for p in bold_paths if os.path.exists(p)), None)
    if reg and bold:
        try:
            pdfmetrics.registerFont(TTFont("DejaVuSans", reg))
            pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", bold))
            FONT_REGULAR, FONT_BOLD = "DejaVuSans", "DejaVuSans-Bold"
        except Exception:
            pass  # keep Helvetica fallback


_register_fonts()

# ------------------------------------------------------------------
# Page config
# ------------------------------------------------------------------
st.set_page_config(page_title="Payslip Generator", page_icon="🧾", layout="wide")


# ==================================================================
# Optional password gate (secondary layer only)
# ==================================================================
def check_password() -> bool:
    """Return True if access is allowed.

    If no APP_PASSWORD secret is set, the gate is disabled and the app
    relies entirely on platform-level access control (recommended).
    """
    try:
        configured = st.secrets.get("APP_PASSWORD", None)
    except Exception:
        configured = None

    if not configured:
        # On a public deployment this means anyone with the link can use the
        # tool. Warn loudly; setting APP_PASSWORD in Secrets removes this.
        st.warning(
            "No APP_PASSWORD is set, so anyone with the link can use this app. "
            "Add APP_PASSWORD in the app's Secrets to restrict access.",
            icon="⚠️")
        return True

    if st.session_state.get("auth_ok"):
        return True

    st.markdown("#### Restricted tool")
    with st.form("login", clear_on_submit=True):
        pw = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Enter")
    if submitted:
        # Constant-time comparison to avoid timing side-channels.
        if hmac.compare_digest(str(pw), str(configured)):
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


# ==================================================================
# Value helpers
# ==================================================================
def safe_num(value) -> float:
    """Coerce any cell to a float. Blanks, NaN and non-numeric text -> 0.0.

    Also tolerates values like "1,200.50" or "₦3,000" by stripping
    non-numeric characters before parsing.
    """
    if value is None:
        return 0.0
    try:
        if pd.isna(value):
            return 0.0
    except (TypeError, ValueError):
        pass
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[^\d.\-]", "", str(value))
    if cleaned in ("", "-", ".", "-."):
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def is_blank(value) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ""


def sanitize_filename(name: str) -> str:
    name = str(name).strip() or "employee"
    name = re.sub(r"[^\w\-. ]", "", name)  # drop path separators / odd chars
    name = re.sub(r"\s+", "_", name)
    return (name[:80] or "employee")


def p(text: str, style) -> Paragraph:
    """Build a reportlab Paragraph with XML-escaped text (injection-safe)."""
    return Paragraph(xml_escape(str(text)), style)


# ------------------------------------------------------------------
# Amount in words (self-contained, no external dependency)
# ------------------------------------------------------------------
_ONES = ["", "one", "two", "three", "four", "five", "six", "seven", "eight",
         "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen",
         "sixteen", "seventeen", "eighteen", "nineteen"]
_TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy",
         "eighty", "ninety"]
_SCALE = ["", "thousand", "million", "billion", "trillion"]


def _under_thousand(n: int) -> str:
    out = []
    if n >= 100:
        out.append(_ONES[n // 100] + " hundred")
        n %= 100
    if n >= 20:
        out.append(_TENS[n // 10])
        if n % 10:
            out.append(_ONES[n % 10])
    elif n > 0:
        out.append(_ONES[n])
    return " ".join(out)


def number_to_words(amount: float, currency_word: str = "",
                    minor_word: str = "cents") -> str:
    """Convert a monetary amount to words, e.g. 1250.75 ->
    'One thousand two hundred fifty naira and seventy five kobo'."""
    if amount < 0:
        return "minus " + number_to_words(-amount, currency_word, minor_word)

    whole = int(amount)
    minor = int(round((amount - whole) * 100))
    if minor == 100:
        whole += 1
        minor = 0

    if whole == 0:
        words = "zero"
    else:
        parts = []
        n = whole
        while n > 0:
            parts.append(n % 1000)
            n //= 1000
        chunks = []
        for i in range(len(parts) - 1, -1, -1):
            if parts[i]:
                chunk = _under_thousand(parts[i])
                if _SCALE[i]:
                    chunk += " " + _SCALE[i]
                chunks.append(chunk)
        words = " ".join(chunks)

    result = words.strip()
    if currency_word:
        result = f"{result} {currency_word}"
    if minor:
        result += f" and {_under_thousand(minor)} {minor_word}"
    return (result[:1].upper() + result[1:]).strip()


# ==================================================================
# PDF generation
# ==================================================================
def build_payslip_pdf(row, cfg) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=32, leftMargin=32, topMargin=28, bottomMargin=28,
        title="Payslip", author=cfg["company_name"],
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("CompanyTitle", parent=styles["Title"],
                                 fontName=FONT_BOLD, fontSize=18,
                                 alignment=TA_CENTER, spaceAfter=2)
    sub_style = ParagraphStyle("Sub", parent=styles["Normal"], fontSize=9.5,
                               fontName=FONT_REGULAR, alignment=TA_CENTER,
                               textColor=colors.grey, spaceAfter=10)
    label_style = ParagraphStyle("Lbl", parent=styles["Normal"], fontSize=10)
    net_style = ParagraphStyle("Net", parent=styles["Normal"], fontSize=14,
                               fontName=FONT_BOLD, alignment=TA_RIGHT)
    words_style = ParagraphStyle("Words", parent=styles["Normal"], fontSize=9,
                                 textColor=colors.HexColor("#444444"),
                                 alignment=TA_RIGHT, spaceBefore=2)
    foot_style = ParagraphStyle("Foot", parent=styles["Normal"], fontSize=8,
                                textColor=colors.grey, alignment=TA_CENTER,
                                spaceBefore=14)

    sym = cfg["currency_symbol"]
    elements = []

    # --- Header: optional logo + company name ---
    if cfg.get("logo_bytes"):
        try:
            bio = io.BytesIO(cfg["logo_bytes"])
            iw, ih = ImageReader(bio).getSize()
            target_h = 16 * 2.83465  # 16mm in points
            target_w = target_h * (iw / ih) if ih else target_h
            bio.seek(0)
            logo = RLImage(bio, width=target_w, height=target_h)
            logo.hAlign = "CENTER"
            elements.append(logo)
            elements.append(Spacer(1, 4))
        except Exception:
            pass  # a bad logo must never break payslip generation

    elements.append(p(cfg["company_name"], title_style))
    meta = f"Payslip for {cfg['pay_period']}"
    if cfg.get("pay_date_str"):
        meta += f"  |  Pay date: {cfg['pay_date_str']}"
    elements.append(p(meta, sub_style))

    # --- Employee info block ---
    emp_name = "" if is_blank(row.get(cfg["name_col"])) else row.get(cfg["name_col"])
    info_rows = [["Employee Name", str(emp_name)]]
    for col in cfg["info_cols"]:
        val = row.get(col, "")
        info_rows.append([str(col), "" if is_blank(val) else str(val)])

    info_tbl = Table(info_rows, colWidths=[150, 300])
    info_tbl.setStyle(TableStyle([
        ("FONT", (0, 0), (0, -1), FONT_BOLD, 10),
        ("FONT", (1, 0), (1, -1), FONT_REGULAR, 10),
        ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#333333")),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
    ]))
    elements.append(info_tbl)
    elements.append(Spacer(1, 12))

    # --- Earnings table ---
    earn_data = [["Earnings", "Amount"]]
    total_earnings = 0.0
    for col in cfg["earnings_cols"]:
        v = safe_num(row.get(col, 0))
        total_earnings += v
        earn_data.append([str(col), f"{sym}{v:,.2f}"])
    earn_data.append(["Total Earnings", f"{sym}{total_earnings:,.2f}"])

    # --- Deductions table ---
    ded_data = [["Deductions", "Amount"]]
    total_deductions = 0.0
    for col in cfg["deduction_cols"]:
        v = safe_num(row.get(col, 0))
        total_deductions += v
        ded_data.append([str(col), f"{sym}{v:,.2f}"])
    ded_data.append(["Total Deductions", f"{sym}{total_deductions:,.2f}"])

    def style_money_table(tbl):
        tbl.setStyle(TableStyle([
            ("FONT", (0, 0), (-1, -1), FONT_REGULAR, 10),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BBBBBB")),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E4374")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONT", (0, 0), (-1, 0), FONT_BOLD, 10),
            ("FONT", (0, -1), (-1, -1), FONT_BOLD, 10),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EDEFF7")),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2),
             [colors.white, colors.HexColor("#F7F8FB")]),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))

    earn_tbl = Table(earn_data, colWidths=[220, 140])
    ded_tbl = Table(ded_data, colWidths=[220, 140])
    style_money_table(earn_tbl)
    style_money_table(ded_tbl)

    two_col = Table([[earn_tbl, ded_tbl]], colWidths=[370, 370])
    two_col.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    elements.append(two_col)
    elements.append(Spacer(1, 16))

    # --- Net pay ---
    computed_net = total_earnings - total_deductions
    if cfg["net_col"]:
        net = safe_num(row.get(cfg["net_col"], computed_net))
    else:
        net = computed_net

    net_tbl = Table(
        [[p("Net Payment", ParagraphStyle("nl", parent=net_style,
            alignment=TA_LEFT)),
          Paragraph(f"{sym}{net:,.2f}", net_style)]],
        colWidths=[560, 180],
    )
    net_tbl.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (-1, 0), 1, colors.HexColor("#2E4374")),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(net_tbl)
    elements.append(Paragraph(
        number_to_words(net, cfg["currency_word"], cfg["minor_word"]) + " only.",
        words_style))

    if cfg.get("footer_note"):
        elements.append(p(cfg["footer_note"], foot_style))

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


# ==================================================================
# Validation
# ==================================================================
def validate(df, cfg):
    """Return a list of human-readable warnings; never raises."""
    warns = []
    name_col = cfg["name_col"]
    blank_names = int(df[name_col].apply(is_blank).sum())
    if blank_names:
        warns.append(f"{blank_names} row(s) have a blank "
                     f"'{name_col}'. They will be named 'employee'.")

    numeric_cols = cfg["earnings_cols"] + cfg["deduction_cols"]
    if cfg["net_col"]:
        numeric_cols.append(cfg["net_col"])
    for col in numeric_cols:
        coerced = pd.to_numeric(df[col], errors="coerce")
        bad = int(coerced.isna().sum() - df[col].apply(is_blank).sum())
        if bad > 0:
            warns.append(f"Column '{col}' has {bad} non-numeric value(s); "
                         "they will be treated as 0.")

    if cfg["net_col"]:
        mismatches = 0
        for _, r in df.iterrows():
            te = sum(safe_num(r.get(c, 0)) for c in cfg["earnings_cols"])
            td = sum(safe_num(r.get(c, 0)) for c in cfg["deduction_cols"])
            provided = safe_num(r.get(cfg["net_col"], 0))
            if abs((te - td) - provided) > 0.01:
                mismatches += 1
        if mismatches:
            warns.append(f"{mismatches} row(s): the '{cfg['net_col']}' column "
                         "does not equal earnings minus deductions. The "
                         "provided value will be printed as-is.")
    return warns


# ==================================================================
# UI
# ==================================================================
if not check_password():
    st.stop()

st.title("🧾 Employee Payslip Generator")

with st.sidebar:
    st.header("Payslip settings")
    company_name = st.text_input("Company name", value="Your Company Ltd")
    col_a, col_b = st.columns(2)
    with col_a:
        currency_symbol = st.text_input("Currency symbol", value="₦", max_chars=4)
    with col_b:
        pay_date = st.date_input("Pay date", value=date.today())
    currency_word = st.text_input("Currency in words", value="naira")
    minor_word = st.text_input("Minor unit in words", value="kobo")
    pay_period = st.text_input("Pay period", value=date.today().strftime("%B %Y"))
    logo_file = st.file_uploader("Company logo (optional)",
                                 type=["png", "jpg", "jpeg"])
    footer_note = st.text_area(
        "Footer note",
        value="This is a system-generated payslip and is confidential.")
    st.divider()
    if st.button("🧹 Clear session and wipe data"):
        st.session_state.clear()
        st.rerun()

st.subheader("1. Upload salary file")
uploaded = st.file_uploader("Upload an .xlsx or .csv (one row per employee)",
                            type=["xlsx", "csv"])

if not uploaded:
    st.info("Upload a spreadsheet to begin. Data stays in memory and is "
            "never written to disk.")
    st.stop()

# --- Read file defensively ---
try:
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)
except Exception as exc:
    st.error(f"Could not read the file. Check that it is a valid "
             f"spreadsheet. ({type(exc).__name__})")
    st.stop()

df.columns = [str(c).strip() for c in df.columns]
df = df.dropna(how="all")  # drop fully empty rows
if df.empty:
    st.error("The file has no data rows.")
    st.stop()

st.subheader("2. Preview")
st.caption(f"{len(df):,} row(s), {len(df.columns)} column(s).")
st.dataframe(df, height=240)

# --- Column mapping ---
st.subheader("3. Map your columns")
cols = list(df.columns)


def guess_name_col(columns):
    for c in columns:
        if "name" in c.lower():
            return c
    return columns[0]


def is_numeric_col(c):
    return bool(pd.to_numeric(df[c], errors="coerce").notna().any())


name_default = guess_name_col(cols)
name_col = st.selectbox("Employee name column", cols,
                        index=cols.index(name_default))

remaining = [c for c in cols if c != name_col]
earn_default = [c for c in ["Basic", "Housing", "Transport"] if c in remaining]
net_matches = [c for c in remaining
               if c.lower().replace(" ", "") in
               ("netpayment", "netpay", "nettotal", "takehome", "net")]
net_default = net_matches[0] if net_matches else "(none / calculate it)"
auto_ded = [c for c in remaining
            if is_numeric_col(c) and c not in earn_default
            and c not in net_matches]

c1, c2 = st.columns(2)
with c1:
    earnings_cols = st.multiselect("Earnings columns", remaining,
                                   default=earn_default)
with c2:
    deduction_cols = st.multiselect("Deduction columns", remaining,
                                    default=auto_ded)

c3, c4 = st.columns(2)
with c3:
    net_choice = st.selectbox(
        "Net pay column", ["(none / calculate it)"] + remaining,
        index=(["(none / calculate it)"] + remaining).index(net_default)
        if net_default in remaining else 0)
    net_col = None if net_choice.startswith("(none") else net_choice
with c4:
    info_cols = st.multiselect(
        "Extra info to show (ID, title, dept, etc.)",
        [c for c in remaining if c not in earnings_cols
         and c not in deduction_cols and c != net_col],
        default=[])

# --- Conflict check ---
assignments = {}
conflicts = set()
for role, group in (("earnings", earnings_cols), ("deductions", deduction_cols),
                    ("info", info_cols),
                    ("net", [net_col] if net_col else [])):
    for c in group:
        if c in assignments:
            conflicts.add(c)
        assignments[c] = role

if conflicts:
    st.error(f"These column(s) are assigned to more than one role: "
             f"{', '.join(sorted(conflicts))}. Fix the mapping above.")
    st.stop()

if not earnings_cols and not deduction_cols:
    st.warning("Select at least one earnings or deduction column.")
    st.stop()

cfg = {
    "company_name": company_name,
    "currency_symbol": currency_symbol,
    "currency_word": currency_word,
    "minor_word": minor_word,
    "pay_period": pay_period,
    "pay_date_str": pay_date.strftime("%d %b %Y") if pay_date else "",
    "logo_bytes": logo_file.getvalue() if logo_file else None,
    "footer_note": footer_note,
    "name_col": name_col,
    "earnings_cols": earnings_cols,
    "deduction_cols": deduction_cols,
    "info_cols": info_cols,
    "net_col": net_col,
}

# --- Validation report ---
warnings = validate(df, cfg)
if warnings:
    with st.expander(f"⚠️ {len(warnings)} data warning(s) found", expanded=True):
        for w in warnings:
            st.write("• " + w)

# --- Totals summary ---
sum_earn = sum(safe_num(df.loc[i, c]) for i in df.index
               for c in earnings_cols)
sum_ded = sum(safe_num(df.loc[i, c]) for i in df.index
              for c in deduction_cols)
sum_net = sum_earn - sum_ded
m1, m2, m3, m4 = st.columns(4)
m1.metric("Employees", f"{len(df):,}")
m2.metric("Total earnings", f"{currency_symbol}{sum_earn:,.2f}")
m3.metric("Total deductions", f"{currency_symbol}{sum_ded:,.2f}")
m4.metric("Total net", f"{currency_symbol}{sum_net:,.2f}")

# ==================================================================
# Generate
# ==================================================================
st.subheader("4. Generate payslips")
if st.button("Generate payslips", type="primary"):
    progress = st.progress(0.0, text="Generating payslips...")
    results = []  # list of (filename, bytes)
    name_counts = {}
    total = len(df)
    for pos, (_, row) in enumerate(df.iterrows(), start=1):
        base = sanitize_filename(row.get(name_col, "employee"))
        count = name_counts.get(base, 0)
        name_counts[base] = count + 1
        filename = (f"{base}_Payslip.pdf" if count == 0
                    else f"{base}_Payslip_{count + 1}.pdf")
        results.append((filename, build_payslip_pdf(row, cfg)))
        progress.progress(pos / total, text=f"Generated {pos} of {total}...")

    # Build ZIP once, from the already-generated bytes.
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, data in results:
            zf.writestr(fname, data)
    zip_buffer.seek(0)

    st.session_state["payslips"] = results
    st.session_state["payslips_zip"] = zip_buffer.getvalue()
    progress.empty()
    st.success(f"Generated {total} payslip(s).")

# --- Downloads (persist across reruns; generated only once) ---
if st.session_state.get("payslips"):
    results = st.session_state["payslips"]
    st.download_button(
        "⬇️ Download all payslips (ZIP)",
        data=st.session_state["payslips_zip"],
        file_name="Payslips.zip",
        mime="application/zip",
        type="primary",
    )
    with st.expander(f"Individual downloads ({len(results)})"):
        for fname, data in results:
            st.download_button(
                f"Download {fname}", data=data, file_name=fname,
                mime="application/pdf", key=f"dl_{fname}")
