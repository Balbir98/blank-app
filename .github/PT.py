# streamlit_app.py
import io
import re
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st


# ----------------------------
# Helpers
# ----------------------------
def norm_text(x: str) -> str:
    """Normalize text for matching (case/space/punctuation tolerant)."""
    if pd.isna(x):
        return ""
    x = str(x).strip().lower()
    x = re.sub(r"\s+", " ", x)
    x = re.sub(r"[^\w\s&-]", "", x)
    return x


def yyyymm_to_month_year(val) -> str:
    """Convert YYYYMM (e.g., 202605) to 'May 2026'."""
    if pd.isna(val):
        return ""
    s = str(val).strip().replace(".0", "")
    if not re.fullmatch(r"\d{6}", s):
        return s
    dt = datetime.strptime(s, "%Y%m")
    return dt.strftime("%B %Y")


def month_sort_key(val) -> int:
    """Sort key for YYYYMM. Unknown formats go last."""
    if pd.isna(val):
        return 99999999
    s = str(val).strip().replace(".0", "")
    if re.fullmatch(r"\d{6}", s):
        return int(s)
    return 99999999


def first_name(full_name: str) -> str:
    if not full_name:
        return ""
    parts = str(full_name).strip().split()
    return parts[0] if parts else ""


def build_email_body_html(broker_first_name: str, lender_name: str, month_lines: list[str]) -> str:
    """
    HTML email body:
    - Segoe UI font
    - Month lines bold + red
    - Specific sentence bold
    - Objection lines (in quotes) bold; explanation lines normal
    """
    # Month lines: bold + red
    month_html = "".join(
        [f'<div style="font-weight:700; color:#C00000;">{ml}</div>' for ml in month_lines]
    )

    return f"""
<html>
  <body style="font-family: 'Segoe UI', SegoeUI, Arial, sans-serif; font-size: 11pt;">
    <p>Hi {broker_first_name},</p>

    <p>Hope you’re well?</p>

    <p>
      Great news - I’ve received some useful data from {lender_name} regarding your upcoming renewals;
      <strong>the following number of potential
      product transfers are due with {lender_name} in the coming months and a new rate can now be secured:</strong>
    </p>

    <p>
      {month_html}
    </p>

    <p>All you need to do is call your {lender_name} Business Development Manager for them to confirm the client name(s).</p>

    <p>
      I wanted to share this data with you, so you can plan your customer conversations with plenty of time, and share some tips for
      customer resistance (as we always hear this!):
    </p>

    <p><strong>• “I’m waiting to see what rates do.”</strong><br>
    No problem - you’ll monitor the whole market and get alerts if rates drop, so they won’t miss anything.</p>

    <p><strong>• “I don’t have time to look into it.”</strong><br>
    You’ll handle everything and keep them updated. Zero hassle for them.</p>

    <p><strong>• “I’ll think about it in the New Year.”</strong><br>
    January gets busy and rates can move - locking in a plan now gives them clarity and avoids last-minute stress.</p>

    <p><strong>• “I just want to stay with my current lender.”</strong><br>
    Perfect - you can check their current options and compare the whole market, so they know they’re getting the best outcome.</p>

    <p>If you need anything else or want to talk through your approach, feel free to reach out.</p>
  </body>
</html>
""".strip()


def make_eml_outlook_draft(to_email: str, subject: str, html_body: str) -> bytes:
    """
    Outlook-friendly unsent draft (.eml):
    - X-Unsent: 1 => opens in compose mode (unsent draft)
    - HTML body for formatting
    NOTE: This does NOT send emails.
    """
    to_email = "" if to_email is None else str(to_email).strip()

    msg = (
        "X-Unsent: 1\n"
        f"To: {to_email}\n"
        f"Subject: {subject}\n"
        "MIME-Version: 1.0\n"
        'Content-Type: text/html; charset="utf-8"\n'
        "Content-Transfer-Encoding: 8bit\n"
        "\n"
        f"{html_body}\n"
    )
    return msg.encode("utf-8")


def read_any(file) -> pd.DataFrame:
    name = (file.name or "").lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)


# ----------------------------
# Lender config (easy to extend)
# ----------------------------
LENDER_CONFIG = {
    "Santander": {
        "lender_required_cols": ["Broker Name", "Firm", "Maturity Month", "Volume"],
        "subject": "Santander upcoming product transfers",
        "display_name": "Santander",
    },
    # Add future lenders here...
}


# ----------------------------
# UI
# ----------------------------
st.set_page_config(page_title="PT Communications – Draft Generator", layout="wide")
st.title("PT Communications – Email Draft Generator")
st.caption("Creates Outlook-friendly **unsent draft .eml files** in a ZIP. Nothing is sent automatically.")

lender_name = st.selectbox("Step 1 — Select Provider", list(LENDER_CONFIG.keys()))
st.divider()

st.subheader(f"Step 2 — Upload files for {lender_name}")

col1, col2 = st.columns(2)
with col1:
    lender_file = st.file_uploader(
        f"Upload {lender_name} data",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=False,
    )
with col2:
    zoho_file = st.file_uploader(
        "Upload Zoho data",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=False,
    )

st.divider()

if lender_file and zoho_file:
    config = LENDER_CONFIG[lender_name]

    # Load
    try:
        df_lender = read_any(lender_file)
        df_zoho = read_any(zoho_file)
    except Exception as e:
        st.error(f"Error reading files: {e}")
        st.stop()

    # Validate columns
    lender_required = config["lender_required_cols"]
    zoho_required = ["Full Name", "AR Firm Name", "Email (AR Active advisers)"]

    missing_lender = [c for c in lender_required if c not in df_lender.columns]
    missing_zoho = [c for c in zoho_required if c not in df_zoho.columns]

    if missing_lender:
        st.error(f"{lender_name} file missing columns: {missing_lender}")
        st.stop()
    if missing_zoho:
        st.error(f"Zoho file missing columns: {missing_zoho}")
        st.stop()

    # Normalize keys for matching
    df_lender = df_lender.copy()
    df_zoho = df_zoho.copy()

    df_lender["__broker_key"] = df_lender["Broker Name"].map(norm_text)
    df_lender["__firm_key"] = df_lender["Firm"].map(norm_text)

    df_zoho["__broker_key"] = df_zoho["Full Name"].map(norm_text)
    df_zoho["__firm_key"] = df_zoho["AR Firm Name"].map(norm_text)

    df_zoho["__email"] = df_zoho["Email (AR Active advisers)"].astype(str).str.strip()
    df_zoho = df_zoho[df_zoho["__email"].str.contains(r"@", na=False)].copy()

    email_lookup = (
        df_zoho.drop_duplicates(subset=["__broker_key", "__firm_key"])
        .set_index(["__broker_key", "__firm_key"])["__email"]
        .to_dict()
    )

    # Ensure numeric volume
    df_lender["Volume"] = pd.to_numeric(df_lender["Volume"], errors="coerce").fillna(0).astype(int)

    # Aggregate per broker+firm+month (handles multiple rows correctly)
    grouped = (
        df_lender.groupby(
            ["Broker Name", "Firm", "__broker_key", "__firm_key", "Maturity Month"],
            dropna=False,
        )["Volume"]
        .sum()
        .reset_index()
    )

    # Sort months chronologically
    grouped["__month_sort"] = grouped["Maturity Month"].map(month_sort_key)
    grouped = grouped.sort_values(["Broker Name", "Firm", "__month_sort"], ascending=True)

    broker_groups = grouped.groupby(["Broker Name", "Firm", "__broker_key", "__firm_key"], dropna=False)

    # Generate ZIP of drafts
    subject = config["subject"]
    lender_display = config.get("display_name", lender_name)

    zip_buffer = io.BytesIO()
    created = 0
    unmatched = 0

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for (broker_name, firm, broker_key, firm_key), sub in broker_groups:
            # If lookup fails, To stays blank (still create the draft)
            to_email = email_lookup.get((broker_key, firm_key), "")
            if not to_email:
                unmatched += 1

            # Build ALL month lines for this broker
            month_lines = []
            for _, r in sub.iterrows():
                vol = int(r["Volume"])
                if vol <= 0:
                    continue
                month_label = yyyymm_to_month_year(r["Maturity Month"])
                month_lines.append(f"{vol} in {month_label}")

            # Skip if no positive volumes
            if not month_lines:
                continue

            html_body = build_email_body_html(
                broker_first_name=first_name(broker_name),
                lender_name=lender_display,
                month_lines=month_lines,
            )

            eml_bytes = make_eml_outlook_draft(
                to_email=to_email,
                subject=subject,
                html_body=html_body,
            )

            safe_broker = re.sub(r"[^\w\s-]", "", str(broker_name)).strip().replace(" ", "_")
            safe_firm = re.sub(r"[^\w\s-]", "", str(firm)).strip().replace(" ", "_")
            filename = f"{lender_name}_{safe_broker}_{safe_firm}.eml"

            zf.writestr(filename, eml_bytes)
            created += 1

    zip_buffer.seek(0)

    st.success(f"Created {created} draft(s). ({unmatched} had no email match — To left blank.)")

    st.download_button(
        "Download ZIP of Outlook draft .eml files",
        data=zip_buffer.getvalue(),
        file_name=f"{lender_name}_draft_emails.zip",
        mime="application/zip",
    )

else:
    st.info("Select a provider and upload both files to generate draft emails.")
