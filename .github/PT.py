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
    if pd.isna(x):
        return ""
    x = str(x).strip().lower()
    x = re.sub(r"\s+", " ", x)
    x = re.sub(r"[^\w\s&-]", "", x)
    return x


def yyyymm_to_month_year(val) -> str:
    if pd.isna(val):
        return ""
    s = str(val).strip().replace(".0", "")
    if not re.fullmatch(r"\d{6}", s):
        return s
    dt = datetime.strptime(s, "%Y%m")
    return dt.strftime("%B %Y")


def month_sort_key(val) -> int:
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
    - Confidentiality classification tag at top
    - Month lines bold + red
    - Key sentence bold
    - Objection lines (in quotes) bold; explanation lines normal
    - Signature + legal disclaimer appended
    """
    month_html = "".join(
        [f'<div style="font-weight:700; color:#C00000; margin:6px 0;">{ml}</div>' for ml in month_lines]
    )

    return f"""
<html>
  <body style="font-family:'Segoe UI', SegoeUI, Arial, sans-serif; font-size:11pt; color:#111;">

    <!-- Classification -->
    <p style="margin:0 0 18px 0;">
      <strong>Classification – <span style="color:#C00000;">Confidential</span></strong>
    </p>

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

    <!-- Signature -->
    <p style="margin-top:24px;">Kind Regards,</p>

    <p style="color:#ED7D31; font-weight:700; margin:0 0 2px 0;">
      Robyn Truslove | Mortgage Development Manager
    </p>

    <p style="margin:0 0 12px 0;">
      The Right Mortgage &amp; Protection Network
    </p>

    <p style="color:#ED7D31; font-weight:700; margin:0 0 12px 0;">
      TRUST. RESPECT. PARTNERSHIP. OPPORTUNITY.
    </p>

    <p style="margin:0 0 12px 0;">
      Phone:&nbsp;01564 732 744<br>
      Web:&nbsp;<a href="https://www.therightmortgage.co.uk">www.therightmortgage.co.uk</a><br>
      Email:&nbsp;<a href="mailto:robyn.truslove@therightmortgage.co.uk">robyn.truslove@therightmortgage.co.uk</a><br>
      St Johns Court, 70 St John’s Close, Knowle, B93 0NH
    </p>

    <!-- Legal -->
    <hr style="border:none; border-top:1px solid #ddd; margin:24px 0;">

    <p style="font-size:9pt; margin:0 0 10px 0;">
      This email and the information it contains may be privileged and/or confidential. It is for the intended addressee(s) only. The unauthorised use, disclosure or copying of this email, or any information it contains is prohibited and could in certain circumstances be a criminal offence. If you are not the intended recipient, please notify <a href="mailto:info@therightmortgage.co.uk">info@therightmortgage.co.uk</a> immediately and delete the message from your system.
    </p>

    <p style="font-size:9pt; margin:0 0 10px 0;">
      Please note that The Right Mortgage does not enter into any form of contract by means of Internet email. None of the staff of The Right Mortgage is authorised to enter into contracts on behalf of the company in this way.  All contracts to which The Right Mortgage is a party are documented by other means.
    </p>

    <p style="font-size:9pt; margin:0 0 10px 0;">
      The Right Mortgage monitors emails to ensure its systems operate effectively and to minimise the risk of viruses. Whilst it has taken reasonable steps to scan this email, it does not accept liability for any virus that it may contain.
    </p>

    <p style="font-size:9pt; margin:0 0 10px 0;">
      Head Office: St Johns Court, 70 St Johns Close, Knowle, Solihull, B93 0NH. Registered in England no. 08130498
    </p>

    <p style="font-size:9pt; margin:0;">
      The Right Mortgage &amp; Protection Network is a trading style of The Right Mortgage Limited, which is authorised and regulated by the Financial Conduct Authority
    </p>

  </body>
</html>
""".strip()


def make_eml_outlook_draft(to_email: str, subject: str, html_body: str) -> bytes:
    """
    Outlook-friendly unsent draft (.eml):
    X-Unsent: 1 => opens in compose mode (unsent draft)
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
# Lender config
# ----------------------------
LENDER_CONFIG = {
    "Santander": {
        "lender_required_cols": ["Broker Name", "Firm", "Maturity Month", "Volume"],
        "subject": "Santander upcoming product transfers",
        "display_name": "Santander",
    },
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
        f"Upload {lender_name} provider data",
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

    # Validate columns (provider/lender file)
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

    # Clean Zoho to rows with valid email only (we still want broker-only lookup though)
    df_zoho_valid = df_zoho[df_zoho["__email"].str.contains(r"@", na=False)].copy()

    # Exact lookup dict: (broker_key, firm_key) -> email
    email_lookup_exact = (
        df_zoho_valid.drop_duplicates(subset=["__broker_key", "__firm_key"])
        .set_index(["__broker_key", "__firm_key"])["__email"]
        .to_dict()
    )

    # Broker-only fallback: broker_key -> first email found for that broker
    broker_only_lookup = (
        df_zoho_valid.drop_duplicates(subset=["__broker_key"])  # keep first per broker
        .set_index(["__broker_key"])["__email"]
        .to_dict()
    )

    # Total unique advisers in provider file (by Broker Name + Firm combination)
    provider_unique_count = df_lender[["Broker Name", "Firm"]].drop_duplicates().shape[0]

    # Ensure Volume numeric
    df_lender["Volume"] = pd.to_numeric(df_lender["Volume"], errors="coerce").fillna(0).astype(int)

    # Aggregate provider rows by broker+firm+month (keeps multiple months)
    grouped = (
        df_lender.groupby(
            ["__broker_key", "__firm_key", "Broker Name", "Firm", "Maturity Month"],
            dropna=False,
        )["Volume"]
        .sum()
        .reset_index()
    )

    # Sort months
    grouped["__month_sort"] = grouped["Maturity Month"].map(month_sort_key)
    grouped = grouped.sort_values(["__broker_key", "__firm_key", "__month_sort"], ascending=True)

    # Build per-broker+firm groups
    broker_groups = grouped.groupby(["__broker_key", "__firm_key", "Broker Name", "Firm"], dropna=False)

    # Prepare ZIP
    subject = config["subject"]
    lender_display = config.get("display_name", lender_name)

    zip_buffer = io.BytesIO()
    manifest = []
    unmatched_list = []

    created = 0
    unmatched_count = 0

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for (broker_key, firm_key, broker_name, firm), sub in broker_groups:
            # Try exact lookup (name + firm)
            to_email = email_lookup_exact.get((broker_key, firm_key), "")
            # Fallback to broker-only
            if not to_email:
                to_email = broker_only_lookup.get(broker_key, "")

            if not to_email:
                unmatched_count += 1
                unmatched_list.append({"Broker Name": broker_name, "Firm": firm})

            # Build month lines for this broker+firm (ALL rows)
            month_lines = []
            for _, r in sub.iterrows():
                vol = int(r["Volume"])
                if vol <= 0:
                    continue
                month_label = yyyymm_to_month_year(r["Maturity Month"])
                month_lines.append(f"{vol} in {month_label}")

            # If no positive volumes, still count as present but skip draft
            if not month_lines:
                continue

            # make html body
            html_body = build_email_body_html(
                broker_first_name=first_name(broker_name),
                lender_name=lender_display,
                month_lines=month_lines,
            )

            eml_bytes = make_eml_outlook_draft(to_email=to_email, subject=subject, html_body=html_body)

            safe_broker = re.sub(r"[^\w\s-]", "", str(broker_name)).strip().replace(" ", "_")
            safe_firm = re.sub(r"[^\w\s-]", "", str(firm)).strip().replace(" ", "_")
            filename = f"{lender_name}_{safe_broker}_{safe_firm}.eml"

            zf.writestr(filename, eml_bytes)
            created += 1

            manifest.append(
                {
                    "Broker Name": broker_name,
                    "Firm": firm,
                    "Email (To)": to_email,
                    "Draft File": filename,
                    "Lines": "; ".join(month_lines),
                }
            )

    zip_buffer.seek(0)

    # Show results & download
    st.subheader("Summary")
    st.markdown(
        f"- Unique advisers (provider file): **{provider_unique_count}**  \n"
        f"- Drafts created: **{created}**  \n"
        f"- Advisers with no matched email (To left blank): **{unmatched_count}**"
    )

    if manifest:
        st.subheader("Draft manifest")
        st.dataframe(pd.DataFrame(manifest), use_container_width=True)

    if unmatched_list:
        st.subheader("Lookup misses (no email found)")
        st.dataframe(pd.DataFrame(unmatched_list).drop_duplicates(), use_container_width=True)

    st.download_button(
        "Download ZIP of Outlook draft .eml files",
        data=zip_buffer.getvalue(),
        file_name=f"{lender_name}_draft_emails.zip",
        mime="application/zip",
    )

else:
    st.info("Select a provider and upload both files to generate draft emails.")
