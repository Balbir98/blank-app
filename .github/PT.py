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
    x = re.sub(r"[^\w\s&-]", "", x)  # keep letters/numbers/_ and spaces/&/-
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


def first_name(full_name: str) -> str:
    if not full_name:
        return ""
    parts = str(full_name).strip().split()
    return parts[0] if parts else ""


def build_email_body(broker_first_name: str, lender_name: str, month_lines: list[str]) -> str:
    """Plain-text body formatted to mimic your screenshots."""
    lines = []
    lines.append(f"Hi {broker_first_name},")
    lines.append("")
    lines.append("Hope you’re well?")
    lines.append("")
    lines.append(
        f"Great news - I’ve received some useful data from {lender_name} regarding your upcoming renewals; "
        f"the following number of potential product transfers are due with {lender_name} in the coming months "
        f"and a new rate can now be secured:"
    )
    lines.append("")
    for ml in month_lines:
        lines.append(ml)
    lines.append("")
    lines.append(
        f"All you need to do is call your {lender_name} Business Development Manager for them to confirm the client name(s)."
    )
    lines.append("")
    lines.append(
        "I wanted to share this data with you, so you can plan your customer conversations with plenty of time, "
        "and share some tips for customer resistance (as we always hear this!):"
    )
    lines.append("")
    lines.append("• “I’m waiting to see what rates do.”")
    lines.append("No problem - you’ll monitor the whole market and get alerts if rates drop, so they won’t miss anything.")
    lines.append("• “I don’t have time to look into it.”")
    lines.append("You’ll handle everything and keep them updated. Zero hassle for them.")
    lines.append("• “I’ll think about it in the New Year.”")
    lines.append("January gets busy and rates can move - locking in a plan now gives them clarity and avoids last-minute stress.")
    lines.append("• “I just want to stay with my current lender.”")
    lines.append("Perfect - you can check their current options and compare the whole market, so they know they’re getting the best outcome.")
    lines.append("")
    lines.append("If you need anything else or want to talk through your approach, feel free to reach out.")
    return "\n".join(lines)


def make_eml(to_email: str, subject: str, body: str) -> bytes:
    """
    Create a simple UTF-8 plain-text .eml draft.
    NOTE: This does NOT send emails. It only generates draft files.
    If lookup fails, To: will be blank.
    """
    to_email = "" if to_email is None else str(to_email).strip()
    msg = (
        f"To: {to_email}\n"
        f"Subject: {subject}\n"
        f"MIME-Version: 1.0\n"
        f'Content-Type: text/plain; charset="utf-8"\n'
        f"Content-Transfer-Encoding: 8bit\n"
        f"\n"
        f"{body}\n"
    )
    return msg.encode("utf-8")


def read_any(file) -> pd.DataFrame:
    name = (file.name or "").lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)


# ----------------------------
# Lender-specific configuration
# ----------------------------
LENDER_CONFIG = {
    "Santander": {
        "lender_required_cols": ["Broker Name", "Firm", "Maturity Month", "Volume"],
        "subject": "Santander upcoming product transfers",
        # in case Santander wording should be different in the body later
        "display_name": "Santander",
    },
    # Add more lenders later:
    # "Halifax": {...},
}


# ----------------------------
# Streamlit UI
# ----------------------------
st.set_page_config(page_title="PT Communications – Email Draft Generator", layout="wide")
st.title("PT Communications – Product Transfer Email Draft Generator")

st.caption("This app generates **draft .eml files only**. It does **not** send emails.")

lender_name = st.selectbox("Step 1 — Select Lender", list(LENDER_CONFIG.keys()))

st.divider()

st.subheader(f"Step 2 — Upload files for {lender_name}")

col1, col2 = st.columns(2)
with col1:
    lender_file = st.file_uploader(
        f"Upload {lender_name} Lender Data",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=False,
    )
with col2:
    zoho_file = st.file_uploader(
        "Upload Zoho Data",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=False,
    )

st.divider()

if lender_file and zoho_file:
    config = LENDER_CONFIG[lender_name]

    # ----------------------------
    # Load files
    # ----------------------------
    try:
        df_lender = read_any(lender_file)
        df_zoho = read_any(zoho_file)
    except Exception as e:
        st.error(f"Error reading files: {e}")
        st.stop()

    # ----------------------------
    # Validate columns
    # ----------------------------
    lender_required = config["lender_required_cols"]

    # New Zoho headers you gave:
    zoho_required = ["Full Name", "AR Firm Name", "Email (AR Active advisers)"]

    missing_lender = [c for c in lender_required if c not in df_lender.columns]
    missing_zoho = [c for c in zoho_required if c not in df_zoho.columns]

    if missing_lender:
        st.error(f"{lender_name} lender data missing columns: {missing_lender}")
        st.stop()
    if missing_zoho:
        st.error(f"Zoho data missing columns: {missing_zoho}")
        st.stop()

    # ----------------------------
    # Normalize / prep for matching
    # ----------------------------
    df_lender = df_lender.copy()
    df_zoho = df_zoho.copy()

    df_lender["__broker_key"] = df_lender["Broker Name"].map(norm_text)
    df_lender["__firm_key"] = df_lender["Firm"].map(norm_text)

    df_zoho["__broker_key"] = df_zoho["Full Name"].map(norm_text)
    df_zoho["__firm_key"] = df_zoho["AR Firm Name"].map(norm_text)

    # Standardize email column name
    df_zoho["__email"] = df_zoho["Email (AR Active advisers)"].astype(str).str.strip()
    df_zoho = df_zoho[df_zoho["__email"].str.contains(r"@", na=False)].copy()

    email_lookup = (
        df_zoho.drop_duplicates(subset=["__broker_key", "__firm_key"])
        .set_index(["__broker_key", "__firm_key"])["__email"]
        .to_dict()
    )

    # ----------------------------
    # Aggregate lender volumes per broker per month
    # ----------------------------
    df_lender["Volume"] = pd.to_numeric(df_lender["Volume"], errors="coerce").fillna(0).astype(int)

    grouped = (
        df_lender.groupby(
            ["Broker Name", "Firm", "__broker_key", "__firm_key", "Maturity Month"],
            dropna=False
        )["Volume"]
        .sum()
        .reset_index()
    )

    def month_sort_key(x):
        try:
            s = str(x).strip().replace(".0", "")
            return int(s) if re.fullmatch(r"\d{6}", s) else 99999999
        except Exception:
            return 99999999

    grouped["__month_sort"] = grouped["Maturity Month"].map(month_sort_key)
    grouped = grouped.sort_values(["Broker Name", "Firm", "__month_sort"], ascending=True)

    broker_groups = grouped.groupby(["Broker Name", "Firm", "__broker_key", "__firm_key"], dropna=False)

    # ----------------------------
    # Generate drafts
    # ----------------------------
    subject = config["subject"]
    lender_display = config.get("display_name", lender_name)

    zip_buffer = io.BytesIO()
    manifest_rows = []
    unmatched_rows = []

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for (broker_name, firm, broker_key, firm_key), sub in broker_groups:
            to_email = email_lookup.get((broker_key, firm_key), "")
            if not to_email:
                unmatched_rows.append({"Broker Name": broker_name, "Firm": firm})

            month_lines = []
            for _, r in sub.iterrows():
                vol = int(r["Volume"])
                if vol <= 0:
                    continue
                month_label = yyyymm_to_month_year(r["Maturity Month"])
                month_lines.append(f"{vol} in {month_label}")

            # Skip if nothing to say
            if not month_lines:
                continue

            body = build_email_body(
                broker_first_name=first_name(broker_name),
                lender_name=lender_display,
                month_lines=month_lines,
            )
            eml_bytes = make_eml(to_email=to_email, subject=subject, body=body)

            safe_broker = re.sub(r"[^\w\s-]", "", str(broker_name)).strip().replace(" ", "_")
            safe_firm = re.sub(r"[^\w\s-]", "", str(firm)).strip().replace(" ", "_")
            filename = f"{lender_name}_{safe_broker}_{safe_firm}.eml"

            zf.writestr(filename, eml_bytes)

            manifest_rows.append(
                {
                    "Broker Name": broker_name,
                    "Firm": firm,
                    "Email (To)": to_email,
                    "Lender": lender_name,
                    "Draft File": filename,
                    "Lines": "; ".join(month_lines),
                }
            )

    zip_buffer.seek(0)

    # ----------------------------
    # Output / QA
    # ----------------------------
    left, right = st.columns([2, 1])

    with left:
        st.subheader("Drafts created")
        st.write(f"Created **{len(manifest_rows)}** email draft(s). (Drafts only — nothing is sent.)")

        if manifest_rows:
            st.dataframe(pd.DataFrame(manifest_rows), use_container_width=True)

        st.download_button(
            label="Download ZIP of .eml drafts",
            data=zip_buffer.getvalue(),
            file_name=f"{lender_name}_email_drafts.zip",
            mime="application/zip",
        )

    with right:
        st.subheader("Lookup misses (To left blank)")
        if unmatched_rows:
            st.warning(
                "These brokers were found in the lender file, but no matching (Name + Firm) email was found in Zoho. "
                "Drafts were still created with an empty To: field."
            )
            st.dataframe(pd.DataFrame(unmatched_rows).drop_duplicates(), use_container_width=True)
        else:
            st.success("All brokers matched to an email address.")
else:
    st.info("Select a lender and upload both files to generate drafts.")
