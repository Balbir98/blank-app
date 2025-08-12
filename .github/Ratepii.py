# streamlit_app.py
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Mortgage Rate Review PII", layout="centered")

st.title("Mortgage Rate Review PII")
st.write(
    "Upload **two** Acre exports: the **Mortgage Rate Review** report and the **Combined Case Report**. "
    "Then click **Generate report** to download a merged CSV that includes PII fields."
)

# --- UI: file uploads ---
col1, col2 = st.columns(2)
with col1:
    rr_file = st.file_uploader(
        "Mortgage Rate Review report",
        type=["csv", "xlsx", "xls"],
        key="rr_file",
        help="Upload the Mortgage Rate Review export (CSV or Excel)."
    )
with col2:
    cc_file = st.file_uploader(
        "Combined Case Report",
        type=["csv", "xlsx", "xls"],
        key="cc_file",
        help="Upload the Combined Case Report export (CSV or Excel)."
    )

st.divider()

# --- constants / config ---
# PII headers to bring from Combined Case Report (we'll select a subset for final output)
EXTRA_HEADERS = [
    "Full name",  # computed from First name + Last name
    "Dob", "Address1", "Address2", "Address3", "Posttown", "Postcode", "County", "Country",
    "Email address", "Mobile phone", "Home phone", "Work phone",
    "Created year", "Created month", "Created week", "Created at",
    "Case type", "Regulated", "Case status", "Mortgage status",
    "Mortgage amount", "Property value", "Term", "Term unit", "N clients", "Ltv",
]

# canonical names we expect to find in the Combined Case Report
CC_REQUIRED_BASE = [
    "First name", "Last name", "Dob", "Address1", "Address2", "Address3", "Posttown",
    "Postcode", "County", "Country", "Email address", "Mobile phone", "Home phone",
    "Work phone", "Created year", "Created month", "Created week", "Created at",
    "Case type", "Regulated", "Case status", "Mortgage status", "Mortgage amount",
    "Property value", "Term", "Term unit", "N clients", "Ltv"
]

JOIN_KEY = "Case id"  # expected in both reports

# Final output columns (and order)
OUTPUT_ORDER = [
    "Advisor name",
    "Case id",
    "Case URL",
    "Mtg completion date",
    "Mortgage id",
    "Lender name",
    "Status",
    "Initial rate",
    "Initial rate end date",
    "Current reminder date",
    "Reminder status",
    "Full name",
    "Dob",
    "Address1",
    "Address2",
    "Address3",
    "Posttown",
    "Postcode",
    "County",
    "Country",
    "Email address",
    "Mobile phone",
    "Home phone",
    "Work phone",
    "Created at",
    "Case type",
    "Case status",
    "Property value",
    "N clients",
]

# The subset of columns expected to come from Rate Review (we'll map them case-insensitively)
RR_EXPECTED_COLS = [
    "Advisor name",
    "Case id",
    "Mtg completion date",
    "Mortgage id",
    "Lender name",
    "Status",
    "Initial rate",
    "Initial rate end date",
    "Current reminder date",
    "Reminder status",
]

def read_any(file) -> pd.DataFrame:
    if file is None:
        return None
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(file)
    raise ValueError("Unsupported file type. Please upload CSV or Excel.")

def norm(s: str) -> str:
    """Normalize a column name for matching (lowercase, strip, collapse spaces/underscores)."""
    return " ".join(str(s).strip().replace("_", " ").split()).lower()

def build_lookup(cols):
    """Create a case-insensitive lookup {normalized_name: original_name} for a DataFrame's columns."""
    return {norm(c): c for c in cols}

def find_col(lookup: dict, target_name: str, alt_variants=None):
    """
    Find a column in lookup by a canonical name with some common variants.
    Returns original column name or None.
    """
    candidates = [target_name] + (alt_variants or [])
    t = target_name
    if " id" in t.lower():
        candidates += [t.replace(" id", " ID"), t.replace(" id", " Id")]
    for c in candidates:
        hit = lookup.get(norm(c))
        if hit:
            return hit
    return None

def safe_merge(rr: pd.DataFrame, cc: pd.DataFrame):
    # Lookups
    rr_lk = build_lookup(rr.columns)
    cc_lk = build_lookup(cc.columns)

    # Join keys
    rr_key = find_col(rr_lk, JOIN_KEY, alt_variants=["Case ID", "case id", "CaseId"])
    cc_key = find_col(cc_lk, JOIN_KEY, alt_variants=["Case ID", "case id", "CaseId"])
    if rr_key is None or cc_key is None:
        missing_side = []
        if rr_key is None:
            missing_side.append("Mortgage Rate Review")
        if cc_key is None:
            missing_side.append("Combined Case Report")
        raise KeyError(
            f"Couldn't find '{JOIN_KEY}' in: {', '.join(missing_side)}. "
            f"Make sure both reports include a '{JOIN_KEY}' column."
        )

    # Map CC columns we need
    cc_map = {}
    missing_in_cc = []
    for base in CC_REQUIRED_BASE:
        col = find_col(cc_lk, base)
        if col is None:
            missing_in_cc.append(base)
        else:
            cc_map[base] = col  # canonical -> actual col name in CC

    # Compute Full name from First/Last
    first_col = cc_map.get("First name")
    last_col = cc_map.get("Last name")

    cc_subset = cc[[cc_key] + [v for k, v in cc_map.items() if k not in ("First name", "Last name")]].copy()

    if first_col is None or last_col is None:
        cc_subset["Full name"] = ""
    else:
        cc_subset["Full name"] = (
            cc[first_col].fillna("").astype(str).str.strip() + " " +
            cc[last_col].fillna("").astype(str).str.strip()
        ).str.strip()

    # Build a source dict by canonical name
    source = {k: cc[v] for k, v in cc_map.items() if k not in ("First name", "Last name")}
    source["Full name"] = cc_subset["Full name"]

    # Rebuild CC subset with canonical headers we care about (only those used in final OUTPUT_ORDER)
    needed_from_cc = list(set(OUTPUT_ORDER) & set(EXTRA_HEADERS))
    cc_out = pd.DataFrame()
    cc_out[cc_key] = cc[cc_key]
    for h in needed_from_cc:
        series = source.get(h)
        if series is not None:
            cc_out[h] = series
        else:
            cc_out[h] = pd.NA

    # Merge (left join keeps all RR rows; note: duplicates in CC can still expand rows)
    merged = rr.merge(
        cc_out,
        left_on=rr_key,
        right_on=cc_key,
        how="left",
        indicator=True
    )

    # Build Case URL based on RR key column
    case_series = merged[rr_key].astype(str).str.strip()
    merged["Case URL"] = "https://crm.myac.re/cases/" + case_series + "/overview"

    # Map RR expected columns to their actual names
    rr_mapped_cols = {}
    for c in RR_EXPECTED_COLS:
        hit = find_col(rr_lk, c)
        if hit:
            rr_mapped_cols[c] = hit

    # Ensure "Case id" in output maps to the RR key's display name "Case id"
    # If the RR key column isn't literally named "Case id", we add a canonical "Case id" column.
    if "Case id" not in merged.columns:
        merged["Case id"] = merged[rr_key]

    # Assemble final columns in OUTPUT_ORDER
    final_df = pd.DataFrame()
    for col in OUTPUT_ORDER:
        if col in merged.columns:
            final_df[col] = merged[col]
        elif col in rr_mapped_cols:
            final_df[col] = merged[rr_mapped_cols[col]]
        else:
            # If not found, create blank column to preserve order
            final_df[col] = pd.NA

    # Metrics
    matched = int((merged["_merge"] == "both").sum()) if "_merge" in merged.columns else None
    total = len(merged)
    if "_merge" in merged.columns:
        merged = merged.drop(columns=["_merge"])

    return final_df, {
        "rr_key": rr_key,
        "cc_key": cc_key,
        "matched": matched,
        "total": total,
        "missing_in_cc": missing_in_cc
    }

def generate_download(df: pd.DataFrame, default_name="mortgage_rate_review_pii.csv"):
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="⬇️ Download Mortgage Rate Review PII CSV",
        data=csv_bytes,
        file_name=default_name,
        mime="text/csv",
        use_container_width=True
    )

# --- Action button ---
if st.button("Generate report", type="primary", use_container_width=True):
    if rr_file is None or cc_file is None:
        st.error("Please upload **both** files before generating the report.")
    else:
        try:
            rr_df = read_any(rr_file)
            cc_df = read_any(cc_file)

            if rr_df is None or rr_df.empty:
                st.error("The Mortgage Rate Review report appears to be empty.")
            elif cc_df is None or cc_df.empty:
                st.error("The Combined Case Report appears to be empty.")
            else:
                with st.spinner("Merging reports…"):
                    output_df, stats = safe_merge(rr_df, cc_df)

                st.success(
                    f"Report is ready! Matched {stats['matched']} of {stats['total']} rows by **{JOIN_KEY}**."
                )
                if stats["missing_in_cc"]:
                    st.warning(
                        "Some PII fields were missing in the Combined Case Report and were left blank: "
                        + ", ".join(stats["missing_in_cc"])
                    )

                generate_download(output_df)

        except KeyError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Something went wrong while generating the report: {e}")
