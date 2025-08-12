# streamlit_app.py
import io
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
# Extra headers to append from the Combined Case Report (in your requested order)
EXTRA_HEADERS = [
    "Full name",  # computed from First name + Last name
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
    "Created year",
    "Created month",
    "Created week",
    "Created at",
    "Case type",
    "Regulated",
    "Case status",
    "Mortgage status",
    "Mortgage amount",
    "Property value",
    "Term",
    "Term unit",
    "N clients",
    "Ltv",
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
    lk = {}
    for c in cols:
        lk[norm(c)] = c
    return lk


def find_col(lookup: dict, target_name: str, alt_variants=None):
    """
    Find a column in lookup by a canonical name with some common variants.
    Returns original column name or None.
    """
    candidates = [target_name]
    if alt_variants:
        candidates.extend(alt_variants)

    # common harmless variants (ID/id/Id, etc.)
    t = target_name
    if " id" in t.lower():
        candidates.append(t.replace(" id", " ID"))
        candidates.append(t.replace(" id", " Id"))

    for c in candidates:
        hit = lookup.get(norm(c))
        if hit:
            return hit
    return None


def safe_merge(rr: pd.DataFrame, cc: pd.DataFrame):
    # Build lookups
    rr_lk = build_lookup(rr.columns)
    cc_lk = build_lookup(cc.columns)

    # Find join key in both
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

    # Prepare Combined Case subset with computed Full name
    # Map base fields
    cc_map = {}
    missing_in_cc = []
    for base in CC_REQUIRED_BASE:
        col = find_col(cc_lk, base)
        if col is None:
            missing_in_cc.append(base)
        else:
            cc_map[base] = col  # canonical -> actual

    # Compute Full name (even if First/Last missing, we’ll fill with empty strings)
    first_col = cc_map.get("First name")
    last_col = cc_map.get("Last name")

    cc_subset = cc[[cc_key] + [v for k, v in cc_map.items() if k not in ("First name", "Last name")]].copy()

    # create Full name
    if first_col is None or last_col is None:
        cc_subset["Full name"] = ""
    else:
        cc_subset["Full name"] = (
            cc[first_col].fillna("").astype(str).str.strip() + " " +
            cc[last_col].fillna("").astype(str).str.strip()
        ).str.strip()

    # Reorder and rename cc_subset columns to the requested EXTRA_HEADERS order
    # Start by ensuring all headers exist; if some are missing we create empty columns
    cc_subset_renamed = pd.DataFrame()
    cc_subset_renamed[cc_key] = cc_subset[cc_key]

    # build a temp dict of source series by canonical key
    source = {k: cc[v] for k, v in cc_map.items() if k not in ("First name", "Last name")}
    source["Full name"] = cc_subset["Full name"]

    for h in EXTRA_HEADERS:
        if h == "Full name":
            cc_subset_renamed[h] = source["Full name"]
        else:
            # map canonical to canonical for renaming
            series = source.get(h)
            if series is not None:
                cc_subset_renamed[h] = series
            else:
                # create empty column if missing
                cc_subset_renamed[h] = pd.Series([None] * len(cc_subset_renamed))

    # Merge (left join to keep all Rate Review rows)
    merged = rr.merge(
        cc_subset_renamed,
        left_on=rr_key,
        right_on=cc_key,
        how="left",
        indicator=True
    )

    # Build output with all Rate Review columns first, then append EXTRA_HEADERS
    rr_cols = rr.columns.tolist()
    # ensure we don't duplicate join key from right
    drop_cols = {cc_key}
    keep_extras = [c for c in EXTRA_HEADERS if c not in rr_cols and c not in drop_cols]

    out_cols = rr_cols + keep_extras
    merged = merged.drop(columns=[c for c in merged.columns if c not in out_cols and c not in rr_cols], errors="ignore")
    merged = merged[out_cols]

    # Metrics
    matched = int((merged["_merge"] == "both").sum()) if "_merge" in merged.columns else None
    total = len(merged)
    if "_merge" in merged.columns:
        merged = merged.drop(columns=["_merge"])

    return merged, {
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

                # Info / warnings
                st.success(
                    f"Report is ready! Matched {stats['matched']} of {stats['total']} "
                    f"rows by **{JOIN_KEY}**."
                )
                if stats["missing_in_cc"]:
                    st.warning(
                        "Some PII fields were missing in the Combined Case Report and were left blank: "
                        + ", ".join(stats["missing_in_cc"])
                    )

                generate_download(output_df)

                # Optional: show preview
                with st.expander("Preview first 50 rows"):
                    st.dataframe(output_df.head(50), use_container_width=True)

        except KeyError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Something went wrong while generating the report: {e}")
