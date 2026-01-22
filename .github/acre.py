import streamlit as st
import pandas as pd
import io
import csv
import time

st.set_page_config(page_title="Boolean Cleaner", layout="wide")
st.title("CSV Boolean Cleaner")

uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])
if uploaded_file:

    @st.cache_data
    def load_csv(file):
        dtype_map = {"Adviser ID": str, "Firm ID": str}
        try:
            return pd.read_csv(file, dtype=dtype_map)
        except Exception:
            file.seek(0)
            raw_bytes = file.read()
            try:
                raw_text = raw_bytes.decode("utf-8")
            except UnicodeDecodeError:
                raw_text = raw_bytes.decode("latin-1")

            dialect = csv.Sniffer().sniff(raw_text[:10_000])
            sep = dialect.delimiter
            return pd.read_csv(io.StringIO(raw_text), sep=sep, engine="python", dtype=dtype_map)

    df = load_csv(uploaded_file)

    # Dates to force as TEXT dd/MM/yyyy
    date_cols = [
        "Application Date",
        "Effective Date",
        "Benefit End Date",
        "Created Date",
        "Last Updated",
        "Earliest Version Date",
        "Older Version Date",
    ]

    def to_ddmmyyyy_text(series: pd.Series) -> pd.Series:
        """
        Force TEXT dd/MM/yyyy.
        - Handles dd/MM/yyyy (+ optional time)
        - Handles YYYY-MM-DD (+ optional time)
        - Handles Excel serials
        - Blanks placeholders and unparseable values
        """
        s = series.fillna("").astype(str).str.strip()

        # Treat placeholder strings as blanks
        s = s.mask(s.str.lower().isin({"nan", "none", "null", "nat"}), "")

        # Parse pass 1 (UK dayfirst) – catches dd/MM/yyyy and many mixed cases
        parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)

        # Parse pass 2 (ISO) for anything still not parsed but containing '-'
        need_iso = parsed.isna() & s.str.contains("-", na=False)
        if need_iso.any():
            parsed.loc[need_iso] = pd.to_datetime(s[need_iso], errors="coerce")

        # Parse pass 3 (Excel serial numbers)
        still_nat = parsed.isna() & s.ne("")
        if still_nat.any():
            nums = pd.to_numeric(s[still_nat].str.replace(",", ""), errors="coerce")
            has_num = nums.notna()
            if has_num.any():
                parsed.loc[still_nat[still_nat].index[has_num]] = (
                    pd.to_datetime("1899-12-30") + pd.to_timedelta(nums[has_num], unit="D")
                )

        # Final: format ONLY valid parsed dates. Anything else -> blank (prevents Zoho import errors)
        out = parsed.dt.strftime("%d/%m/%Y")
        out = out.where(parsed.notna(), "")
        return out.astype(str)

    # Apply strict dd/MM/yyyy text conversion to all listed date columns
    for col in date_cols:
        if col in df.columns:
            df[col] = to_ddmmyyyy_text(df[col])

    # Sort by Application Date ascending (blanks last)
    if "Application Date" in df.columns:
        app_dt = pd.to_datetime(df["Application Date"], errors="coerce", dayfirst=True)
        df["__app_sort__"] = app_dt
        df = df.sort_values("__app_sort__", ascending=True, na_position="last").drop(columns="__app_sort__")

    # Boolean cleanup (unchanged)
    bool_cols = [
        "High Risk", "Whole Of Life", "In Trust", "BTL", "Adverse",
        "Self Cert", "Off Panel", "Introduced?", "Been Checked?",
        "App2 Blank?", "Lending into retirement?", "Second Charge?"
    ]

    def transform_bool(val):
        if val == "t":
            return True
        elif val == "f":
            return False
        else:
            return val

    prog = st.progress(0)
    status = st.empty()

    cols_to_do = [c for c in bool_cols if c in df.columns]
    total_steps = len(cols_to_do) + 1
    step = 0

    for col in cols_to_do:
        df[col] = df[col].apply(transform_bool)
        step += 1
        prog.progress(step / total_steps)
        status.text(f"Transforming “{col}” — ≈ {total_steps-step}s remaining")

    status.text("Generating cleaned CSV…")

    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

    step += 1
    prog.progress(step / total_steps)
    status.text("All done!")

    st.download_button(
        label="Download cleaned CSV",
        data=csv_bytes,
        file_name="cleaned_data.csv",
        mime="text/csv",
    )

else:
    st.info("Please upload a CSV file to get started.")
