# streamlit_app.py
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
        # 1) Try standard UTF-8, comma
        try:
            return pd.read_csv(file)
        except Exception:
            # 2) Sniff delimiter & python engine
            file.seek(0)
            raw = file.read().decode("utf-8", errors="replace")
            dialect = csv.Sniffer().sniff(raw[:10_000])
            try:
                return pd.read_csv(io.StringIO(raw), sep=dialect.delimiter, engine="python")
            except Exception:
                # 3) Fallback to Latin-1
                file.seek(0)
                return pd.read_csv(file, encoding="latin-1", sep=dialect.delimiter, engine="python")

    df = load_csv(uploaded_file)

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
            return val  # leave blanks/others

    # Set up progress UI
    prog = st.progress(0)
    status = st.empty()

    # Count steps: one per column + one to write CSV
    cols_to_do = [c for c in bool_cols if c in df.columns]
    total_steps = len(cols_to_do) + 1
    step = 0

    # 1) Transform each column
    for col in cols_to_do:
        start = time.perf_counter()
        df[col] = df[col].apply(transform_bool)
        step += 1
        elapsed = time.perf_counter() - start
        # assume roughly 1s per step for estimation
        remaining_secs = max(0, total_steps - step)
        prog.progress(step / total_steps)
        status.text(f"Transforming “{col}” — ≈ {remaining_secs}s remaining")

    # 2) Generate CSV bytes
    status.text("Generating cleaned CSV…")
    csv_bytes = df.to_csv(index=False).encode("utf-8")
    step += 1
    prog.progress(step / total_steps)
    status.text("All done!")

    # Download button
    st.download_button(
        label="Download cleaned CSV",
        data=csv_bytes,
        file_name="cleaned_data.csv",
        mime="text/csv",
    )

else:
    st.info("Please upload a CSV file to get started.")
