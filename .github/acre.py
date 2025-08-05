# streamlit_app.py
import streamlit as st
import pandas as pd
import io
import csv

st.set_page_config(page_title="Boolean Cleaner", layout="wide")
st.title("CSV Boolean Cleaner")

uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])
if uploaded_file:

    @st.cache_data
    def load_csv(file):
        # 1) Try standard UTF-8, comma
        try:
            return pd.read_csv(file)
        except Exception as e1:
            # 2) Fallback: sniff delimiter & python engine
            file.seek(0)
            raw = file.read().decode("utf-8", errors="replace")
            sniffer = csv.Sniffer()
            dialect = sniffer.sniff(raw[:10_000])
            sep = dialect.delimiter
            try:
                return pd.read_csv(io.StringIO(raw), sep=sep, engine="python")
            except Exception:
                # 3) Last resort: Latin-1
                file.seek(0)
                return pd.read_csv(file, encoding="latin-1", sep=sep, engine="python")

    df = load_csv(uploaded_file)

    # The columns to transform
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
            return val  # leave blanks or other values unchanged

    for col in bool_cols:
        if col in df.columns:
            df[col] = df[col].apply(transform_bool)

    st.subheader("Preview of Cleaned Data")
    st.dataframe(df)

    # CSV download
    csv_data = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download cleaned CSV",
        data=csv_data,
        file_name="cleaned_data.csv",
        mime="text/csv",
    )
else:
    st.info("Please upload a CSV file to get started.")
