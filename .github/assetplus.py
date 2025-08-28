# app.py
import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Adviser Report Builder", layout="centered")

REQ = ["Advisor name", "Amount due", "Status", "Type", "Due date"]

# ---------- helpers ----------
def normalise(df: pd.DataFrame) -> pd.DataFrame:
    want = {c.lower(): c for c in REQ}
    ren = {}
    for c in df.columns:
        lc = str(c).strip().lower()
        if lc in want and c != want[lc]:
            ren[c] = want[lc]
    return df.rename(columns=ren)

def assert_required(df: pd.DataFrame):
    missing = [c for c in REQ if c not in df.columns]
    if missing:
        st.error("Missing required columns: " + ", ".join(missing))
        st.stop()

def parse_due_date(series: pd.Series) -> pd.Series:
    # Month-first first (works for ISO 'YYYY-MM-DD HH:MM:SS'), then fall back to day-first
    d = pd.to_datetime(series, errors="coerce", dayfirst=False)
    m = d.isna()
    if m.any():
        d2 = pd.to_datetime(series[m], errors="coerce", dayfirst=True)
        d.loc[m] = d2
    return d

def add_year_month_cols(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Due date"] = parse_due_date(out["Due date"])
    out["Year"] = out["Due date"].dt.year
    out["MonthNum"] = out["Due date"].dt.month
    out["Month"] = out["Due date"].dt.strftime("%b")
    return out

def mask_due(df: pd.DataFrame) -> pd.Series:
    return df["Status"].astype(str).str.strip().str.casefold().eq("due")

def pivot_fallback_table(df: pd.DataFrame) -> pd.DataFrame:
    # Build a pivot-like summary with Status='Due'
    d = df.loc[mask_due(df)].copy()
    pvt = pd.pivot_table(
        d,
        index=["Advisor name"],
        columns=["Year", "MonthNum", "Month"],
        values="Amount due",
        aggfunc="sum",
        fill_value=0,
    )
    # Sort by Year then MonthNum, then flatten to "YYYY Mon"
    if isinstance(pvt.columns, pd.MultiIndex):
        pvt = pvt.sort_index(axis=1, level=[0, 1])
        pvt.columns = [f"{y} {m}" for (y, mn, m) in pvt.columns.to_list()]
    else:
        pvt = pvt.sort_index(axis=1)
    return pvt.reset_index()

def add_unique_sheet(wb, writer, base_name: str):
    existing = {name.lower() for name in writer.sheets.keys()}
    name = base_name
    i = 1
    while name.lower() in existing:
        name = f"{base_name}_{i}"
        i += 1
    ws = wb.add_worksheet(name)
    writer.sheets[name] = ws
    return ws, name

def write_workbook_for_adviser(df_all_adv_rows: pd.DataFrame, adviser: str) -> bytes:
    # Keep ALL rows for this adviser in Data
    df = add_year_month_cols(df_all_adv_rows.copy())
    df["Amount due"] = pd.to_numeric(df["Amount due"], errors="coerce").fillna(0)

    import xlsxwriter
    from xlsxwriter.utility import xl_col_to_name

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as xw:
        # --- Data sheet
        df.to_excel(xw, sheet_name="Data", index=False)
        wb = xw.book
        wsD = xw.sheets["Data"]
        rows, cols = df.shape
        wsD.autofilter(0, 0, rows, cols - 1)
        for i in range(cols):
            wsD.set_column(i, i, 16)
        # Hide helper cols (Year, MonthNum, Month) if present at end
        if cols >= 3:
            wsD.set_column(cols - 3, cols - 1, None, None, {"hidden": 1})

        # --- Summary sheet (unique, created once)
        wsS, sname = add_unique_sheet(wb, xw, "Summary")
        wsS.write("A1", f"Summary for: {adviser}")

        can_pivot = hasattr(wb, "add_pivot_table")
        if can_pivot:
            try:
                data_ref = f"Data!A1:{xl_col_to_name(cols - 1)}{rows + 1}"
                wb.add_pivot_table({
                    "data": data_ref,
                    "name": "AdvisorPivot",
                    "worksheet": wsS,
                    "origin": "A3",
                    "rows":    [{"field": "Advisor name"}],
                    # Use MonthNum as a hidden ordering helper between Year and Month:
                    "columns": [{"field": "Year"}, {"field": "MonthNum"}, {"field": "Month"}],
                    "filters": [
                        {"field": "Status", "items": ["Due"]},
                        {"field": "Type"},
                    ],
                    "values":  [{"field": "Amount due", "function": "sum"}],
                })
                # (Excel will show a thin MonthNum header row; it's only for sorting.)
            except Exception:
                p = pivot_fallback_table(df)
                p.to_excel(xw, sheet_name=sname, startrow=2, index=False)
                wsS.set_column(0, 0, 22)
        else:
            p = pivot_fallback_table(df)
            p.to_excel(xw, sheet_name=sname, startrow=2, index=False)
            wsS.set_column(0, 0, 22)

    out.seek(0)
    return out.read()

# ---------- UI ----------
st.title("Adviser Report Builder")

up = st.file_uploader("Upload data (.xlsx or .csv)", type=["xlsx","csv"])
if up:
    df = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
    df = normalise(df)
    assert_required(df)

    advisers = sorted(df["Advisor name"].dropna().astype(str).unique())
    selected = st.multiselect("Select advisers (one workbook per adviser)", advisers)

    if st.button("Generate report(s)", type="primary", disabled=(len(selected) == 0)):
        if len(selected) == 1:
            adv = selected[0]
            adv_df_all = df[df["Advisor name"].astype(str) == adv]
            if adv_df_all.empty:
                adv_df_all = df.iloc[0:0]
            xlsx = write_workbook_for_adviser(adv_df_all, adv)
            st.download_button(
                f"Download {adv}.xlsx",
                data=xlsx,
                file_name=f"{adv}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            mem = io.BytesIO()
            with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for adv in selected:
                    adv_df_all = df[df["Advisor name"].astype(str) == adv]
                    if adv_df_all.empty:
                        adv_df_all = df.iloc[0:0]
                    xlsx = write_workbook_for_adviser(adv_df_all, adv)
                    zf.writestr(f"{adv}.xlsx", xlsx)
            mem.seek(0)
            st.download_button(
                f"Download {len(selected)} reports (.zip)",
                data=mem.getvalue(),
                file_name="adviser_reports.zip",
                mime="application/zip",
            )

st.caption("Each workbook: Data = ALL rows for that adviser; Summary = Sum(Amount due) by Year → Month, default Status='Due' with Type filter adjustable.")
