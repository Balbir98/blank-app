import re
import os
from io import BytesIO
import zipfile
import pandas as pd
import streamlit as st

# Excel handling / styling
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font
from openpyxl.utils import get_column_letter, coordinate_to_tuple

# ---------------------------
# Utilities: robust readers
# ---------------------------

def _read_any_table(uploaded_file, preferred_sheet_name=None):
    """
    Read CSV or Excel into a DataFrame.
    """
    name = uploaded_file.name.lower()
    ext = os.path.splitext(name)[1]
    if ext in [".csv", ".txt"]:
        try:
            return pd.read_csv(uploaded_file, engine="python", sep=None)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, engine="python", sep=None, encoding="latin-1")
    elif ext in [".xlsx", ".xls"]:
        if preferred_sheet_name:
            try:
                return pd.read_excel(uploaded_file, sheet_name=preferred_sheet_name)
            except Exception:
                uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, sheet_name=0)
    else:
        try:
            return pd.read_csv(uploaded_file, engine="python", sep=None)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, engine="python", sep=None, encoding="latin-1")

# ---------------------------
# Transformation Logic
# ---------------------------

ID_COLS = ['Random ID', 'Provider Name', 'Name', 'Phone', 'Email']

def _is_event_label(x):
    """Heuristic: is a subheader a date/month/quarter string?"""
    if pd.isna(x):
        return False
    s = str(x).strip()
    months3 = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    if any(s.lower().startswith(m) for m in months3):
        return True
    if s in ['Q1','Q2','Q3','Q4','Monthly','Quarterly']:
        return True
    if re.search(r'\d', s):  # e.g., "4th February - Midlands"
        return True
    return False

def transform(form_df: pd.DataFrame, costs_df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert Zoho's wide export to normalized rows and join Cost from MOF sheet.
    - “Option” rule is applied for both named and unnamed columns.
    """
    if form_df.shape[0] < 2:
        return pd.DataFrame(columns=[
            'Random ID','Provider Name','Name','Phone','Email',
            'Type','Event Date (if applicable)','Product','Cost'
        ])

    subheaders = form_df.iloc[0]
    data_rows = form_df.iloc[1:].reset_index(drop=True)

    records = []
    for ridx, row in data_rows.iterrows():
        current_type = None
        for j, col in enumerate(form_df.columns):
            val = row[col]
            if not str(col).startswith('Unnamed'):
                if col not in ID_COLS and col not in ['Added Time', 'Referrer Name', 'Task Owner']:
                    current_type = col
            if pd.isna(val):
                continue

            sub = subheaders.iloc[j] if j < len(subheaders) else None
            text_val = str(val).strip()

            if not str(col).startswith('Unnamed'):
                # Named column: apply Option rule too (bug fix)
                if col in ID_COLS or col in ['Added Time', 'Referrer Name', 'Task Owner']:
                    continue
                if re.search(r'\boption(s)?\b', text_val, flags=re.I) and not pd.isna(sub):
                    prod = str(sub).strip()
                    evt = None if not _is_event_label(sub) else str(sub).strip()
                    if _is_event_label(sub):
                        prod = text_val
                else:
                    prod = text_val
                    evt = str(sub).strip() if _is_event_label(sub) else None
                records.append({'_ridx': ridx, 'Type': current_type,
                                'Event Date (if applicable)': evt, 'Product': prod})
            else:
                # Unnamed column path
                if re.search(r'\boption(s)?\b', text_val, flags=re.I):
                    prod = str(sub).strip() if not pd.isna(sub) else text_val
                    evt = None
                else:
                    if _is_event_label(sub):
                        evt = str(sub).strip()
                        prod = text_val
                    else:
                        evt = None
                        prod = str(sub).strip() if not pd.isna(sub) else text_val
                if current_type is None:
                    continue
                records.append({'_ridx': ridx, 'Type': current_type,
                                'Event Date (if applicable)': evt, 'Product': prod})

    out = pd.DataFrame.from_records(records)

    # Attach ID columns
    for c in ID_COLS:
        out[c] = out['_ridx'].map(data_rows[c]) if c in data_rows.columns else None

    out = out[['Random ID','Provider Name','Name','Phone','Email',
               'Type','Event Date (if applicable)','Product']]

    # Join Cost
    def _norm(s): return None if pd.isna(s) else str(s).strip()
    if not set(['Type','Product','Cost']).issubset(costs_df.columns):
        raise ValueError("Cost sheet must contain columns: Type, Product, Cost")
    costs2 = costs_df.copy()
    costs2['Type_norm'] = costs2['Type'].apply(_norm)
    costs2['Product_norm'] = costs2['Product'].apply(_norm)
    out['Type_norm'] = out['Type'].apply(_norm)
    out['Product_norm'] = out['Product'].apply(_norm)

    out = out.merge(
        costs2[['Type_norm','Product_norm','Cost']],
        on=['Type_norm','Product_norm'], how='left'
    ).drop(columns=['Type_norm','Product_norm'])

    out = out.sort_values(['Random ID','Type','Event Date (if applicable)','Product']).reset_index(drop=True)
    return out

# ---------------------------
# Template helpers
# ---------------------------

def _sanitize_name(name: str) -> str:
    """Remove commas and extra whitespace."""
    if pd.isna(name):
        return ""
    s = str(name).replace(",", " ")
    return re.sub(r"\s+", " ", s).strip()

def _find_table_header_row(ws):
    """
    Find a row containing 'Type' and 'Product'.
    Return (header_row_index, header_col_index, header_map[original header text -> col]).
    """
    max_row = min(ws.max_row, 200)
    max_col = min(ws.max_column, 80)
    for r in range(1, max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        for c in range(1, max_col):
            v = row_vals[c-1]
            if v is not None and str(v).strip().lower() == "type":
                nxt = row_vals[c] if c < max_col else None
                if nxt is not None and str(nxt).strip().lower() == "product":
                    # Build header map from this row
                    hmap = {}
                    for c2 in range(1, max_col + 1):
                        v2 = ws.cell(r, c2).value
                        if v2 is None:
                            continue
                        key = str(v2).strip()
                        hmap[key] = c2
                    return r, c, hmap
    raise RuntimeError("Could not find the table header row with 'Type' and 'Product'.")

def _get_col(hmap, key, aliases=()):
    """Find a column by exact header or any alias (case-insensitive)."""
    if key in hmap:
        return hmap[key]
    lowered = {k.lower(): v for k, v in hmap.items()}
    if key.lower() in lowered:
        return lowered[key.lower()]
    for a in aliases:
        if a in hmap:
            return hmap[a]
        if a.lower() in lowered:
            return lowered[a.lower()]
    return None

# -------- NEW: label/value helpers for summary rows --------

def _label_cell(ws, label_text):
    """
    Find (row, col) of the label cell using forgiving match (case/colon/spacing).
    Searches the whole sheet (reasonable caps applied).
    """
    target = re.sub(r'[:\u00A0]+$', '', str(label_text)).strip().lower()
    max_r = min(ws.max_row, 800)
    max_c = min(ws.max_column, 120)
    for rr in range(1, max_r + 1):
        for cc in range(1, max_c + 1):
            v = ws.cell(rr, cc).value
            if v is None:
                continue
            s = re.sub(r'[:\u00A0]+$', '', str(v)).strip().lower()
            if s == target:
                return rr, cc
    return None, None

def _value_cell_right(ws, r, c, span=10):
    """
    Return the FIRST editable value cell to the right of (r,c),
    skipping literal '£' or '-' placeholders.
    This lands you in the visible value column (e.g., J).
    """
    for cc in range(c + 1, c + 1 + span):
        v = ws.cell(r, cc).value
        if isinstance(v, str) and v.strip() in {"£", "-"}:
            continue
        if v is None or isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")):
            return ws.cell(r, cc)
    return ws.cell(r, c + 1)

# ---------------------------
# Template population
# ---------------------------

def _populate_template_bytes(template_bytes: bytes, cleaned: pd.DataFrame, costs_df: pd.DataFrame) -> BytesIO:
    """
    Returns a ZIP containing one populated template per Provider.
    - Fixed cells:
        B7 = Provider Name
        B9 = Main contact (Name sans comma)
        G9 = Phone
        I9, B11, I15 = Email
    - Table rows filled and styled; Total = Qty * Charge.
    - Borders apply through 'Notes' and 'When to Invoice' if present.
    - Headers styled: Segoe UI 12 bold white.
    - Charge & Total cells formatted as Accounting (GBP).
    - Summary block:
        * Total Package = SUM(Total column)
        * Discount = (leave as is or 0 if empty)
        * Total Package Price = Total Package - Discount
        * VAT = Total Package Price / 5
        * Overall Package Price = Total Package Price + VAT  <-- now forced with helpers
    """
    # Optional F2F mapping from cost sheet
    f2f_map = {}
    if 'F2F or Online?' in costs_df.columns:
        def _n(s): return None if pd.isna(s) else str(s).strip()
        tmp = costs_df[['Product','F2F or Online?']].copy()
        tmp['Product_norm'] = tmp['Product'].apply(_n)
        f2f_map = dict(zip(tmp['Product_norm'], tmp['F2F or Online?']))

    groups = cleaned.groupby('Provider Name', dropna=False)

    zip_buf = BytesIO()
    dotted = Side(style='dotted')
    border = Border(top=dotted, bottom=dotted, left=dotted, right=dotted)
    font10 = Font(name="Segoe UI", size=10)
    header_font = Font(name="Segoe UI", size=12, bold=True, color="FFFFFF")
    ACC_FMT = '_-£* #,##0.00_-;_-£* -#,##0.00_-;_-£* "-"??_-;_-@_-'

    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for provider, dfp in groups:
            wb = load_workbook(BytesIO(template_bytes))
            ws = wb.active  # first sheet

            # Fixed cells
            provider_val = "" if pd.isna(provider) else str(provider)
            name_val = _sanitize_name(dfp['Name'].iloc[0] if 'Name' in dfp.columns and len(dfp) > 0 else "")
            phone_val = "" if 'Phone' not in dfp.columns else ("" if pd.isna(dfp['Phone'].iloc[0]) else str(dfp['Phone'].iloc[0]))
            email_val = "" if 'Email' not in dfp.columns else ("" if pd.isna(dfp['Email'].iloc[0]) else str(dfp['Email'].iloc[0]))

            ws["B7"]  = provider_val
            ws["B9"]  = name_val
            ws["G9"]  = phone_val
            ws["I9"]  = email_val
            ws["B11"] = email_val
            ws["I15"] = email_val

            # Find table and headers
            hdr_row, hdr_col, hmap = _find_table_header_row(ws)

            c_Type   = _get_col(hmap, "Type")
            c_Prod   = _get_col(hmap, "Product")
            c_Month  = _get_col(hmap, "Month", aliases=("Event Month","Month of"))
            c_Date   = _get_col(hmap, "Date")
            c_F2F    = _get_col(hmap, "F2F or Online?", aliases=("F2F or Online","F2F/Online"))
            c_Qty    = _get_col(hmap, "Qty", aliases=("Quantity",))
            c_Charge = _get_col(hmap, "Charge", aliases=("Cost","Price"))
            c_Total  = _get_col(hmap, "Total")
            c_Notes  = _get_col(hmap, "Notes")
            c_When   = _get_col(hmap, "When to Invoice", aliases=("When To Invoice","When-to-Invoice"))

            # Header styling (font only; template handles fill color)
            for c in [c_Type,c_Prod,c_Month,c_Date,c_F2F,c_Qty,c_Charge,c_Total,c_Notes,c_When]:
                if c:
                    ws.cell(hdr_row, c).font = header_font

            start_row = hdr_row + 1
            n = len(dfp)

            # Insert rows to match number of items
            if n > 1:
                ws.insert_rows(start_row + 1, amount=n - 1)

            # Write only to mapped columns
            for i, (_, r) in enumerate(dfp.iterrows()):
                rr = start_row + i
                typ   = r.get("Type", "")
                prod  = r.get("Product", "")
                month = r.get("Event Date (if applicable)", "")
                qty   = 1
                charge = r.get("Cost", None)
                prod_key = None if pd.isna(prod) else str(prod).strip()
                f2f_val = f2f_map.get(prod_key, "")

                if c_Type:   ws.cell(rr, c_Type,   typ)
                if c_Prod:   ws.cell(rr, c_Prod,   prod)
                if c_Month:  ws.cell(rr, c_Month,  month)
                if c_Date:   ws.cell(rr, c_Date,   None)         # blank
                if c_F2F:    ws.cell(rr, c_F2F,    f2f_val)
                if c_Qty:    ws.cell(rr, c_Qty,    qty)

                # Charge (numeric), formatted as Accounting
                if c_Charge:
                    ws.cell(rr, c_Charge, charge)
                    ws.cell(rr, c_Charge).number_format = ACC_FMT

                # Total = Qty * Charge (numeric), formatted as Accounting
                if c_Total:
                    try:
                        total_val = (qty or 0) * (float(charge) if charge not in [None, ""] else 0.0)
                    except Exception:
                        total_val = None
                    ws.cell(rr, c_Total, total_val)
                    ws.cell(rr, c_Total).number_format = ACC_FMT

            # Borders + body font (Segoe UI 10)
            last_row = start_row + max(n - 1, 0)
            table_cols = [c for c in [c_Type,c_Prod,c_Month,c_Date,c_F2F,c_Qty,c_Charge,c_Total,c_Notes,c_When] if c]
            if not table_cols:
                table_cols = [hdr_col, hdr_col+1]  # minimal safety
            last_col = max(table_cols)
            first_col = min(table_cols)

            # Header borders
            for c in range(first_col, last_col + 1):
                cell = ws.cell(hdr_row, c)
                cell.border = border

            # Body borders + font
            for r in range(start_row, last_row + 1):
                for c in range(first_col, last_col + 1):
                    cell = ws.cell(r, c)
                    cell.border = border
                    cell.font = font10

            # ---------- Summary block ----------
            total_col = c_Total if c_Total else c_Charge
            ACC = ACC_FMT

            tp_coord = disc_coord = tpp_coord = vat_coord = None

            if total_col:
                sum_rng = f"{get_column_letter(total_col)}{start_row}:{get_column_letter(total_col)}{last_row}"

                # Total Package
                r_tp, c_tp = _label_cell(ws, "Total Package")
                if r_tp and c_tp:
                    tp_cell = _value_cell_right(ws, r_tp, c_tp)
                    tp_cell.value = f"=SUM({sum_rng})"
                    tp_cell.number_format = ACC
                    tp_coord = tp_cell.coordinate

                # Discount (keep value if present, else set to 0)
                r_d, c_d = _label_cell(ws, "Discount")
                if r_d and c_d:
                    disc_cell = _value_cell_right(ws, r_d, c_d)
                    if disc_cell.value is None:
                        disc_cell.value = 0
                    disc_cell.number_format = ACC
                    disc_coord = disc_cell.coordinate

                # Total Package Price = Total Package - Discount
                r_tpp, c_tpp = _label_cell(ws, "Total Package Price")
                if r_tpp and c_tpp and tp_coord and disc_coord:
                    tpp_cell = _value_cell_right(ws, r_tpp, c_tpp)
                    tpp_cell.value = f"={tp_coord}-{disc_coord}"
                    tpp_cell.number_format = ACC
                    tpp_coord = tpp_cell.coordinate

                # VAT = Total Package Price / 5
                r_vat, c_vat = _label_cell(ws, "VAT")
                if r_vat and c_vat and tpp_coord:
                    vat_cell = _value_cell_right(ws, r_vat, c_vat)
                    vat_cell.value = f"={tpp_coord}/5"
                    vat_cell.number_format = ACC
                    vat_coord = vat_cell.coordinate

                # ---- FORCE Overall Package Price = (TPP) + (VAT) ----
                r_opp, c_opp = _label_cell(ws, "Overall Package Price")
                if r_opp and c_opp and tpp_coord and vat_coord:
                    opp_cell = _value_cell_right(ws, r_opp, c_opp)
                    opp_cell.value = None  # clear legacy like =J27+J29
                    opp_cell.value = f"={tpp_coord}+{vat_coord}"
                    opp_cell.number_format = ACC

            # Save workbook into ZIP under templates/
            out_bytes = BytesIO()
            wb.save(out_bytes)
            out_bytes.seek(0)
            safe_provider = re.sub(r'[^A-Za-z0-9 _.-]+', '_', provider_val or "Unknown_Provider")
            zf.writestr(f"templates/{safe_provider}.xlsx", out_bytes.getvalue())

    zip_buf.seek(0)
    return zip_buf

# ---------------------------
# Streamlit App
# ---------------------------

st.set_page_config(page_title="Zoho Forms → Cleaned Output + Templates", layout="wide")
st.title("Zoho Forms → Cleaned Output + Populated Templates")

st.markdown("""
Upload your files and click **Submit** to get a ZIP containing:
- **templates/** → one populated workbook **per Provider Name**  
- **data/cleaned_output.xlsx** → your single cleaned dataset
""")

c1, c2, c3 = st.columns(3)
with c1:
    form_file = st.file_uploader("Zoho Forms export (.csv/.xlsx/.xls)", type=["csv","xlsx","xls","txt"])
with c2:
    cost_file = st.file_uploader("MOF Cost Sheet (.csv/.xlsx/.xls)", type=["csv","xlsx","xls","txt"])
with c3:
    template_file = st.file_uploader("Template (Excel, optional)", type=["xlsx","xls"])

if st.button("Submit"):
    if not form_file or not cost_file:
        st.error("Please upload both the Zoho Forms export and the MOF Cost Sheet.")
    else:
        with st.spinner("Processing..."):
            # Read inputs
            try:
                form_df = _read_any_table(form_file, preferred_sheet_name="Form")
            except Exception as e:
                st.exception(RuntimeError(f"Failed to read the Zoho Forms export: {e}"))
                st.stop()
            try:
                costs_df = _read_any_table(cost_file)
            except Exception as e:
                st.exception(RuntimeError(f"Failed to read the MOF Cost Sheet: {e}"))
                st.stop()

            # Transform
            try:
                cleaned = transform(form_df, costs_df)
            except Exception as e:
                st.exception(e)
                st.stop()

            # Build results.zip with templates (if provided) + data/cleaned_output.xlsx
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                # Add cleaned_output.xlsx
                cleaned_bytes = BytesIO()
                cleaned.to_excel(cleaned_bytes, index=False)
                zf.writestr("data/cleaned_output.xlsx", cleaned_bytes.getvalue())

                # Optional: templates
                if template_file is not None:
                    try:
                        template_bytes = template_file.read()
                        tpl_zip = _populate_template_bytes(template_bytes, cleaned, costs_df)
                        # Copy files from tpl_zip into templates/ folder in our main ZIP
                        with zipfile.ZipFile(tpl_zip, 'r') as tplzf:
                            for info in tplzf.infolist():
                                zf.writestr(info.filename, tplzf.read(info.filename))
                    except Exception as e:
                        st.exception(RuntimeError(f"Template population failed: {e}"))

            zip_buf.seek(0)

            # Summary & download
            st.success(f"Done. Cleaned {len(cleaned)} rows.")
            st.dataframe(cleaned.head(100), use_container_width=True)
            missing_costs = cleaned['Cost'].isna().sum()
            if missing_costs > 0:
                st.warning(f"{missing_costs} row(s) have no Cost match. Ensure (Type, Product) exist in the MOF Cost Sheet.")

            st.download_button(
                "Download results.zip",
                data=zip_buf.getvalue(),
                file_name="results.zip",
                mime="application/zip"
            )

st.markdown("---")
st.caption("OPP is now explicitly set to (Total Package Price + VAT) in the first editable value cell to the right of its label, so it follows the table as rows are inserted.")
