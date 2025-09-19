import re
import os
from io import BytesIO
import zipfile
import pandas as pd
import streamlit as st

# Excel handling / styling
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font
from openpyxl.utils import get_column_letter

# ===========================
# Utilities: robust readers
# ===========================
def _read_any_table(uploaded_file, preferred_sheet_name=None):
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

# ===========================
# Type -> Product overrides
# ===========================
TYPE_TO_PRODUCT = {
    "Compliance Update Sponsorship": "Monthly Compliance Bulletin",
    "Podcast": "Podcasts",
    "Directly Authorised Club Site Takeover": "Directly Authorised Club Site Takeover",
    "Full Takeover (DA Club and Adviser Site)": "Full Takeover (DA Club and Adviser Site)",
    "Network Adviser Site Takeover": "Network Adviser Site Takeover",
    "Promotional Emails": "Promotional Email",
    "Social Media Post Share": "Social Media Post",
    "Training Video": "Online training video (provider produces and edits)",
    "Video adverts": "Video advert",
    "Product Focus emails": "Product Focus emails",

    # === New explicit 1:1 mappings you asked for ===
    "Compliance Webinar Sponsorship": "Compliance Webinar Sponsorship",
    "Full Adviser Site Takeover (Network and DA Club)": "Full Adviser Site Takeover (Network and DA Club)",
}
_TYPE_OVERRIDE_LC = {k.casefold().strip(): v for k, v in TYPE_TO_PRODUCT.items()}

def _apply_type_overrides(df: pd.DataFrame) -> pd.DataFrame:
    if 'Type' not in df.columns or 'Product' not in df.columns:
        return df
    def _override(row):
        t = str(row.get('Type', '')).casefold().strip()
        return _TYPE_OVERRIDE_LC.get(t, row.get('Product'))
    df = df.copy()
    df['Product'] = df.apply(_override, axis=1)
    return df

# ===========================
# Helper: loose column picker
# ===========================
def _pick(colnames, *candidates):
    lowmap = {str(c).strip().lower(): c for c in colnames}
    for cand in candidates:
        if cand is None:
            continue
        key = str(cand).strip().lower()
        if key in lowmap:
            return lowmap[key]
    # try contains-style (looser)
    lc = [str(c).lower() for c in colnames]
    for cand in candidates:
        if not cand:
            continue
        ck = str(cand).strip().lower()
        for i, name in enumerate(lc):
            if ck == name:
                return colnames[i]
    return None

# ===========================
# Wishlist Transformation
# ===========================
# New repeated header set (in cleaned output)
REPEATED_FIRST = [
    'Random ID','Provider Name','Name','Phone','Email',
    'Events Name','Events Email',
    'Marketing Publications Name','Marketing Publications Email',
    'Copy Name','Copy Email',
    'When To Invoice',
    'Invoice Name','Invoice Email',
]

# Try to carry the long free-text notes
NOTES_SOURCE_HEADERS = [
    "Please provide any further notes you may have or want to have considered with this form:",
    "Further notes", "Any further notes"
]

# ID/admin-like columns we should copy from the Zoho wide export rows (case-insensitive match)
ID_COLS_WISHLIST = [
    'Random ID','Provider Name','Name','Phone','Email',
    'Events Name','Events Email',
    'Marketing Publications Name','Marketing Publications Email',
    'Copy Name','Copy Email',
    'When To Invoice',
    'Invoice Name','Invoice Email',
] + NOTES_SOURCE_HEADERS

def _is_event_label(x):
    if pd.isna(x):
        return False
    s = str(x).strip()
    months3 = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    if any(s.lower().startswith(m) for m in months3): return True
    if s in ['Q1','Q2','Q3','Q4','Monthly','Quarterly']: return True
    if re.search(r'\d', s): return True
    return False

# Simple location normaliser used for Regional Roadshow special-case
_LOCATION_SET = {
    "north","south","east","west","midlands","central","scotland","wales","northern ireland",
    "north east","north west","south east","south west","london","yorkshire","humberside"
}
def _looks_like_location(s: str) -> bool:
    if not isinstance(s, str): s = str(s or "")
    return s.strip().casefold() in _LOCATION_SET

def transform_wishlist(form_df: pd.DataFrame, costs_df: pd.DataFrame) -> pd.DataFrame:
    """
    Parse Zoho's wide export (row 0 = subheaders) → long rows,
    then join Cost (+F2F) and include the new repeated fields.

    Changes:
      * Product lookups to MOF Cost are by Product only (not Type+Product).
      * Special handling for "Regional Roadshow Event (...)" so locations go to Event Date.
      * Type→Product overrides applied BEFORE cost merge.
    """
    if form_df.shape[0] < 2:
        return pd.DataFrame(columns=REPEATED_FIRST + ['Type','Event Date (if applicable)','Product','Cost','F2F or Online?'])

    subheaders = form_df.iloc[0]
    data_rows = form_df.iloc[1:].reset_index(drop=True)

    # Normalize lookup for ID columns (case-insensitive)
    col_lc_map = {str(c).strip().lower(): c for c in form_df.columns}
    wanted_cols = [col_lc_map.get(c.lower()) for c in ID_COLS_WISHLIST if col_lc_map.get(c.lower())]
    # Extract the free text notes column name, if present
    notes_col = None
    for cand in NOTES_SOURCE_HEADERS:
        c = col_lc_map.get(cand.strip().lower())
        if c:
            notes_col = c
            break

    records = []
    for ridx, row in data_rows.iterrows():
        current_type = None
        for j, col in enumerate(form_df.columns):
            val = row[col]
            if not str(col).startswith('Unnamed'):
                # new "Type block" when we hit a titled column that's not admin/ID
                if col not in wanted_cols and col not in ['Added Time','Referrer Name','Task Owner']:
                    current_type = col
            if pd.isna(val):
                continue

            sub = subheaders.iloc[j] if j < len(subheaders) else None
            text_val = str(val).strip()

            if not str(col).startswith('Unnamed'):
                if col in wanted_cols or col in ['Added Time','Referrer Name','Task Owner']:
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

    # Attach repeated fields from data_rows
    for c in REPEATED_FIRST + ([notes_col] if notes_col else []):
        if c and c in data_rows.columns:
            out[c] = out['_ridx'].map(data_rows[c])
        else:
            if c in REPEATED_FIRST:
                out[c] = None

    # Thin to our key outputs before joining cost
    out = out[['_ridx'] + REPEATED_FIRST + ['Type','Event Date (if applicable)','Product']].copy()

    # ---------- Special handling: Regional Roadshow Events ----------
    def _is_regional_roadshow(t: str) -> bool:
        if not isinstance(t, str): t = str(t or "")
        t_l = t.casefold()
        return t_l.startswith("regional roadshow event (february)") or \
               t_l.startswith("regional roadshow event (june)") or \
               t_l.startswith("regional roadshow event (september/october)") or \
               t_l.startswith("regional roadshow event")

    mask_rre = out['Type'].apply(_is_regional_roadshow).fillna(False)
    if mask_rre.any():
        # Move location-like Products into the Event Date field (if it's empty/NA),
        # leaving Product free for the actual offering lines (Presenter, Stand Only, etc).
        loc_mask = mask_rre & out['Product'].apply(_looks_like_location)
        need_date = loc_mask & (out['Event Date (if applicable)'].isna() | (out['Event Date (if applicable)'].astype(str).str.strip() == ""))
        out.loc[need_date, 'Event Date (if applicable)'] = out.loc[need_date, 'Product']
        # When the row was purely a location line, clear Product so it won't try to cost-match.
        out.loc[need_date, 'Product'] = None

    # ---------- Apply Type->Product overrides BEFORE cost merge ----------
    out = _apply_type_overrides(out)

    # ---- Join Cost (+ F2F) by PRODUCT ONLY ----
    def _norm(s): return None if pd.isna(s) else str(s).strip()
    if 'Product' not in costs_df.columns or 'Cost' not in costs_df.columns:
        raise ValueError("Cost sheet must contain columns: Product, Cost (Type optional but ignored for lookup)")

    costs2 = costs_df.copy()
    costs2['Product_norm'] = costs2['Product'].apply(_norm)

    # Deduplicate costs on Product to avoid exploding joins
    bring_cols = ['Product_norm','Cost']
    f2f_col_name = None
    for cand in ['F2F or Online?', 'F2F or Online', 'F2F/Online']:
        if cand in costs2.columns:
            bring_cols.append(cand)
            f2f_col_name = cand
            break
    costs2 = costs2[bring_cols].drop_duplicates(subset=['Product_norm'], keep='first')

    out['Product_norm'] = out['Product'].apply(_norm)

    out = out.merge(costs2, on=['Product_norm'], how='left').drop(columns=['Product_norm'])

    # Final column order
    final_cols = REPEATED_FIRST + ['Type','Event Date (if applicable)','Product','Cost']
    if f2f_col_name:
        out = out.rename(columns={f2f_col_name: 'F2F or Online?'})
        final_cols += ['F2F or Online?']

    out = out[final_cols].sort_values(['Random ID','Type','Event Date (if applicable)','Product']).reset_index(drop=True)

    return out

# ===========================
# Confirmed Items Transform
# ===========================
def transform_confirmed(confirmed_df: pd.DataFrame) -> pd.DataFrame:
    cols = list(confirmed_df.columns)

    # Repeated fields
    c_rand   = _pick(cols, "Random ID","RandomID","ID")
    c_prov   = _pick(cols, "Provider Name","Provider")
    c_name   = _pick(cols, "Name","Main Contact","Contact Name for Content")
    c_phone  = _pick(cols, "Phone","Contact Phone Number")
    c_email  = _pick(cols, "Email","Contact Email")

    c_ev_n   = _pick(cols, "Events Name")
    c_ev_e   = _pick(cols, "Events Email")
    c_mp_n   = _pick(cols, "Marketing Publications Name")
    c_mp_e   = _pick(cols, "Marketing Publications Email")
    c_cp_n   = _pick(cols, "Copy Name")
    c_cp_e   = _pick(cols, "Copy Email")
    c_wti    = _pick(cols, "When To Invoice","When to Invoice")
    c_inv_n  = _pick(cols, "Invoice Name")
    c_inv_e  = _pick(cols, "Invoice Email")

    # Line fields
    c_type   = _pick(cols, "Type Of Event","Type of Event","Type")
    c_prod   = _pick(cols, "Product Ordered","Product")
    c_month  = _pick(cols, "Event Date (If Applicable)","Event Date (if applicable)","Month","Date")
    c_cost   = _pick(cols, "Product Cost","Cost","Charge")
    c_f2f    = _pick(cols, "F2F or Online?","F2F or Online","F2F/Online")

    # Notes long text (optional)
    notes_col = _pick(cols, *NOTES_SOURCE_HEADERS)

    required = [
        ("Provider Name", c_prov),
        ("Type", c_type),
        ("Product", c_prod),
        ("Cost", c_cost),
    ]
    miss = [lab for lab, col in required if col is None]
    if miss:
        raise ValueError(f"Confirmed list missing required columns: {', '.join(miss)}")

    out = pd.DataFrame({
        'Random ID': confirmed_df[c_rand] if c_rand else "",
        'Provider Name': confirmed_df[c_prov],
        'Name': confirmed_df[c_name] if c_name else "",
        'Phone': confirmed_df[c_phone] if c_phone else "",
        'Email': confirmed_df[c_email] if c_email else "",
        'Events Name': confirmed_df[c_ev_n] if c_ev_n else "",
        'Events Email': confirmed_df[c_ev_e] if c_ev_e else "",
        'Marketing Publications Name': confirmed_df[c_mp_n] if c_mp_n else "",
        'Marketing Publications Email': confirmed_df[c_mp_e] if c_mp_e else "",
        'Copy Name': confirmed_df[c_cp_n] if c_cp_n else "",
        'Copy Email': confirmed_df[c_cp_e] if c_cp_e else "",
        'When To Invoice': confirmed_df[c_wti] if c_wti else "",
        'Invoice Name': confirmed_df[c_inv_n] if c_inv_n else "",
        'Invoice Email': confirmed_df[c_inv_e] if c_inv_e else "",
        'Type': confirmed_df[c_type],
        'Event Date (if applicable)': confirmed_df[c_month] if c_month else "",
        'Product': confirmed_df[c_prod],
        'Cost': confirmed_df[c_cost],
        'F2F or Online?': confirmed_df[c_f2f] if c_f2f else "",
    })
    # Apply overrides (kept here too in case someone pastes odd "Type" into confirmed sheet)
    out = _apply_type_overrides(out)
    # Keep order
    final_cols = REPEATED_FIRST + ['Type','Event Date (if applicable)','Product','Cost','F2F or Online?']
    out = out[final_cols].sort_values(['Provider Name','Type','Event Date (if applicable)','Product'],
                                      na_position='last').reset_index(drop=True)
    # Keep notes as attribute for template stage (aggregate later)
    if notes_col:
        out['_notes_long_'] = confirmed_df[notes_col]
    else:
        out['_notes_long_'] = ""
    return out

# ===========================
# Template helpers (shared)
# ===========================
def _sanitize_name(name: str) -> str:
    if pd.isna(name): return ""
    s = str(name).replace(",", " ")
    return re.sub(r"\s+", " ", s).strip()

def _find_table_header_row(ws):
    max_row = min(ws.max_row, 200)
    max_col = min(ws.max_column, 80)
    for r in range(1, max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        for c in range(1, max_col):
            v = row_vals[c-1]
            if v is not None and str(v).strip().lower() == "type":
                nxt = row_vals[c] if c < max_col else None
                if nxt is not None and str(nxt).strip().lower() == "product":
                    hmap = {}
                    for c2 in range(1, max_col + 1):
                        v2 = ws.cell(r, c2).value
                        if v2 is not None:
                            hmap[str(v2).strip()] = c2
                    return r, c, hmap
    raise RuntimeError("Could not find the table header row with 'Type' and 'Product'.")

def _get_col(hmap, key, aliases=()):
    if key in hmap: return hmap[key]
    lowered = {k.lower(): v for k, v in hmap.items()}
    if key.lower() in lowered: return lowered[key.lower()]
    for a in aliases:
        if a in hmap: return hmap[a]
        if a.lower() in lowered: return lowered[a.lower()]
    return None

def _canon_label(s: str) -> str:
    return re.sub(r'[^a-z0-9]+', '', str(s).lower())

def _find_label_in_column(ws, label_text, col_idx, row_start, row_end):
    tgt = _canon_label(label_text)
    row_end = min(row_end, ws.max_row)
    for rr in range(row_start, row_end + 1):
        v = ws.cell(rr, col_idx).value
        if v is None:
            continue
        if tgt in _canon_label(v):
            return rr, col_idx
    return None, None

def _first_value_cell_right(ws, r, c, try_two=True):
    for cc in [c + 1, c + 2] if try_two else [c + 1]:
        v = ws.cell(r, cc).value
        if isinstance(v, str) and v.strip() == "£":
            continue
        return ws.cell(r, cc)
    return ws.cell(r, c + 1)

# ===========================
# Template population (shared) — NEW template layout
# ===========================
def _populate_template_bytes(template_bytes: bytes, cleaned: pd.DataFrame, costs_df: pd.DataFrame | None) -> BytesIO:
    # Build a product->F2F map if we have a cost sheet (Wishlist mode)
    f2f_map = {}
    if costs_df is not None and 'F2F or Online?' in costs_df.columns:
        def _n(s): return None if pd.isna(s) else str(s).strip()
        tmp = costs_df[['Product','F2F or Online?']].copy()
        tmp['Product_norm'] = tmp['Product'].apply(_n)
        f2f_map = dict(zip(tmp['Product_norm'], tmp['F2F or Online?']))

    zip_buf = BytesIO()
    dotted = Side(style='dotted')
    header_font = Font(name="Segoe UI", size=12, bold=True, color="000000")
    ACC_FMT = '_-£* #,##0.00_-;_-£* -#,##0.00_-;_-£* "-"??_-;_-@_-'

    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for provider, dfp in cleaned.groupby('Provider Name', dropna=False):
            wb = load_workbook(BytesIO(template_bytes))
            ws = wb.active

            # Fixed cells (new positions)
            prov = "" if pd.isna(provider) else str(provider)
            nm   = _sanitize_name(dfp['Name'].iloc[0] if 'Name' in dfp.columns and len(dfp) else "")
            ph   = "" if 'Phone' not in dfp.columns else ("" if pd.isna(dfp['Phone'].iloc[0]) else str(dfp['Phone'].iloc[0]))
            em   = "" if 'Email' not in dfp.columns else ("" if pd.isna(dfp['Email'].iloc[0]) else str(dfp['Email'].iloc[0]))
            wti  = dfp['When To Invoice'].iloc[0] if 'When To Invoice' in dfp.columns else ""

            ws['B4'] = prov
            ws['B6'] = nm
            ws['D4'] = wti
            ws['D6'] = ph
            ws['E6'] = em

            # Secondary contacts
            def _first(dfcol):
                return (dfcol.iloc[0] if dfcol is not None and len(dfp) else "")
            if 'Events Name' in dfp.columns: ws['B10'] = _first(dfp['Events Name'])
            if 'Events Email' in dfp.columns: ws['B12'] = _first(dfp['Events Email'])
            if 'Marketing Publications Name' in dfp.columns: ws['D10'] = _first(dfp['Marketing Publications Name'])
            if 'Marketing Publications Email' in dfp.columns: ws['D12'] = _first(dfp['Marketing Publications Email'])
            if 'Invoice Name' in dfp.columns: ws['F10'] = _first(dfp['Invoice Name'])
            if 'Invoice Email' in dfp.columns: ws['F12'] = _first(dfp['Invoice Email'])
            if 'Copy Name' in dfp.columns: ws['H10'] = _first(dfp['Copy Name'])
            if 'Copy Email' in dfp.columns: ws['H12'] = _first(dfp['Copy Email'])

            # Find table header
            hdr_row, hdr_col, hmap = _find_table_header_row(ws)
            c_Type   = _get_col(hmap, "Type")
            c_Prod   = _get_col(hmap, "Product")
            c_Det    = _get_col(hmap, "Details")
            c_Date   = _get_col(hmap, "Date")
            c_Charge = _get_col(hmap, "Charge", aliases=("Cost","Price"))
            c_Qty    = _get_col(hmap, "Qty", aliases=("Quantity",))
            c_Total  = _get_col(hmap, "Total")
            c_Notes  = _get_col(hmap, "Notes")

            # Make header font bold (optional)
            for c in [c_Type,c_Prod,c_Det,c_Date,c_Charge,c_Qty,c_Total,c_Notes]:
                if c: ws.cell(hdr_row, c).font = header_font

            # Insert rows for n items
            start_row = hdr_row + 1
            n = len(dfp)
            if n > 1:
                ws.insert_rows(start_row + 1, amount=n - 1)

            # Fill rows
            for i, (_, r) in enumerate(dfp.iterrows()):
                rr = start_row + i
                typ   = r.get("Type", "")
                prod  = r.get("Product", "")
                datev = r.get("Event Date (if applicable)", "")
                charge = r.get("Cost", None)
                qty = 1
                # Details column: F2F/Online (from cleaned or fallback from cost map)
                details = r.get("F2F or Online?", "")
                if not details:
                    pk = None if pd.isna(prod) else str(prod).strip()
                    if pk in f2f_map:
                        details = f2f_map.get(pk, "")

                if c_Type:   ws.cell(rr, c_Type, typ)
                if c_Prod:   ws.cell(rr, c_Prod, prod)
                if c_Det:    ws.cell(rr, c_Det, details)
                if c_Date:   ws.cell(rr, c_Date, datev)
                if c_Qty:    ws.cell(rr, c_Qty, qty)
                if c_Charge:
                    ws.cell(rr, c_Charge, charge).number_format = ACC_FMT
                if c_Total:
                    try:
                        total_val = (qty or 0) * (float(charge) if charge not in [None, ""] else 0.0)
                    except Exception:
                        total_val = None
                    ws.cell(rr, c_Total, total_val).number_format = ACC_FMT
                if c_Notes:
                    ws.cell(rr, c_Notes, None)

            # Borders across table (dotted)
            last_row = start_row + max(n - 1, 0)
            table_cols = [c for c in [c_Type,c_Prod,c_Det,c_Date,c_Charge,c_Qty,c_Total,c_Notes] if c]
            first_col, last_col = min(table_cols), max(table_cols)
            dotted_b = Border(top=dotted, bottom=dotted, left=dotted, right=dotted)
            for r in range(hdr_row, last_row + 1):
                for c in range(first_col, last_col + 1):
                    ws.cell(r, c).border = dotted_b

            # ---------- Summary block (labels to the left, values to the right) ----------
            ACC = ACC_FMT
            sum_col = c_Total if c_Total else c_Charge
            if sum_col:
                sum_rng = f"{get_column_letter(sum_col)}{start_row}:{get_column_letter(sum_col)}{last_row}"

                # Detect label column from "Total Package"
                label_col = None
                for c in range(1, ws.max_column + 1):
                    for rr in range(last_row + 1, min(ws.max_row, last_row + 200) + 1):
                        v = ws.cell(rr, c).value
                        if v and _canon_label(v) == _canon_label("Total Package"):
                            label_col = c
                            break
                    if label_col:
                        break
                if not label_col:
                    label_col = 8  # fallback H if needed

                value_col = label_col + 1
                search_start = last_row + 1
                search_end   = min(ws.max_row, last_row + 200)

                # Total Package
                r_tp, _ = _find_label_in_column(ws, "Total Package", label_col, search_start, search_end)
                tp_coord = None
                if r_tp:
                    tp_cell = _first_value_cell_right(ws, r_tp, label_col)
                    tp_cell.value = f"=SUM({sum_rng})"
                    tp_cell.number_format = ACC
                    tp_coord = tp_cell.coordinate

                # Discount
                r_disc, _ = _find_label_in_column(ws, "Discount", label_col, search_start, search_end)
                disc_coord = None
                if r_disc:
                    disc_cell = _first_value_cell_right(ws, r_disc, label_col)
                    if disc_cell.value is None:
                        disc_cell.value = 0
                    disc_cell.number_format = ACC
                    disc_coord = disc_cell.coordinate

                # Total Package Price = TP - Discount
                r_tpp, _ = _find_label_in_column(ws, "Total Package Price", label_col, search_start, search_end)
                tpp_coord = None
                if r_tpp and tp_coord and disc_coord:
                    tpp_cell = _first_value_cell_right(ws, r_tpp, label_col)
                    tpp_cell.value = f"={tp_coord}-{disc_coord}"
                    tpp_cell.number_format = ACC
                    tpp_coord = tpp_cell.coordinate

                # VAT = TPP / 5
                r_vat, _ = _find_label_in_column(ws, "VAT", label_col, search_start, search_end)
                vat_coord = None
                if r_vat and tpp_coord:
                    vat_cell = _first_value_cell_right(ws, r_vat, label_col)
                    vat_cell.value = f"={tpp_coord}/5"
                    vat_cell.number_format = ACC
                    vat_coord = vat_cell.coordinate

                # OPP = TPP + VAT
                r_opp, _ = _find_label_in_column(ws, "Overall Package Price", label_col, search_start, search_end)
                if not r_opp and r_vat:
                    r_opp = r_vat + 2
                if r_opp and tpp_coord and vat_coord:
                    opp_cell = ws.cell(r_opp, value_col)
                    opp_cell.value = f"={tpp_coord}+{vat_coord}"
                    opp_cell.number_format = ACC

            # ---------- Long "Notes" box ----------
            notes_text = ""
            if '_notes_long_' in dfp.columns:
                notes_text = "\n\n".join(
                    [str(x).strip() for x in dfp['_notes_long_'].fillna("").unique() if str(x).strip()]
                )
            else:
                for cand in NOTES_SOURCE_HEADERS:
                    if cand in dfp.columns:
                        notes_text = "\n\n".join(
                            [str(x).strip() for x in dfp[cand].fillna("").unique() if str(x).strip()]
                        )
                        break

            if notes_text:
                found_r = None
                for rr in range(last_row + 1, min(ws.max_row, last_row + 300) + 1):
                    v = ws.cell(rr, 2).value  # column B = 2
                    if isinstance(v, str) and _canon_label(v) == "notes":
                        found_r = rr
                        break
                if found_r:
                    target_r = found_r + 1
                    for cc in [2, 3, 4]:
                        ws.cell(target_r, cc).value = notes_text
                        break

            # Save provider file into zip
            out_bytes = BytesIO()
            wb.save(out_bytes)
            out_bytes.seek(0)
            safe_provider = re.sub(r'[^A-Za-z0-9 _.-]+', '_', prov or "Unknown_Provider")
            zf.writestr(f"templates/{safe_provider}.xlsx", out_bytes.getvalue())

    zip_buf.seek(0)
    return zip_buf

# ===========================
# Streamlit App
# ===========================
st.set_page_config(page_title="MOF Automation", layout="wide")
st.title("MOF Automation")

# Button styles (red submit, green download)
st.markdown("""
<style>
div.stButton > button:first-child { background-color: #d90429 !important; color: white !important; }
div.stDownloadButton > button:first-child { background-color: #2a9d8f !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
Upload your files and click **Submit** to get a ZIP containing:

**MOF Templates** → one populated workbook per Provider Name  
**Cleaned Data** → your single cleaned dataset (Wishlist mode)
""")

mode = st.selectbox("Choose mode", ["Wishlist", "Confirmed Items"])

# ---------------- Wishlist UI ----------------
if mode == "Wishlist":
    c1, c2, c3 = st.columns(3)
    with c1:
        form_file = st.file_uploader("Zoho Forms export (.csv/.xlsx/.xls)", type=["csv","xlsx","xls","txt"], key="form_wishlist")
    with c2:
        cost_file = st.file_uploader("MOF Cost Sheet (.csv/.xlsx/.xls)", type=["csv","xlsx","xls","txt"], key="cost_wishlist")
    with c3:
        template_file = st.file_uploader("New Template (Excel)", type=["xlsx","xls"], key="tpl_wishlist")

    if st.button("Submit", key="submit_wishlist"):
        if not form_file or not cost_file or not template_file:
            st.error("Please upload the Zoho Forms export, MOF Cost Sheet, and the new Template.")
        else:
            with st.spinner("Processing wishlist..."):
                try:
                    form_df = _read_any_table(form_file, preferred_sheet_name="Form")
                    costs_df = _read_any_table(cost_file)
                    cleaned = transform_wishlist(form_df, costs_df)
                except Exception as e:
                    st.exception(e)
                    st.stop()

                # Build results.zip with templates + cleaned data
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                    # cleaned_output.xlsx
                    cleaned_bytes = BytesIO()
                    cleaned.to_excel(cleaned_bytes, index=False)
                    zf.writestr("data/cleaned_output.xlsx", cleaned_bytes.getvalue())

                    # templates/*
                    try:
                        template_bytes = template_file.read()
                        tpl_zip = _populate_template_bytes(template_bytes, cleaned, costs_df)
                        with zipfile.ZipFile(tpl_zip, 'r') as tplzf:
                            for info in tplzf.infolist():
                                zf.writestr(info.filename, tplzf.read(info.filename))
                    except Exception as e:
                        st.exception(RuntimeError(f"Template population failed: {e}"))

                zip_buf.seek(0)
                st.success(f"Done. Cleaned {len(cleaned)} rows.")

                # Missing costs warning
                missing_costs = cleaned['Cost'].isna().sum()
                if missing_costs > 0:
                    st.warning(f"{missing_costs} row(s) have no Cost match by Product. Ensure Product exists in the MOF Cost Sheet.")
                    with st.expander("Preview rows missing Cost"):
                        st.dataframe(
                            cleaned[cleaned['Cost'].isna()][
                                ['Provider Name','Type','Event Date (if applicable)','Product']
                            ].head(500)
                        )

                st.download_button(
                    "Download Now!",
                    data=zip_buf.getvalue(),
                    file_name="results.zip",
                    mime="application/zip",
                    key="dl_wishlist"
                )

# ---------------- Confirmed Items UI ----------------
else:
    c1, c2 = st.columns(2)
    with c1:
        confirmed_file = st.file_uploader("Confirmed Items export (.csv/.xlsx/.xls)", type=["csv","xlsx","xls","txt"], key="confirmed_list")
    with c2:
        template_file_c = st.file_uploader("New Template (Excel)", type=["xlsx","xls"], key="tpl_confirmed")

    if st.button("Submit", key="submit_confirmed"):
        if not confirmed_file or not template_file_c:
            st.error("Please upload the Confirmed Items export and the new Template.")
        else:
            with st.spinner("Processing confirmed items..."):
                try:
                    confirmed_df = _read_any_table(confirmed_file)
                    cleaned_confirmed = transform_confirmed(confirmed_df)
                except Exception as e:
                    st.exception(e)
                    st.stop()

                # Build results.zip with templates only
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                    try:
                        template_bytes = template_file_c.read()
                        tpl_zip = _populate_template_bytes(template_bytes, cleaned_confirmed, costs_df=None)
                        with zipfile.ZipFile(tpl_zip, 'r') as tplzf:
                            for info in tplzf.infolist():
                                zf.writestr(info.filename, tplzf.read(info.filename))
                    except Exception as e:
                        st.exception(RuntimeError(f"Template population failed: {e}"))

                zip_buf.seek(0)
                st.success(f"Done. Generated {cleaned_confirmed['Provider Name'].nunique()} provider template(s).")

                st.download_button(
                    "Download Now!",
                    data=zip_buf.getvalue(),
                    file_name="confirmed_results.zip",
                    mime="application/zip",
                    key="dl_confirmed"
                )
