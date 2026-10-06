"""2026–27 Marketing Order Form generator.

Run with: streamlit run MOF26-27.py
Upload the Zoho export, the 2027 MOF Cost Sheet, and the 2027 order form template.
"""

from __future__ import annotations

import re
import zipfile
from io import BytesIO

import pandas as pd
from openpyxl.utils import get_column_letter
from xml.dom import minidom


START_ROW = 15
BASE_LAST_ROW = 43
UNSELECTED = {"", "no", "false", "0", "none", "n/a", "not selected", "unchecked", "nan"}
MONTHS = {"jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "sept", "oct", "nov", "dec",
          "january", "february", "march", "april", "june", "july", "august", "september", "october",
          "november", "december", "q1", "q2", "q3", "q4"}
FULL_MONTHS = {
    "jan": "January", "feb": "February", "mar": "March", "apr": "April", "may": "May",
    "jun": "June", "jul": "July", "aug": "August", "sep": "September",
    "sept": "September", "oct": "October", "nov": "November", "dec": "December",
}
NOTES = (
    "Please provide any feedback on our Marketing & Opportunities 2027 Pack and webinar:",
    "Please provide any further notes you may have or want to have considered with this form:",
)
CONTACTS = {
    "Random ID": ("Random ID",),
    "Provider Name": ("Provider Name",),
    "Name": ("Your Name", "Name"),
    "Phone": ("Phone",),
    "Email": ("Email",),
    "Events Name": ("Contact Name for Events Correspondence", "Events Name"),
    "Events Email": ("Contact Email for Events Correspondence", "Events Email"),
    "Marketing Publications Name": ("Contact Name for Marketing Correspondence", "Marketing Publications Name"),
    "Marketing Publications Email": ("Contact Email for Marketing Correspondence", "Marketing Publications Email"),
    "Copy Name": ("Copy Name",),
    "Copy Email": ("Copy Email",),
    "When To Invoice": ("When To Invoice",),
    "Invoice Name": ("Invoice Name",),
    "Invoice Email": ("Invoice Email",),
    "Added Time": ("Added Time",),
}


def clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u00a0", " ")).strip()


def key(value) -> str:
    """Match spacing/case and a hyphen with variable surrounding spaces."""
    return re.sub(r"\s*-\s*", "-", clean_text(value)).casefold()


def choices(value: str) -> list[str]:
    """Split commas between choices, never commas inside parentheses."""
    parts, current, depth = [], [], 0
    for ch in clean_text(value):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        if ch == "," and depth == 0:
            part = "".join(current).strip()
            if part:
                parts.append(part)
            current = []
        else:
            current.append(ch)
    part = "".join(current).strip()
    if part:
        parts.append(part)
    return parts


def quantity(value: str) -> int | None:
    s = clean_text(value)
    match = re.match(r"^(?:quantity\s*[:=]?\s*)?(\d+)(?:\.0+)?(?:\s*(?:ticket|tickets|place|places))?$", s, re.I)
    if match:
        return int(match.group(1))
    return None


def full_month(value: str) -> str:
    value = clean_text(value)
    return FULL_MONTHS.get(value.casefold(), value)


def read_form(upload) -> pd.DataFrame:
    raw = upload.getvalue() if hasattr(upload, "getvalue") else upload
    filename = getattr(upload, "name", "export.csv").lower()
    if filename.endswith(".xlsx"):
        df = pd.read_excel(BytesIO(raw), dtype=str, keep_default_na=False)
    else:
        try:
            df = pd.read_csv(BytesIO(raw), dtype=str, keep_default_na=False, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(BytesIO(raw), dtype=str, keep_default_na=False, encoding="cp1252")
    if df.empty or "Provider Name" not in df.columns:
        raise ValueError("The Zoho export must include its column headings and at least one submission.")
    # Zoho exports a second row containing the subheading for each multi-choice field.
    first = df.iloc[0]
    is_subheader = clean_text(first.get("Provider Name", "")) == "" and any(
        clean_text(first.get(c, "")) for c in df.columns if c not in CONTACTS
    )
    df.attrs["subheaders"] = first.copy() if is_subheader else pd.Series("", index=df.columns)
    df.attrs["subheader_row"] = bool(is_subheader)
    return df


def read_costs(upload) -> pd.DataFrame:
    raw = upload.getvalue() if hasattr(upload, "getvalue") else upload
    if getattr(upload, "name", "").lower().endswith((".csv", ".txt")):
        df = pd.read_csv(BytesIO(raw))
    else:
        df = pd.read_excel(BytesIO(raw), sheet_name=0)
    df.columns = [clean_text(c) for c in df.columns]
    if not {"Type", "Event", "Cost"}.issubset(df.columns):
        raise ValueError("The cost file needs Type, Event, and Cost columns.")
    df = df.loc[df["Type"].notna() & df["Event"].notna(), ["Type", "Event", "Cost"]].copy()
    df["Type"] = df["Type"].map(clean_text)
    df["Event"] = df["Event"].map(clean_text)
    df["Cost"] = pd.to_numeric(df["Cost"], errors="coerce")
    df = df.loc[(df["Type"] != "") & (df["Event"] != "")].reset_index(drop=True)
    if df.empty:
        raise ValueError("The cost file has no usable Type/Event rows.")
    df["_key"] = list(zip(df["Type"].map(key), df["Event"].map(key)))
    return df


def _cost_match(costs: pd.DataFrame, typ: str, event: str):
    matches = costs.loc[costs["_key"].map(lambda x: x == (key(typ), key(event)))]
    if matches.empty:
        return typ, event, None, "No Cost match"
    first = matches.iloc[0]
    prices = matches["Cost"].dropna().unique()
    status = ("No price in Cost Sheet" if not len(prices) else
              "Matched" if len(prices) == 1 else "Conflicting prices in Cost Sheet")
    return first["Type"], first["Event"], first["Cost"] if pd.notna(first["Cost"]) else None, status


def _specialist_type(event: str) -> str:
    e = key(event)
    if "protection specialist masterclass" in e:
        return "Specialist Events (Protection Masterclass)"
    if "income protection summit" in e:
        return "Specialist Events (Income Protection)"
    if "beyond prime" in e:
        return "Specialist Events (Beyond Prime)"
    if "buy to let" in e or "buy-to-let" in e:
        return "Specialist Events (Buy-to-Let)"
    return "Specialist Events"


def _month_choice(value: str) -> bool:
    return all(key(x) in MONTHS for x in choices(value))


def _get(row: pd.Series, *names: str) -> str:
    for name in names:
        if name in row.index and clean_text(row[name]):
            return clean_text(row[name])
    return ""


def transform_wishlist(form: pd.DataFrame, costs: pd.DataFrame) -> pd.DataFrame:
    subheaders = form.attrs.get("subheaders", pd.Series("", index=form.columns))
    submissions = form.iloc[1:] if form.attrs.get("subheader_row") else form
    contact_names = {n for variants in CONTACTS.values() for n in variants}
    ignore = contact_names | set(NOTES) | {"Referrer Name", "Task Owner"}
    records = []
    cost_by_type = {}
    for _, c in costs.iterrows():
        cost_by_type.setdefault(key(c["Type"]), []).append(c["Event"])

    for submission_no, (_, row) in enumerate(submissions.iterrows(), start=1):
        contact = {name: _get(row, *variants) for name, variants in CONTACTS.items()}
        dt = pd.to_datetime(contact["Added Time"], errors="coerce", dayfirst=False)
        contact["Added Time"] = dt.strftime("%d/%m/%Y") if pd.notna(dt) else contact["Added Time"]
        parent = ""

        def add(typ: str, event: str, date: str = "", qty: int = 1, source: str = ""):
            if not typ or not event:
                return
            matched_type, matched_event, price, status = _cost_match(costs, typ, event)
            records.append({
                "Submission": submission_no, **contact,
                "Type": matched_type, "Event": matched_event,
                "Event Date (if applicable)": full_month(date), "Qty": qty,
                "Cost": price, "Line Total": price * qty if price is not None else None,
                "Match Status": status, "Zoho Field": source,
                "_note_q1": _get(row, NOTES[0]), "_note_q2": _get(row, NOTES[1]),
            })

        for col in form.columns:
            is_unnamed = str(col).startswith("Unnamed:")
            if not is_unnamed and col not in ignore:
                parent = clean_text(col)
            if col in ignore or not parent:
                continue
            value = clean_text(row[col])
            if key(value) in UNSELECTED:
                continue
            sub = clean_text(subheaders.get(col, ""))
            source = f"{parent} / {sub}" if sub else parent

            if parent == "Additional Gala Ticket" or parent == "Additional Daytime Ticket":
                add("National Training Event & Awards Gala Dinner", parent, qty=quantity(value) or 1, source=source)
            elif parent.startswith("Regional Sales Roadshows"):
                typ = f"{parent} - {sub}" if sub else parent
                for option in choices(value):
                    if key(option) != "select":
                        add(typ, option, source=source)
            elif parent == "Specialist Events":
                for option in choices(value):
                    dated = re.search(r"\(([^()]*)\)\s*$", option)
                    add(_specialist_type(option), sub,
                        date=dated.group(1) if dated else option, source=source)
            elif parent.startswith("Sales & Development Webinars"):
                qty = quantity(value) or 1
                add("Sales & Development Webinars (20 Minute Presentation)",
                    "Presentation slot (20 minute presentation)", qty=qty, source=source)
            elif parent == "Business Leader Growth Forums - Formerly Peer Group Meetings":
                for option in choices(value):
                    add("Business Leader Growth Forums", option, source=source)
            elif key(value) == "select":
                # In these Zoho columns, the option is the subheading and Select means checked.
                add(parent, sub or (cost_by_type.get(key(parent)) or [parent])[0], source=source)
            elif sub and _month_choice(value):
                # Each chosen month or quarter is a separate order item.
                for month in choices(value):
                    add(parent, sub, date=month, source=source)
            elif not sub and quantity(value) is not None:
                add(parent, (cost_by_type.get(key(parent)) or [parent])[0], qty=quantity(value), source=source)
            elif sub and parent == "The Right Academy Induction Courses":
                for option in choices(value):
                    add(parent, option, source=source)
            elif sub and key(value) in {key(x) for x in cost_by_type.get(key(parent), [])}:
                add(parent, value, source=source)
            elif sub and parent in {"National Training Event & Awards Gala Dinner",
                                     "Summit & Gala Dinner: Private Medical Insurance",
                                     "Later Life Lending Workshop - May",
                                     "Later Life Lending Conference & Gala Dinner - December",
                                     "The Right DA Club Annual Conference",
                                     "Accreditation and Reaccreditation Events"}:
                add(parent, sub, source=source)
            elif sub:
                for option in choices(value):
                    add(parent, option, source=source)
            else:
                for option in choices(value):
                    add(parent, option, source=source)

    return pd.DataFrame.from_records(records)


# The supplied template contains Excel array-formula metadata, shared formulas,
# shapes and x14 validations. openpyxl rewrites these on save. Preserve the
# original OOXML package and edit only the relevant worksheet XML instead.
SPARE_ROWS = 50
CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _direct(parent, tag):
    return next((n for n in parent.childNodes if n.nodeType == n.ELEMENT_NODE and n.tagName == tag), None)


def _children(parent, tag):
    return [n for n in parent.childNodes if n.nodeType == n.ELEMENT_NODE and n.tagName == tag]


def _put_text(parent, tag, value):
    node = _direct(parent, tag)
    if node is None:
        node = parent.ownerDocument.createElement(tag)
        parent.appendChild(node)
    while node.firstChild:
        node.removeChild(node.firstChild)
    node.appendChild(parent.ownerDocument.createTextNode(str(value)))
    return node


def _column_number(letter):
    number = 0
    for ch in letter:
        number = number * 26 + ord(ch) - ord("A") + 1
    return number


def _get_row(sheet_data, number):
    for row in _children(sheet_data, "row"):
        rr = int(row.getAttribute("r"))
        if rr == number:
            return row
        if rr > number:
            new = sheet_data.ownerDocument.createElement("row")
            new.setAttribute("r", str(number))
            sheet_data.insertBefore(new, row)
            return new
    new = sheet_data.ownerDocument.createElement("row")
    new.setAttribute("r", str(number))
    sheet_data.appendChild(new)
    return new


def _get_cell(sheet_data, address):
    col, rr = CELL_RE.fullmatch(address).groups()
    row = _get_row(sheet_data, int(rr))
    target = _column_number(col)
    for cell in _children(row, "c"):
        cc = _column_number(CELL_RE.fullmatch(cell.getAttribute("r")).group(1))
        if cc == target:
            return cell
        if cc > target:
            new = sheet_data.ownerDocument.createElement("c")
            new.setAttribute("r", address)
            row.insertBefore(new, cell)
            return new
    new = sheet_data.ownerDocument.createElement("c")
    new.setAttribute("r", address)
    row.appendChild(new)
    return new


def _set_cell(sheet_data, address, value, *, formula=False, style=None, array=False):
    cell = _get_cell(sheet_data, address)
    for child in list(cell.childNodes):
        cell.removeChild(child)
    for attr in ("t", "cm", "vm"):
        if cell.hasAttribute(attr):
            cell.removeAttribute(attr)
    if style is not None:
        cell.setAttribute("s", str(style))
    if value is None or value == "":
        return
    if formula:
        f = _put_text(cell, "f", value.lstrip("="))
        if array:
            f.setAttribute("t", "array")
            f.setAttribute("ref", address)
            cell.setAttribute("cm", "1")
        cell.appendChild(cell.ownerDocument.createElement("v"))
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        _put_text(cell, "v", value)
    else:
        cell.setAttribute("t", "inlineStr")
        inline = cell.ownerDocument.createElement("is")
        t = _put_text(inline, "t", value)
        if str(value) != str(value).strip():
            t.setAttribute("xml:space", "preserve")
        cell.appendChild(inline)


def _formula_e(rr):
    return (f'IF(OR(A{rr}="",B{rr}=""),"",IFERROR(_xlfn.XLOOKUP(1,'
            f"('Cost Sheet'!$A$2:$A$1000=A{rr})*('Cost Sheet'!$B$2:$B$1000=B{rr}),"
            "'Cost Sheet'!$C$2:$C$1000),\"\"))")


def _shift_footer(sheet_data, delta):
    if not delta:
        return
    for row in reversed(_children(sheet_data, "row")):
        rr = int(row.getAttribute("r"))
        if rr < 45:
            continue
        row.setAttribute("r", str(rr + delta))
        for cell in _children(row, "c"):
            old = cell.getAttribute("r")
            col = CELL_RE.fullmatch(old).group(1)
            cell.setAttribute("r", f"{col}{rr + delta}")


def _named_ranges(wb_doc, types, events_by_type):
    defined = wb_doc.getElementsByTagName("definedNames")[0]
    for name in list(_children(defined, "definedName")):
        if name.getAttribute("name").startswith("MOF"):
            defined.removeChild(name)
    def add(name, target):
        node = wb_doc.createElement("definedName")
        node.setAttribute("name", name)
        node.appendChild(wb_doc.createTextNode(target))
        defined.appendChild(node)
    add("MOFTypes", f"'Cost Sheet'!$J$2:$J${len(types)+1}")
    add("MOFEmpty", "'Cost Sheet'!$J$1000")
    for i, events in enumerate(events_by_type, start=1):
        col = get_column_letter(i + 10)
        add(f"MOFEvents_{i:03d}", f"'Cost Sheet'!${col}$2:${col}${len(events)+1}")
    for node in _children(defined, "definedName"):
        if node.getAttribute("name") == "_xlnm.Print_Area" and node.getAttribute("localSheetId") == "1":
            # Updated by populate_template once the footer's new row is known.
            return node
    return None


def _cost_helpers(cost_doc, costs):
    root = cost_doc.documentElement
    sheet_data = _direct(root, "sheetData")
    for rr in range(2, max(103, len(costs) + 2)):
        for col in "ABC":
            _set_cell(sheet_data, f"{col}{rr}", None)
    for rr, (_, rec) in enumerate(costs.iterrows(), start=2):
        _set_cell(sheet_data, f"A{rr}", rec["Type"])
        _set_cell(sheet_data, f"B{rr}", rec["Event"])
        _set_cell(sheet_data, f"C{rr}", float(rec["Cost"]) if pd.notna(rec["Cost"]) else None)
    types = sorted(costs["Type"].unique(), key=str.casefold)
    events_by_type = []
    _set_cell(sheet_data, "J1", "Type dropdown")
    for i, typ in enumerate(types, start=2):
        _set_cell(sheet_data, f"J{i}", typ)
        events = costs.loc[costs["Type"] == typ, "Event"].drop_duplicates().tolist()
        events_by_type.append(events)
        col = get_column_letter(i + 9)
        for rr, event in enumerate(events, start=2):
            _set_cell(sheet_data, f"{col}{rr}", event)
    cols = _direct(root, "cols")
    hidden = cost_doc.createElement("col")
    hidden.setAttribute("min", "10")
    hidden.setAttribute("max", str(10 + len(types)))
    hidden.setAttribute("hidden", "1")
    hidden.setAttribute("width", "8")
    hidden.setAttribute("customWidth", "1")
    cols.appendChild(hidden)
    _direct(root, "dimension").setAttribute("ref", f"A1:{get_column_letter(10+len(types))}{max(102,len(costs)+1)}")
    return types, events_by_type


def _data_validations(sheet_doc, last_item):
    root = sheet_doc.documentElement
    # Keep the template's x14 invoice dropdown at D4. The old A/B validations
    # reference one Event list for every row; replace those with standard ones.
    ext = _direct(root, "extLst")
    if ext is not None:
        for data in ext.getElementsByTagName("x14:dataValidations"):
            for dv in list(_children(data, "x14:dataValidation")):
                sqref = dv.getElementsByTagName("xm:sqref")
                if sqref and sqref[0].firstChild and sqref[0].firstChild.nodeValue.startswith(("A15", "B15")):
                    data.removeChild(dv)
            data.setAttribute("count", str(len(_children(data, "x14:dataValidation"))))
    old = _direct(root, "dataValidations")
    if old is not None:
        root.removeChild(old)
    dvs = sheet_doc.createElement("dataValidations")
    dvs.setAttribute("count", "2")
    for rng, source in [
        (f"A15:A{last_item}", "MOFTypes"),
        (f"B15:B{last_item}",
         'INDIRECT(IFERROR("MOFEvents_"&TEXT(MATCH($A15,MOFTypes,0),"000"),"MOFEmpty"))'),
    ]:
        dv = sheet_doc.createElement("dataValidation")
        dv.setAttribute("type", "list")
        dv.setAttribute("allowBlank", "1")
        dv.setAttribute("showErrorMessage", "1")
        dv.setAttribute("sqref", rng)
        _put_text(dv, "formula1", source)
        dvs.appendChild(dv)
    # OOXML worksheet order places dataValidations before pageMargins.
    root.insertBefore(dvs, _direct(root, "pageMargins"))


def _set_notes(notes_doc, first):
    data = _direct(notes_doc.documentElement, "sheetData")
    for addr, field in [("A2", NOTES[0]), ("A3", NOTES[1]),
                        ("B2", first.get("_note_q1", "")), ("B3", first.get("_note_q2", ""))]:
        _set_cell(data, addr, field)
    _direct(notes_doc.documentElement, "dimension").setAttribute("ref", "A1:H3")


def populate_template(template_bytes: bytes, rows: pd.DataFrame, costs: pd.DataFrame) -> bytes:
    with zipfile.ZipFile(BytesIO(template_bytes)) as original:
        data = {name: original.read(name) for name in original.namelist()}
        info = {z.filename: z for z in original.infolist()}
    required = ("xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml",
                "xl/worksheets/sheet4.xml", "xl/workbook.xml")
    if not all(name in data for name in required):
        raise ValueError("The supplied 2027 template has a different sheet structure.")
    sheet = minidom.parseString(data["xl/worksheets/sheet2.xml"])
    cost_sheet = minidom.parseString(data["xl/worksheets/sheet4.xml"])
    notes = minidom.parseString(data["xl/worksheets/sheet3.xml"])
    workbook = minidom.parseString(data["xl/workbook.xml"])

    n = len(rows)
    last_item = max(44, START_ROW + n + SPARE_ROWS - 1)
    delta = last_item - 44
    root = sheet.documentElement
    sheet_data = _direct(root, "sheetData")
    _shift_footer(sheet_data, delta)
    first = rows.iloc[0] if n else {}
    fields = {"B4": "Provider Name", "B6": "Name", "D4": "When To Invoice",
              "D6": "Phone", "F6": "Email", "B10": "Events Name", "B12": "Events Email",
              "D10": "Marketing Publications Name", "D12": "Marketing Publications Email",
              "F10": "Invoice Name", "F12": "Invoice Email", "H10": "Copy Name", "H12": "Copy Email"}
    for address, field in fields.items():
        _set_cell(sheet_data, address, first.get(field, "") if n else "")
    for offset in range(last_item - START_ROW + 1):
        rr = START_ROW + offset
        item = rows.iloc[offset] if offset < n else None
        for col, value in (
            ("A", item["Type"] if item is not None else None),
            ("B", item["Event"] if item is not None else None),
            ("C", "CHECK COST" if item is not None and item["Match Status"] != "Matched" else None),
            ("D", item["Event Date (if applicable)"] if item is not None else None),
            ("F", int(item["Qty"]) if item is not None else None),
        ):
            _set_cell(sheet_data, f"{col}{rr}", value, style=6 if rr > 44 else None)
        if rr >= 44:
            _set_cell(sheet_data, f"E{rr}", _formula_e(rr), formula=True, style=32, array=True)
            _set_cell(sheet_data, f"G{rr}", f'IF(F{rr}="","",E{rr}*F{rr})', formula=True, style=33)
    footer = 45 + delta
    _set_cell(sheet_data, f"H{footer}", f"SUM(G{START_ROW}:G{last_item})", formula=True)
    _set_cell(sheet_data, f"H{footer+4}", f"H{footer}-H{footer+2}", formula=True)
    _set_cell(sheet_data, f"H{footer+6}", f"H{footer+4}/5", formula=True)
    _set_cell(sheet_data, f"H{footer+8}", f"H{footer+4}+H{footer+6}", formula=True)
    _direct(root, "dimension").setAttribute("ref", f"A1:I{footer+8}")
    _data_validations(sheet, last_item)

    types, events = _cost_helpers(cost_sheet, costs)
    area = _named_ranges(workbook, types, events)
    if area is not None:
        while area.firstChild:
            area.removeChild(area.firstChild)
        area.appendChild(workbook.createTextNode(f"'Marketing Order Form'!$A$2:$H${footer+8}"))
    calc = workbook.getElementsByTagName("calcPr")
    if calc:
        calc[0].setAttribute("fullCalcOnLoad", "1")
        calc[0].setAttribute("forceFullCalc", "1")
    _set_notes(notes, first)
    data["xl/worksheets/sheet2.xml"] = sheet.toxml(encoding="utf-8")
    data["xl/worksheets/sheet3.xml"] = notes.toxml(encoding="utf-8")
    data["xl/worksheets/sheet4.xml"] = cost_sheet.toxml(encoding="utf-8")
    data["xl/workbook.xml"] = workbook.toxml(encoding="utf-8")
    # A formula chain from the old template has stale footer coordinates.
    data.pop("xl/calcChain.xml", None)
    rels_path = "xl/_rels/workbook.xml.rels"
    if rels_path in data:
        rels = minidom.parseString(data[rels_path])
        for rel in list(rels.getElementsByTagName("Relationship")):
            if rel.getAttribute("Type").endswith("/calcChain"):
                rel.parentNode.removeChild(rel)
        data[rels_path] = rels.toxml(encoding="utf-8")
    ct_path = "[Content_Types].xml"
    ct = minidom.parseString(data[ct_path])
    for node in list(ct.getElementsByTagName("Override")):
        if node.getAttribute("PartName") == "/xl/calcChain.xml":
            node.parentNode.removeChild(node)
    data[ct_path] = ct.toxml(encoding="utf-8")
    output = BytesIO()
    with zipfile.ZipFile(output, "w") as z:
        for name, contents in data.items():
            z.writestr(info[name], contents)
    return output.getvalue()


def result_zip(form: pd.DataFrame, costs: pd.DataFrame, template_bytes: bytes):
    clean = transform_wishlist(form, costs)
    if clean.empty:
        raise ValueError("No selected items were found in this Zoho export.")
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        clean_export = clean.drop(columns=["_note_q1", "_note_q2"])
        bio = BytesIO()
        clean_export.to_excel(bio, index=False)
        archive.writestr("data/cleaned_output.xlsx", bio.getvalue())
        for submission, group in clean.groupby("Submission", sort=True):
            provider = clean_text(group.iloc[0]["Provider Name"]) or "Unknown Provider"
            name = clean_text(group.iloc[0]["Name"]) or "Unknown Contact"
            safe = lambda s: re.sub(r"[^\w .-]", "_", s)[:75]
            filename = f"templates/{submission:03d} - {safe(provider)} - {safe(name)} - WISHLIST.xlsx"
            archive.writestr(filename, populate_template(template_bytes, group, costs))
    return output.getvalue(), clean_export


def main():
    import streamlit as st
    st.set_page_config(page_title="MOF 26–27", layout="wide")
    st.title("MOF 26–27")
    st.write("Upload the Zoho Forms export, 2027 cost sheet, and your 2027 Excel template. "
             "The ZIP contains one workbook per submission and a cleaned data file.")
    a, b, c = st.columns(3)
    form_file = a.file_uploader("Zoho Forms export", type=["csv", "xlsx"])
    costs_file = b.file_uploader("MOF Cost Sheet 2027", type=["xlsx", "csv"])
    template_file = c.file_uploader("Wishlist MSA Order Form Template 2027", type=["xlsx"])
    if st.button("Generate order forms", type="primary"):
        if not all((form_file, costs_file, template_file)):
            st.error("Please upload all three files.")
            return
        try:
            form = read_form(form_file)
            costs = read_costs(costs_file)
            data, clean = result_zip(form, costs, template_file.getvalue())
        except Exception as exc:
            st.exception(exc)
            return
        st.success(f"Created {clean['Submission'].nunique()} order form(s) from {len(clean)} item lines.")
        missing = clean.loc[clean["Match Status"] != "Matched"]
        if not missing.empty:
            st.warning(f"{len(missing)} line(s) need a cost-sheet review. They remain in the output and are marked CHECK COST.")
            st.dataframe(missing[["Submission", "Provider Name", "Type", "Event", "Event Date (if applicable)",
                                  "Match Status", "Zoho Field"]], hide_index=True)
        with st.expander("Preview cleaned lines"):
            st.dataframe(clean, hide_index=True)
        st.download_button("Download results ZIP", data=data, file_name="MOF26-27-results.zip",
                           mime="application/zip")


if __name__ == "__main__":
    main()
