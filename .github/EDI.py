import io
import re
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st

st.set_page_config(page_title="EDI to CSV App!", page_icon="📄", layout="wide")

st.title("Aviva COMTFR EDI to CSV converter")
st.caption("Uploads raw OpenText/EDIFACT files with no extension and converts Aviva CHD commission lines to CSV.")


def read_uploaded_file(uploaded_file) -> str:
    """Read an uploaded file even when the browser reports it as application/octet-stream."""
    data = uploaded_file.read()
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="ignore")


def clean_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"^End-of-Header:\s*\n", "", text, flags=re.IGNORECASE | re.MULTILINE)
    # Some copied files include whitespace/newlines inside the EDI payload. Keep spaces in names,
    # but remove line breaks that can split segments.
    return text.replace("\n", "")


def split_edifact(value: str, delimiter: str) -> List[str]:
    """Split EDIFACT text on a delimiter, ignoring delimiters escaped with ?.

    Example: O?'LOUGHLIN will not be split at the apostrophe because the
    apostrophe is released/escaped by ?.
    """
    if value is None:
        return []

    parts = []
    current = []
    release_next = False

    for char in str(value):
        if release_next:
            current.append(char)
            release_next = False
            continue

        if char == "?":
            current.append(char)
            release_next = True
            continue

        if char == delimiter:
            parts.append("".join(current))
            current = []
        else:
            current.append(char)

    parts.append("".join(current))
    return parts


def parse_segments(text: str) -> List[str]:
    text = clean_text(text)
    return [seg.strip() for seg in split_edifact(text, "'") if seg.strip()]


def tokenise(segment: str) -> Tuple[str, List[str]]:
    parts = split_edifact(segment, "+")
    return parts[0].strip(), parts[1:]


def parse_composite(value: str) -> List[str]:
    return split_edifact(value, ":") if value is not None else []


def strip_leading_zeros(value: Optional[str]) -> str:
    if value is None:
        return ""
    value = str(value)
    stripped = value.lstrip("0")
    return stripped if stripped else ("0" if value else "")


def parse_date(value: Optional[str]) -> str:
    if not value:
        return ""
    value = str(value)
    if re.fullmatch(r"\d{8}", value):
        return f"{value[0:4]}-{value[4:6]}-{value[6:8]}"
    if re.fullmatch(r"\d{6}", value):
        return f"20{value[0:2]}-{value[2:4]}-{value[4:6]}"
    return value




def unescape_edifact(value: Optional[str]) -> str:
    """Remove EDIFACT release characters from text values, especially names.

    In EDIFACT, ? is used to escape reserved characters. For example:
    O?'LOUGHLIN becomes O'LOUGHLIN.
    """
    if value is None:
        return ""
    return (
        str(value)
        .replace("??", "?")
        .replace("?'", "'")
        .replace("?:", ":")
        .replace("?+", "+")
        .replace("?\r", "")
        .replace("?\n", "")
    )


def format_pounds_from_pence(value: Optional[str]) -> str:
    """Convert pence values from the EDI file into a GBP display value."""
    if value is None or str(value).strip() == "":
        return ""
    cleaned = str(value).strip().replace(",", "")
    try:
        amount = float(cleaned) / 100
    except ValueError:
        return str(value)
    return f"£{amount:,.2f}"


def build_aviva_output(df: pd.DataFrame) -> pd.DataFrame:
    """Return the Aviva statement columns requested by the team."""
    if df.empty:
        return df

    required_columns = [
        "provider_detected", "nad_pa", "policy_reference", "surname", "forename",
        "amount_qualifier", "amount", "premium_amount"
    ]
    for col in required_columns:
        if col not in df.columns:
            df[col] = ""

    output = pd.DataFrame({
        "Provider Name": df["provider_detected"].replace({"Aviva (tentative: seen in sample)": "Aviva"}),
        "Agency Code": df["nad_pa"],
        "Policy Reference": df["policy_reference"],
        "Surname": df["surname"],
        "First Name": df["forename"],
        "Product Version": df["amount_qualifier"],
        "Amount": df["amount"].apply(format_pounds_from_pence),
        "Premium Amount": df["premium_amount"].apply(format_pounds_from_pence),
    })
    return output


def parse_unb(fields: List[str]) -> Dict[str, str]:
    # UNB+UNOA:1+sender+recipient+YYMMDD:HHMM+control_ref...
    out = {
        "syntax": "", "sender": "", "recipient": "", "datetime_raw": "", "control_ref": ""
    }
    if len(fields) > 0:
        out["syntax"] = fields[0]
    if len(fields) > 1:
        out["sender"] = fields[1]
    if len(fields) > 2:
        out["recipient"] = strip_leading_zeros(fields[2])
    if len(fields) > 3:
        out["datetime_raw"] = fields[3]
    if len(fields) > 4:
        out["control_ref"] = fields[4]
    return out


def parse_bgm(fields: List[str]) -> Dict[str, str]:
    out = {"bgm_code": "", "bgm_reference": "", "payment_date_raw": "", "payment_date": ""}
    if len(fields) > 0:
        out["bgm_code"] = fields[0]
    if len(fields) > 1:
        out["bgm_reference"] = fields[1]
    for field in fields:
        parts = parse_composite(field)
        if len(parts) >= 2 and parts[0] == "PYD":
            out["payment_date_raw"] = parts[1]
            out["payment_date"] = parse_date(parts[1])
    return out


def parse_nad(fields: List[str]) -> Tuple[str, str]:
    if len(fields) < 2:
        return "", ""
    qual, value = fields[0], fields[1]
    if qual == "IN":
        value = strip_leading_zeros(value)
    return qual, value


def parse_rff(fields: List[str]) -> Dict[str, str]:
    refs = {}
    for field in fields:
        if ":" in field:
            key, value = field.split(":", 1)
            refs[key] = value
    return refs


def parse_pol(fields: List[str]) -> Dict[str, str]:
    out = {
        "basis_of_sale": "", "party_qualifier": "", "name_format": "",
        "surname": "", "forename": "", "full_name": "", "product_code": ""
    }

    # Common examples:
    # L&G: POL+59+PH+U:SMITH A+XYZ
    # Aviva sample: POL++PH+F:Miller:Paul
    if fields and re.fullmatch(r"\d{2}", fields[0] or ""):
        out["basis_of_sale"] = fields[0]

    for i, field in enumerate(fields):
        if field in ("PH", "MR", "PA", "IN"):
            out["party_qualifier"] = field
            if i + 1 < len(fields):
                # Names can contain EDIFACT release characters, e.g. O?'LOUGHLIN:JOSEPH.
                # Unescape before splitting so escaped reserved characters are treated as text.
                name_field = unescape_edifact(fields[i + 1])
                name_parts = parse_composite(name_field)
                if len(name_parts) >= 1:
                    out["name_format"] = unescape_edifact(name_parts[0])
                if len(name_parts) >= 2:
                    out["surname"] = unescape_edifact(name_parts[1])
                if len(name_parts) >= 3:
                    out["forename"] = unescape_edifact(name_parts[2])
                out["full_name"] = " ".join([p for p in [out["forename"], out["surname"]] if p]) or out["surname"]
            break

    # Use the final short alphanumeric token as product code if present and not the name field.
    if fields:
        last = fields[-1]
        if last and ":" not in last and last not in ("PH", "MR", "PA", "IN") and not re.fullmatch(r"\d{2}", last):
            out["product_code"] = last
    return out


def parse_chd(fields: List[str]) -> Dict[str, str]:
    out = {
        "amount_qualifier": "", "amount": "", "currency": "",
        "currency_qualifier": "", "charge_type": "", "due_date_raw": "", "due_date": "",
        "due_date_format": "", "premium_type": "", "premium_amount": "", "premium_currency": ""
    }
    if fields:
        c516 = parse_composite(fields[0])
        if len(c516) > 0:
            out["amount_qualifier"] = c516[0]
        if len(c516) > 1:
            out["amount"] = c516[1]
        if len(c516) > 2:
            out["currency"] = c516[2]

    if len(fields) > 1:
        c876 = parse_composite(fields[1])
        if len(c876) > 0:
            out["currency_qualifier"] = c876[0]
        if len(c876) > 1:
            out["charge_type"] = c876[1]

    cdd_idx = None
    for idx, field in enumerate(fields[2:], start=2):
        if field.startswith("CDD:"):
            cdd_idx = idx
            cdd = parse_composite(field)
            if len(cdd) > 1:
                out["due_date_raw"] = cdd[1]
                out["due_date"] = parse_date(cdd[1])
            if len(cdd) > 2:
                out["due_date_format"] = cdd[2]
            break

    if cdd_idx is not None and cdd_idx + 1 < len(fields):
        prem = parse_composite(fields[cdd_idx + 1])
        if len(prem) > 0:
            out["premium_type"] = strip_leading_zeros(prem[0])
        if len(prem) > 1:
            out["premium_amount"] = prem[1]
        if len(prem) > 2:
            out["premium_currency"] = prem[2]
    return out


def parse_pdt(fields: List[str]) -> Dict[str, str]:
    return {
        "pdt_code_1": fields[0] if len(fields) > 0 else "",
        "pdt_code_2": strip_leading_zeros(fields[1]) if len(fields) > 1 else "",
    }


def parse_cnt(fields: List[str]) -> Dict[str, str]:
    out = {"cnt_count_type": "", "cnt_count": "", "cnt_amount_type": "", "cnt_amount": ""}
    for field in fields:
        parts = parse_composite(field)
        if len(parts) >= 2:
            if parts[0] == "CTN":
                out["cnt_count_type"] = parts[0]
                out["cnt_count"] = parts[1]
            elif parts[0] == "CAM":
                out["cnt_amount_type"] = parts[0]
                out["cnt_amount"] = parts[1]
    return out


def detect_provider(nads: Dict[str, str], sender: str, recipient: str) -> str:
    # We do not have reliable provider tags yet. This gives a helpful hint without depending on it.
    # Add new mappings here once OpenText/provider identifiers are confirmed.
    known_ids = {
        "649443": "Aviva",
    }
    for value in [nads.get("BO", ""), nads.get("PA", ""), sender, recipient]:
        if value in known_ids:
            return known_ids[value]
    return "Unknown"


def convert_comtfr(text: str) -> pd.DataFrame:
    rows: List[Dict[str, str]] = []
    segments = parse_segments(text)

    envelope = {"syntax": "", "sender": "", "recipient": "", "datetime_raw": "", "control_ref": ""}
    message = {}
    nads: Dict[str, str] = {}
    current_gis = ""
    current_policy = ""
    current_ifn_by_policy: Dict[str, str] = {}
    current_pol: Dict[str, str] = {}
    message_row_indexes: Dict[str, List[int]] = {}

    def reset_message():
        return {
            "unh_number": "", "message_type": "", "bgm_code": "", "bgm_reference": "",
            "payment_date_raw": "", "payment_date": "", "unt_segment_count": "", "unt_control_ref": ""
        }

    message = reset_message()

    for segment in segments:
        tag, fields = tokenise(segment)

        if tag == "UNB":
            envelope = parse_unb(fields)

        elif tag == "UNH":
            message = reset_message()
            nads = {}
            current_gis = ""
            current_policy = ""
            current_ifn_by_policy = {}
            current_pol = {}
            if len(fields) > 0:
                message["unh_number"] = fields[0]
            if len(fields) > 1:
                message["message_type"] = fields[1]
            message_row_indexes.setdefault(message["unh_number"], [])

        elif tag == "BGM":
            message.update(parse_bgm(fields))

        elif tag == "NAD":
            qual, value = parse_nad(fields)
            if qual:
                nads[qual] = value

        elif tag == "GIS":
            current_gis = fields[0] if fields else ""

        elif tag == "RFF":
            refs = parse_rff(fields)
            if "POL" in refs:
                current_policy = refs["POL"]
                current_pol = {}
            if "IFN" in refs and current_policy:
                current_ifn_by_policy[current_policy] = refs["IFN"]

        elif tag == "POL":
            current_pol = parse_pol(fields)

        elif tag == "CHD":
            chd = parse_chd(fields)
            row = {
                "provider_detected": detect_provider(nads, envelope.get("sender", ""), envelope.get("recipient", "")),
                "interchange_control_ref": envelope.get("control_ref", ""),
                "unb_sender": envelope.get("sender", ""),
                "unb_recipient": envelope.get("recipient", ""),
                "unb_datetime_raw": envelope.get("datetime_raw", ""),
                "unh_number": message.get("unh_number", ""),
                "message_type": message.get("message_type", ""),
                "bgm_code": message.get("bgm_code", ""),
                "bgm_reference": message.get("bgm_reference", ""),
                "payment_date_raw": message.get("payment_date_raw", ""),
                "payment_date": message.get("payment_date", ""),
                "nad_in": nads.get("IN", ""),
                "nad_bo": nads.get("BO", ""),
                "nad_pa": nads.get("PA", ""),
                "gis_code": current_gis,
                "policy_reference": current_policy,
                "ifn_reference": current_ifn_by_policy.get(current_policy, ""),
                **current_pol,
                **chd,
                "pdt_code_1": "",
                "pdt_code_2": "",
                "cnt_count": "",
                "cnt_amount": "",
                "unt_segment_count": "",
                "unt_control_ref": "",
            }
            rows.append(row)
            message_row_indexes.setdefault(message.get("unh_number", ""), []).append(len(rows) - 1)

        elif tag == "PDT":
            pdt = parse_pdt(fields)
            indexes = message_row_indexes.get(message.get("unh_number", ""), [])
            if indexes:
                rows[indexes[-1]].update(pdt)

        elif tag == "CNT":
            cnt = parse_cnt(fields)
            for idx in message_row_indexes.get(message.get("unh_number", ""), []):
                rows[idx]["cnt_count"] = cnt.get("cnt_count", "")
                rows[idx]["cnt_amount"] = cnt.get("cnt_amount", "")

        elif tag == "UNT":
            count = fields[0] if len(fields) > 0 else ""
            control = fields[1] if len(fields) > 1 else ""
            for idx in message_row_indexes.get(message.get("unh_number", ""), []):
                rows[idx]["unt_segment_count"] = count
                rows[idx]["unt_control_ref"] = control

    return pd.DataFrame(rows)


with st.expander("What this app accepts", expanded=True):
    st.write(
        "This uploader deliberately has no file-type restriction, so files whose type only shows as "
        "`file` or `application/octet-stream` are accepted. It currently parses COMTFR-style EDIFACT "
        "data and outputs the Aviva statement fields requested by the team."
    )

uploaded_files = st.file_uploader(
    "Upload raw OpenText file(s)",
    type=None,
    accept_multiple_files=True,
    help="No extension is required. TXT, EDI, DAT, and extensionless files are all accepted.",
)

if uploaded_files:
    all_frames = []
    for uploaded in uploaded_files:
        text = read_uploaded_file(uploaded)
        df_one = convert_comtfr(text)
        if not df_one.empty:
            df_one.insert(0, "source_file", uploaded.name)
            all_frames.append(df_one)

    if not all_frames:
        st.warning("No CHD commission rows were found. Check whether the file is COMTFR EDIFACT and contains CHD segments.")
    else:
        df = pd.concat(all_frames, ignore_index=True)
        output_df = build_aviva_output(df)
        st.success(f"Parsed {len(output_df):,} commission row(s) from {len(uploaded_files)} file(s).")

        st.subheader("Preview")
        st.dataframe(output_df, use_container_width=True)

        csv_bytes = output_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Download CSV",
            data=csv_bytes,
            file_name="aviva_comtfr_converted.csv",
            mime="text/csv",
            type="primary",
        )

        with st.expander("Column notes"):
            st.markdown(
                "- This version outputs the Aviva statement fields requested by the team only.\n"
                "- One CSV row is created for each `CHD` commission/charge line.\n"
                "- `Amount` and `Premium Amount` are converted from pence to pounds and formatted with a £ symbol."
            )
else:
    st.info("Upload one or more raw files to convert them.")
