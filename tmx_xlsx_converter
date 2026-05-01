import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from datetime import datetime
from io import BytesIO
import os

# ============================================================
# Streamlit page setup
# ============================================================

st.set_page_config(
    page_title="LangOps Converter",
    page_icon="🔁",
    layout="centered"
)

# ============================================================
# Language options
# ============================================================

SOURCE_LANGUAGES = {
    "English (US) - en-US": "en-US",
    "English (UK) - en-GB": "en-GB",
    "German (Germany) - de-DE": "de-DE",
    "German (Switzerland) - de-CH": "de-CH",
    "French (France) - fr-FR": "fr-FR",
    "French (Canada) - fr-CA": "fr-CA",
    "Spanish (Spain) - es-ES": "es-ES",
    "Spanish (Mexico) - es-MX": "es-MX",
    "Spanish (Worldwide) - es-WW": "es-WW",
    "Italian - it-IT": "it-IT",
    "Portuguese (Brazil) - pt-BR": "pt-BR",
    "Portuguese (Portugal) - pt-PT": "pt-PT",
    "Dutch - nl-NL": "nl-NL",
    "Japanese - ja-JP": "ja-JP",
    "Korean - ko-KR": "ko-KR",
    "Chinese (Simplified) - zh-CN": "zh-CN",
    "Chinese (Traditional) - zh-TW": "zh-TW",
    "Arabic - ar-SA": "ar-SA",
    "Polish - pl-PL": "pl-PL",
    "Czech - cs-CZ": "cs-CZ",
    "Swedish - sv-SE": "sv-SE",
    "Danish - da-DK": "da-DK",
    "Norwegian - nb-NO": "nb-NO",
    "Finnish - fi-FI": "fi-FI",
    "Turkish - tr-TR": "tr-TR",
    "Russian - ru-RU": "ru-RU",
}

TARGET_LANGUAGES = {
    "German (Germany) - de-DE": "de-DE",
    "German (Switzerland) - de-CH": "de-CH",
    "French (France) - fr-FR": "fr-FR",
    "French (Canada) - fr-CA": "fr-CA",
    "Spanish (Spain) - es-ES": "es-ES",
    "Spanish (Mexico) - es-MX": "es-MX",
    "Spanish (Worldwide) - es-WW": "es-WW",
    "Arabic - ar-SA": "ar-SA",
    "Korean - ko-KR": "ko-KR",
    "Japanese - ja-JP": "ja-JP",
    "Chinese (Simplified) - zh-CN": "zh-CN",
    "Chinese (Traditional) - zh-TW": "zh-TW",
    "Italian - it-IT": "it-IT",
    "Portuguese (Brazil) - pt-BR": "pt-BR",
    "Portuguese (Portugal) - pt-PT": "pt-PT",
    "Dutch - nl-NL": "nl-NL",
    "Polish - pl-PL": "pl-PL",
    "Czech - cs-CZ": "cs-CZ",
    "Swedish - sv-SE": "sv-SE",
    "Danish - da-DK": "da-DK",
    "Norwegian - nb-NO": "nb-NO",
    "Finnish - fi-FI": "fi-FI",
    "Turkish - tr-TR": "tr-TR",
    "Russian - ru-RU": "ru-RU",
    "English (US) - en-US": "en-US",
    "English (UK) - en-GB": "en-GB",
}

DEFAULT_SOURCE = "English (US) - en-US"
DEFAULT_TARGET = "German (Germany) - de-DE"

XML_LANG = "{http://www.w3.org/XML/1998/namespace}lang"

# ============================================================
# TMX templates
# ============================================================

TMX_TEMPLATE = """<?xml version="1.0" encoding="utf-8"?>
<tmx version="1.4">
    <header creationtool="LangOps Converter"
            creationtoolversion="1.0"
            segtype="block"
            o-tmf="OTC"
            adminlang="en-US"
            srclang="{source_lang}"
            datatype="unknown"
            creationdate="{creationdate}" />
    <body>
{entries}
    </body>
</tmx>"""

TU_TEMPLATE = """        <tu>
            <prop type="Txt::Domain">sales_central</prop>
            <prop type="Txt::Product">sales_transcreation</prop>
            <prop type="Txt::Origin">{origin}</prop>
            <tuv xml:lang="{source_lang}">
                <seg>{source}</seg>
            </tuv>
            <tuv xml:lang="{target_lang}">
                <seg>{target}</seg>
            </tuv>
        </tu>"""

# ============================================================
# Helpers
# ============================================================

def normalize_lang(lang_code):
    return str(lang_code).strip().lower()


def lang_matches(requested, found):
    requested = normalize_lang(requested)
    found = normalize_lang(found)

    if requested == found:
        return True

    if found.startswith(requested + "-") or requested.startswith(found + "-"):
        return True

    return False


def get_seg_text(seg_elem):
    if seg_elem is None:
        return ""
    return "".join(seg_elem.itertext()).strip()


def load_input_file(uploaded_file):
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file)

    if file_name.endswith(".csv"):
        encodings_to_try = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
        separators_to_try = [None, ",", ";", "\t", "|"]

        content = uploaded_file.getvalue()

        last_error = None

        for enc in encodings_to_try:
            for sep in separators_to_try:
                try:
                    buffer = BytesIO(content)
                    if sep is None:
                        df = pd.read_csv(buffer, encoding=enc, sep=None, engine="python")
                    else:
                        df = pd.read_csv(buffer, encoding=enc, sep=sep)
                    return df
                except Exception as e:
                    last_error = e

        raise ValueError(f"CSV file could not be read. Last error: {last_error}")

    raise ValueError("Unsupported file format. Please upload an .xlsx or .csv file.")


def get_source_target_columns(df):
    original_columns = list(df.columns)
    normalized_columns = [str(col).strip().lower() for col in original_columns]

    source_idx = None
    target_idx = None

    source_candidates = ["source", "src", "english", "en", "source text"]
    target_candidates = ["target", "trg", "translation", "german", "de", "target text"]

    for i, col in enumerate(normalized_columns):
        if source_idx is None and col in source_candidates:
            source_idx = i
        if target_idx is None and col in target_candidates:
            target_idx = i

    if source_idx is None or target_idx is None:
        if len(df.columns) < 2:
            raise ValueError(
                "The file must contain at least 2 columns or headers named 'source' and 'target'."
            )
        source_idx = 0
        target_idx = 1

    return df.iloc[:, source_idx], df.iloc[:, target_idx]


# ============================================================
# XLSX / CSV -> TMX
# ============================================================

def create_tmx(uploaded_file, source_lang, target_lang):
    df = load_input_file(uploaded_file)
    source_col, target_col = get_source_target_columns(df)

    entries = []
    base_filename = uploaded_file.name

    for index, (source_val, target_val) in enumerate(zip(source_col, target_col), start=1):
        source = "" if pd.isna(source_val) else str(source_val).strip()
        target = "" if pd.isna(target_val) else str(target_val).strip()

        source = escape(source)
        target = escape(target)

        origin = f"imported_from/{base_filename}/row{index}.spl"

        entry = TU_TEMPLATE.format(
            origin=escape(origin),
            source_lang=source_lang,
            target_lang=target_lang,
            source=source,
            target=target,
        )
        entries.append(entry)

    creationdate = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    tmx_content = TMX_TEMPLATE.format(
        source_lang=source_lang,
        creationdate=creationdate,
        entries="\n".join(entries),
    )

    return tmx_content.encode("utf-8")


# ============================================================
# TMX -> XLSX
# ============================================================

def tmx_to_xlsx(uploaded_file, source_lang, target_lang):
    tree = ET.parse(uploaded_file)
    root = tree.getroot()

    tus = root.findall(".//tu")

    if not tus:
        raise ValueError("No translation units <tu> were found in the TMX file.")

    rows = []

    for tu in tus:
        source_text = ""
        target_text = ""

        tuvs = tu.findall("./tuv")

        for tuv in tuvs:
            tuv_lang = tuv.attrib.get(XML_LANG, "").strip()
            seg = tuv.find("./seg")
            seg_text = get_seg_text(seg)

            if not source_text and lang_matches(source_lang, tuv_lang):
                source_text = seg_text

            if not target_text and lang_matches(target_lang, tuv_lang):
                target_text = seg_text

        if source_text or target_text:
            rows.append({
                "source": source_text,
                "target": target_text
            })

    if not rows:
        raise ValueError(
            f"No matching translation units found for source '{source_lang}' and target '{target_lang}'."
        )

    df = pd.DataFrame(rows, columns=["source", "target"])

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="TMX Export")

        ws = writer.sheets["TMX Export"]
        ws.column_dimensions["A"].width = 60
        ws.column_dimensions["B"].width = 60

    output.seek(0)
    return output.getvalue()


# ============================================================
# Streamlit UI
# ============================================================

st.title("🔁 LangOps Converter")
st.caption("Convert XLSX / CSV files to TMX or convert TMX files back to XLSX.")

mode = st.radio(
    "Conversion Mode",
    ["XLSX / CSV → TMX", "TMX → XLSX"],
    horizontal=True
)

col1, col2 = st.columns(2)

with col1:
    source_label = st.selectbox(
        "Source Language",
        list(SOURCE_LANGUAGES.keys()),
        index=list(SOURCE_LANGUAGES.keys()).index(DEFAULT_SOURCE)
    )

with col2:
    target_label = st.selectbox(
        "Target Language",
        list(TARGET_LANGUAGES.keys()),
        index=list(TARGET_LANGUAGES.keys()).index(DEFAULT_TARGET)
    )

source_lang = SOURCE_LANGUAGES[source_label]
target_lang = TARGET_LANGUAGES[target_label]

st.divider()

if mode == "XLSX / CSV → TMX":
    uploaded_file = st.file_uploader(
        "Upload XLSX or CSV file",
        type=["xlsx", "csv"]
    )

    st.info("Preferred headers: source and target. If no matching headers exist, Column A and Column B will be used.")

    if uploaded_file is not None:
        if st.button("Create TMX", type="primary"):
            try:
                tmx_bytes = create_tmx(uploaded_file, source_lang, target_lang)

                st.success("TMX file created successfully.")

                st.download_button(
                    label="Download TMX",
                    data=tmx_bytes,
                    file_name="ready_to_import.tmx",
                    mime="application/xml"
                )

            except Exception as e:
                st.error(f"An error occurred: {e}")

else:
    uploaded_file = st.file_uploader(
        "Upload TMX file",
        type=["tmx"]
    )

    st.info("The output will be an Excel file with two columns: source and target.")

    if uploaded_file is not None:
        if st.button("Create XLSX", type="primary"):
            try:
                xlsx_bytes = tmx_to_xlsx(uploaded_file, source_lang, target_lang)

                base_name = os.path.splitext(uploaded_file.name)[0]
                output_name = f"{base_name}_source_target.xlsx"

                st.success("XLSX file created successfully.")

                st.download_button(
                    label="Download XLSX",
                    data=xlsx_bytes,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except ET.ParseError as e:
                st.error(f"The TMX file could not be parsed: {e}")

            except Exception as e:
                st.error(f"An error occurred: {e}")

st.divider()
st.caption("LangOps Converter · Streamlit Cloud version")
