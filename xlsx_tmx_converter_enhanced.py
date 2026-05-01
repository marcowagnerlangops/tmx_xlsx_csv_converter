import customtkinter as ctk
import tkinter.filedialog as fd
from tkinter import messagebox
import pandas as pd
import os
from datetime import datetime
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

# ============================================================
# LangOps Branding / Theme
# ============================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_BG = "#0D1117"
CARD_BG = "#161B22"
INPUT_BG = "#0F141A"
BORDER = "#2F3A4A"
TEXT_MAIN = "#F3F4F6"
TEXT_MUTED = "#9CA3AF"
ACCENT = "#3B82F6"
ACCENT_HOVER = "#2563EB"
SUCCESS = "#10B981"

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

XML_LANG = "{http://www.w3.org/XML/1998/namespace}lang"

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

    # Match broader code against regional code, e.g. de == de-DE
    if found.startswith(requested + "-") or requested.startswith(found + "-"):
        return True

    return False

def get_seg_text(seg_elem):
    if seg_elem is None:
        return ""
    return "".join(seg_elem.itertext()).strip()

# ============================================================
# File loading helpers
# ============================================================

def load_input_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".xlsx":
        return pd.read_excel(file_path)

    if ext == ".csv":
        encodings_to_try = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
        separators_to_try = [None, ",", ";", "\t", "|"]

        last_error = None
        for enc in encodings_to_try:
            for sep in separators_to_try:
                try:
                    if sep is None:
                        df = pd.read_csv(file_path, encoding=enc, sep=None, engine="python")
                    else:
                        df = pd.read_csv(file_path, encoding=enc, sep=sep)
                    return df
                except Exception as e:
                    last_error = e

        raise ValueError(f"CSV file could not be read. Last error: {last_error}")

    raise ValueError("Unsupported file format. Please select an .xlsx, .csv, or .tmx file.")

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

def create_tmx(file_path, output_folder, source_lang, target_lang):
    try:
        df = load_input_file(file_path)
        source_col, target_col = get_source_target_columns(df)

        entries = []
        base_filename = os.path.basename(file_path)

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

        output_file = os.path.join(output_folder, "ready_to_import.tmx")
        with open(output_file, "w", encoding="utf-8", newline="\n") as tmx_file:
            tmx_file.write(tmx_content)

        messagebox.showinfo("Success", f"TMX file created successfully:\n\n{output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n\n{e}")

# ============================================================
# TMX -> XLSX
# ============================================================

def tmx_to_xlsx(file_path, output_folder, source_lang, target_lang):
    try:
        tree = ET.parse(file_path)
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

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_file = os.path.join(output_folder, f"{base_name}_source_target.xlsx")

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="TMX Export")

            ws = writer.sheets["TMX Export"]
            ws.column_dimensions["A"].width = 60
            ws.column_dimensions["B"].width = 60

        messagebox.showinfo("Success", f"XLSX file created successfully:\n\n{output_file}")

    except ET.ParseError as e:
        messagebox.showerror("TMX Parse Error", f"The TMX file could not be parsed:\n\n{e}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n\n{e}")

# ============================================================
# GUI actions
# ============================================================

def select_file():
    current_mode = mode_var.get()

    if current_mode == "XLSX / CSV → TMX":
        filetypes = [
            ("Supported files", "*.xlsx *.csv"),
            ("Excel files", "*.xlsx"),
            ("CSV files", "*.csv"),
        ]
    else:
        filetypes = [
            ("TMX files", "*.tmx"),
        ]

    file_path = fd.askopenfilename(filetypes=filetypes)
    if file_path:
        file_entry.delete(0, ctk.END)
        file_entry.insert(0, file_path)

def select_folder():
    folder_path = fd.askdirectory()
    if folder_path:
        folder_entry.delete(0, ctk.END)
        folder_entry.insert(0, folder_path)

def update_mode_ui(selected_mode):
    file_entry.delete(0, ctk.END)

    if selected_mode == "XLSX / CSV → TMX":
        title_label.configure(text="LangOps XLSX / CSV to TMX Converter")
        subtitle_label.configure(
            text="Convert XLSX or CSV files into TMX format using source/target headers or column A/B fallback"
        )
        input_label.configure(text="Select Input File")
        info_label.configure(
            text="Accepted inputs: XLSX, CSV | Preferred headers: source, target | Fallback: first 2 columns"
        )
        action_button.configure(text="Create TMX")
    else:
        title_label.configure(text="LangOps TMX to XLSX Converter")
        subtitle_label.configure(
            text="Convert TMX files into Excel files with source and target columns"
        )
        input_label.configure(text="Select TMX File")
        info_label.configure(
            text="Accepted input: TMX | Output: Excel with columns 'source' and 'target'"
        )
        action_button.configure(text="Create XLSX")

def start_processing():
    file_path = file_entry.get().strip()
    output_folder = folder_entry.get().strip()

    if not os.path.isfile(file_path):
        messagebox.showwarning("Missing file", "Please select a valid input file.")
        return

    if not os.path.isdir(output_folder):
        messagebox.showwarning("Missing folder", "Please select a valid destination folder.")
        return

    source_lang = SOURCE_LANGUAGES[source_lang_var.get()]
    target_lang = TARGET_LANGUAGES[target_lang_var.get()]
    current_mode = mode_var.get()

    if current_mode == "XLSX / CSV → TMX":
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in [".xlsx", ".csv"]:
            messagebox.showwarning("Invalid file", "Please select an XLSX or CSV file.")
            return
        create_tmx(file_path, output_folder, source_lang, target_lang)

    elif current_mode == "TMX → XLSX":
        ext = os.path.splitext(file_path)[1].lower()
        if ext != ".tmx":
            messagebox.showwarning("Invalid file", "Please select a TMX file.")
            return
        tmx_to_xlsx(file_path, output_folder, source_lang, target_lang)

# ============================================================
# GUI
# ============================================================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("LangOps Converter")
        self.geometry("760x800")
        self.configure(fg_color=APP_BG)

        self.main_frame = ctk.CTkFrame(
            self,
            fg_color=CARD_BG,
            corner_radius=18,
            border_width=1,
            border_color=BORDER
        )
        self.main_frame.pack(padx=24, pady=24, fill="both", expand=True)

        self.mode_var = ctk.StringVar(value="XLSX / CSV → TMX")
        self.source_lang_var = ctk.StringVar(value=DEFAULT_SOURCE)
        self.target_lang_var = ctk.StringVar(value=DEFAULT_TARGET)

        self.build_ui()

    def build_ui(self):
        mode_label = ctk.CTkLabel(
            self.main_frame,
            text="Conversion Mode",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=TEXT_MAIN
        )
        mode_label.pack(pady=(20, 8))

        self.mode_selector = ctk.CTkSegmentedButton(
            self.main_frame,
            values=["XLSX / CSV → TMX", "TMX → XLSX"],
            variable=self.mode_var,
            command=self.update_mode_ui,
            height=38,
            fg_color=INPUT_BG,
            selected_color=ACCENT,
            selected_hover_color=ACCENT_HOVER,
            unselected_color=INPUT_BG,
            unselected_hover_color="#1A2230",
            text_color="white"
        )
        self.mode_selector.pack(padx=30, fill="x")

        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="LangOps XLSX / CSV to TMX Converter",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=TEXT_MAIN
        )
        self.title_label.pack(pady=(20, 6))

        self.subtitle_label = ctk.CTkLabel(
            self.main_frame,
            text="Convert XLSX or CSV files into TMX format using source/target headers or column A/B fallback",
            font=ctk.CTkFont(size=13),
            text_color=TEXT_MUTED
        )
        self.subtitle_label.pack(pady=(0, 20))

        self.input_label = ctk.CTkLabel(
            self.main_frame,
            text="Select Input File",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=TEXT_MAIN
        )
        self.input_label.pack(anchor="w", padx=30, pady=(6, 6))

        file_row = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        file_row.pack(fill="x", padx=30, pady=(0, 14))

        self.file_entry = ctk.CTkEntry(
            file_row,
            width=520,
            height=38,
            fg_color=INPUT_BG,
            border_color=BORDER,
            text_color=TEXT_MAIN
        )
        self.file_entry.pack(side="left", padx=(0, 10), fill="x", expand=True)

        ctk.CTkButton(
            file_row,
            text="Browse",
            width=120,
            height=38,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            text_color="white",
            command=self.select_file
        ).pack(side="left")

        ctk.CTkLabel(
            self.main_frame,
            text="Select Destination Folder",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=TEXT_MAIN
        ).pack(anchor="w", padx=30, pady=(6, 6))

        folder_row = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        folder_row.pack(fill="x", padx=30, pady=(0, 14))

        self.folder_entry = ctk.CTkEntry(
            folder_row,
            width=520,
            height=38,
            fg_color=INPUT_BG,
            border_color=BORDER,
            text_color=TEXT_MAIN
        )
        self.folder_entry.pack(side="left", padx=(0, 10), fill="x", expand=True)

        ctk.CTkButton(
            folder_row,
            text="Browse",
            width=120,
            height=38,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            text_color="white",
            command=self.select_folder
        ).pack(side="left")

        lang_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        lang_frame.pack(fill="x", padx=30, pady=(8, 10))

        source_col = ctk.CTkFrame(lang_frame, fg_color="transparent")
        source_col.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkLabel(
            source_col,
            text="Source Language",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=TEXT_MAIN
        ).pack(anchor="w", pady=(0, 6))

        self.source_menu = ctk.CTkOptionMenu(
            source_col,
            values=list(SOURCE_LANGUAGES.keys()),
            variable=self.source_lang_var,
            height=38,
            fg_color=ACCENT,
            button_color=ACCENT,
            button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=CARD_BG,
            dropdown_hover_color="#1F2937",
            text_color="white"
        )
        self.source_menu.pack(fill="x")

        target_col = ctk.CTkFrame(lang_frame, fg_color="transparent")
        target_col.pack(side="left", fill="x", expand=True, padx=(10, 0))

        ctk.CTkLabel(
            target_col,
            text="Target Language",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=TEXT_MAIN
        ).pack(anchor="w", pady=(0, 6))

        self.target_menu = ctk.CTkOptionMenu(
            target_col,
            values=list(TARGET_LANGUAGES.keys()),
            variable=self.target_lang_var,
            height=38,
            fg_color=ACCENT,
            button_color=ACCENT,
            button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=CARD_BG,
            dropdown_hover_color="#1F2937",
            text_color="white"
        )
        self.target_menu.pack(fill="x")

        self.info_label = ctk.CTkLabel(
            self.main_frame,
            text="Accepted inputs: XLSX, CSV | Preferred headers: source, target | Fallback: first 2 columns",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED
        )
        self.info_label.pack(pady=(18, 8))

        self.action_button = ctk.CTkButton(
            self.main_frame,
            text="Create TMX",
            command=self.start_processing,
            height=42,
            width=220,
            fg_color=SUCCESS,
            hover_color="#059669",
            text_color="white",
            font=ctk.CTkFont(size=15, weight="bold")
        )
        self.action_button.pack(pady=(8, 24))

    def select_file(self):
        current_mode = self.mode_var.get()

        if current_mode == "XLSX / CSV → TMX":
            filetypes = [
                ("Supported files", "*.xlsx *.csv"),
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
            ]
        else:
            filetypes = [("TMX files", "*.tmx")]

        file_path = fd.askopenfilename(filetypes=filetypes)
        if file_path:
            self.file_entry.delete(0, ctk.END)
            self.file_entry.insert(0, file_path)

    def select_folder(self):
        folder_path = fd.askdirectory()
        if folder_path:
            self.folder_entry.delete(0, ctk.END)
            self.folder_entry.insert(0, folder_path)

    def update_mode_ui(self, selected_mode):
        self.file_entry.delete(0, ctk.END)

        if selected_mode == "XLSX / CSV → TMX":
            self.title_label.configure(text="LangOps XLSX / CSV to TMX Converter")
            self.subtitle_label.configure(
                text="Convert XLSX or CSV files into TMX format using source/target headers or column A/B fallback"
            )
            self.input_label.configure(text="Select Input File")
            self.info_label.configure(
                text="Accepted inputs: XLSX, CSV | Preferred headers: source, target | Fallback: first 2 columns"
            )
            self.action_button.configure(text="Create TMX")
        else:
            self.title_label.configure(text="LangOps TMX to XLSX Converter")
            self.subtitle_label.configure(
                text="Convert TMX files into Excel files with source and target columns"
            )
            self.input_label.configure(text="Select TMX File")
            self.info_label.configure(
                text="Accepted input: TMX | Output: Excel with columns 'source' and 'target'"
            )
            self.action_button.configure(text="Create XLSX")

    def start_processing(self):
        file_path = self.file_entry.get().strip()
        output_folder = self.folder_entry.get().strip()

        if not os.path.isfile(file_path):
            messagebox.showwarning("Missing file", "Please select a valid input file.")
            return

        if not os.path.isdir(output_folder):
            messagebox.showwarning("Missing folder", "Please select a valid destination folder.")
            return

        source_lang = SOURCE_LANGUAGES[self.source_lang_var.get()]
        target_lang = TARGET_LANGUAGES[self.target_lang_var.get()]
        current_mode = self.mode_var.get()

        if current_mode == "XLSX / CSV → TMX":
            ext = os.path.splitext(file_path)[1].lower()
            if ext not in [".xlsx", ".csv"]:
                messagebox.showwarning("Invalid file", "Please select an XLSX or CSV file.")
                return
            create_tmx(file_path, output_folder, source_lang, target_lang)

        elif current_mode == "TMX → XLSX":
            ext = os.path.splitext(file_path)[1].lower()
            if ext != ".tmx":
                messagebox.showwarning("Invalid file", "Please select a TMX file.")
                return
            tmx_to_xlsx(file_path, output_folder, source_lang, target_lang)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()