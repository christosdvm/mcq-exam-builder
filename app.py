
import re
import random
import zipfile
import os
import tempfile
import hashlib
from io import BytesIO
from difflib import SequenceMatcher
from pathlib import Path
from typing import List, Dict, Tuple

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from PIL import Image
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from exam_versions import (
    POSITION_DEPENDENT_OPTION_RE,
    LETTERS,
    allowed_letters_for_count,
    answer_columns_for_count,
    build_version_questions,
    build_versions,
    clean_correct,
    has_position_dependent_options,
    normalise_cell,
    version_label,
)
from question_importer import (
    dataframe_for_editor,
    import_all_sources,
    insert_preserved_word_drawing,
    split_image_files,
    validation_issue_location,
    validate_reviewed_questions,
    word_drawing_fingerprint,
)
from source_formats import SUPPORTED_EXTENSIONS

APP_VERSION = "v15.4.0"
APP_AUTHOR = os.getenv("MCQ_BUILDER_AUTHOR", "Dr Christos I. Karagiannis").strip()

BASE_COLUMNS = ["No", "Question"]
IMAGE_COLUMN = "Image"
OPTION_PREFIX = "Option "
CORRECT_COLUMN = "Correct"
MAX_OPTIONS_SUPPORTED = 6
MAX_IMAGE_PIXELS = 50_000_000


st.set_page_config(
    page_title="MCQ Exam Builder",
    page_icon=str(Path(__file__).parent / "assets" / "favicon.png"),
    layout="wide",
    initial_sidebar_state="collapsed",
)


def inject_app_styles():
    """Keep the presentation layer separate from exam processing."""
    css = Path(__file__).with_name("ui_styles.css").read_text(encoding="utf-8")
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


# -----------------------------
# Template helpers
# -----------------------------

def make_safe_source_prefix(filename: str) -> str:
    name = filename.rsplit(".", 1)[0]
    name = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")
    return name or "MCQ_FILE"


def make_internal_id(source_file: str, no_value, row_number: int) -> str:
    prefix = make_safe_source_prefix(source_file)
    no_text = "" if pd.isna(no_value) else str(no_value).strip()
    if no_text:
        try:
            f = float(no_text)
            if f.is_integer():
                no_text = str(int(f))
        except (TypeError, ValueError):
            pass
        return f"{prefix}-{no_text}"
    return f"{prefix}-ROW{row_number}"


def make_template_xlsx(option_count: int = 6) -> BytesIO:
    """
    v15 universal template.
    Always provides Option 1–6.
    Lecturers leave unused option columns blank.
    """
    option_cols = answer_columns_for_count(MAX_OPTIONS_SUPPORTED)
    columns = BASE_COLUMNS + [IMAGE_COLUMN] + option_cols + [CORRECT_COLUMN]

    wb = Workbook()
    ws = wb.active
    ws.title = "MCQs"

    ws.append(columns)

    # A–D sample
    ws.append([
        "1",
        "Which structure is the functional unit of the kidney?",
        "",
        "Nephron",
        "Alveolus",
        "Hepatocyte",
        "Osteon",
        "",
        "",
        "A"
    ])

    # A–B sample
    ws.append([
        "2",
        "Melatonin is produced by the pineal gland.",
        "",
        "True",
        "False",
        "",
        "",
        "",
        "",
        "A"
    ])

    # A–F sample
    ws.append([
        "3",
        "Which of the following is a gland of the endocrine system?",
        "",
        "Thyroid",
        "Adrenal",
        "Pineal",
        "Pituitary",
        "Pancreas",
        "Ovary",
        "A"
    ])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    widths = {
        "A": 10,
        "B": 65,
        "C": 28,
        "D": 30,
        "E": 30,
        "F": 30,
        "G": 30,
        "H": 30,
        "I": 30,
        "J": 12,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in ws.iter_rows(min_row=2, max_row=500, min_col=1, max_col=len(columns)):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Correct column is J
    dv = DataValidation(type="list", formula1='"A,B,C,D,E,F"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("J2:J500")
    ws.freeze_panes = "A2"

    instructions = wb.create_sheet("Instructions")
    instructions["A1"] = "MCQ Exam Builder Universal Template v15"
    instructions["A1"].font = Font(size=16, bold=True, color="1F4E78")
    instructions["A3"] = "Fill one row per MCQ."
    instructions["A4"] = "The No column is optional. If left blank, the app will use the Excel row order."
    instructions["A5"] = "The Image column is optional. For multiple images, separate exact filenames with |."
    instructions["A6"] = "Use Option 1–6. Leave unused option columns blank."
    instructions["A7"] = "The app automatically detects how many options each question has."
    instructions["A8"] = "Examples:"
    instructions["A9"] = "2 filled options = A–B"
    instructions["A10"] = "4 filled options = A–D"
    instructions["A11"] = "6 filled options = A–F"
    instructions["A12"] = "Options must be consecutive. Do not fill Option 4 if Option 3 is blank."
    instructions["A13"] = "Correct must match the available options: A for Option 1, B for Option 2, etc."
    instructions["A15"] = "Columns:"
    instructions["A15"].font = Font(bold=True)

    for i, col in enumerate(columns, start=16):
        note = " (optional)" if col in ["No", IMAGE_COLUMN] else ""
        instructions[f"A{i}"] = col + note
    instructions.column_dimensions["A"].width = 110

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def validate_image_bytes(image_name: str, data: bytes) -> str:
    """Return an error message for invalid or excessively large raster images."""
    if not data:
        return f"Image file is empty: {image_name}"
    try:
        with Image.open(BytesIO(data)) as image:
            width, height = image.size
            if width <= 0 or height <= 0:
                return f"Image has invalid dimensions: {image_name}"
            if width * height > MAX_IMAGE_PIXELS:
                return f"Image exceeds the 50-megapixel safety limit: {image_name}"
            image.verify()
    except Exception as exc:  # noqa: BLE001 - Pillow exposes several decoder-specific exceptions.
        return f"Image cannot be read: {image_name} ({type(exc).__name__})"
    return ""


def detect_answer_columns(columns: List[str]) -> List[str]:
    """
    v11 detects universal template columns:
    Option 1, Option 2, ..., Option 6.

    For backwards compatibility, it can also read older Option 1, Option 2, etc.
    """
    columns_clean = [str(c).strip() for c in columns]

    option_found = []
    for col in columns_clean:
        m = re.fullmatch(r"Option\s*([1-9][0-9]*)", col, flags=re.IGNORECASE)
        if m:
            number = int(m.group(1))
            if 1 <= number <= MAX_OPTIONS_SUPPORTED:
                option_found.append((number, col))

    if option_found:
        option_found = sorted(option_found, key=lambda x: x[0])
        return [col for _, col in option_found]

    # Backwards compatibility with v3-v10 templates.
    answer_found = []
    for col in columns_clean:
        m = re.fullmatch(r"Answer_([A-Z])", col, flags=re.IGNORECASE)
        if m:
            letter = m.group(1).upper()
            if letter in LETTERS[:MAX_OPTIONS_SUPPORTED]:
                answer_found.append((LETTERS.index(letter) + 1, col))

    answer_found = sorted(answer_found, key=lambda x: x[0])
    return [col for _, col in answer_found]


def detect_active_option_count(row, option_cols: List[str]) -> Tuple[int, List[str]]:
    """
    Determines how many options are active for one question.
    Options must be consecutive from Option 1.
    """
    filled = [bool(normalise_cell(row.get(col, ""))) for col in option_cols]
    active_count = 0
    gap_found = False
    gap_details = []

    for i, is_filled in enumerate(filled):
        if is_filled and not gap_found:
            active_count += 1
        elif not is_filled:
            gap_found = True
        elif is_filled and gap_found:
            gap_details.append(option_cols[i])

    return active_count, gap_details

def read_mcq_excel(uploaded_file) -> Tuple[pd.DataFrame, List[Dict]]:
    df = pd.read_excel(uploaded_file)
    df.columns = [normalise_cell(c) for c in df.columns]

    issues = []
    filename = uploaded_file.name

    option_cols = detect_answer_columns(df.columns.tolist())

    missing_base = [col for col in ["Question", "Correct"] if col not in df.columns]
    if missing_base:
        issues.append({
            "Source file": filename,
            "Row": "",
            "Internal_ID": "",
            "Severity": "ERROR",
            "Issue": f"Missing required columns: {', '.join(missing_base)}"
        })
        return pd.DataFrame(), issues

    if len(option_cols) < 2:
        issues.append({
            "Source file": filename,
            "Row": "",
            "Internal_ID": "",
            "Severity": "ERROR",
            "Issue": "At least two option columns are required, e.g. Option 1 and Option 2."
        })
        return pd.DataFrame(), issues

    # Ensure uploaded template has option columns continuous from Option 1.
    expected_cols = answer_columns_for_count(len(option_cols))
    # For older Option 1 templates, do not force exact names; only force count/order.
    using_universal = all(str(c).lower().startswith("option") for c in option_cols)
    if using_universal and option_cols != expected_cols:
        issues.append({
            "Source file": filename,
            "Row": "",
            "Internal_ID": "",
            "Severity": "ERROR",
            "Issue": f"Option columns must be continuous from Option 1. Found: {', '.join(option_cols)}. Expected: {', '.join(expected_cols)}."
        })
        return pd.DataFrame(), issues

    if "No" not in df.columns:
        df["No"] = ""
    if IMAGE_COLUMN not in df.columns:
        df[IMAGE_COLUMN] = ""

    required_for_this_file = ["No", "Question", IMAGE_COLUMN] + option_cols + ["Correct"]
    df = df[required_for_this_file].copy()
    df["Source_File"] = filename

    df = df.dropna(how="all")

    internal_ids = []
    active_counts = []
    answer_columns_used = []

    for idx, row in df.iterrows():
        excel_row = int(idx) + 2
        internal_id = make_internal_id(filename, row.get("No", ""), excel_row)
        internal_ids.append(internal_id)

        question = normalise_cell(row["Question"])
        correct = clean_correct(row["Correct"])

        active_count, gap_details = detect_active_option_count(row, option_cols)
        active_counts.append(active_count)
        active_option_cols = option_cols[:active_count]
        answer_columns_used.append(",".join(active_option_cols))

        if not question:
            issues.append({
                "Source file": filename,
                "Row": excel_row,
                "Internal_ID": internal_id,
                "Severity": "ERROR",
                "Issue": "Missing Question text"
            })

        image_name = normalise_cell(row.get(IMAGE_COLUMN, ""))
        if image_name:
            valid_ext = image_name.lower().endswith((".png", ".jpg", ".jpeg"))
            if not valid_ext:
                issues.append({
                    "Source file": filename,
                    "Row": excel_row,
                    "Internal_ID": internal_id,
                    "Severity": "WARNING",
                    "Issue": f"Image filename '{image_name}' has unsupported extension. Use PNG, JPG, or JPEG."
                })

        if active_count < 2:
            issues.append({
                "Source file": filename,
                "Row": excel_row,
                "Internal_ID": internal_id,
                "Severity": "ERROR",
                "Issue": "At least two filled options are required."
            })

        if gap_details:
            issues.append({
                "Source file": filename,
                "Row": excel_row,
                "Internal_ID": internal_id,
                "Severity": "ERROR",
                "Issue": f"Options must be consecutive. Filled option(s) after a blank: {', '.join(gap_details)}."
            })

        allowed = allowed_letters_for_count(max(active_count, 0))
        if correct not in allowed:
            issues.append({
                "Source file": filename,
                "Row": excel_row,
                "Internal_ID": internal_id,
                "Severity": "ERROR",
                "Issue": f"Invalid Correct value: '{row['Correct']}'. This question has {active_count} option(s), so use one of: {', '.join(allowed) if allowed else 'none'}."
            })

        option_values = [normalise_cell(row[col]).lower() for col in active_option_cols if normalise_cell(row[col])]
        if len(option_values) != len(set(option_values)):
            issues.append({
                "Source file": filename,
                "Row": excel_row,
                "Internal_ID": internal_id,
                "Severity": "WARNING",
                "Issue": "Duplicate answer option text within the same question"
            })

        df.at[idx, "Correct"] = correct

    df["Internal_ID"] = internal_ids
    df["Option_Count"] = active_counts
    df["Answer_Columns"] = answer_columns_used

    # Normalize all option columns into Option 1..Option 6 for internal use.
    normalized = pd.DataFrame()
    normalized["No"] = df["No"]
    normalized["Question"] = df["Question"]
    normalized[IMAGE_COLUMN] = df[IMAGE_COLUMN]

    for i in range(MAX_OPTIONS_SUPPORTED):
        target_col = f"Option {i+1}"
        if i < len(option_cols):
            normalized[target_col] = df[option_cols[i]]
        else:
            normalized[target_col] = ""

    normalized["Correct"] = df["Correct"]
    normalized["Option_Count"] = df["Option_Count"]
    normalized["Answer_Columns"] = df["Answer_Columns"]
    normalized["Source_File"] = df["Source_File"]
    normalized["Internal_ID"] = df["Internal_ID"]

    final_cols = ["No", "Question", IMAGE_COLUMN] + answer_columns_for_count(MAX_OPTIONS_SUPPORTED) + [
        "Correct", "Option_Count", "Answer_Columns", "Source_File", "Internal_ID"
    ]

    return normalized[final_cols], issues

def validate_all(files) -> Tuple[pd.DataFrame, pd.DataFrame]:
    all_dfs, all_issues = [], []
    for uploaded_file in files:
        uploaded_file.seek(0)
        df, issues = read_mcq_excel(uploaded_file)
        if not df.empty:
            all_dfs.append(df)
        all_issues.extend(issues)

    if all_dfs:
        combined = pd.concat(all_dfs, ignore_index=True)
    else:
        combined = pd.DataFrame()

    if not combined.empty:
        ids = combined["Internal_ID"].astype(str).str.strip()
        duplicates = ids[ids.duplicated(keep=False) & ids.ne("")]
        for internal_id in sorted(set(duplicates)):
            rows = combined.index[ids == internal_id].tolist()
            sources = combined.loc[rows, "Source_File"].tolist()
            all_issues.append({
                "Source file": "; ".join(sources),
                "Row": "",
                "Internal_ID": internal_id,
                "Severity": "WARNING",
                "Issue": "Duplicate internal ID. This usually means duplicated No values in the same file or the same file uploaded twice."
            })

    report = pd.DataFrame(all_issues)
    if report.empty:
        report = pd.DataFrame(columns=["Source file", "Row", "Internal_ID", "Severity", "Issue"])
    return combined, report


def has_blocking_errors(report_df: pd.DataFrame) -> bool:
    return not report_df.empty and (report_df["Severity"] == "ERROR").any()


def normalise_validation_report(report_df: pd.DataFrame) -> pd.DataFrame:
    columns = ["Source file", "Row", "Internal_ID", "Severity", "Issue"]
    if report_df is None or report_df.empty:
        return pd.DataFrame(columns=["Severity", "Location", "Issue", "Source file", "Row", "Internal_ID"])
    output = report_df.copy()
    for column in columns:
        if column not in output.columns:
            output[column] = ""
        output[column] = output[column].fillna("").astype(str)
    output["Location"] = output.apply(validation_issue_location, axis=1)
    return output[["Severity", "Location", "Issue", "Source file", "Row", "Internal_ID"]]


def make_preview_fingerprint(frame: pd.DataFrame, settings_payload, image_map: Dict[str, bytes]) -> str:
    """Return a stable signature so previews cannot silently outlive their inputs."""
    digest = hashlib.sha256()
    if frame is not None and not frame.empty:
        normalised = frame.fillna("").astype(str)
        digest.update(pd.util.hash_pandas_object(normalised, index=True).values.tobytes())
    digest.update(repr(settings_payload).encode("utf-8"))
    for image_name in sorted(image_map):
        digest.update(image_name.encode("utf-8"))
        digest.update(word_drawing_fingerprint(image_map[image_name]))
    return digest.hexdigest()


# -----------------------------
# Formatting preference helpers
# -----------------------------

def get_default_formatting():
    return {
        "font_name": "Arial",
        "font_size": 12,
        "title_font_size": 14,
        "answer_key_font_size": 10,
        "top_margin_cm": 1.7,
        "bottom_margin_cm": 1.7,
        "left_margin_cm": 1.8,
        "right_margin_cm": 1.8,
        "option_indent_cm": 0.5,
        "space_before_question_pt": 6,
        "space_after_question_pt": 3,
        "space_after_option_pt": 0,
        "blank_line_between_questions": True,
        "bold_questions": True,
        "show_page_numbers": True,
        "image_width_cm": 12.0,
        "show_version_in_header": True,
        "show_source_in_answer_key": True,
        "answer_display": "Separate answer key only"
    }


def set_run_font(run, formatting: Dict, size_key: str = "font_size", bold=None, italic=None):
    run.font.name = formatting.get("font_name", "Arial")
    run.font.size = Pt(float(formatting.get(size_key, formatting.get("font_size", 12))))
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic




def set_paragraph_keep_together(paragraph):
    """Try to keep an MCQ paragraph with the following paragraph and avoid internal page splits."""
    paragraph.paragraph_format.keep_together = True
    paragraph.paragraph_format.keep_with_next = True
    pPr = paragraph._p.get_or_add_pPr()
    if pPr.find(qn("w:keepLines")) is None:
        pPr.append(OxmlElement("w:keepLines"))
    if pPr.find(qn("w:keepNext")) is None:
        pPr.append(OxmlElement("w:keepNext"))


def set_paragraph_keep_lines_only(paragraph):
    """Avoid splitting this paragraph internally, but allow the block to end after it."""
    paragraph.paragraph_format.keep_together = True
    pPr = paragraph._p.get_or_add_pPr()
    if pPr.find(qn("w:keepLines")) is None:
        pPr.append(OxmlElement("w:keepLines"))


# -----------------------------
# DOCX export helpers
# -----------------------------

def set_doc_defaults(document: Document, formatting: Dict):
    style = document.styles["Normal"]
    style.font.name = formatting.get("font_name", "Arial")
    style.font.size = Pt(float(formatting.get("font_size", 12)))
    rPr = style.element.rPr
    if rPr is not None and rPr.rFonts is not None:
        rPr.rFonts.set(qn("w:eastAsia"), formatting.get("font_name", "Arial"))


def add_page_number(paragraph):
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = paragraph.add_run()
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)


def add_doc_header(document: Document, title: str, version: str, formatting: Dict, include_answer_note: bool = False):
    p = document.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r = p.add_run(title)
    set_run_font(r, formatting, size_key="title_font_size", bold=True)

    if formatting.get("show_version_in_header", True):
        p2 = document.add_paragraph()
        p2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        r2 = p2.add_run(f"Version {version}")
        set_run_font(r2, formatting, size_key="font_size", bold=True)

    if include_answer_note:
        p3 = document.add_paragraph()
        p3.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        r3 = p3.add_run("Answer Key")
        set_run_font(r3, formatting, size_key="font_size", bold=True)

    document.add_paragraph("")


def add_image_to_document(document: Document, image_name: str, image_map: Dict[str, bytes], formatting: Dict):
    if not image_name or not image_map:
        return

    key = image_name.strip().lower()
    if key not in image_map:
        p = document.add_paragraph()
        run = p.add_run(f"[Image missing: {image_name}]")
        run.italic = True
        set_run_font(run, formatting, size_key="answer_key_font_size", italic=True)
        return

    image_data = image_map[key]
    p = document.add_paragraph()
    set_paragraph_keep_together(p)
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run()
    try:
        if insert_preserved_word_drawing(document, run, image_data):
            return
    except Exception:  # noqa: BLE001 - fall back to the validated preview raster.
        pass

    suffix = Path(image_name).suffix.lower() or ".png"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(image_data)
        tmp_path = tmp.name

    try:
        try:
            run.add_picture(tmp_path, width=Cm(float(formatting.get("image_width_cm", 12.0))))
        except Exception as exc:  # noqa: BLE001 - python-docx delegates to multiple image decoders.
            run.text = f"[Image could not be rendered: {image_name}]"
            set_run_font(run, formatting, size_key="answer_key_font_size", italic=True)
            run.add_text(f" ({type(exc).__name__})")
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def safe_excel_value(value):
    """Prevent user-controlled text from becoming an Excel formula."""
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def safe_excel_dataframe(frame: pd.DataFrame) -> pd.DataFrame:
    if frame is None:
        return pd.DataFrame()
    output = frame.copy()
    for column in output.columns:
        output[column] = output[column].map(safe_excel_value)
    return output


def create_exam_docx(questions: List[Dict], title: str, version: str, show_answers: bool = False, image_map: Dict[str, bytes] = None, formatting: Dict = None) -> BytesIO:
    formatting = formatting or get_default_formatting()
    document = Document()
    set_doc_defaults(document, formatting)
    section = document.sections[0]
    section.top_margin = Cm(float(formatting.get("top_margin_cm", 1.7)))
    section.bottom_margin = Cm(float(formatting.get("bottom_margin_cm", 1.7)))
    section.left_margin = Cm(float(formatting.get("left_margin_cm", 1.8)))
    section.right_margin = Cm(float(formatting.get("right_margin_cm", 1.8)))
    if formatting.get("show_page_numbers", True) and section.footer.paragraphs:
        add_page_number(section.footer.paragraphs[0])

    add_doc_header(document, title, version, formatting, include_answer_note=show_answers)

    for i, q in enumerate(questions, start=1):
        p = document.add_paragraph()
        set_paragraph_keep_together(p)
        p.paragraph_format.space_before = Pt(float(formatting.get("space_before_question_pt", 6)))
        p.paragraph_format.space_after = Pt(float(formatting.get("space_after_question_pt", 3)))
        run = p.add_run(f"{i}. {q['Question']}")
        set_run_font(run, formatting, size_key="font_size", bold=bool(formatting.get("bold_questions", True)))

        for image_name in q.get("Images") or split_image_files(q.get("Image", "")):
            add_image_to_document(document, image_name, image_map, formatting)

        for opt_index, opt in enumerate(q["Options"]):
            p_opt = document.add_paragraph()
            p_opt.paragraph_format.left_indent = Cm(float(formatting.get("option_indent_cm", 0.5)))
            p_opt.paragraph_format.space_after = Pt(float(formatting.get("space_after_option_pt", 0)))
            if opt_index < len(q["Options"]) - 1 or show_answers:
                set_paragraph_keep_together(p_opt)
            else:
                set_paragraph_keep_lines_only(p_opt)
            run_opt = p_opt.add_run(f"{opt['letter']}. {opt['text']}")
            set_run_font(run_opt, formatting, size_key="font_size")

        if show_answers:
            p_ans = document.add_paragraph()
            set_paragraph_keep_lines_only(p_ans)
            p_ans.paragraph_format.left_indent = Cm(float(formatting.get("option_indent_cm", 0.5)))
            run_ans = p_ans.add_run(f"ANSWER: {q['Correct']}")
            set_run_font(run_ans, formatting, size_key="font_size", bold=True)

        if formatting.get("blank_line_between_questions", True) and i < len(questions):
            document.add_paragraph("")

    output = BytesIO()
    document.save(output)
    output.seek(0)
    return output


def create_answer_key_docx(questions: List[Dict], title: str, version: str, formatting: Dict = None) -> BytesIO:
    formatting = formatting or get_default_formatting()
    document = Document()
    set_doc_defaults(document, formatting)
    section = document.sections[0]
    section.top_margin = Cm(float(formatting.get("top_margin_cm", 1.7)))
    section.bottom_margin = Cm(float(formatting.get("bottom_margin_cm", 1.7)))
    section.left_margin = Cm(float(formatting.get("left_margin_cm", 1.8)))
    section.right_margin = Cm(float(formatting.get("right_margin_cm", 1.8)))
    if formatting.get("show_page_numbers", True) and section.footer.paragraphs:
        add_page_number(section.footer.paragraphs[0])
    add_doc_header(document, title, version, formatting, include_answer_note=True)

    # Compact answer sheet: numbering and correct answer are next to each other.
    table = document.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "Question No."
    hdr[1].text = "Correct Answer"

    for i, q in enumerate(questions, start=1):
        cells = table.add_row().cells
        cells[0].text = str(i)
        cells[1].text = q["Correct"]

    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                for run in p.runs:
                    set_run_font(run, formatting, size_key="answer_key_font_size")

    # Optional audit details, kept below the compact answer sheet.
    if formatting.get("show_source_in_answer_key", True):
        document.add_paragraph("")
        p = document.add_paragraph()
        r = p.add_run("Audit details")
        set_run_font(r, formatting, size_key="font_size", bold=True)

        audit_table = document.add_table(rows=1, cols=4)
        audit_table.style = "Table Grid"
        hdr = audit_table.rows[0].cells
        hdr[0].text = "Question No."
        hdr[1].text = "Internal ID"
        hdr[2].text = "Options"
        hdr[3].text = "Source File"

        for i, q in enumerate(questions, start=1):
            cells = audit_table.add_row().cells
            cells[0].text = str(i)
            cells[1].text = q["Internal_ID"]
            cells[2].text = f"A–{LETTERS[q['Option_Count']-1]}"
            cells[3].text = q["Source_File"]

        for row in audit_table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        set_run_font(run, formatting, size_key="answer_key_font_size")

    output = BytesIO()
    document.save(output)
    output.seek(0)
    return output


def create_mapping_report(versions: Dict[str, List[Dict]]) -> BytesIO:
    rows = []
    for version, questions in versions.items():
        for new_no, q in enumerate(questions, start=1):
            rows.append({
                "Version": version,
                "New Question No.": new_no,
                "Internal_ID": q["Internal_ID"],
                "Original No": q["No"],
                "Options": f"A–{LETTERS[q['Option_Count']-1]}",
                "Correct Answer": q["Correct"],
                "Original Correct Answer": q["Original_Correct"],
                "Answer Shuffle Skipped": "Yes" if q.get("Answer_Shuffle_Skipped") else "No",
                "Source File": q["Source_File"]
            })
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_excel_dataframe(pd.DataFrame(rows)).to_excel(writer, sheet_name="Version Mapping", index=False)
    output.seek(0)
    return output


def create_validation_report(report_df: pd.DataFrame, combined_df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_excel_dataframe(report_df).to_excel(writer, sheet_name="Validation Issues", index=False)
        summary = pd.DataFrame([
            ["Files imported", combined_df["Source_File"].nunique() if not combined_df.empty else 0],
            ["Questions imported", len(combined_df)],
            ["Errors", int((report_df["Severity"] == "ERROR").sum()) if not report_df.empty else 0],
            ["Warnings", int((report_df["Severity"] == "WARNING").sum()) if not report_df.empty else 0],
        ], columns=["Metric", "Value"])
        safe_excel_dataframe(summary).to_excel(writer, sheet_name="Summary", index=False)
        if not combined_df.empty:
            safe_excel_dataframe(combined_df).to_excel(writer, sheet_name="Imported Questions", index=False)
    output.seek(0)
    return output


def create_zip_output_from_versions(
    versions: Dict[str, List[Dict]],
    combined_df,
    report_df,
    exam_title,
    include_master_with_answers,
    image_map: Dict[str, bytes] = None,
    formatting: Dict = None,
    export_options: Dict = None,
    kahoot_time_limit: int = 20,
    blueprint_df: pd.DataFrame = None,
) -> BytesIO:
    """Build an exam package from explicit reviewed inputs and settings."""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Validation_Report.xlsx", create_validation_report(report_df, combined_df).getvalue())
        if blueprint_df is not None:
            zf.writestr("Blueprint_Report.xlsx", make_blueprint_report(blueprint_df).getvalue())

        export_options = export_options or {
            "word": True,
            "answer_key": True,
            "aiken": False,
            "kahoot": False,
            "export_summary": True
        }

        for label, questions in versions.items():
            answer_display = (formatting or get_default_formatting()).get("answer_display", "Separate answer key only")
            exam_show_answers = answer_display == "Answers below each question"

            if export_options.get("word", True):
                zf.writestr(
                    f"Exam_{label}.docx",
                    create_exam_docx(questions, exam_title, label, exam_show_answers, image_map=image_map, formatting=formatting).getvalue()
                )

            if export_options.get("answer_key", True) and answer_display in ["Separate answer key only", "Both answer key and answers below each question"]:
                zf.writestr(
                    f"Answer_Key_{label}.docx",
                    create_answer_key_docx(questions, exam_title, label, formatting=formatting).getvalue()
                )

            if include_master_with_answers or answer_display == "Both answer key and answers below each question":
                zf.writestr(
                    f"Exam_{label}_WITH_ANSWERS.docx",
                    create_exam_docx(questions, exam_title, label, True, image_map=image_map, formatting=formatting).getvalue()
                )

            if export_options.get("aiken", False):
                zf.writestr(f"Aiken_Blackboard_Version_{label}.txt", create_aiken_txt(questions, exam_title, label).getvalue())

            if export_options.get("kahoot", False):
                kahoot_xlsx, _ = create_kahoot_xlsx(questions, default_time_limit=kahoot_time_limit)
                zf.writestr(f"Kahoot_Import_Version_{label}.xlsx", kahoot_xlsx.getvalue())

        zf.writestr("Version_Mapping_Report.xlsx", create_mapping_report(versions).getvalue())
        if export_options.get("export_summary", True):
            zf.writestr("Export_Summary.xlsx", create_export_summary_excel(versions, export_options).getvalue())
    zip_buffer.seek(0)
    return zip_buffer




# -----------------------------
# Smart validation helpers
# -----------------------------

def normalize_for_similarity(text: str) -> str:
    text = normalise_cell(text).lower()
    text = re.sub(r"[^a-z0-9α-ωάέήίόύώϊϋΐΰ ]+", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def similarity_ratio(a: str, b: str) -> float:
    a_norm = normalize_for_similarity(a)
    b_norm = normalize_for_similarity(b)
    if not a_norm or not b_norm:
        return 0.0
    return SequenceMatcher(None, a_norm, b_norm).ratio()


def add_smart_validation_issues(
    combined_df: pd.DataFrame,
    report_df: pd.DataFrame,
    similarity_threshold: float = 0.90,
    long_question_limit: int = 350,
    long_answer_limit: int = 180,
    enable_fuzzy_duplicates: bool = True
) -> pd.DataFrame:
    """
    Adds non-blocking smart validation warnings.
    This does not replace structural validation; it adds quality-control warnings.
    """
    if combined_df is None or combined_df.empty:
        return report_df

    issues = report_df.to_dict("records") if report_df is not None and not report_df.empty else []

    # Exact duplicate question text
    norm_questions = combined_df["Question"].apply(normalize_for_similarity)
    duplicate_norms = norm_questions[norm_questions.duplicated(keep=False) & norm_questions.ne("")]
    for norm_q in sorted(set(duplicate_norms)):
        rows = combined_df.index[norm_questions == norm_q].tolist()
        sources = combined_df.loc[rows, "Source_File"].tolist()
        ids = combined_df.loc[rows, "Internal_ID"].tolist()
        issues.append({
            "Source file": "; ".join(sources),
            "Row": "",
            "Internal_ID": "; ".join(ids),
            "Severity": "WARNING",
            "Issue": "Exact duplicate question text detected"
        })

    # Similar-question detection is bounded so a very large upload cannot make
    # the lecturer interface unresponsive.
    if enable_fuzzy_duplicates:
        fuzzy_limit = 1500
        if len(combined_df) > fuzzy_limit:
            issues.append({
                "Source file": "All uploaded files",
                "Row": "",
                "Internal_ID": "",
                "Severity": "WARNING",
                "Issue": f"Fuzzy duplicate scan skipped above {fuzzy_limit} questions; exact duplicates were still checked"
            })
        else:
            records = combined_df[["Question", "Internal_ID", "Source_File"]].copy()
            records["Normalized"] = norm_questions
            items = records.to_dict("records")
            fuzzy_findings = 0
            fuzzy_findings_limit = 500
            fuzzy_truncated = False
            for i in range(len(items)):
                left = items[i]["Normalized"]
                if not left:
                    continue
                for j in range(i + 1, len(items)):
                    right = items[j]["Normalized"]
                    if not right or left == right:
                        continue
                    length_upper_bound = 2 * min(len(left), len(right)) / (len(left) + len(right))
                    if length_upper_bound < similarity_threshold:
                        continue
                    ratio = SequenceMatcher(None, left, right).ratio()
                    if ratio >= similarity_threshold:
                        issues.append({
                            "Source file": f"{items[i]['Source_File']}; {items[j]['Source_File']}",
                            "Row": "",
                            "Internal_ID": f"{items[i]['Internal_ID']}; {items[j]['Internal_ID']}",
                            "Severity": "WARNING",
                            "Issue": f"Very similar questions detected ({ratio:.0%} similarity)"
                        })
                        fuzzy_findings += 1
                        if fuzzy_findings >= fuzzy_findings_limit:
                            fuzzy_truncated = True
                            break
                if fuzzy_truncated:
                    break
            if fuzzy_truncated:
                issues.append({
                    "Source file": "All uploaded files",
                    "Row": "",
                    "Internal_ID": "",
                    "Severity": "WARNING",
                    "Issue": f"Fuzzy duplicate findings limited to the first {fuzzy_findings_limit} matches"
                })

    # Long question and long answer warnings
    for _, row in combined_df.iterrows():
        question = normalise_cell(row["Question"])
        if len(question) > long_question_limit:
            issues.append({
                "Source file": row["Source_File"],
                "Row": "",
                "Internal_ID": row["Internal_ID"],
                "Severity": "WARNING",
                "Issue": f"Question is very long ({len(question)} characters; limit {long_question_limit})"
            })

        option_count = int(row["Option_Count"])
        for col in answer_columns_for_count(option_count):
            answer = normalise_cell(row[col])
            if len(answer) > long_answer_limit:
                issues.append({
                    "Source file": row["Source_File"],
                    "Row": "",
                    "Internal_ID": row["Internal_ID"],
                    "Severity": "WARNING",
                    "Issue": f"{col} is very long ({len(answer)} characters; limit {long_answer_limit})"
                })

        # Duplicate answer options after stronger normalization
        answer_values = [normalize_for_similarity(row[col]) for col in answer_columns_for_count(option_count)]
        answer_values = [a for a in answer_values if a]
        if len(answer_values) != len(set(answer_values)):
            issues.append({
                "Source file": row["Source_File"],
                "Row": "",
                "Internal_ID": row["Internal_ID"],
                "Severity": "WARNING",
                "Issue": "Duplicate or near-identical answer option text within the same question"
            })

        raw_answer_values = [normalise_cell(row[col]) for col in answer_columns_for_count(option_count)]
        if has_position_dependent_options(raw_answer_values):
            issues.append({
                "Source file": row["Source_File"],
                "Row": "",
                "Internal_ID": row["Internal_ID"],
                "Severity": "WARNING",
                "Issue": "Position-dependent option detected; answer order will be preserved for this question"
            })

    smart_report = pd.DataFrame(issues)
    if smart_report.empty:
        smart_report = pd.DataFrame(columns=["Source file", "Row", "Internal_ID", "Severity", "Issue"])
    return smart_report

# -----------------------------
# Blueprint helpers
# -----------------------------

def select_questions_by_blueprint(combined_df: pd.DataFrame, blueprint_counts: Dict[str, int], seed: int = 12345) -> pd.DataFrame:
    """
    Selects requested number of questions from each source file.
    If requested number equals or exceeds available questions, all questions are used.
    Selection is deterministic for reproducibility.
    """
    # Deliberately deterministic for reproducible blueprint sampling.
    rng = random.Random(seed)  # nosec B311
    selected_indices = []

    for source_file, requested in blueprint_counts.items():
        source_df = combined_df[combined_df["Source_File"] == source_file]
        available_indices = list(source_df.index)

        if requested >= len(available_indices):
            selected_indices.extend(available_indices)
        elif requested <= 0:
            continue
        else:
            selected_indices.extend(rng.sample(available_indices, requested))

    selected_df = combined_df.loc[selected_indices].copy().reset_index(drop=True)
    return selected_df


def make_blueprint_report(blueprint_df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_excel_dataframe(blueprint_df).to_excel(writer, sheet_name="Blueprint", index=False)
    output.seek(0)
    return output


def make_blueprint_dataframe(combined_df: pd.DataFrame, blueprint_counts: Dict[str, int]) -> pd.DataFrame:
    rows = []
    for source_file in sorted(combined_df["Source_File"].unique()):
        available = int((combined_df["Source_File"] == source_file).sum())
        selected = int(blueprint_counts.get(source_file, available))
        rows.append({
            "Source File": source_file,
            "Available Questions": available,
            "Selected Questions": min(selected, available),
            "Requested Questions": selected,
            "Status": "OK" if selected <= available else "Requested more than available; all available used"
        })
    return pd.DataFrame(rows)




# -----------------------------
# Multi-format export helpers
# -----------------------------

def truncate_text(text: str, max_len: int) -> str:
    text = normalise_cell(text)
    if len(text) <= max_len:
        return text
    return text[: max_len - 1] + "…"


def create_aiken_txt(questions: List[Dict], title: str, version: str) -> BytesIO:
    """
    Blackboard/Aiken-style plain text.
    Supports variable answer option counts.
    """
    lines = []
    lines.append(f"{title} - Version {version}")
    lines.append("")

    for q in questions:
        lines.append(q["Question"])
        for opt in q["Options"]:
            lines.append(f"{opt['letter']}. {opt['text']}")
        lines.append(f"ANSWER: {q['Correct']}")
        lines.append("")

    output = BytesIO()
    output.write("\\n".join(lines).encode("utf-8-sig"))
    output.seek(0)
    return output


def create_kahoot_xlsx(questions: List[Dict], default_time_limit: int = 20) -> Tuple[BytesIO, pd.DataFrame]:
    """
    Creates Kahoot import Excel using the user's preferred format:
    Question, Answer 1-4, Time limit (sec), Correct answer(s).

    Kahoot supports up to 4 answers in this template.
    Questions with >4 answer options are skipped and reported.
    Questions with 2-3 answers are padded with blanks.
    Text is truncated to Kahoot limits:
    - Question max 120 chars
    - Answers max 75 chars
    """
    rows = []
    skipped = []

    for i, q in enumerate(questions, start=1):
        option_count = int(q["Option_Count"])

        if option_count > 4:
            skipped.append({
                "Question No.": i,
                "Internal ID": q["Internal_ID"],
                "Reason": f"Kahoot template supports up to 4 answers; this question has {option_count}.",
                "Source File": q["Source_File"]
            })
            continue

        options = q["Options"]
        answers = [opt["text"] for opt in options]
        while len(answers) < 4:
            answers.append("")

        correct_letter = q["Correct"]
        correct_number = LETTERS.index(correct_letter) + 1

        rows.append({
            "Question - max 120 characters": truncate_text(q["Question"], 120),
            "Answer 1 - max 75 characters": truncate_text(answers[0], 75),
            "Answer 2 - max 75 characters": truncate_text(answers[1], 75),
            "Answer 3 - max 75 characters": truncate_text(answers[2], 75),
            "Answer 4 - max 75 characters": truncate_text(answers[3], 75),
            "Time limit (sec)": default_time_limit,
            "Correct answer(s)": correct_number
        })

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_excel_dataframe(pd.DataFrame(rows)).to_excel(writer, index=False, sheet_name="Kahoot")
        skipped_df = pd.DataFrame(skipped)
        if skipped_df.empty:
            skipped_df = pd.DataFrame(columns=["Question No.", "Internal ID", "Reason", "Source File"])
        safe_excel_dataframe(skipped_df).to_excel(writer, index=False, sheet_name="Skipped")
    output.seek(0)

    return output, pd.DataFrame(skipped)


def make_kahoot_limit_report(versions: Dict[str, List[Dict]]) -> pd.DataFrame:
    rows = []
    for version, questions in versions.items():
        for i, q in enumerate(questions, start=1):
            if len(normalise_cell(q["Question"])) > 120:
                rows.append({
                    "Version": version,
                    "Question No.": i,
                    "Internal ID": q["Internal_ID"],
                    "Issue": f"Question exceeds Kahoot limit: {len(normalise_cell(q['Question']))}/120 characters"
                })
            for opt in q["Options"]:
                if len(normalise_cell(opt["text"])) > 75:
                    rows.append({
                        "Version": version,
                        "Question No.": i,
                        "Internal ID": q["Internal_ID"],
                        "Issue": f"Answer {opt['letter']} exceeds Kahoot limit: {len(normalise_cell(opt['text']))}/75 characters"
                    })
            if int(q["Option_Count"]) > 4:
                rows.append({
                    "Version": version,
                    "Question No.": i,
                    "Internal ID": q["Internal_ID"],
                    "Issue": f"Kahoot export skipped: {q['Option_Count']} options"
                })
    if not rows:
        return pd.DataFrame(columns=["Version", "Question No.", "Internal ID", "Issue"])
    return pd.DataFrame(rows)


def create_export_summary_excel(versions: Dict[str, List[Dict]], export_options: Dict) -> BytesIO:
    rows = []
    for label, questions in versions.items():
        kahoot_skipped = sum(1 for q in questions if int(q["Option_Count"]) > 4)
        rows.append({
            "Version": label,
            "Questions": len(questions),
            "Word exam": "Yes" if export_options.get("word", True) else "No",
            "Answer key": "Yes" if export_options.get("answer_key", True) else "No",
            "Aiken TXT": "Yes" if export_options.get("aiken", False) else "No",
            "Kahoot XLSX": "Yes" if export_options.get("kahoot", False) else "No",
            "Kahoot skipped questions": kahoot_skipped
        })

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_excel_dataframe(pd.DataFrame(rows)).to_excel(writer, index=False, sheet_name="Export Summary")
        safe_excel_dataframe(make_kahoot_limit_report(versions)).to_excel(
            writer, index=False, sheet_name="Kahoot Warnings"
        )
    output.seek(0)
    return output


# -----------------------------
# Preview helpers
# -----------------------------

def make_answer_key_dataframe(questions: List[Dict]) -> pd.DataFrame:
    return pd.DataFrame([
        {
            "Question No.": i,
            "Correct": q["Correct"],
            "Internal ID": q["Internal_ID"],
            "Source": q["Source_File"],
            "Images": " | ".join(q.get("Images") or split_image_files(q.get("Image", ""))),
            "Options": f"A–{LETTERS[q['Option_Count']-1]}"
        }
        for i, q in enumerate(questions, start=1)
    ])


def make_mapping_dataframe(versions: Dict[str, List[Dict]]) -> pd.DataFrame:
    rows = []
    for version, questions in versions.items():
        for new_no, q in enumerate(questions, start=1):
            rows.append({
                "Version": version,
                "New Question No.": new_no,
                "Internal ID": q["Internal_ID"],
                "Original No": q["No"],
                "Correct": q["Correct"],
                "Original Correct": q["Original_Correct"],
                "Answer Shuffle Skipped": "Yes" if q.get("Answer_Shuffle_Skipped") else "No",
                "Source": q["Source_File"],
                "Images": " | ".join(q.get("Images") or split_image_files(q.get("Image", ""))),
                "Options": f"A–{LETTERS[q['Option_Count']-1]}"
            })
    return pd.DataFrame(rows)


def show_student_preview(questions: List[Dict], max_preview: int, search: str, image_map: Dict[str, bytes] = None):
    shown = 0
    search_lower = search.strip().lower()

    for i, q in enumerate(questions, start=1):
        image_names = q.get("Images") or split_image_files(q.get("Image", ""))
        searchable = " ".join(
            [q["Question"], *image_names] + [opt["text"] for opt in q["Options"]] + [q["Internal_ID"], q["Source_File"]]
        ).lower()

        if search_lower and search_lower not in searchable:
            continue

        with st.container(border=True):
            st.markdown(f"**{i}. {q['Question']}**")
            for image_name in image_names:
                img_key = image_name.strip().lower()
                if image_map and img_key in image_map:
                    st.image(image_map[img_key], caption=image_name, width=450)
                else:
                    st.warning(f"Image missing: {image_name}")

            st.markdown("  \n".join(f"{opt['letter']}. {opt['text']}" for opt in q["Options"]))

            if q.get("Answer_Shuffle_Skipped"):
                st.caption("Answer order preserved because this question contains a position-dependent option.")


        shown += 1

        if shown >= max_preview:
            break

    if shown == 0:
        st.warning("No questions matched the search.")
    elif search_lower:
        st.info(f"Showing {shown} matching question(s).")
    else:
        st.info(f"Showing first {shown} question(s).")

def show_answer_preview(questions: List[Dict], search: str):
    df = make_answer_key_dataframe(questions)
    if search.strip():
        mask = df.astype(str).apply(lambda col: col.str.lower().str.contains(search.strip().lower(), na=False)).any(axis=1)
        df = df[mask]
    st.dataframe(df, width="stretch")


def show_statistics(versions: Dict[str, List[Dict]], combined_df: pd.DataFrame, shuffle_questions: bool, shuffle_answers: bool):
    rows = []
    for label, qs in versions.items():
        counts = {}
        for q in qs:
            fmt = f"A–{LETTERS[q['Option_Count']-1]}"
            counts[fmt] = counts.get(fmt, 0) + 1
        rows.append({
            "Version": label,
            "Questions": len(qs),
            "Formats": ", ".join([f"{k}: {v}" for k, v in counts.items()]),
            "Shuffle question order": "Yes" if shuffle_questions else "No",
            "Shuffle answer options": "Yes" if shuffle_answers else "No"
        })
    st.dataframe(pd.DataFrame(rows), width="stretch")

    st.subheader("Questions by source file")
    source_counts = combined_df["Source_File"].value_counts().reset_index()
    source_counts.columns = ["Source File", "Questions"]
    st.dataframe(source_counts, width="stretch")

    st.subheader("Option formats in imported bank")
    option_counts = combined_df["Option_Count"].value_counts().sort_index()
    fmt_df = pd.DataFrame({
        "Options": [f"A–{LETTERS[int(i)-1]}" for i in option_counts.index],
        "Questions": option_counts.values
    })
    st.dataframe(fmt_df, width="stretch")



# -----------------------------
# Lecturer mode UI helpers
# -----------------------------

def make_safe_filename(text: str) -> str:
    text = normalise_cell(text)
    text = re.sub(r"[^A-Za-z0-9Α-Ωα-ωάέήίόύώϊϋΐΰ_-]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "MCQ_Exam"


def show_progress(current_step: int, has_questions=False, has_package=False):
    """One accessible navigation bar; link only to sections present on this run."""
    steps = [
        ("Upload", "start-building", True),
        ("Review questions", "review-questions", has_questions),
        ("Configure &amp; create", "configure-exam", has_questions),
        ("Download", "download-package", has_package),
    ]
    markers = []
    for i, (label, anchor, available) in enumerate(steps, start=1):
        complete = i < current_step
        state = "is-active" if i == current_step else "is-complete" if complete else ""
        number = "✓" if complete else str(i)
        current = ' aria-current="step"' if i == current_step else ""
        content = f'<span class="workflow-number">{number}</span><span>{label}</span>'
        if available:
            markers.append(f'<a class="workflow-step {state}" href="#{anchor}"{current}>{content}</a>')
        else:
            markers.append(f'<span class="workflow-step {state}" aria-disabled="true"{current}>{content}</span>')
    st.markdown(
        f'<nav class="workflow-shell" aria-label="Exam workflow">{"".join(markers)}</nav>',
        unsafe_allow_html=True,
    )


def clear_current_exam():
    """Remove all exam-specific data and reset uploader widgets."""
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.session_state.clear_confirmation = True


def render_build_settings():
    """Render configuration only after a question bank has been imported."""
    st.markdown("---")
    st.header("3 — Configure & create", anchor="configure-exam")
    st.caption("Choose your versions and output files. Formatting and extra checks are optional.")

    col_basic_1, col_basic_2, col_basic_3 = st.columns(3)
    with col_basic_1:
        exam_title = st.text_input("Exam title", value="MCQ Examination")
    with col_basic_2:
        num_versions = st.number_input("Number of exam versions", min_value=1, max_value=50, value=1, step=1)
    with col_basic_3:
        naming = st.radio("Version labels", ["Letters", "Numbers"], horizontal=True)

    col_shuffle_a, col_shuffle_b = st.columns(2)
    with col_shuffle_a:
        shuffle_questions = st.checkbox("Shuffle question order", value=True)
    with col_shuffle_b:
        shuffle_answers = st.checkbox("Shuffle answer options", value=True)
    defaults = get_default_formatting()
    values = {
        "include_master_with_answers": True, "enable_fuzzy_duplicates": True,
        "similarity_threshold": 0.90, "long_question_limit": 350, "long_answer_limit": 180,
        "font_name": defaults["font_name"], "font_size": defaults["font_size"],
        "title_font_size": defaults["title_font_size"], "answer_key_font_size": defaults["answer_key_font_size"],
        "top_margin_cm": defaults["top_margin_cm"], "bottom_margin_cm": defaults["bottom_margin_cm"],
        "left_margin_cm": defaults["left_margin_cm"], "right_margin_cm": defaults["right_margin_cm"],
        "option_indent_cm": defaults["option_indent_cm"], "space_before_question_pt": defaults["space_before_question_pt"],
        "space_after_question_pt": defaults["space_after_question_pt"], "space_after_option_pt": defaults["space_after_option_pt"],
        "blank_line_between_questions": defaults["blank_line_between_questions"], "bold_questions": defaults["bold_questions"],
        "show_page_numbers": defaults["show_page_numbers"], "show_version_in_header": defaults["show_version_in_header"],
        "show_source_in_answer_key": defaults["show_source_in_answer_key"], "image_width_cm": defaults["image_width_cm"],
        "answer_display": defaults["answer_display"], "export_word": True, "export_answer_key": True,
        "export_aiken": False, "export_kahoot": False, "kahoot_time_limit": 20,
        "export_summary": True, "use_blueprint_advanced": False,
    }
    with st.expander("Files to include", expanded=True):
        st.caption("Word exams and answer keys are selected by default. Add learning-platform formats if needed.")
        document_column, platform_column = st.columns(2)
        with document_column:
            values["export_word"] = st.checkbox("Export Word exam files", value=True)
            values["export_answer_key"] = st.checkbox("Export Word answer keys", value=True)
            values["include_master_with_answers"] = st.checkbox("Include lecturer copies with answers", value=True,
                help="Adds a separate copy of each exam with the answers below the questions.")
            values["export_summary"] = st.checkbox("Include export summary report", value=True)
        with platform_column:
            values["export_aiken"] = st.checkbox("Export Blackboard/Aiken TXT", value=False)
            values["export_kahoot"] = st.checkbox("Export Kahoot import Excel", value=False)
            values["kahoot_time_limit"] = st.selectbox("Kahoot time limit (seconds)", [5, 10, 20, 30, 60, 90, 120, 240], index=2,
                disabled=not values["export_kahoot"])
    with st.expander("Document formatting", expanded=False):
        type_tab, layout_tab, detail_tab = st.tabs(["Typography", "Page layout", "Answers & details"])
        with type_tab:
            values["font_name"] = st.selectbox("Font", ["Arial", "Calibri", "Times New Roman"])
            values["font_size"] = st.number_input("Main font size", 8, 16, defaults["font_size"], 1)
            values["title_font_size"] = st.number_input("Title font size", 10, 20, defaults["title_font_size"], 1)
            values["answer_key_font_size"] = st.number_input("Answer key table font size", 8, 14, defaults["answer_key_font_size"], 1)
        with layout_tab:
            for key, label in [("top_margin_cm", "Top"), ("bottom_margin_cm", "Bottom"), ("left_margin_cm", "Left"), ("right_margin_cm", "Right")]:
                values[key] = st.number_input(f"{label} margin (cm)", 0.5, 4.0, defaults[key], 0.1)
            values["option_indent_cm"] = st.number_input("Answer option indent (cm)", 0.0, 2.0, defaults["option_indent_cm"], 0.1)
            values["space_before_question_pt"] = st.number_input("Space before question (pt)", 0, 24, defaults["space_before_question_pt"], 1)
            values["space_after_question_pt"] = st.number_input("Space after question (pt)", 0, 24, defaults["space_after_question_pt"], 1)
            values["space_after_option_pt"] = st.number_input("Space after each option (pt)", 0, 12, defaults["space_after_option_pt"], 1)
        with detail_tab:
            for key, label in [("blank_line_between_questions", "Blank line between questions"), ("bold_questions", "Bold question text"), ("show_page_numbers", "Show page numbers"), ("show_version_in_header", "Show version in header"), ("show_source_in_answer_key", "Show source file in answer key")]:
                values[key] = st.checkbox(label, value=defaults[key])
            values["image_width_cm"] = st.number_input("Default image width in Word (cm)", 3.0, 18.0, defaults["image_width_cm"], 0.5)
            values["answer_display"] = st.selectbox("Answer display in exported files", ["Separate answer key only", "Answers below each question", "Both answer key and answers below each question"])
    with st.expander("Question selection & validation", expanded=False):
        st.markdown("### Question selection")
        values["use_blueprint_advanced"] = st.checkbox("Choose question count by source file", value=False)
        st.markdown("### Smart validation")
        values["enable_fuzzy_duplicates"] = st.checkbox("Detect very similar questions", value=True)
        values["similarity_threshold"] = st.slider("Similarity threshold", 0.80, 0.98, 0.90, 0.01)
        values["long_question_limit"] = st.number_input("Question-length warning (characters)", 100, 1000, 350, 25)
        values["long_answer_limit"] = st.number_input("Answer-length warning (characters)", 50, 500, 180, 10)

    values.update({"exam_title": exam_title, "num_versions": int(num_versions), "naming": naming,
                   "shuffle_questions": shuffle_questions, "shuffle_answers": shuffle_answers})
    return values


def make_exam_summary(versions: Dict[str, List[Dict]], combined_df: pd.DataFrame, settings: Dict):
    first_version = next(iter(versions.values())) if versions else []
    image_count = sum(1 for q in first_version if q.get("Images") or q.get("Image"))
    c1, c2, c3 = st.columns(3)
    c1.metric("Questions", len(first_version))
    c2.metric("Versions", len(versions))
    c3.metric("With images", image_count)


# -----------------------------
# UI
# -----------------------------

inject_app_styles()

if "preview_versions" not in st.session_state:
    st.session_state.preview_versions = None
if "preview_settings" not in st.session_state:
    st.session_state.preview_settings = None
if "combined_df" not in st.session_state:
    st.session_state.combined_df = None
if "report_df" not in st.session_state:
    st.session_state.report_df = None
if "blueprint_df" not in st.session_state:
    st.session_state.blueprint_df = None
if "image_map" not in st.session_state:
    st.session_state.image_map = {}
if "formatting" not in st.session_state:
    st.session_state.formatting = get_default_formatting()
if "shuffle_seed" not in st.session_state:
    st.session_state.shuffle_seed = 10000
if "selected_df_for_preview" not in st.session_state:
    st.session_state.selected_df_for_preview = None
if "preview_input_fingerprint" not in st.session_state:
    st.session_state.preview_input_fingerprint = None


st.markdown('<div class="brand-line">Exam preparation for educators</div>', unsafe_allow_html=True)
st.title("MCQ Exam Builder")
st.markdown(
    '<p class="hero-copy">Create checked exam versions from the question files you already use. '
    'Review flagged items, shuffle questions and answer choices, and download editable papers with matching keys.</p>'
    '<div class="trust-line"><span>No generative-AI processing</span> '
    '<span>No exam database</span></div>',
    unsafe_allow_html=True,
)
progress_slot = st.empty()

if st.session_state.pop("clear_confirmation", False):
    st.success("The current exam has been cleared from this active session.")

st.header("1 — Upload your question bank", anchor="start-building")
st.caption("Each uploaded file should contain the questions, answer options, and correct answers.")

upload_column, starter_column = st.columns([2, 1], gap="large")
with upload_column:
    uploaded_files = st.file_uploader(
        "Question files",
        type=SUPPORTED_EXTENSIONS,
        accept_multiple_files=True,
        help="Upload one or more question files. For document or text files, an explicit answer key can appear at the end of the same file. Answer keys uploaded as separate files are not matched to questions.",
        key="question_file_uploader",
    )
    st.caption("DOCX preserves embedded images best. Converted documents and text-based imports may need extra review.")
    with st.expander("Add referenced question images"):
        st.caption("For spreadsheet, text, Markdown, or JSON sources, upload PNG or JPG files that match the image filenames in your question bank.")
        image_files = st.file_uploader(
            "Referenced images", type=["png", "jpg", "jpeg"], accept_multiple_files=True,
            help="Use the exact filenames referenced in your questions.", key="referenced_image_uploader",
        )
with starter_column:
    with st.container(border=True):
        st.subheader("Your exam package")
        st.markdown("- Editable Word exams\n- Matching answer keys\n- Shuffled versions in one ZIP")
        st.caption("Choose extra formats after reviewing your questions.")
    with st.expander("Need a question template?"):
        st.caption("Start with the optional Excel template. Supports 2–6 answer options per question.")
        st.download_button(
            "Download Excel template", data=make_template_xlsx(),
            file_name="MCQ_Universal_Template_v15.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with st.expander("How it works & supported formats"):
    st.markdown(
        "1. **Upload** question files that include their correct answers.\n"
        "2. **Review** the imported text, answers, images, and flagged items.\n"
        "3. **Configure & create** your exam versions and output files. "
        "Choose **Create exam package**, or **Preview exam first** to inspect and optionally reshuffle the result.\n"
        "4. **Download** the complete exam package as a ZIP."
    )
    st.markdown(
        "**Documents:** DOCX, DOC, DOCM, DOTX, DOTM, ODT, OTT, FODT, RTF. "
        "Legacy Word and OpenDocument files are converted and require review (20 MB limit).\n\n"
        "**Spreadsheets:** XLSX, XLS, XLSM, XLTX, XLTM, ODS, CSV, TSV. "
        "Include Question, Option 1–6, and Correct columns; leave unused option columns blank.\n\n"
        "**Text:** TXT, MD, MARKDOWN, HTML, HTM. Use numbered questions, labelled options, "
        "and explicit answers. Visual bold/highlight answer cues are supported in Word, not text imports.\n\n"
        "**PDF:** selectable-text questions with explicit answers (up to 50 MB and 200 pages). "
        "Scanned pages and PDFs with images or drawn graphics need the original document or an OCR/text export.\n\n"
        "**Question banks:** JSON, JSONL, and Moodle XML single-answer MCQs or true/false questions. "
        "Multiple-answer, partial-credit, and embedded-media XML questions are not supported."
    )
    st.caption("JSON: use a questions array with question, options (2–6 strings), and correct (A–F or a 1-based number). "
               "JSONL uses one question object per line. Text-based imports require UTF-8 or UTF-16 encoding. "
               "Separate answer-key files are not automatically matched to question files.")


with st.expander("Privacy & data handling"):
    st.markdown(
        "**Processing.** Uploaded files are processed on the application server, not solely in your browser. "
        "The application does not send exam content to generative-AI services and has no exam database, "
        "shared content cache or analytics integration.\n\n"
        "**Retention.** Imported questions, images and generated packages are held in session memory. "
        "Document conversion and Word image export also use temporary files, removed after processing. "
        "An interrupted server can delay cleanup; session memory and temporary storage remain subject to "
        "the hosting platform's lifecycle. Closing a browser tab is not a guarantee of immediate deletion.\n\n"
        "**Hosted use.** The public app runs on the hosting provider's infrastructure. Ordinary operational "
        "logs may be retained; the application does not intentionally log exam content. Check your "
        "institution's policy before uploading confidential or restricted examination material.\n\n"
        "**Do not upload student personal data, grades or candidate identifiers.** "
        "Download the files you need before ending the session."
    )

if uploaded_files or st.session_state.get("combined_df") is not None:
    st.button(
        "Clear current exam and start again",
        on_click=clear_current_exam,
        type="secondary",
        help="Clears uploaded files, imported questions, images, previews, settings, and generated packages from this session.",
    )

manual_image_map = {}
if image_files:
    for img in image_files:
        img.seek(0)
        image_blob = img.read()
        image_error = validate_image_bytes(img.name, image_blob)
        if image_error:
            st.error(image_error)
            continue
        manual_image_map[img.name.strip().lower()] = image_blob
    if manual_image_map:
        st.success(f"{len(manual_image_map)} valid additional image file(s) uploaded.")




# UPLOAD_BLOCK_MOVED_TO_STEP_2


if uploaded_files:
    # Reuse successful imports within this session; document conversion should not
    # run again on every editor or settings interaction. No shared/persistent cache.
    upload_signature = tuple((item.name, hashlib.sha256(item.getvalue()).hexdigest()) for item in uploaded_files)
    if st.session_state.get("import_signature") != upload_signature:
        with st.spinner("Reading your question files…"):
            result = import_all_sources(uploaded_files)
        st.session_state.import_result = result
        st.session_state.import_signature = upload_signature if not has_blocking_errors(result[1]) else None
    cached_df, cached_report, cached_images = st.session_state.import_result
    imported_df, import_report_df = cached_df.copy(deep=True), cached_report.copy(deep=True)
    extracted_image_map = dict(cached_images)
    merged_image_map = {}
    extracted_image_errors = []
    for image_name, image_blob in extracted_image_map.items():
        image_error = validate_image_bytes(image_name, image_blob)
        if image_error:
            extracted_image_errors.append(image_error)
        else:
            merged_image_map[image_name] = image_blob
    merged_image_map.update(manual_image_map)
    st.session_state.image_map = merged_image_map

    st.divider()
    st.header("2 — Review questions", anchor="review-questions")
    st.write(
        "The importer has converted every source into a consistent structure. Check the question text, answer "
        "options and correct-answer letter. Converted images, extracted text, and uncertain answers are flagged for confirmation."
    )

    validation_slot = st.container()
    review_acknowledged = True
    if not imported_df.empty:
        source_signature = abs(hash(upload_signature))
        editor_df = dataframe_for_editor(imported_df)
        flagged_review_count = int(imported_df["Needs_Review"].fillna(False).astype(bool).sum())
        if flagged_review_count:
            st.markdown('<div id="required-review"></div>', unsafe_allow_html=True)
            st.warning(
                f"**Action required:** {flagged_review_count} imported question(s) need your confirmation. "
                "Before creating or previewing the exam package, review the flagged rows and tick the confirmation below."
            )
        reviewed_editor_df = st.data_editor(
            editor_df,
            key=f"question_review_{source_signature}",
            width="stretch",
            height=min(520, 80 + 36 * len(editor_df)),
            num_rows="dynamic",
            hide_index=True,
            column_config={
                "Question": st.column_config.TextColumn("Question", width="large", required=True),
                "Correct": st.column_config.SelectboxColumn(
                    "Correct",
                    help="Select the letter of the correct option.",
                    options=["", "A", "B", "C", "D", "E", "F"],
                    required=False,
                    width="small",
                ),
                "Image_Files": st.column_config.TextColumn(
                    "Images",
                    help="One filename, or several filenames separated with |.",
                    width="medium",
                ),
                "Needs_Review": st.column_config.CheckboxColumn("Needs review", width="small"),
                "Import_Warnings": st.column_config.TextColumn("Importer notes", width="large"),
            },
            disabled=[
                "Needs_Review",
                "Import_Method",
                "Answer_Evidence",
                "Confidence",
                "Import_Warnings",
                "Source_File",
                "Internal_ID",
            ],
        )
        if flagged_review_count:
            with st.container(border=True):
                st.subheader("Required action — confirm flagged questions")
                st.write(
                    f"The importer flagged **{flagged_review_count} question(s)** because their answers "
                    "or imported content need verification. Check every row marked **Needs review** above."
                )
                review_acknowledged = st.checkbox(
                    f"I confirm that I reviewed all {flagged_review_count} flagged question(s), their content, and correct answers",
                    value=False,
                    key=f"review_acknowledged_{source_signature}",
                    help="This confirmation is required before the exam package can be created or previewed.",
                )
                if review_acknowledged:
                    st.success("Review confirmed. You can now create or preview the exam package, provided there are no blocking errors.")
                else:
                    st.error("Confirm your review above before creating or previewing the exam package.")
        combined_df, review_report_df = validate_reviewed_questions(reviewed_editor_df)
    else:
        combined_df = imported_df
        review_report_df = pd.DataFrame(
            columns=["Source file", "Row", "Internal_ID", "Severity", "Issue"]
        )

    settings_ui = render_build_settings()
    exam_title = settings_ui["exam_title"]
    num_versions = settings_ui["num_versions"]
    naming = settings_ui["naming"]
    shuffle_questions = settings_ui["shuffle_questions"]
    shuffle_answers = settings_ui["shuffle_answers"]
    include_master_with_answers = settings_ui["include_master_with_answers"]
    enable_fuzzy_duplicates = settings_ui["enable_fuzzy_duplicates"]
    similarity_threshold = settings_ui["similarity_threshold"]
    long_question_limit = settings_ui["long_question_limit"]
    long_answer_limit = settings_ui["long_answer_limit"]
    use_blueprint_advanced = settings_ui["use_blueprint_advanced"]
    kahoot_time_limit = settings_ui["kahoot_time_limit"]
    export_options = {
        "word": settings_ui["export_word"], "answer_key": settings_ui["export_answer_key"],
        "aiken": settings_ui["export_aiken"], "kahoot": settings_ui["export_kahoot"],
        "export_summary": settings_ui["export_summary"],
    }
    formatting = {
        key: settings_ui[key] for key in [
            "font_name", "font_size", "title_font_size", "answer_key_font_size",
            "top_margin_cm", "bottom_margin_cm", "left_margin_cm", "right_margin_cm",
            "option_indent_cm", "space_before_question_pt", "space_after_question_pt",
            "space_after_option_pt", "blank_line_between_questions", "bold_questions",
            "show_page_numbers", "image_width_cm", "show_version_in_header",
            "show_source_in_answer_key", "answer_display",
        ]
    }
    st.session_state.formatting = formatting

    # Editable row-level findings are recalculated from the review table. Word
    # importer notes describe source features (such as image composition or a
    # crop fallback), so they must retain their original question location.
    persistent_import_report_df = import_report_df
    if not import_report_df.empty:
        has_row = import_report_df["Row"].fillna("").astype(str).str.strip().ne("")
        has_internal_id = import_report_df["Internal_ID"].fillna("").astype(str).str.strip().ne("")
        importer_notes = set(
            imported_df.get("Import_Warnings", pd.Series(dtype=str))
            .fillna("")
            .astype(str)
            .str.strip()
        ) - {""}
        is_source_note = import_report_df["Issue"].fillna("").astype(str).str.strip().isin(importer_notes)
        persistent_import_report_df = import_report_df.loc[
            ~(has_row & has_internal_id & ~is_source_note)
        ].copy()

    image_report_df = pd.DataFrame(
        [
            {
                "Source file": "Embedded Word image",
                "Row": "",
                "Internal_ID": "",
                "Severity": "ERROR",
                "Issue": message,
            }
            for message in extracted_image_errors
        ]
    )
    report_frames = [
        frame
        for frame in [persistent_import_report_df, review_report_df, image_report_df]
        if not frame.empty
    ]
    report_df = (
        pd.concat(report_frames, ignore_index=True)
        if report_frames
        else pd.DataFrame(columns=["Source file", "Row", "Internal_ID", "Severity", "Issue"])
    )
    report_df = add_smart_validation_issues(
        combined_df,
        report_df,
        similarity_threshold=similarity_threshold,
        long_question_limit=int(long_question_limit),
        long_answer_limit=int(long_answer_limit),
        enable_fuzzy_duplicates=enable_fuzzy_duplicates
    )

    if not combined_df.empty:
        extra_issues = report_df.to_dict("records") if not report_df.empty else []
        uploaded_image_names = set(st.session_state.get("image_map", {}).keys())
        for _, row in combined_df.iterrows():
            for image_name in split_image_files(row.get("Image_Files") or row.get("Image", "")):
                if image_name.lower() not in uploaded_image_names:
                    extra_issues.append({
                        "Source file": row["Source_File"],
                        "Row": "",
                        "Internal_ID": row["Internal_ID"],
                        "Severity": "ERROR",
                        "Issue": f"Image referenced but not uploaded: {image_name}"
                    })
        report_df = pd.DataFrame(extra_issues) if extra_issues else report_df
    report_df = normalise_validation_report(report_df)
    st.session_state.combined_df = combined_df
    st.session_state.report_df = report_df

    with validation_slot:
        correct_count = int(combined_df["Correct"].fillna("").astype(str).str.strip().ne("").sum()) if not combined_df.empty else 0
        image_question_count = int(combined_df["Image_Files"].fillna("").astype(str).str.strip().ne("").sum()) if not combined_df.empty else 0
        st.markdown(
            f"**{len(combined_df)} questions imported** · **{correct_count} correct answers identified** · "
            f"**{image_question_count} questions contain images**"
        )
        attention_report = report_df.loc[report_df["Severity"].isin(["ERROR", "WARNING"])]
        import_notes = report_df.loc[report_df["Severity"] == "INFO"]
        for _, note in import_notes.iterrows():
            st.info(f"{note['Source file']}: {note['Issue']}")
        attention_count = len(attention_report)
        if attention_count:
            error_count = int((report_df["Severity"] == "ERROR").sum())
            warning_count = int((report_df["Severity"] == "WARNING").sum())
            finding_parts = []
            if error_count:
                finding_parts.append(f"{error_count} blocking error{'s' if error_count != 1 else ''}")
            if warning_count:
                finding_parts.append(f"{warning_count} warning{'s' if warning_count != 1 else ''}")
            attention_verb = "needs" if attention_count == 1 else "need"
            attention_message = (
                f"⚠️ {' and '.join(finding_parts)} {attention_verb} your attention. "
                "The affected question or row is listed in **Validation details** immediately below."
            )
            if error_count:
                st.error(attention_message)
            else:
                st.warning(attention_message)

            st.subheader("Validation details — affected questions and locations")
            st.caption(
                "Use the Location column to find each item in the uploaded source. "
                "It includes the source file, question or row number, and stable question ID when available."
            )
            st.dataframe(attention_report, width="stretch", hide_index=True)

        with st.expander("Import details & issue counts"):
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Source files", len(uploaded_files))
            c2.metric("Questions", len(combined_df))
            c3.metric("Errors", int((report_df["Severity"] == "ERROR").sum()) if not report_df.empty else 0)
            c4.metric("Warnings", int((report_df["Severity"] == "WARNING").sum()) if not report_df.empty else 0)
            if not report_df.empty:
                issue_summary = report_df.groupby(["Severity", "Issue"]).size().reset_index(name="Count")
                st.dataframe(issue_summary, width="stretch", hide_index=True)
            if not combined_df.empty:
                counts = combined_df["Option_Count"].value_counts().sort_index()
                st.caption("Answer options: " + " · ".join(f"{int(n)} options: {count} questions" for n, count in counts.items()))

        if not attention_report.empty:
            if has_blocking_errors(report_df):
                st.error("Validation errors found. Fix them before creating or previewing the exam package.")
            else:
                st.info("Validation warnings do not prevent package creation, but review the locations listed above.")
        else:
            st.success("No validation issues detected.")

    blueprint_counts = {}
    selected_df = combined_df.copy()

    if use_blueprint_advanced:
        st.subheader("Question selection by source file")
        st.write("Choose how many questions to include from each uploaded file.")
        source_files = sorted(combined_df["Source_File"].unique())
        for source_file in source_files:
            available = int((combined_df["Source_File"] == source_file).sum())
            blueprint_counts[source_file] = st.number_input(
                f"{source_file} — available {available}",
                min_value=0,
                max_value=available,
                value=available,
                step=1,
                key=f"blueprint_{source_file}"
            )
        selected_df = select_questions_by_blueprint(combined_df, blueprint_counts, seed=12345)
    else:
        for source_file in sorted(combined_df["Source_File"].unique()) if not combined_df.empty else []:
            available = int((combined_df["Source_File"] == source_file).sum())
            blueprint_counts[source_file] = available

    if not combined_df.empty:
        blueprint_df = make_blueprint_dataframe(combined_df, blueprint_counts)
        if use_blueprint_advanced:
            st.dataframe(blueprint_df, width="stretch")
        if use_blueprint_advanced:
            st.caption(f"Questions included in the exam: {len(selected_df)} of {len(combined_df)}")
        st.session_state.blueprint_df = blueprint_df

    preview_fingerprint_payload = {
        "exam_title": exam_title,
        "num_versions": int(num_versions),
        "naming": naming,
        "shuffle_questions": bool(shuffle_questions),
        "shuffle_answers": bool(shuffle_answers),
        "include_master_with_answers": bool(include_master_with_answers),
        "formatting": formatting,
        "export_options": export_options,
        "kahoot_time_limit": int(kahoot_time_limit),
        "blueprint_counts": blueprint_counts,
    }
    current_preview_fingerprint = make_preview_fingerprint(
        selected_df,
        preview_fingerprint_payload,
        st.session_state.get("image_map", {}),
    )
    if (
        st.session_state.preview_versions
        and st.session_state.preview_input_fingerprint != current_preview_fingerprint
    ):
        st.session_state.preview_versions = None
        st.session_state.preview_settings = None
        st.session_state.selected_df_for_preview = None
        st.session_state.preview_input_fingerprint = None
        st.info("The questions or settings changed. Create the package again or generate a new preview before downloading.")

    st.subheader("Package summary")
    selected_exports = []
    if export_options.get("word"): selected_exports.append("Word exams")
    if export_options.get("answer_key"): selected_exports.append("Word answer keys")
    if export_options.get("aiken"): selected_exports.append("Blackboard/Aiken")
    if export_options.get("kahoot"): selected_exports.append("Kahoot")
    with st.container(border=True):
        st.write(exam_title)
        st.caption(f"{len(selected_df)} questions · {int(num_versions)} version(s) · "
                   f"{', '.join(selected_exports) if selected_exports else 'Reports only'}")
        st.caption(f"Shuffle questions: {'On' if shuffle_questions else 'Off'} · "
                   f"Shuffle answers: {'On' if shuffle_answers else 'Off'} · "
                   f"{formatting['font_name']} {formatting['font_size']} pt")

    if not has_blocking_errors(report_df) and review_acknowledged and not selected_df.empty:
        direct_column, preview_column = st.columns(2)
        with direct_column:
            direct_download_clicked = st.button(
                "Create exam package",
                type="primary",
                use_container_width=True,
                help="Create the final package now without opening the optional on-screen preview.",
            )
            st.caption("Prepare your files, then download the ZIP below.")
        with preview_column:
            preview_clicked = st.button(
                "Preview exam first",
                type="secondary",
                use_container_width=True,
                help="Inspect every version and optionally re-shuffle before downloading.",
            )
            st.caption("Optional: inspect every exam version and answer key before downloading.")

        if direct_download_clicked or preview_clicked:
            review_before_download = bool(preview_clicked)
            st.session_state.shuffle_seed = 10000
            versions = build_versions(
                selected_df,
                num_versions=int(num_versions),
                naming=naming,
                shuffle_questions=shuffle_questions,
                shuffle_answers=shuffle_answers,
                base_seed=st.session_state.shuffle_seed
            )
            st.session_state.preview_versions = versions
            st.session_state.preview_settings = {
                "exam_title": exam_title,
                "num_versions": int(num_versions),
                "naming": naming,
                "shuffle_questions": shuffle_questions,
                "shuffle_answers": shuffle_answers,
                "include_master_with_answers": include_master_with_answers,
                "formatting": st.session_state.formatting,
                "export_options": export_options,
                "kahoot_time_limit": int(kahoot_time_limit),
                "review_before_download": review_before_download
            }
            st.session_state.combined_df = selected_df
            st.session_state.selected_df_for_preview = selected_df
            st.session_state.preview_input_fingerprint = current_preview_fingerprint
            if review_before_download:
                st.success("Preview ready. Review it below and reshuffle if needed before downloading.")
                st.markdown("[Go to preview ↓](#preview-exam) · [Go to download ↓](#download-package)")
            else:
                st.markdown("[Go to your download ↓](#download-package)")
    elif not has_blocking_errors(report_df) and not review_acknowledged and not selected_df.empty:
        st.error(
            "Package creation and preview are unavailable because the required review has not been confirmed. "
            "Use the link below to return to the confirmation panel."
        )
        st.markdown("[↑ Go to the required review confirmation](#required-review)")
        locked_direct_column, locked_preview_column = st.columns(2)
        with locked_direct_column:
            st.button("Create exam package — review required", disabled=True)
        with locked_preview_column:
            st.button("Preview exam first — review required", disabled=True)
    elif not has_blocking_errors(report_df) and selected_df.empty and not combined_df.empty:
        st.warning("No questions selected. Increase the question count for at least one source file.")

else:
    st.session_state.image_map = manual_image_map
    st.session_state.preview_versions = None
    st.session_state.preview_settings = None
    st.session_state.selected_df_for_preview = None
    st.session_state.preview_input_fingerprint = None



if st.session_state.preview_versions and (st.session_state.preview_settings or {}).get("review_before_download", True):
    st.markdown("---")
    st.subheader("Exam preview", anchor="preview-exam")
    st.caption("Optional: check your generated exam before downloading. You can reshuffle below.")

    versions = st.session_state.preview_versions
    settings = st.session_state.preview_settings or {}
    labels = list(versions.keys())

    shuffle_action, shuffle_note = st.columns([1, 3])
    with shuffle_action:
        if st.button("Reshuffle all versions", type="secondary"):
            selected_df_current = st.session_state.get("selected_df_for_preview")
            if selected_df_current is not None:
                st.session_state.shuffle_seed += 1000
                st.session_state.preview_versions = build_versions(
                    selected_df_current,
                    num_versions=settings.get("num_versions", len(labels)),
                    naming=settings.get("naming", "Letters"),
                    shuffle_questions=settings.get("shuffle_questions", True),
                    shuffle_answers=settings.get("shuffle_answers", True),
                    base_seed=st.session_state.shuffle_seed
                )
                st.rerun()
    with shuffle_note:
        st.caption("The download uses the exact question and answer order shown here. "
                   "Reshuffle to create a different arrangement.")

    versions = st.session_state.preview_versions
    labels = list(versions.keys())

    # Compact version selection. For a small number of versions, use tabs-like segmented control.
    # For many versions, use a narrow dropdown instead of a full-width selector.
    col_version, col_summary = st.columns([1, 4])
    with col_version:
        if len(labels) <= 4:
            selected_version = st.radio("Version", labels, horizontal=True, label_visibility="visible")
        else:
            selected_version = st.selectbox("Version", labels, label_visibility="visible")
    selected_questions = versions[selected_version]

    with col_summary:
        make_exam_summary(versions, st.session_state.combined_df, settings)

    st.caption(f"Version {selected_version} · Arrangement {int((st.session_state.shuffle_seed - 10000) / 1000) + 1} · "
               f"{settings.get('formatting', {}).get('font_name', 'Arial')} "
               f"{settings.get('formatting', {}).get('font_size', 12)} pt in Word")

    review_tabs = st.tabs(["Student view", "Answer key", "Question mapping", "Statistics"])

    with review_tabs[0]:
        st.subheader(f"Student View — Version {selected_version}")
        col_a, col_b = st.columns([2, 1])
        with col_a:
            search = st.text_input("Search in this version", value="", key=f"search_{selected_version}")
        with col_b:
            max_preview = st.number_input(
                "Questions to show",
                min_value=1,
                max_value=max(1, min(500, len(selected_questions))),
                value=max(1, min(50, len(selected_questions))),
                step=1,
            )
        show_student_preview(selected_questions, max_preview=int(max_preview), search=search, image_map=st.session_state.get("image_map", {}))

    with review_tabs[1]:
        st.subheader(f"Answer Key — Version {selected_version}")
        answer_search = st.text_input("Search answer key", value="", key=f"answer_search_{selected_version}")
        show_answer_preview(selected_questions, answer_search)

    with review_tabs[2]:
        st.subheader("Version Mapping")
        mapping_df = make_mapping_dataframe(versions)
        st.dataframe(mapping_df[mapping_df["Version"] == selected_version], width="stretch")
        with st.expander("Show mapping for all versions"):
            st.dataframe(mapping_df, width="stretch")

    with review_tabs[3]:
        st.subheader("Statistics")
        show_statistics(
            versions,
            st.session_state.combined_df,
            settings.get("shuffle_questions", False),
            settings.get("shuffle_answers", False)
        )
        if settings.get("export_options", {}).get("kahoot", False):
            st.subheader("Kahoot export warnings")
            kahoot_warnings = make_kahoot_limit_report(versions)
            if kahoot_warnings.empty:
                st.success("No Kahoot warnings detected.")
            else:
                st.warning("Some questions may be truncated or skipped in Kahoot export. Check details below.")
                st.dataframe(kahoot_warnings, width="stretch")

    st.markdown("---")
    st.header("4 — Download", anchor="download-package")
    st.write("The exam package will use exactly the versions and answer order shown in the preview above.")

    # One-click download: generate the ZIP before rendering the download button.
    # This avoids the previous two-step workflow: button -> generate -> second download button.
    with st.spinner("Preparing your exam files…"):
        zip_output = create_zip_output_from_versions(
            versions=versions,
            combined_df=st.session_state.combined_df,
            report_df=st.session_state.report_df,
            exam_title=settings.get("exam_title", "MCQ Examination"),
            include_master_with_answers=settings.get("include_master_with_answers", True),
            image_map=st.session_state.get("image_map", {}),
            formatting=settings.get("formatting", st.session_state.get("formatting", get_default_formatting())),
            export_options=settings.get("export_options", {"word": True, "answer_key": True}),
            kahoot_time_limit=settings.get("kahoot_time_limit", 20),
            blueprint_df=st.session_state.get("blueprint_df"),
        )

    st.success("Your exam package is ready")
    st.download_button(
        "Download exam package (.zip)",
        data=zip_output,
        file_name=f"{make_safe_filename(settings.get('exam_title', 'MCQ_Exam'))}_Package.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )


if st.session_state.preview_versions and not (st.session_state.preview_settings or {}).get("review_before_download", True):
    st.markdown("---")
    st.header("4 — Download", anchor="download-package")
    st.write("The selected exam files have been created and packaged in a ZIP file.")
    direct_settings = st.session_state.preview_settings or {}
    with st.spinner("Preparing your exam files…"):
        direct_zip_output = create_zip_output_from_versions(
            versions=st.session_state.preview_versions,
            combined_df=st.session_state.combined_df,
            report_df=st.session_state.report_df,
            exam_title=direct_settings.get("exam_title", "MCQ Examination"),
            include_master_with_answers=direct_settings.get("include_master_with_answers", True),
            image_map=st.session_state.get("image_map", {}),
            formatting=direct_settings.get("formatting", st.session_state.get("formatting", get_default_formatting())),
            export_options=direct_settings.get("export_options", {"word": True, "answer_key": True}),
            kahoot_time_limit=direct_settings.get("kahoot_time_limit", 20),
            blueprint_df=st.session_state.get("blueprint_df"),
        )
    st.success("Your exam package is ready")
    st.download_button(
        "Download exam package (.zip)",
        data=direct_zip_output,
        file_name=f"{make_safe_filename(direct_settings.get('exam_title', 'MCQ_Exam'))}_Package.zip",
        mime="application/zip",
        type="primary",
    )


with progress_slot:
    has_package = bool(st.session_state.preview_versions)
    current_step = 1
    if uploaded_files:
        current_step = 2 if has_blocking_errors(report_df) or not review_acknowledged or combined_df.empty else 3
    if has_package:
        current_step = 4
    show_progress(current_step, bool(uploaded_files), has_package)


st.markdown(
    f'<div class="app-footer"><span><strong>MCQ Exam Builder</strong> · {APP_VERSION}</span>'
    f'<span>Created by <strong>{APP_AUTHOR}</strong></span></div>',
    unsafe_allow_html=True,
)

with st.expander("About"):
    st.markdown("**From existing question banks to checked exam versions**")
    st.write(
        "Preparing several versions of an examination means more than rearranging questions. "
        "Answer keys must follow each shuffle, figures must stay with their questions, and "
        "inconsistent source files need review before they become final papers. "
        "MCQ Exam Builder brings these tasks into one workflow using the material lecturers already have."
    )
    st.write(
        "The software checks structure and answer mappings, flags detected ambiguity, and generates "
        "reproducible arrangements. It does not write questions or judge academic correctness; "
        "the lecturer reviews the content and approves the examination."
    )
    st.markdown("**Created and maintained by Dr Christos I. Karagiannis**")
    st.caption("DVM, MSc, Dip. ECAWBM(BM), MRCVS · Veterinary behavioural specialist, Assistant Professor and programme coordinator")
    st.write(
        "My teaching and programme coordination work shapes this project: which preparation tasks "
        "need support, which errors matter, and where a lecturer must remain in control."
    )
    st.caption("An independent project. Not an institutional examination platform or endorsement.")

footer_links = []
if os.getenv("MCQ_BUILDER_GITHUB_URL"):
    footer_links.append(f"[GitHub]({os.getenv('MCQ_BUILDER_GITHUB_URL')})")
if os.getenv("MCQ_BUILDER_LINKEDIN_URL"):
    footer_links.append(f"[LinkedIn]({os.getenv('MCQ_BUILDER_LINKEDIN_URL')})")
if footer_links:
    st.markdown(" · ".join(footer_links))
