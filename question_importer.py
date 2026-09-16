from __future__ import annotations

import hashlib
import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from PIL import Image, ImageOps


MAX_OPTIONS = 6
LETTERS = list("ABCDEF")
OPTION_COLUMNS = [f"Option {i}" for i in range(1, MAX_OPTIONS + 1)]
CORE_COLUMNS = ["No", "Question", "Image", *OPTION_COLUMNS, "Correct"]
META_COLUMNS = [
    "Image_Files",
    "Option_Count",
    "Answer_Columns",
    "Source_File",
    "Internal_ID",
    "Import_Method",
    "Answer_Evidence",
    "Confidence",
    "Needs_Review",
    "Import_Warnings",
]
NORMALIZED_COLUMNS = CORE_COLUMNS + META_COLUMNS
MAX_SOURCE_BYTES = 200 * 1024 * 1024
MAX_OFFICE_MEMBERS = 10_000
MAX_OFFICE_UNCOMPRESSED_BYTES = 1_000 * 1024 * 1024


QUESTION_ALIASES = {
    "question",
    "question text",
    "question_text",
    "stem",
    "item",
    "mcq",
    "prompt",
}
CORRECT_ALIASES = {
    "correct",
    "correct answer",
    "correct_answer",
    "answer key",
    "answer_key",
    "key",
    "answer",
}
NUMBER_ALIASES = {"no", "number", "question no", "question number", "question id", "id"}
IMAGE_ALIASES = {"image", "image file", "image filename", "picture", "figure", "asset"}


def validate_office_container(data: bytes, filename: str) -> None:
    """Reject malformed or unreasonably expanded modern Office containers."""
    if len(data) > MAX_SOURCE_BYTES:
        raise ValueError("File exceeds the 200 MB source limit.")
    suffix = Path(filename).suffix.lower()
    if suffix not in {".docx", ".docm", ".dotx", ".dotm", ".xlsx", ".xlsm", ".xltx", ".xltm", ".odt", ".ott", ".ods"}:
        return
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            members = archive.infolist()
            if len(members) > MAX_OFFICE_MEMBERS:
                raise ValueError("Office file contains an unreasonable number of internal parts.")
            expanded_size = sum(member.file_size for member in members)
            if expanded_size > MAX_OFFICE_UNCOMPRESSED_BYTES:
                raise ValueError("Office file expands beyond the 1 GB safety limit.")
            if any(member.file_size < 0 or member.compress_size < 0 for member in members):
                raise ValueError("Office file contains invalid internal part sizes.")
    except zipfile.BadZipFile as exc:
        raise ValueError("File is not a valid modern Office container.") from exc


def normalise_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).replace("\u00a0", " ").strip()


def normalise_header(value) -> str:
    text = normalise_text(value).lower()
    text = re.sub(r"[\s\-]+", " ", text)
    return text.strip()


def safe_prefix(filename: str) -> str:
    if filename.endswith(']') and ' [' in filename:
        source, sheet = filename.rsplit(' [', 1)
        stem = Path(source).stem + '_' + sheet[:-1] + '_' + hashlib.sha256(sheet.encode()).hexdigest()[:8]
    else:
        stem = Path(filename).stem
    value = re.sub(r"[^A-Za-z0-9]+", "_", stem).strip("_")
    return value or "MCQ_SOURCE"


def internal_id(filename: str, number, row_number: int) -> str:
    number_text = normalise_text(number)
    if number_text:
        try:
            parsed = float(number_text)
            if parsed.is_integer():
                number_text = str(int(parsed))
        except (TypeError, ValueError):
            pass
        return f"{safe_prefix(filename)}-{number_text}"
    return f"{safe_prefix(filename)}-ROW{row_number}"


def issue(filename: str, row, internal: str, severity: str, message: str) -> Dict:
    return {
        "Source file": filename,
        "Row": row,
        "Internal_ID": internal,
        "Severity": severity,
        "Issue": message,
    }


def validation_issue_location(finding) -> str:
    """Return a lecturer-facing location for a validation finding."""
    source = normalise_text(finding.get("Source file"))
    row = normalise_text(finding.get("Row"))
    internal = normalise_text(finding.get("Internal_ID"))
    parts = []
    if source:
        parts.append(source)
    if row:
        parts.append(f"question/row {row}")
    if internal:
        parts.append(f"ID {internal}")
    return " · ".join(parts) or "General validation"


def split_image_files(value) -> List[str]:
    if isinstance(value, list):
        return [normalise_text(v) for v in value if normalise_text(v)]
    text = normalise_text(value)
    if not text:
        return []
    return [part.strip() for part in text.split("|") if part.strip()]


def _option_column_info(header: str) -> Optional[Tuple[int, str]]:
    raw = normalise_text(header)
    clean = normalise_header(raw)
    patterns = [
        r"(?:option|answer|choice)\s*([1-6])$",
        r"(?:option|answer|choice)\s*([a-f])$",
        r"answer[_\s]*([a-f])$",
        r"^([a-f])$",
    ]
    for pattern in patterns:
        match = re.fullmatch(pattern, clean, flags=re.IGNORECASE)
        if not match:
            continue
        token = match.group(1).upper()
        number = int(token) if token.isdigit() else LETTERS.index(token) + 1
        return number, raw
    return None


def _find_alias(columns: Sequence[str], aliases: set[str]) -> Optional[str]:
    for column in columns:
        if normalise_header(column) in aliases:
            return column
    return None


def _excel_header_score(values: Iterable) -> int:
    headers = [normalise_header(v) for v in values]
    score = 0
    if any(h in QUESTION_ALIASES for h in headers):
        score += 8
    if any(h in CORRECT_ALIASES for h in headers):
        score += 5
    score += min(6, sum(1 for h in headers if _option_column_info(h))) * 2
    if any(h in NUMBER_ALIASES for h in headers):
        score += 1
    return score


def _find_excel_layouts(data: bytes):
    layouts, skipped = [], []
    with pd.ExcelFile(io.BytesIO(data)) as workbook:
        for sheet_name in workbook.sheet_names:
            raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None,
                                nrows=40, keep_default_na=False, dtype=object)
            score, header_row = max(
                ((_excel_header_score(raw.iloc[i].tolist()), i) for i in range(min(20, len(raw)))),
                default=(0, 0), key=lambda item: item[0],
            )
            if score < 12:
                skipped.append(sheet_name)
                continue
            raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None,
                                keep_default_na=False, dtype=object)
            frame = raw.iloc[header_row + 1:].reset_index(drop=True).copy()
            frame.columns = [normalise_text(c) for c in raw.iloc[header_row]]
            layouts.append((sheet_name, header_row, frame))
    if not layouts:
        raise ValueError("Could not identify a question table with Question, option and Correct columns.")
    return layouts, skipped


def _normalise_correct_value(value, option_values: List[str]) -> Tuple[str, str, bool]:
    text = normalise_text(value)
    if not text:
        return "", "Correct answer not provided", True

    compact = text.upper().strip()
    compact = re.sub(r"^(?:CORRECT\s*ANSWER|ANSWER|OPTION|CHOICE)\s*[:=\-]?\s*", "", compact)
    compact = re.sub(r"^[()\[\]{}:;.\-\s]+|[()\[\]{}:;.\-\s]+$", "", compact)

    if re.search(r"[,/&+]", compact) or re.search(r"\bAND\b", compact):
        candidates = re.findall(r"\b([A-F])\b", compact)
        if len(set(candidates)) > 1:
            return ",".join(dict.fromkeys(candidates)), "Multiple correct answers detected", True

    if compact in LETTERS:
        return compact, "Explicit correct-answer value", False
    if compact.isdigit() and 1 <= int(compact) <= MAX_OPTIONS:
        return LETTERS[int(compact) - 1], "Numeric correct-answer value", False

    matches = [index for index, option in enumerate(option_values)
               if normalise_text(option).casefold() == text.casefold()]
    if len(matches) > 1:
        return "", "Correct-answer text matches multiple options; choose a letter", True
    if matches:
        return LETTERS[matches[0]], "Correct answer matched by option text", False

    return compact, "Unrecognised correct-answer value", True


def _validate_and_finalize_row(
    record: Dict,
    filename: str,
    row_number: int,
    issues: List[Dict],
    strict_correct: bool = False,
) -> Dict:
    internal = normalise_text(record.get("Internal_ID")) or internal_id(
        filename, record.get("No", ""), row_number
    )
    record["Internal_ID"] = internal
    record["Source_File"] = normalise_text(record.get("Source_File")) or filename
    record["Question"] = normalise_text(record.get("Question"))

    option_values = [normalise_text(record.get(column, "")) for column in OPTION_COLUMNS]
    active_count = 0
    gap = False
    gap_columns = []
    for idx, value in enumerate(option_values):
        if value and not gap:
            active_count += 1
        elif not value:
            gap = True
        elif value and gap:
            gap_columns.append(OPTION_COLUMNS[idx])

    correct, evidence, needs_review = _normalise_correct_value(record.get("Correct", ""), option_values)
    if normalise_text(record.get("Answer_Evidence")):
        evidence = normalise_text(record.get("Answer_Evidence"))
    record["Correct"] = correct
    record["Option_Count"] = active_count
    record["Answer_Columns"] = ",".join(OPTION_COLUMNS[:active_count])
    record["Answer_Evidence"] = evidence

    if not record["Question"]:
        issues.append(issue(filename, row_number, internal, "ERROR", "Missing question text."))
    if active_count < 2:
        issues.append(issue(filename, row_number, internal, "ERROR", "At least two answer options are required."))
    if gap_columns:
        issues.append(
            issue(
                filename,
                row_number,
                internal,
                "ERROR",
                "Options must be consecutive; content was found after a blank option.",
            )
        )

    allowed = LETTERS[:active_count]
    if correct not in allowed:
        needs_review = True
        if correct or strict_correct:
            message = (
                f"Correct answer must be one of {', '.join(allowed) if allowed else 'the available options'}."
                if correct
                else f"Choose the correct answer ({', '.join(allowed)}) before building the exam."
            )
            issues.append(issue(filename, row_number, internal, "ERROR", message))

    folded = [v.casefold() for v in option_values[:active_count] if v]
    if len(folded) != len(set(folded)):
        issues.append(issue(filename, row_number, internal, "WARNING", "Duplicate option text detected."))

    image_files = split_image_files(record.get("Image_Files") or record.get("Image"))
    record["Image_Files"] = "|".join(image_files)
    record["Image"] = image_files[0] if image_files else ""
    record["Needs_Review"] = bool(record.get("Needs_Review")) or needs_review
    if not normalise_text(record.get("Confidence")):
        record["Confidence"] = "Review required" if record["Needs_Review"] else "High"
    record.setdefault("Import_Warnings", "")
    record.setdefault("Import_Method", "")
    for column in NORMALIZED_COLUMNS:
        record.setdefault(column, "")
    return record


def read_excel_questions(data: bytes, filename: str) -> Tuple[pd.DataFrame, List[Dict], Dict[str, bytes]]:
    issues: List[Dict] = []
    try:
        validate_office_container(data, filename)
        if zipfile.is_zipfile(io.BytesIO(data)):
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                if any(name.startswith(('xl/media/', 'xl/embeddings/')) for name in archive.namelist()):
                    raise ValueError('Spreadsheet contains embedded images or objects. Use DOCX or image filenames with separately uploaded images.')
                for name in archive.namelist():
                    if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                        root = parse_xml(archive.read(name))
                        if root.xpath(".//*[local-name()='f']"):
                            raise ValueError('Spreadsheet contains formula cells. Paste their calculated results as values before uploading; cached results may be missing or outdated.')
        layouts, skipped = _find_excel_layouts(data)
    except Exception as exc:  # noqa: BLE001 - pandas engines raise format-specific exceptions.
        return pd.DataFrame(columns=NORMALIZED_COLUMNS), [
            issue(filename, "", "", "ERROR", f"Excel import failed: {exc}")
        ], {}

    frames = []
    for sheet_name, header_row, frame in layouts:
        source = f'{filename} [{sheet_name}]' if len(layouts) > 1 else filename
        imported, findings, _ = read_question_table(frame, source, header_row, sheet_name)
        frames.append(imported)
        issues.extend(findings)
    if skipped:
        issues.append(issue(filename, '', '', 'INFO',
                            'No question table detected in these sheets: ' + ', '.join(skipped)))
    return pd.concat(frames, ignore_index=True), issues, {}


def read_question_table(frame, filename, header_row=0, sheet_name="Questions", method_label=None):
    """Normalize a parsed table using the same answer and question validation."""
    issues: List[Dict] = []
    columns = list(frame.columns)
    headers = [normalise_header(c) for c in columns if normalise_text(c)]
    if len(headers) != len(set(headers)):
        raise ValueError('Duplicate table headers detected. Use one unique header for each column.')
    for aliases, role in [(QUESTION_ALIASES, 'question'), (CORRECT_ALIASES, 'correct-answer')]:
        if sum(header in aliases for header in headers) > 1:
            raise ValueError(f'Multiple columns identify the {role} field. Keep one authoritative column.')
    if any(re.fullmatch(r"(?:(?:option|answer|choice)\s*)?(?:[g-z]|[7-9]|\d{2,})", str(c).strip(), re.I) for c in columns):
        raise ValueError("The app supports a maximum of six answer options per question. Remove extra option columns before importing.")
    question_col = _find_alias(columns, QUESTION_ALIASES)
    correct_col = _find_alias(columns, CORRECT_ALIASES)
    number_col = _find_alias(columns, NUMBER_ALIASES)
    image_col = _find_alias(columns, IMAGE_ALIASES)

    option_info = []
    for column in columns:
        detected = _option_column_info(column)
        if detected:
            option_info.append((detected[0], column))
    if len({n for n, _ in option_info}) != len(option_info):
        raise ValueError("More than one column maps to the same answer option. Use unique Option 1–6 headers.")
    option_info = sorted(option_info)
    if option_info and [n for n, _ in option_info] != list(range(1, len(option_info) + 1)):
        raise ValueError("Answer columns must be consecutive, starting with Option 1 or A. An option column is missing.")

    if question_col is None or correct_col is None or len(option_info) < 2:
        return pd.DataFrame(columns=NORMALIZED_COLUMNS), [
            issue(
                filename,
                header_row + 1,
                "",
                "ERROR",
                "The detected question table must include a question column, at least two option columns and a correct-answer column.",
            )
        ], {}

    exact_template = (
        normalise_header(question_col) == "question"
        and normalise_header(correct_col) == "correct"
        and [number for number, _ in option_info] == list(range(1, len(option_info) + 1))
    )
    method = method_label or ("Universal Excel template" if exact_template else "Flexible Excel column mapping")

    records = []
    for idx, row in frame.iterrows():
        question = normalise_text(row.get(question_col, ""))
        option_values = [normalise_text(row.get(column, "")) for _, column in option_info]
        if not question and not any(option_values) and not normalise_text(row.get(correct_col, "")):
            continue

        excel_row = int(idx) + header_row + 2
        record = {
            "No": normalise_text(row.get(number_col, "")) if number_col else "",
            "Question": question,
            "Image": normalise_text(row.get(image_col, "")) if image_col else "",
            "Correct": row.get(correct_col, ""),
            "Source_File": filename,
            "Import_Method": method,
            "Answer_Evidence": "Correct-answer field" if method_label else "Excel correct-answer column",
            "Confidence": "High" if exact_template else "Review recommended",
            "Needs_Review": not exact_template,
            "Import_Warnings": "" if exact_template else f"Mapped from sheet '{sheet_name}', header row {header_row + 1}.",
        }
        for option_index in range(MAX_OPTIONS):
            record[OPTION_COLUMNS[option_index]] = option_values[option_index] if option_index < len(option_values) else ""
        record["Image_Files"] = record["Image"]
        record = _validate_and_finalize_row(record, filename, excel_row, issues)
        records.append(record)

    return pd.DataFrame(records, columns=NORMALIZED_COLUMNS), issues, {}


@dataclass
class LineRecord:
    paragraph_index: int
    text: str
    bold_ratio: float
    highlighted: bool
    underlined: bool
    coloured: bool
    num_id: Optional[int]
    num_level: Optional[int]
    style: str
    images: List[Tuple[str, bytes]] = field(default_factory=list)
    drawing_count: int = 0
    warnings: List[str] = field(default_factory=list)


@dataclass
class ParsedOption:
    text: str
    source_label: str = ""
    highlighted: bool = False
    bold_ratio: float = 0.0
    underlined: bool = False
    coloured: bool = False


@dataclass
class ParsedQuestion:
    number: str
    text: str
    paragraph_index: int
    num_id: Optional[int]
    options: List[ParsedOption] = field(default_factory=list)
    images: List[Tuple[str, bytes]] = field(default_factory=list)
    drawing_count: int = 0
    option_num_id: Optional[int] = None
    warnings: List[str] = field(default_factory=list)
    explicit_answer: str = ""
    answer_evidence: str = ""


class WordDrawingAsset(bytes):
    """Raster-compatible Word drawing retaining its complete OOXML figure."""

    def __new__(
        cls,
        preview_blob: bytes,
        drawing_xml: bytes,
        relationships: Dict[str, Tuple[str, bytes]],
    ):
        instance = super().__new__(cls, preview_blob)
        instance.drawing_xml = drawing_xml
        instance.relationships = relationships
        return instance


def word_drawing_fingerprint(data: bytes) -> bytes:
    """Hash both the preview raster and any preserved Word drawing payload."""
    digest = hashlib.sha256(bytes(data))
    if isinstance(data, WordDrawingAsset):
        digest.update(data.drawing_xml)
        for relationship_id in sorted(data.relationships):
            name, blob = data.relationships[relationship_id]
            digest.update(relationship_id.encode("utf-8"))
            digest.update(name.encode("utf-8"))
            digest.update(hashlib.sha256(blob).digest())
    return digest.digest()


def insert_preserved_word_drawing(document: Document, run, data: bytes) -> bool:
    """Insert a preserved Word drawing as a stable inline figure."""
    if not isinstance(data, WordDrawingAsset):
        return False

    drawing = parse_xml(data.drawing_xml)
    for blip in drawing.xpath(".//a:blip"):
        old_relationship_id = blip.get(qn("r:embed"))
        if not old_relationship_id or old_relationship_id not in data.relationships:
            continue
        _, blob = data.relationships[old_relationship_id]
        new_relationship_id, _ = document.part.get_or_add_image(io.BytesIO(blob))
        blip.set(qn("r:embed"), new_relationship_id)
        # Some Office image-effect extensions point to proprietary fallback
        # parts. The primary raster, its crop and all grouped shapes remain.
        extension_list = blip.find(qn("a:extLst"))
        if extension_list is not None:
            blip.remove(extension_list)

    if drawing.tag == qn("wp:anchor"):
        inline = OxmlElement("wp:inline")
        for attribute in ("distT", "distB", "distL", "distR"):
            inline.set(attribute, "0")
        for child_name in (
            "wp:extent",
            "wp:effectExtent",
            "wp:docPr",
            "wp:cNvGraphicFramePr",
            "a:graphic",
        ):
            child = drawing.find(qn(child_name))
            if child is not None:
                inline.append(child)
        drawing = inline

    document_properties = drawing.find(qn("wp:docPr"))
    if document_properties is not None:
        existing_ids = []
        for existing in document.element.xpath(".//wp:docPr"):
            try:
                existing_ids.append(int(existing.get("id", "0")))
            except ValueError:
                continue
        next_id = max(existing_ids, default=0) + 1
        document_properties.set("id", str(next_id))
        document_properties.set("name", f"Imported Word figure {next_id}")

    run._r.add_drawing(drawing)
    return True


def _clamp_crop_value(value: int) -> int:
    return max(0, min(100_000, int(value)))


def _parse_drawingml_crop_value(value) -> int:
    """Parse an OOXML percentage, where 100000 represents 100 percent."""
    if value is None:
        return 0
    text = str(value).strip()
    if not text:
        return 0
    try:
        if text.endswith("%"):
            return _clamp_crop_value(round(float(text[:-1]) * 1000))
        return _clamp_crop_value(round(float(text)))
    except ValueError:
        return 0


def _parse_vml_crop_value(value) -> int:
    """Parse legacy VML crop fractions into DrawingML percentage units."""
    if value is None:
        return 0
    text = str(value).strip().lower()
    if not text:
        return 0
    try:
        if text.endswith("f"):
            return _clamp_crop_value(round(float(text[:-1]) / 65_536 * 100_000))
        if text.endswith("%"):
            return _clamp_crop_value(round(float(text[:-1]) * 1000))
        parsed = float(text)
        if 0 <= parsed <= 1:
            parsed *= 100_000
        return _clamp_crop_value(round(parsed))
    except ValueError:
        return 0


def _drawingml_crop_rect(blip) -> Tuple[int, int, int, int]:
    blip_fill = blip.getparent()
    source_rect = blip_fill.find(qn("a:srcRect")) if blip_fill is not None else None
    if source_rect is None:
        return (0, 0, 0, 0)
    return tuple(
        _parse_drawingml_crop_value(source_rect.get(side))
        for side in ("l", "t", "r", "b")
    )


def _vml_crop_rect(image_data) -> Tuple[int, int, int, int]:
    return tuple(
        _parse_vml_crop_value(image_data.get(attribute))
        for attribute in ("cropleft", "croptop", "cropright", "cropbottom")
    )


def _apply_word_image_crop(
    original_name: str,
    blob: bytes,
    crop_rect: Tuple[int, int, int, int],
) -> Tuple[str, bytes, str]:
    """Render Word's display crop into the extracted raster image bytes."""
    left, top, right, bottom = crop_rect
    if not any(crop_rect):
        return original_name, blob, ""
    if left + right >= 100_000 or top + bottom >= 100_000:
        return original_name, blob, f"Invalid Word crop values for {original_name}; the full image was retained."

    try:
        with Image.open(io.BytesIO(blob)) as source:
            source_format = (source.format or "PNG").upper()
            rendered = ImageOps.exif_transpose(source)
            width, height = rendered.size
            box = (
                round(width * left / 100_000),
                round(height * top / 100_000),
                round(width * (100_000 - right) / 100_000),
                round(height * (100_000 - bottom) / 100_000),
            )
            if box[2] <= box[0] or box[3] <= box[1]:
                raise ValueError("crop has no visible area")
            cropped = rendered.crop(box)
            output = io.BytesIO()
            save_kwargs = {}
            if source.info.get("icc_profile"):
                save_kwargs["icc_profile"] = source.info["icc_profile"]
            if source_format in {"JPG", "JPEG"}:
                source_format = "JPEG"
                if cropped.mode not in {"RGB", "L"}:
                    cropped = cropped.convert("RGB")
                save_kwargs["quality"] = 95
            elif source_format != "PNG":
                source_format = "PNG"
                original_name = str(Path(original_name).with_suffix(".png"))
            cropped.save(output, format=source_format, **save_kwargs)
            return original_name, output.getvalue(), ""
    except Exception as exc:  # noqa: BLE001 - a crop failure must fall back to the source image.
        return (
            original_name,
            blob,
            f"Word image crop could not be preserved for {original_name} ({type(exc).__name__}); the full image was retained.",
        )


def _preserved_word_drawings(paragraph, document: Document):
    """Return one bytes-compatible asset per top-level DrawingML figure."""
    assets = []
    warnings = []
    for drawing in paragraph._p.xpath(".//wp:inline | .//wp:anchor"):
        drawing_copy = parse_xml(drawing.xml)
        relationships: Dict[str, Tuple[str, bytes]] = {}
        preview_name = "word_figure.png"
        preview_blob = b""
        source_blips = drawing.xpath(".//a:blip")
        copied_blips = drawing_copy.xpath(".//a:blip")
        for blip, copied_blip in zip(source_blips, copied_blips):
            relationship_id = blip.get(qn("r:embed"))
            if not relationship_id and blip.get(qn("r:link")):
                raise ValueError('Word contains a linked image. Embed the image in Word before uploading.')
            if not relationship_id or relationship_id not in document.part.related_parts:
                raise ValueError('Word image data is missing. Embed the original image in Word before uploading.')
            part = document.part.related_parts[relationship_id]
            original_name = Path(str(part.partname)).name
            rendered_name, rendered_blob, crop_warning = _apply_word_image_crop(
                original_name, part.blob, _drawingml_crop_rect(blip)
            )
            relationships[relationship_id] = (rendered_name, rendered_blob)
            copied_fill = copied_blip.getparent()
            copied_crop = copied_fill.find(qn("a:srcRect")) if copied_fill is not None else None
            if copied_crop is not None:
                copied_fill.remove(copied_crop)
            if not preview_blob:
                preview_name, preview_blob = rendered_name, rendered_blob
            if crop_warning:
                warnings.append(crop_warning)
        if preview_blob and relationships:
            assets.append(
                (
                    preview_name,
                    WordDrawingAsset(
                        preview_blob,
                        drawing_copy.xml.encode("utf-8"),
                        relationships,
                    ),
                )
            )
    return assets, warnings


def _style_chain(style):
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


def _font_value(run, paragraph, attribute):
    """Resolve direct formatting before character and paragraph style inheritance."""
    fonts = [run.font]
    fonts.extend(style.font for style in _style_chain(run.style))
    fonts.extend(style.font for style in _style_chain(paragraph.style))
    for font in fonts:
        value = getattr(font, attribute)
        if value is not None:
            return value
    return None


class _WordNumbering:
    """Resolve list labels stored outside paragraph text, including restarts."""

    def __init__(self, document):
        try:
            root = document.part.numbering_part.element
        except (KeyError, NotImplementedError):
            # Numbering is optional in valid Word packages (common after RTF conversion).
            root = OxmlElement('w:numbering')
        self.numbers = {int(item.get(qn('w:numId'))): item for item in root.findall(qn('w:num'))}
        self.abstracts = {int(item.get(qn('w:abstractNumId'))): item for item in root.findall(qn('w:abstractNum'))}
        self.counters = {}

    @staticmethod
    def value(element, name, default=None):
        child = element.find(qn('w:' + name)) if element is not None else None
        return child.get(qn('w:val')) if child is not None else default

    def resolve(self, paragraph):
        num_id = level = None
        properties = [paragraph._p.pPr]
        properties.extend(style.element.pPr for style in _style_chain(paragraph.style))
        for props in properties:
            num = props.find(qn('w:numPr')) if props is not None else None
            if num_id is None:
                num_id = self.value(num, 'numId')
            if level is None:
                level = self.value(num, 'ilvl')
        if num_id is None or int(num_id) == 0:
            return None, None, ''
        num_id, level = int(num_id), int(level or 0)
        number = self.numbers.get(num_id)
        abstract_id = self.value(number, 'abstractNumId')
        abstract = self.abstracts.get(int(abstract_id)) if abstract_id is not None else None
        if abstract is None:
            return num_id, level, ''
        levels = {int(item.get(qn('w:ilvl'))): item for item in abstract.findall(qn('w:lvl'))}
        overrides = {int(item.get(qn('w:ilvl'))): item for item in number.findall(qn('w:lvlOverride'))}
        override = overrides.get(level)
        definition = override.find(qn('w:lvl')) if override is not None else None
        if definition is None:
            definition = levels.get(level)
        if definition is None:
            return num_id, level, ''
        start = int(self.value(override, 'startOverride', self.value(definition, 'start', '1')))
        counters = self.counters.setdefault(num_id, {})
        # A list can begin at a child level; Word implicitly starts its ancestors.
        for ancestor in range(level):
            if ancestor not in counters:
                ancestor_override = overrides.get(ancestor)
                counters[ancestor] = int(self.value(ancestor_override, 'startOverride',
                    self.value(levels.get(ancestor), 'start', '1')))
        value = counters.get(level, start - 1) + 1
        counters[level] = value
        for deeper in list(counters):
            restart = int(self.value(levels.get(deeper), 'lvlRestart', str(deeper)))
            if deeper > level and restart and level < restart:
                del counters[deeper]
        fmt = self.value(definition, 'numFmt', '')
        template = self.value(definition, 'lvlText', '')
        # Compound outline numbers cannot safely be treated as MCQ option labels.
        simple_label = len(re.findall(r'%[1-9]', template)) == 1
        if simple_label and fmt == 'decimal' and level == 0:
            return num_id, level, f'{value}. '
        if simple_label and fmt in {'lowerLetter', 'upperLetter'}:
            if not 1 <= value <= 8:
                raise ValueError('Word answer list continues beyond H. Restart the answer list at A for each question.')
            return num_id, level, chr(64 + value) + '. '
        return num_id, level, ''


def _only_horizontal_rules(paragraph):
    """Recognise plain separator lines, never arrows, diagrams or labelled shapes."""
    drawings = paragraph._p.xpath('.//w:drawing | .//w:pict')
    if not drawings or paragraph.text.strip():
        return False
    for drawing in drawings:
        if any(node.tag in {qn('w:t'), qn('a:blip')} or node.tag.endswith('}imagedata') for node in drawing.iter()):
            return False
        if drawing.tag == qn('w:drawing'):
            geometry = list(drawing.iter(qn('a:prstGeom')))
            extents = list(drawing.iter(qn('wp:extent')))
            if (len(geometry) != 1 or geometry[0].get('prst') != 'line'
                    or len(extents) != 1 or extents[0].get('cy') != '0'
                    or any(node.tag in {qn('a:headEnd'), qn('a:tailEnd'), qn('a:custGeom')} for node in drawing.iter())):
                return False
        else:
            lines = drawing.xpath('./*[local-name()="line"]')
            if len(lines) != 1 or len(drawing) != 1:
                return False
            start, end = lines[0].get('from', '').split(','), lines[0].get('to', '').split(',')
            if len(start) != 2 or len(end) != 2 or start[1].strip() != end[1].strip():
                return False
            if any('arrow' in name.lower() for node in lines[0].iter() for name in node.attrib):
                return False
    return True


def _line_records(paragraph, paragraph_index: int, document: Document, numbering=None) -> List[LineRecord]:
    buffers = [
        {
            "text": "",
            "chars": 0,
            "bold": 0,
            "highlighted": False,
            "underlined": False,
            "coloured": False,
        }
    ]

    # Resolve tab boundaries over the whole paragraph, including split runs.
    elements = paragraph._p.xpath('./w:r | ./w:hyperlink/w:r')
    full_text = ''.join(Run(element, paragraph).text for element in elements)
    tab_breaks = {match.start() for match in re.finditer(r'\t+(?=[(\[]?[A-Ha-h]\s*[.)\]:-]\s)', full_text)}
    offset = 0
    # paragraph.runs omits runs inside hyperlinks, including answer text.
    for element in elements:
        run = Run(element, paragraph)
        run_text = ''.join('\n' if offset + index in tab_breaks else char for index, char in enumerate(run.text))
        offset += len(run.text)
        pieces = re.split(r"(\n)", run_text)
        for piece in pieces:
            if piece == "\n":
                buffers.append(
                    {
                        "text": "",
                        "chars": 0,
                        "bold": 0,
                        "highlighted": False,
                        "underlined": False,
                        "coloured": False,
                    }
                )
                continue
            if not piece:
                continue
            length = len(piece.strip())
            current = buffers[-1]
            current["text"] += piece
            current["chars"] += length
            if _font_value(run, paragraph, 'bold') is True:
                current["bold"] += length
            if _font_value(run, paragraph, 'highlight_color') is not None:
                current["highlighted"] = True
            if _font_value(run, paragraph, 'underline'):
                current["underlined"] = True
            if run.font.color is not None and run.font.color.rgb is not None:
                rgb = str(run.font.color.rgb).upper()
                if rgb not in {"000000", "AUTO"}:
                    current["coloured"] = True

    num_id, num_level, list_label = (numbering or _WordNumbering(document)).resolve(paragraph)

    image_items, image_warnings = _preserved_word_drawings(paragraph, document)
    if not image_items:
        seen_images = set()
        for element in paragraph._p.xpath(".//*[local-name()='imagedata']"):
            rid = element.get(qn("r:id"))
            if not rid or rid not in document.part.related_parts:
                raise ValueError('Word contains a missing or linked image. Embed the original image in Word before uploading.')
            part = document.part.related_parts[rid]
            part_name = str(part.partname)
            crop_rect = _vml_crop_rect(element)
            image_key = (part_name, crop_rect)
            if image_key in seen_images:
                continue
            seen_images.add(image_key)
            image_name, image_blob, crop_warning = _apply_word_image_crop(
                Path(part_name).name, part.blob, crop_rect
            )
            image_items.append((image_name, image_blob))
            if crop_warning:
                image_warnings.append(crop_warning)

    drawing_count = len(paragraph._p.xpath(".//w:drawing")) + len(paragraph._p.xpath(".//w:pict"))
    if drawing_count and not image_items and _only_horizontal_rules(paragraph):
        drawing_count = 0
    if drawing_count and not image_items:
        raise ValueError('Word drawing has no supported embedded picture. Convert the drawing to an embedded image before uploading.')
    if image_items and sum(bool(buffer['text'].strip()) for buffer in buffers) > 1:
        raise ValueError('A Word picture shares a paragraph with multiple text lines. Put the picture in a separate paragraph after its question stem before uploading.')
    style = paragraph.style.name if paragraph.style is not None else ""
    records = []
    for buffer_index, buffer in enumerate(buffers):
        text = normalise_text(buffer["text"])
        if buffer_index == 0 and text and list_label and not (
            OPTION_LABEL_RE.match(text) or QUESTION_NUMBER_RE.match(text) or NAMED_QUESTION_NUMBER_RE.match(text)
        ):
            text = list_label + text
        if not text and not image_items:
            continue
        chars = max(1, int(buffer["chars"]))
        records.append(
            LineRecord(
                paragraph_index=paragraph_index,
                text=text,
                bold_ratio=float(buffer["bold"]) / chars,
                highlighted=bool(buffer["highlighted"]),
                underlined=bool(buffer["underlined"]),
                coloured=bool(buffer["coloured"]),
                num_id=num_id,
                num_level=num_level,
                style=style,
                images=image_items if buffer_index == 0 else [],
                drawing_count=drawing_count if buffer_index == 0 else 0,
                warnings=image_warnings if buffer_index == 0 else [],
            )
        )
    return records


OPTION_LABEL_RE = re.compile(r"^\s*[(\[]?([A-Ha-h])\s*[\.\)\]\:\-]\s*(.+)$", re.DOTALL)
# A hyphen is deliberately excluded: numeric answer ranges such as "15-20 seconds"
# and "8.4 %" must remain options rather than becoming questions 15 and 8.
QUESTION_NUMBER_RE = re.compile(r"^\s*(\d{1,4})\s*[\.\)\:]\s+(.+)$", re.DOTALL)
NAMED_QUESTION_NUMBER_RE = re.compile(
    r"^\s*(?:question|q)\s*(\d{1,4})(?:\s*[\.\)\]\:\-\u2013\u2014]\s*|\s+)(.+)$",
    flags=re.IGNORECASE | re.DOTALL,
)
WORD_INLINE_ANSWER_RE = re.compile(
    r"^\s*(?:answer|correct(?:\s+answer)?)\s*[:=\-\u2013\u2014]\s*([A-Fa-f1-6])\s*[\.\)]?\s*$",
    flags=re.IGNORECASE,
)
WORD_ANSWER_KEY_HEADING_RE = re.compile(
    r"^\s*(?:correct\s+answers?|answer\s*key|answers?)\s*:?\s*(.*)$",
    flags=re.IGNORECASE | re.DOTALL,
)
WORD_NUMBERED_KEY_PAIR_RE = re.compile(
    r"(?:^|[|;,])\s*(?:question\s*|q\s*)?(\d{1,4})\s*[\.\)\]\:\-=]\s*([A-Fa-f1-6])\b",
    flags=re.IGNORECASE,
)


def _is_question_line(line: LineRecord) -> bool:
    text = line.text.strip()
    if not text or OPTION_LABEL_RE.match(text) or re.fullmatch(r"[A-Fa-f]", text):
        return False
    if QUESTION_NUMBER_RE.match(text) or NAMED_QUESTION_NUMBER_RE.match(text):
        return True
    if text.endswith("?"):
        return True
    if text.endswith(":") and len(text) >= 15:
        return True
    question_words = r"^(which|what|why|where|when|how|who|choose|select|identify|the term\b|this is\b)"
    if line.bold_ratio >= 0.65 and re.search(question_words, text, flags=re.IGNORECASE):
        return True
    # A bold correct option in a multilevel Word list is still an option. Word
    # stores the visible A/B/C/D marker in numbering.xml rather than paragraph
    # text, so the list level is required to distinguish it from the question.
    if (
        line.bold_ratio >= 0.85
        and line.num_id is not None
        and line.num_level in {None, 0}
        and len(text) >= 20
    ):
        return True
    return False


def _answer_token_to_letter(token: str) -> str:
    compact = token.strip().upper()
    return LETTERS[int(compact) - 1] if compact.isdigit() else compact


def _set_answer(answers, number, answer):
    previous = answers.get(number)
    if previous and previous != answer:
        raise ValueError(f"Conflicting correct answers for question {number}: {previous} and {answer}. Resolve the answer key before importing.")
    answers[number] = answer


def _numbered_answer_pairs(text: str) -> Dict[int, str]:
    # Match the entire key line, never just the prefix of a question stem.
    remainder = WORD_NUMBERED_KEY_PAIR_RE.sub("", text)
    if remainder.strip(" \t\r\n|;,"):
        return {}
    answers = {}
    for number, token in WORD_NUMBERED_KEY_PAIR_RE.findall(text):
        _set_answer(answers, int(number), _answer_token_to_letter(token))
    return answers


def _word_answer_key(lines: List[LineRecord]):
    """Return answers, consumed LINE indices, and whether a bare key is positional."""
    nonempty = [(i, line) for i, line in enumerate(lines) if line.text]
    for heading_index, (line_index, line) in enumerate(nonempty):
        if WORD_INLINE_ANSWER_RE.fullmatch(line.text):
            continue
        heading = WORD_ANSWER_KEY_HEADING_RE.fullmatch(line.text)
        if not heading:
            continue
        answer_key = _numbered_answer_pairs(heading.group(1))
        numbered = bool(answer_key)
        key_lines = {line_index}
        next_bare_answer = 1
        for candidate_index, candidate in nonempty[heading_index + 1:]:
            pairs = _numbered_answer_pairs(candidate.text)
            if pairs:
                numbered = True
                for number, answer in pairs.items():
                    _set_answer(answer_key, number, answer)
                key_lines.add(candidate_index)
                next_bare_answer = max(next_bare_answer, max(pairs) + 1)
                continue
            bare = re.fullmatch(r"\s*([A-Fa-f1-6])\s*", candidate.text)
            if bare:
                while next_bare_answer in answer_key:
                    next_bare_answer += 1
                _set_answer(answer_key, next_bare_answer, _answer_token_to_letter(bare.group(1)))
                next_bare_answer += 1
                key_lines.add(candidate_index)
                continue
            if re.fullmatch(r'\s*[A-Fa-f](?:[\s,;|]+[A-Fa-f])+\s*', candidate.text):
                for token in re.findall(r'[A-Fa-f]', candidate.text):
                    while next_bare_answer in answer_key:
                        next_bare_answer += 1
                    _set_answer(answer_key, next_bare_answer, token.upper())
                    next_bare_answer += 1
                key_lines.add(candidate_index)
                continue
            break
        if answer_key:
            return answer_key, key_lines, not numbered

    trailing = []
    for line_index, line in reversed(nonempty):
        match = re.fullmatch(r"\s*([A-Fa-f])\s*", line.text)
        if not match:
            break
        trailing.append((line_index, match.group(1).upper()))
    trailing.reverse()
    if len(trailing) < 2:
        return {}, set(), False
    return ({index + 1: answer for index, (_, answer) in enumerate(trailing)},
            {line_index for line_index, _ in trailing}, True)


def _resolved_answer(inline, separate, number):
    answers = {}
    if inline:
        _set_answer(answers, number, inline)
    if separate:
        _set_answer(answers, number, separate)
    return answers.get(number, "")


def _deduplicate_images(images: List[Tuple[str, bytes]]) -> List[Tuple[str, bytes]]:
    output = []
    seen = set()
    for original_name, blob in images:
        digest = word_drawing_fingerprint(blob).hex()
        if digest in seen:
            continue
        seen.add(digest)
        output.append((original_name, blob))
    return output


def _structured_word_table_imports(document: Document, filename: str):
    frames = []
    issues = []
    structured_table_ids = set()
    for table_index, table in enumerate(document.tables, start=1):
        rows = [[normalise_text(cell.text) for cell in row.cells] for row in table.rows]
        if not rows:
            continue
        scored_headers = [(_excel_header_score(row), index) for index, row in enumerate(rows[:3])]
        score, header_index = max(scored_headers, default=(0, 0))
        if score < 12:
            continue

        if table._tbl.xpath('.//w:vMerge | .//w:gridSpan'):
            raise ValueError('A Word question table contains merged cells. Unmerge them so each question and option has its own cell.')

        structured_table_ids.add(id(table._tbl))
        headers = rows[header_index]
        width = len(headers)
        data_rows = [(row + [""] * width)[:width] for row in rows[header_index + 1 :]]
        table_frame = pd.DataFrame(data_rows, columns=headers)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            table_frame.to_excel(writer, index=False, sheet_name="Word Table")
        table_source = f"{Path(filename).stem}_Table_{table_index}.xlsx"
        imported, table_issues, _ = read_excel_questions(output.getvalue(), table_source)
        if not imported.empty:
            imported["Source_File"] = filename
            imported["Import_Method"] = f"Structured Word table {table_index}"
            frames.append(imported)
        for item in table_issues:
            item["Source file"] = filename
            item["Issue"] = f"Word table {table_index}: {item['Issue']}"
        issues.extend(table_issues)

        if table._tbl.xpath(".//a:blip") or table._tbl.xpath(".//*[local-name()='imagedata']"):
            issues.append(
                issue(
                    filename,
                    f"Table {table_index}",
                    "",
                    "ERROR",
                    "Structured Word table contains images that could not be assigned safely; move them into question paragraphs or upload them separately.",
                )
            )
    return frames, issues, structured_table_ids


def _iter_body_paragraphs(document: Document, structured_table_ids: set[int]):
    """Yield paragraphs in order, including nested non-structured tables."""
    def walk(container, parent):
        for child in container.iterchildren():
            if child.tag == qn('w:p'):
                yield Paragraph(child, parent)
            elif child.tag == qn('w:tbl') and id(child) not in structured_table_ids:
                table = Table(child, parent)
                seen_cells = set()
                for row in table.rows:
                    for cell in row.cells:
                        if cell._tc in seen_cells:
                            continue
                        seen_cells.add(cell._tc)
                        yield from walk(cell._tc, cell)
    yield from walk(document.element.body, document)


def read_docx_questions(data: bytes, filename: str) -> Tuple[pd.DataFrame, List[Dict], Dict[str, bytes]]:
    issues: List[Dict] = []
    try:
        validate_office_container(data, filename)
        document = Document(io.BytesIO(data))
    except Exception as exc:  # noqa: BLE001 - untrusted DOCX parsers raise varied exceptions.
        return pd.DataFrame(columns=NORMALIZED_COLUMNS), [
            issue(filename, "", "", "ERROR", f"Word import failed: {exc}")
        ], {}

    for expression, message in [
        ('.//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo',
         'Word document contains tracked changes. Accept or reject them in Word and upload the final version.'),
        ('.//m:oMath | .//m:oMathPara',
         'Word document contains equations that cannot be preserved as plain text. Convert them to images in Word before uploading.'),
        ('.//w:sdt | .//w:altChunk',
         'Word document contains content controls or embedded document text. Convert these to ordinary paragraphs before uploading.'),
        ('.//w:object',
         'Word document contains embedded objects. Convert them to ordinary text or images before uploading.'),
    ]:
        if document.element.body.xpath(expression):
            raise ValueError(message)

    table_frames, table_issues, structured_table_ids = _structured_word_table_imports(document, filename)
    issues.extend(table_issues)

    lines: List[LineRecord] = []
    numbering = _WordNumbering(document)
    for paragraph_index, paragraph in enumerate(
        _iter_body_paragraphs(document, structured_table_ids)
    ):
        if _only_horizontal_rules(paragraph):
            issues.append(issue(filename, '', '', 'INFO',
                'A standalone horizontal separator line was omitted from the question content.'))
        lines.extend(_line_records(paragraph, paragraph_index, document, numbering))

    answer_key, key_lines, positional_key = _word_answer_key(lines)
    questions: List[ParsedQuestion] = []
    current: Optional[ParsedQuestion] = None

    def finish_current():
        nonlocal current
        if current is None:
            return
        current.text = normalise_text(current.text)
        if current.text:
            questions.append(current)
        current = None

    for line_index, line in enumerate(lines):
        if line_index in key_lines:
            continue
        text = line.text.strip()
        if re.fullmatch(r'[(\[]?[A-Ha-h]\s*[.)\]:-]', text):
            raise ValueError(f'An option label has no answer text: {text}. Complete the option before importing.')
        labelled_option = OPTION_LABEL_RE.match(text)
        if line.images and labelled_option:
            raise ValueError('A picture belongs to an answer choice. This layout cannot yet preserve choice-specific images; combine the labelled choices into one question figure before uploading.')
        explicit_question = NAMED_QUESTION_NUMBER_RE.match(text) or QUESTION_NUMBER_RE.match(text)
        inline_answer = WORD_INLINE_ANSWER_RE.fullmatch(text)
        is_question = _is_question_line(line)

        if current is not None and inline_answer:
            token = inline_answer.group(1).upper()
            value = LETTERS[int(token) - 1] if token.isdigit() else token
            current.explicit_answer = _resolved_answer(current.explicit_answer, value, current.number)
            current.answer_evidence = "Explicit inline correct-answer line"
            continue

        if (current is not None and not current.options and not explicit_question
                and not labelled_option and line.num_id is None and text
                and not current.text.rstrip().endswith(('?', ':'))):
            current.text = normalise_text(current.text + ' ' + text)
            current.images.extend(line.images)
            current.drawing_count += line.drawing_count
            current.warnings.extend(line.warnings)
            continue

        if is_question:
            finish_current()
            number = explicit_question.group(1) if explicit_question else str(len(questions) + 1)
            question_text = explicit_question.group(2) if explicit_question else text
            current = ParsedQuestion(
                number=number,
                text=question_text,
                paragraph_index=line.paragraph_index,
                num_id=line.num_id,
                images=list(line.images),
                drawing_count=line.drawing_count,
                warnings=list(line.warnings),
            )
            continue

        if current is None:
            continue

        if line.images:
            current.images.extend(line.images)
            current.drawing_count += line.drawing_count
        if line.warnings:
            current.warnings.extend(line.warnings)
        if not text:
            continue

        if labelled_option:
            current.options.append(
                ParsedOption(
                    text=normalise_text(labelled_option.group(2)),
                    source_label=labelled_option.group(1).upper(),
                    highlighted=line.highlighted,
                    bold_ratio=line.bold_ratio,
                    underlined=line.underlined,
                    coloured=line.coloured,
                )
            )
            continue

        if line.num_id is not None:
            if current.option_num_id is None:
                current.option_num_id = line.num_id
            if line.num_id == current.option_num_id:
                current.options.append(
                    ParsedOption(
                        text=text,
                        highlighted=line.highlighted,
                        bold_ratio=line.bold_ratio,
                        underlined=line.underlined,
                        coloured=line.coloured,
                    )
                )
                continue

        if current.options and any(option.source_label for option in current.options):
            current.options[-1].text = normalise_text(current.options[-1].text + ' ' + text)
        else:
            current.options.append(
                ParsedOption(
                    text=text,
                    highlighted=line.highlighted,
                    bold_ratio=line.bold_ratio,
                    underlined=line.underlined,
                    coloured=line.coloured,
                )
            )

    finish_current()

    if not questions and not table_frames:
        return pd.DataFrame(columns=NORMALIZED_COLUMNS), [
            issue(filename, "", "", "ERROR", "No MCQ question blocks could be detected in the Word document.")
        ], {}

    if answer_key and not positional_key:
        unknown_numbers = set(answer_key) - {int(question.number) for question in questions}
        if unknown_numbers:
            labels = ', '.join(map(str, sorted(unknown_numbers)))
            raise ValueError(f'Answer key refers to question numbers not found in the document: {labels}. Check question numbering and the answer key in Word before uploading.')

    if answer_key and len(answer_key) != len(questions):
        issues.append(
            issue(
                filename,
                "",
                "",
                "WARNING",
                f"A trailing answer key with {len(answer_key)} entries was detected for {len(questions)} questions.",
            )
        )

    image_map: Dict[str, bytes] = {}
    records = []
    for question_index, parsed in enumerate(questions, start=1):
        if len(parsed.options) > MAX_OPTIONS:
            raise ValueError(f"Question {parsed.number or question_index} has {len(parsed.options)} options; the maximum is six. No options have been silently removed.")
        record = {
            "No": parsed.number or str(question_index),
            "Question": parsed.text,
            "Correct": "",
            "Source_File": filename,
            "Import_Method": "Structured Word import",
            "Answer_Evidence": "Correct answer not provided",
            "Confidence": "Review required",
            "Needs_Review": True,
            "Import_Warnings": "",
        }
        for option_index in range(MAX_OPTIONS):
            record[OPTION_COLUMNS[option_index]] = (
                parsed.options[option_index].text if option_index < len(parsed.options) else ""
            )

        highlighted = [idx for idx, option in enumerate(parsed.options) if option.highlighted]
        bold_candidates = [idx for idx, option in enumerate(parsed.options) if option.bold_ratio >= 0.65]
        underlined = [idx for idx, option in enumerate(parsed.options) if option.underlined]
        coloured = [idx for idx, option in enumerate(parsed.options) if option.coloured]

        key_number = question_index if positional_key else int(parsed.number)
        explicit = _resolved_answer(parsed.explicit_answer, answer_key.get(key_number, ''), parsed.number)
        if explicit:
            if len(highlighted) == 1 and LETTERS[highlighted[0]] != explicit:
                raise ValueError(f'Conflicting highlighted and explicit correct answers for question {parsed.number}. Resolve the answer markings before importing.')
            record["Correct"] = explicit
            record["Answer_Evidence"] = parsed.answer_evidence if parsed.explicit_answer else 'Separate trailing answer key'
            record["Confidence"] = "High"
            record["Needs_Review"] = False
        elif len(highlighted) == 1:
            record["Correct"] = LETTERS[highlighted[0]]
            record["Answer_Evidence"] = "Highlighted answer option"
            record["Confidence"] = "High"
            record["Needs_Review"] = False
        elif len(bold_candidates) == 1 and parsed.options[bold_candidates[0]].bold_ratio >= 0.95:
            record["Correct"] = LETTERS[bold_candidates[0]]
            record["Answer_Evidence"] = "Entire answer option uniquely bold"
            record["Confidence"] = "High"
            record["Needs_Review"] = False
        elif len(bold_candidates) == 1:
            record["Correct"] = LETTERS[bold_candidates[0]]
            record["Answer_Evidence"] = "Uniquely bold answer option"
            record["Confidence"] = "Review recommended"
            record["Needs_Review"] = True
        elif len(underlined) == 1:
            record["Correct"] = LETTERS[underlined[0]]
            record["Answer_Evidence"] = "Uniquely underlined answer option"
            record["Confidence"] = "Review recommended"
            record["Needs_Review"] = True
        elif len(coloured) == 1:
            record["Correct"] = LETTERS[coloured[0]]
            record["Answer_Evidence"] = "Uniquely coloured answer option"
            record["Confidence"] = "Review recommended"
            record["Needs_Review"] = True
        elif len(highlighted) > 1 or len(bold_candidates) > 1:
            record["Answer_Evidence"] = "Multiple formatted answer candidates"

        unique_images = _deduplicate_images(parsed.images)
        image_names = []
        for image_index, (original_name, blob) in enumerate(unique_images, start=1):
            suffix = Path(original_name).suffix.lower()
            if suffix not in {".png", ".jpg", ".jpeg"}:
                suffix = ".png"
            generated = f"{safe_prefix(filename)}_Q{question_index}_IMG{image_index}{suffix}"
            image_map[generated.lower()] = blob
            image_names.append(generated)
        record["Image_Files"] = "|".join(image_names)
        record["Image"] = image_names[0] if image_names else ""

        warnings = list(dict.fromkeys(parsed.warnings))
        labelled_positions = [
            (index, option.source_label)
            for index, option in enumerate(parsed.options)
            if option.source_label
        ]
        label_mismatch = labelled_positions and any(
            source_label != LETTERS[index]
            for index, source_label in labelled_positions
            if index < len(LETTERS)
        )
        if label_mismatch:
            source_labels = [source_label for _, source_label in labelled_positions]
            warnings.append(
                f"Source labels {', '.join(source_labels)} do not match their option positions. Use consecutive A–F labels before importing."
            )
        if len(unique_images) > 1:
            warnings.append("Multiple images were associated with this question; review their order and annotations.")
        if parsed.drawing_count > 1 and len(unique_images) > 1:
            warnings.append("Complex Word drawing composition detected; overlays or labels may require visual correction.")
        record["Import_Warnings"] = " ".join(warnings)
        if warnings:
            issue_row = parsed.number or question_index
            issue_internal_id = internal_id(filename, issue_row, question_index)
            issues.append(
                    issue(filename, issue_row, issue_internal_id, "ERROR" if label_mismatch else "WARNING", record["Import_Warnings"])
            )

        record = _validate_and_finalize_row(record, filename, question_index, issues)
        records.append(record)

    body_frame = pd.DataFrame(records, columns=NORMALIZED_COLUMNS)
    frames = [*table_frames]
    if not body_frame.empty:
        frames.append(body_frame)
    combined = pd.concat(frames, ignore_index=True) if frames else body_frame
    return combined[NORMALIZED_COLUMNS], issues, image_map


TEXT_QUESTION_RE = re.compile(
    r"^\s*(?:(?:question|q)\s*)?(\d{1,4})\s*[\.\)\:]\s+(.+)$",
    flags=re.IGNORECASE,
)
TEXT_NUMERIC_OPTION_RE = re.compile(r"^\s*([1-6])\s*[\.\)\:]\s+(.+)$", flags=re.DOTALL)
TEXT_ANSWER_RE = re.compile(
    r"^\s*(?:answer|correct(?:\s+answer)?)\s*[:=\-]\s*([A-Fa-f1-6])\s*$",
    flags=re.IGNORECASE,
)
TEXT_IMAGE_RE = re.compile(r"^\s*(?:image|figure|picture)\s*[:=]\s*(.+)$", flags=re.IGNORECASE)


def read_text_questions(data: bytes, filename: str) -> Tuple[pd.DataFrame, List[Dict], Dict[str, bytes]]:
    issues: List[Dict] = []
    from source_formats import decode_text
    text = decode_text(data)
    lines = [normalise_text(line) for line in text.splitlines()]
    line_records = [LineRecord(i, line, 0.0, False, False, False, None, None, '')
                    for i, line in enumerate(lines)]
    answer_key, key_lines, positional_key = _word_answer_key(line_records)
    questions = []
    current = None

    def finish_current():
        nonlocal current
        if current and normalise_text(current.get('Question')):
            questions.append(current)
        current = None

    for line_index, line in enumerate(lines):
        if not line or line_index in key_lines:
            continue
        if re.fullmatch(r'[A-Ha-h]\s*[.)\]:-]', line):
            raise ValueError(f'An option label has no answer text: {line}. Complete the option before importing.')
        if current is not None:
            answer_match = WORD_INLINE_ANSWER_RE.fullmatch(line)
            if answer_match:
                value = _answer_token_to_letter(answer_match.group(1))
                current['Correct'] = _resolved_answer(current['Correct'], value, current['No'])
                continue
            image_match = TEXT_IMAGE_RE.match(line)
            if image_match:
                current['Image_Files'] = '|'.join(split_image_files(image_match.group(1)))
                continue
            labelled = OPTION_LABEL_RE.match(line)
            numeric = TEXT_NUMERIC_OPTION_RE.match(line)
            if labelled:
                expected = LETTERS[len(current['Options'])] if len(current['Options']) < MAX_OPTIONS else ''
                if labelled.group(1).upper() != expected:
                    raise ValueError(f"Question {current['No']}: option labels must be consecutive A–F; labels are missing, reordered, duplicated, or beyond the six-option limit.")
                current['Options'].append(normalise_text(labelled.group(2)))
                current['Labelled'] = True
                continue
            if (numeric and not current['Correct'] and not current['Labelled']
                    and int(numeric.group(1)) == len(current['Options']) + 1):
                current['Options'].append(normalise_text(numeric.group(2)))
                current['Numeric'] = True
                continue

        question_match = TEXT_QUESTION_RE.match(line)
        if (current is not None and not current['Options'] and not question_match
                and not current['Question'].rstrip().endswith(('?', ':'))):
            current['Question'] = normalise_text(current['Question'] + ' ' + line)
            continue
        if question_match or (line.endswith('?') and (current is None or len(current['Options']) >= 2)):
            finish_current()
            current = {'No': question_match.group(1) if question_match else str(len(questions) + 1),
                       'Question': question_match.group(2) if question_match else line,
                       'Options': [], 'Correct': '', 'Image_Files': '', 'Labelled': False, 'Numeric': False}
            continue
        if current is not None:
            if current['Correct']:
                raise ValueError(f"Unexpected text after the answer for question {current['No']}. Number the next question or remove the extra text.")
            if current['Options'] and (current['Labelled'] or current['Numeric']):
                current['Options'][-1] = normalise_text(current['Options'][-1] + ' ' + line)
            else:
                current['Options'].append(line)
    finish_current()
    if not questions:
        return pd.DataFrame(columns=NORMALIZED_COLUMNS), [
            issue(filename, '', '', 'ERROR', 'No MCQ question blocks could be detected in the text file.')
        ], {}
    if answer_key and len(answer_key) != len(questions):
        issues.append(issue(filename, '', '', 'WARNING',
                            f'An answer key with {len(answer_key)} entries was detected for {len(questions)} questions.'))

    records = []
    for question_index, parsed in enumerate(questions, start=1):
        if len(parsed["Options"]) > MAX_OPTIONS:
            raise ValueError(f"Question {parsed.get('No', question_index)} has more than six options. Reduce the options before importing.")
        key_number = question_index if positional_key else int(parsed.get("No", question_index))
        correct = _resolved_answer(parsed.get("Correct"), answer_key.get(key_number, ""), parsed["No"])
        record = {
            "No": parsed.get("No", question_index),
            "Question": parsed.get("Question", ""),
            "Correct": correct,
            "Image_Files": parsed.get("Image_Files", ""),
            "Source_File": filename,
            "Import_Method": "Structured plain-text import",
            "Answer_Evidence": "Inline or separate text answer key" if correct else "Correct answer not provided",
            "Confidence": "High" if correct else "Review required",
            "Needs_Review": not bool(correct),
            "Import_Warnings": "",
        }
        for option_index in range(MAX_OPTIONS):
            record[OPTION_COLUMNS[option_index]] = (
                parsed["Options"][option_index] if option_index < len(parsed["Options"]) else ""
            )
        record = _validate_and_finalize_row(record, filename, question_index, issues)
        records.append(record)

    return pd.DataFrame(records, columns=NORMALIZED_COLUMNS), issues, {}


def import_question_source(uploaded_file) -> Tuple[pd.DataFrame, List[Dict], Dict[str, bytes]]:
    filename = str(getattr(uploaded_file, "name", "uploaded_file"))
    try:
        uploaded_file.seek(0)
        data = uploaded_file.read(MAX_SOURCE_BYTES + 1)
        if not data:
            raise ValueError("The uploaded file is empty.")
        if len(data) > MAX_SOURCE_BYTES:
            raise ValueError("File exceeds the 200 MB source limit.")
        suffix = Path(filename).suffix.lower()
        if suffix in {".xlsx", ".xls"}:
            return read_excel_questions(data, filename)
        if suffix == ".docx":
            return read_docx_questions(data, filename)
        if suffix == ".txt":
            from source_formats import decode_text
            return read_text_questions(decode_text(data).encode("utf-8"), filename)
        from source_formats import read_additional_source
        return read_additional_source(data, filename)
    except Exception as exc:  # Convert malformed/unsupported inputs into visible file-level errors.
        return pd.DataFrame(columns=NORMALIZED_COLUMNS), [
            issue(filename, "", "", "ERROR", f"Import failed: {exc}")
        ], {}


def import_all_sources(files) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, bytes]]:
    frames = []
    all_issues: List[Dict] = []
    images: Dict[str, bytes] = {}
    for uploaded_file in files:
        frame, issues, source_images = import_question_source(uploaded_file)
        all_issues.extend(issues)
        image_renames = {}
        for name, blob in source_images.items():
            candidate = name
            counter = 2
            while candidate in images and images[candidate] != blob:
                stem, suffix = str(Path(name).with_suffix("")), Path(name).suffix
                candidate = f"{stem}_{counter}{suffix}"
                counter += 1
            images[candidate] = blob
            if candidate != name:
                image_renames[name.lower()] = candidate

        if not frame.empty and image_renames:
            frame = frame.copy()
            frame["Image_Files"] = frame["Image_Files"].apply(
                lambda value, rename_map=image_renames: "|".join(
                    rename_map.get(image_name.lower(), image_name)
                    for image_name in split_image_files(value)
                )
            )
            frame["Image"] = frame["Image_Files"].apply(
                lambda value: split_image_files(value)[0] if split_image_files(value) else ""
            )
        if not frame.empty:
            frames.append(frame)

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=NORMALIZED_COLUMNS)
    if not combined.empty:
        ids = combined["Internal_ID"].astype(str).str.strip()
        duplicates = ids[ids.duplicated(keep=False) & ids.ne("")]
        for duplicate in sorted(set(duplicates)):
            all_issues.append(
                issue(
                    "; ".join(combined.loc[ids == duplicate, "Source_File"].astype(str).tolist()),
                    "",
                    duplicate,
                    "WARNING",
                    "Duplicate internal ID detected.",
                )
            )
    report = pd.DataFrame(all_issues)
    if report.empty:
        report = pd.DataFrame(columns=["Source file", "Row", "Internal_ID", "Severity", "Issue"])
    return combined, report, images


def validate_reviewed_questions(frame: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    issues: List[Dict] = []
    records = []
    for index, row in frame.iterrows():
        record = row.to_dict()
        filename = normalise_text(record.get("Source_File")) or "Reviewed questions"
        record["Needs_Review"] = False
        record = _validate_and_finalize_row(
            record,
            filename,
            index + 1,
            issues,
            strict_correct=True,
        )
        allowed = LETTERS[: int(record.get("Option_Count", 0) or 0)]
        if record["Correct"] in allowed:
            record["Needs_Review"] = False
            record["Confidence"] = "Lecturer confirmed"
            if normalise_text(record.get("Answer_Evidence")) in {
                "Correct answer not provided",
                "Unrecognised correct-answer value",
            }:
                record["Answer_Evidence"] = "Lecturer review"
        records.append(record)
    report = pd.DataFrame(issues)
    if report.empty:
        report = pd.DataFrame(columns=["Source file", "Row", "Internal_ID", "Severity", "Issue"])
    return pd.DataFrame(records, columns=NORMALIZED_COLUMNS), report


def dataframe_for_editor(frame: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "No",
        "Question",
        *OPTION_COLUMNS,
        "Correct",
        "Image_Files",
        "Needs_Review",
        "Import_Method",
        "Answer_Evidence",
        "Confidence",
        "Import_Warnings",
        "Source_File",
        "Internal_ID",
    ]
    output = frame.copy()
    for column in columns:
        if column not in output.columns:
            output[column] = ""
    return output[columns]
