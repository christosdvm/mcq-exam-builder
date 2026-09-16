"""Additional MCQ source formats. No network fetching or macro execution."""
from __future__ import annotations

import csv
import io
import json
import os
import re
import shutil
import signal
import subprocess
import tempfile
import threading
import zipfile
from html.parser import HTMLParser
from pathlib import Path

import pandas as pd
from lxml import etree
import question_importer as qi

WORD_VARIANTS = {'.docm', '.dotx', '.dotm'}
CONVERTED_DOCUMENTS = {'.doc', '.dot', '.odt', '.rtf', '.ott', '.fodt'}
EXCEL_VARIANTS = {'.xlsm', '.xltx', '.xltm'}
SUPPORTED_EXTENSIONS = [
    'docx', 'doc', 'dot', 'docm', 'dotx', 'dotm', 'odt', 'ott', 'fodt', 'rtf',
    'xlsx', 'xls', 'xlsm', 'xltx', 'xltm', 'ods', 'csv', 'tsv',
    'txt', 'md', 'markdown', 'html', 'htm', 'pdf', 'json', 'jsonl', 'xml',
]
CONVERSION_LOCK = threading.BoundedSemaphore(1)
MAX_CONVERT_BYTES = 20 * 1024 * 1024
MAX_TEXT_CHARS = 5_000_000
MAX_TABLE_ROWS = 20_000


def decode_text(data):
    """Honor BOMs; never silently replace undecodable question characters."""
    if data.startswith((b'\xff\xfe', b'\xfe\xff')):
        text = data.decode('utf-16')
    else:
        try:
            text = data.decode('utf-8-sig')
        except UnicodeDecodeError as exc:
            raise ValueError('Save this text file using UTF-8 or UTF-16 encoding and upload it again.') from exc
    if '\x00' in text:
        raise ValueError('This file contains binary data, not readable question text.')
    if len(text) > MAX_TEXT_CHARS:
        raise ValueError('Text source exceeds the 5 million character limit.')
    return text


def safe_xml(data):
    if len(data) > MAX_CONVERT_BYTES:
        raise ValueError('XML content exceeds the 20 MB limit.')
    parser = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False)
    root = etree.fromstring(data, parser=parser)
    if root.getroottree().docinfo.doctype:
        raise ValueError('XML files containing a DTD or entity declarations are not supported.')
    return root


def require_review(result, filename, message):
    frame, findings, images = result
    if not frame.empty:
        frame = frame.copy()
        frame['Needs_Review'] = True
        frame['Confidence'] = 'Review required'
        frame['Import_Warnings'] = frame['Import_Warnings'].fillna('').map(
            lambda previous: ' '.join(part for part in [previous, message] if part))
        # File-level warning survives revalidation of editable rows.
        findings = [*findings, qi.issue(filename, '', '', 'WARNING', message)]
    return frame, findings, images


def read_extracted_text(text, filename):
    """Join wrapped stems/options while refusing missing or reordered option labels."""
    lines = []
    option_count = 0
    have_question = False
    in_key = False
    after_answer = False
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if qi.WORD_ANSWER_KEY_HEADING_RE.fullmatch(line) and not qi.WORD_INLINE_ANSWER_RE.fullmatch(line):
            in_key = True
            lines.append(line)
            continue
        if in_key:
            lines.append(line)
            continue
        option = qi.OPTION_LABEL_RE.match(line)
        question = qi.TEXT_QUESTION_RE.match(line)
        if option and have_question:
            label = option.group(1).upper()
            if option_count >= 6 or label != qi.LETTERS[option_count]:
                raise ValueError("Answer choices must be labelled consecutively A–F; an option is missing, reordered, or beyond the six-option limit.")
            option_count += 1
            after_answer = False
            lines.append(line)
        elif question:
            if have_question and option_count < 2:
                raise ValueError("A question has fewer than two labelled answer choices. Use A., B., C., etc. before each option.")
            have_question = True
            option_count = 0
            after_answer = False
            lines.append(line)
        elif qi.TEXT_ANSWER_RE.match(line) or qi.TEXT_IMAGE_RE.match(line):
            lines.append(line)
            after_answer = bool(qi.TEXT_ANSWER_RE.match(line))
        elif have_question:
            if after_answer:
                raise ValueError("Unexpected text after a correct answer. Number each new question and remove page headers/footers before importing.")
            lines[-1] += " " + line
    if have_question and option_count < 2:
        raise ValueError("A question has fewer than two labelled answer choices. Use A., B., C., etc. before each option.")
    return qi.read_text_questions("\n".join(lines).encode(), filename)


def read_word_variant(data, filename):
    qi.validate_office_container(data, filename)
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as target:
        for entry in source.infolist():
            blob = source.read(entry.filename)
            if entry.filename == '[Content_Types].xml':
                root = safe_xml(blob)
                found = False
                for node in root:
                    if node.get('PartName') == '/word/document.xml':
                        node.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
                        found = True
                if not found:
                    raise ValueError('The file does not contain a Word document.')
                blob = etree.tostring(root)
            target.writestr(entry.filename, blob)
    # python-docx reads document parts only; embedded VBA is never executed.
    return qi.read_docx_questions(output.getvalue(), filename)


def convert_document(data, filename):
    """Convert older documents in a disposable profile, with active content off."""
    if len(data) > MAX_CONVERT_BYTES:
        raise ValueError('Legacy Word, RTF, and OpenDocument imports are limited to 20 MB per file.')
    executable = shutil.which('libreoffice') or shutil.which('soffice')
    if not executable:
        raise ValueError('Document conversion is unavailable on this server. Save the file as DOCX and upload it again.')
    suffix = Path(filename).suffix.lower()
    qi.validate_office_container(data, filename)
    if suffix == '.fodt':
        safe_xml(data)
    if not CONVERSION_LOCK.acquire(timeout=5):
        raise ValueError('Another document is being converted. Please try this upload again shortly.')
    try:
        with tempfile.TemporaryDirectory(prefix='mcq_import_') as directory:
            root = Path(directory)
            profile = root / 'profile'
            (profile / 'user').mkdir(parents=True)
            (profile / 'user' / 'registrymodifications.xcu').write_text('''<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse" oor:finalized="true"><value>3</value></prop><prop oor:name="DisableActiveContent" oor:op="fuse" oor:finalized="true"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>2</value></prop><prop oor:name="Field" oor:op="fuse"><value>false</value></prop><prop oor:name="Chart" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Save/Document"><prop oor:name="AutoSave" oor:op="fuse"><value>false</value></prop></item>
</oor:items>''', encoding='utf-8')
            source = root / ('source' + suffix)
            source.write_bytes(data)
            destination = root / 'out'
            destination.mkdir()
            command = [executable, '-env:UserInstallation=' + profile.as_uri(), '--headless', '--nologo',
                       '--nodefault', '--nofirststartwizard', '--norestore', '--convert-to',
                       'docx:Office Open XML Text', '--outdir', str(destination), str(source)]
            process = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                       start_new_session=(os.name == 'posix'))
            try:
                code = process.wait(timeout=45)
            except subprocess.TimeoutExpired as exc:
                if os.name == 'posix':
                    os.killpg(process.pid, signal.SIGKILL)
                else:
                    process.kill()
                process.wait()
                raise ValueError('Document conversion took too long. Save the file as DOCX and upload it again.') from exc
            output = destination / 'source.docx'
            if code or not output.is_file() or not output.stat().st_size:
                raise ValueError('This document could not be converted. Check that it opens normally and is not password-protected, or save it as DOCX.')
            if output.stat().st_size > qi.MAX_SOURCE_BYTES:
                raise ValueError('Converted document exceeds the import size limit.')
            result = qi.read_docx_questions(output.read_bytes(), filename)
        frame, findings, images = result
        if images:
            return require_review(result, filename,
                f'Imported from {suffix.upper().lstrip(".")} with images. Compare image size, position, '
                'crop and labels against the original before confirming your review.')
        if not frame.empty:
            # Conversion alone is not evidence that every question is uncertain.
            # Preserve the importer's actual answer confidence and findings.
            findings = [*findings, qi.issue(filename, '', '', 'INFO',
                f'Imported from {suffix.upper().lstrip(".")}: {len(frame)} questions read. '
                'Review the preview before export.')]
        return frame, findings, images
    finally:
        CONVERSION_LOCK.release()


def table_from_rows(rows, filename, method):
    if len(rows) > MAX_TABLE_ROWS:
        raise ValueError('Question table exceeds the 20,000 row limit.')
    if not rows:
        raise ValueError('No question table was found.')
    candidates = [(qi._excel_header_score(row), i) for i, row in enumerate(rows[:20])]
    score, header = max(candidates, key=lambda item: item[0])
    if score < 12:
        raise ValueError('The table needs Question, at least two Option columns, and Correct headers.')
    columns = [qi.normalise_text(v) for v in rows[header]]
    folded = [qi.normalise_header(v) for v in columns if v]
    if len(folded) != len(set(folded)):
        raise ValueError('The question table has duplicate column headers.')
    if any(re.fullmatch(r'(?:option|answer|choice)\s*(?:[7-9]|\d{2,})', c, re.I) for c in columns):
        raise ValueError('The app supports a maximum of six answer options per question.')
    records = []
    for row in rows[header + 1:]:
        if not any(qi.normalise_text(v) for v in row):
            continue
        if len(row) > len(columns) and any(qi.normalise_text(v) for v in row[len(columns):]):
            raise ValueError('A table row has more fields than its header. Check delimiters and quoting.')
        records.append((list(row) + [''] * len(columns))[:len(columns)])
    frame = pd.DataFrame(records, columns=columns)
    return qi.read_question_table(frame, filename, header_row=header, sheet_name=method, method_label=method)


def read_delimited(data, filename):
    text = decode_text(data)
    if Path(filename).suffix.lower() == '.tsv':
        delimiter = '\t'
    else:
        sample = text[:65536]
        try:
            delimiter = csv.Sniffer().sniff(sample, delimiters=',;\t').delimiter
        except csv.Error:
            delimiter = ','
    rows = list(csv.reader(io.StringIO(text), delimiter=delimiter, strict=True))
    return table_from_rows(rows, filename, 'Delimited question table')


def read_json_questions(data, filename):
    text = decode_text(data)
    def unique_fields(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f'Duplicate JSON field {key!r}. Keep one value per field; answers must not overwrite each other.')
            result[key] = value
        return result
    if Path(filename).suffix.lower() == '.jsonl':
        records = [json.loads(line, object_pairs_hook=unique_fields) for line in text.splitlines() if line.strip()]
    else:
        records = json.loads(text, object_pairs_hook=unique_fields)
        if isinstance(records, dict):
            records = records.get('questions')
    if not isinstance(records, list) or not records or len(records) > MAX_TABLE_ROWS:
        raise ValueError('Use a JSON array of question objects, or an object with a questions array (up to 20,000 records).')
    rows = [qi.CORE_COLUMNS]
    for i, record in enumerate(records, 1):
        if not isinstance(record, dict):
            raise ValueError(f'JSON question {i} must be an object.')
        if 'options' not in record:
            # The universal table schema is also accepted.
            if not all(key in record for key in ['Question', 'Correct', 'Option 1', 'Option 2']):
                raise ValueError(f'JSON question {i} needs question, options, and correct fields.')
            for key in record:
                if (re.fullmatch(r'(?:option|answer|choice)\s*(?:[a-z]|\d+)', key, re.I)
                        and key not in qi.OPTION_COLUMNS):
                    raise ValueError(f'JSON question {i} has an unsupported option field {key!r}. Use only Option 1–6; extra choices cannot be discarded.')
            rows.append([record.get(column, '') for column in qi.CORE_COLUMNS])
            continue
        options = record['options']
        if isinstance(options, dict):
            keys = sorted(options)
            if keys != qi.LETTERS[:len(keys)]:
                raise ValueError(f'JSON question {i}: option keys must be consecutive letters A–F.')
            options = [options[key] for key in keys]
        if not isinstance(options, list) or not 2 <= len(options) <= 6 or not all(isinstance(v, str) for v in options):
            raise ValueError(f'JSON question {i} needs 2–6 text options.')
        if not isinstance(record.get('question'), str) or 'correct' not in record:
            raise ValueError(f'JSON question {i} needs question text and a correct answer.')
        if isinstance(record['correct'], (list, dict, bool)):
            raise ValueError(f'JSON question {i} needs a single correct-answer letter, option text, or 1-based number.')
        rows.append([record.get('number', record.get('id', i)), record['question'], record.get('image', ''),
                     *options, *([''] * (6 - len(options))), record['correct']])
    return table_from_rows(rows, filename, 'JSON question bank')


class StaticHTML(HTMLParser):
    """Extract static text/tables without a browser, scripts, or remote requests."""
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.outside_parts = []
        self.tables = []
        self.table = None
        self.row = None
        self.cell = None
        self.skip = 0
        self.images = False
        self.lists = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag in {'script', 'style', 'noscript', 'head'}:
            self.skip += 1
        if self.skip:
            return
        if tag in {'img', 'svg', 'math', 'object', 'embed', 'iframe'}:
            self.images = True
        if tag in {'ol', 'ul'}:
            self.lists.append([tag, attrs.get('type', '1'), int(attrs.get('start', '1')) if attrs.get('start', '1').isdigit() else 1])
        if tag == 'li' and self.lists:
            kind, style, number = self.lists[-1]
            prefix = (chr(64 + number) if style.lower() == 'a' or len(self.lists) > 1 else str(number)) + '. ' if kind == 'ol' else ''
            self.parts.append('\n' + prefix)
            self.lists[-1][2] += 1
        if tag in {'p', 'div', 'br', 'tr', 'h1', 'h2', 'h3', 'h4', 'pre'}:
            self.parts.append('\n')
            if self.table is None:
                self.outside_parts.append('\n')
            if self.cell is not None:
                self.cell.append('\n')
        if tag == 'table':
            if self.table is not None:
                raise ValueError('Nested HTML tables are not supported; use a flat question table or DOCX.')
            self.table = []
        if tag == 'tr':
            self.row = []
        if tag in {'td', 'th'}:
            if attrs.get('rowspan', '1') != '1' or attrs.get('colspan', '1') != '1':
                raise ValueError('HTML contains merged table cells. Unmerge them so question, options and correct answer stay in separate columns.')
            self.cell = []

    def handle_endtag(self, tag):
        if tag in {'script', 'style', 'noscript', 'head'}:
            self.skip = max(0, self.skip - 1)
            return
        if self.skip:
            return
        if tag in {'ol', 'ul'} and self.lists:
            self.lists.pop()
        if tag in {'td', 'th'} and self.cell is not None:
            if self.row is not None:
                self.row.append(''.join(self.cell).strip())
            self.cell = None
            self.parts.append('\t')
        if tag == 'tr' and self.row is not None:
            if self.table is not None:
                self.table.append(self.row)
            self.row = None
        if tag == 'table' and self.table is not None:
            self.tables.append(self.table)
            self.table = None
        if tag in {'p', 'div', 'li', 'tr', 'h1', 'h2', 'h3', 'h4', 'pre'}:
            self.parts.append('\n')

    def handle_data(self, text):
        if not self.skip:
            self.parts.append(text)
            if self.table is None:
                self.outside_parts.append(text)
            if self.cell is not None:
                self.cell.append(text)


def html_text(text):
    parser = StaticHTML()
    parser.feed(text)
    parser.close()
    if parser.images:
        raise ValueError('HTML contains images, equations, or embedded objects that cannot be preserved. Use DOCX to retain them.')
    return parser


def read_html_questions(data, filename):
    parser = html_text(decode_text(data))
    tables = [rows for rows in parser.tables if any(qi._excel_header_score(row) >= 12 for row in rows[:20])]
    if tables:
        if any(qi.TEXT_QUESTION_RE.match(line.strip()) for line in ''.join(parser.outside_parts).splitlines()):
            raise ValueError('HTML mixes question tables with numbered questions outside the tables. Use one layout per file or upload DOCX.')
        results = [table_from_rows(rows, filename, 'HTML question table') for rows in tables]
        result = (pd.concat([r[0] for r in results], ignore_index=True), [v for r in results for v in r[1]], {})
    else:
        result = read_extracted_text(''.join(parser.parts), filename)
    return require_review(result, filename, 'HTML text import: check all questions and explicit answers. Visual answer formatting is not interpreted.')


def read_markdown_questions(data, filename):
    text = decode_text(data)
    if re.search(r'!\[.*?\]\(', text) or re.search(r'<(?:img|svg|math)\b', text, re.I):
        raise ValueError('Markdown images or equations are not embedded automatically. Use DOCX, or replace images with Image: filename.png and upload the referenced files.')
    lines = []
    for line in text.splitlines():
        line = re.sub(r'^\s{0,3}#{1,6}\s+', '', line)
        line = re.sub(r'^\s*[-*+]\s+(?=[A-Fa-f][.)])', '', line)
        line = re.sub(r'\*\*(.*?)\*\*|__(.*?)__', lambda m: m.group(1) if m.group(1) is not None else m.group(2), line)
        line = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', line)
        if re.match(r'^\s*```', line):
            continue
        lines.append(line)
    if any(re.fullmatch(r'\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?\s*', line) for line in lines):
        if any(qi.TEXT_QUESTION_RE.match(line.strip()) for line in lines if '|' not in line):
            raise ValueError('Markdown mixes a question table with numbered questions outside it. Use one layout per file.')
        rows = [[part.strip() for part in line.strip().strip('|').split('|')] for line in lines if '|' in line and not re.fullmatch(r'[\s|:\-]+', line)]
        result = table_from_rows(rows, filename, 'Markdown question table')
    else:
        result = read_extracted_text('\n'.join(lines), filename)
    return require_review(result, filename, 'Markdown text import: verify the questions and explicit answers; bold or highlighted answers alone are not interpreted.')


def read_ods_questions(data, filename):
    qi.validate_office_container(data, filename)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        root = safe_xml(archive.read('content.xml'))
    ns = {'t': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0', 'x': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'}
    if root.xpath('//*[local-name()="image" or local-name()="object" or local-name()="object-ole"]'):
        raise ValueError('ODS contains embedded images or objects. Use image filenames with separately uploaded images, or upload DOCX.')
    table_ns = '{' + ns['t'] + '}'
    candidates = []
    for sheet in root.findall('.//t:table', ns):
        rows = []
        for node in sheet.findall('t:table-row', ns):
            row = []
            for cell in node:
                if cell.tag not in {table_ns + 'table-cell', table_ns + 'covered-table-cell'}:
                    continue
                if cell.get(table_ns + 'formula') is not None:
                    raise ValueError('ODS contains formula cells. Paste calculated results as values before uploading.')
                if any(cell.get(table_ns + name, '1') != '1'
                       for name in ['number-columns-spanned', 'number-rows-spanned']):
                    raise ValueError('ODS contains merged cells. Unmerge question-table cells before uploading.')
                value = '\n'.join(''.join(p.itertext()) for p in cell.findall('x:p', ns))
                repeats = int(cell.get(table_ns + 'number-columns-repeated', '1'))
                if repeats < 1 or repeats > 1024 or len(row) + repeats > 1024:
                    if not value:
                        break  # Ignore the sheet's unused trailing blank columns.
                    raise ValueError('OpenDocument table expands to too many columns.')
                row.extend([value] * repeats)
            repeats = int(node.get(table_ns + 'number-rows-repeated', '1'))
            if not any(row):
                if repeats > MAX_TABLE_ROWS:
                    continue
            if repeats < 1 or len(rows) + repeats > MAX_TABLE_ROWS:
                raise ValueError('OpenDocument table expands to too many rows.')
            rows.extend([row] * repeats)
        score = max((qi._excel_header_score(row) for row in rows[:20]), default=0)
        if score >= 12:
            candidates.append((sheet.get(table_ns + 'name', 'Sheet'), rows))
    if not candidates:
        raise ValueError('No spreadsheet tables found in this ODS file.')
    results = [table_from_rows(rows, f'{filename} [{name}]' if len(candidates) > 1 else filename,
                               'OpenDocument spreadsheet') for name, rows in candidates]
    return pd.concat([result[0] for result in results], ignore_index=True), [v for result in results for v in result[1]], {}


def read_moodle_xml(data, filename):
    root = safe_xml(data)
    if root.tag != 'quiz':
        raise ValueError('Only Moodle question-bank XML (<quiz>) is supported, not arbitrary XML files.')
    rows = [qi.CORE_COLUMNS]
    for node in root.findall('question'):
        kind = node.get('type')
        if kind == 'category':
            continue
        if kind not in {'multichoice', 'truefalse'}:
            raise ValueError(f'Moodle question type {kind!r} is not supported. Export only single-answer MCQs or true/false questions.')
        if node.findall('.//file'):
            raise ValueError('Moodle XML contains embedded files. Use DOCX to preserve question images.')
        if kind == 'multichoice' and (node.findtext('single', 'true').strip().lower() not in {'true', '1'}):
            raise ValueError('Multiple-correct-answer Moodle questions are not supported.')
        question = ''.join(html_text(node.findtext('questiontext/text', '')).parts).strip()
        answers = node.findall('answer')
        if not 2 <= len(answers) <= 6:
            raise ValueError('Each Moodle question must have 2–6 answer choices.')
        fractions = [float(answer.get('fraction', '0')) for answer in answers]
        if fractions.count(100.0) != 1 or any(v != 0 for v in fractions if v != 100.0):
            raise ValueError('Moodle questions need exactly one 100% answer and no partial-credit or penalty answers.')
        options = [''.join(html_text(answer.findtext('text', '')).parts).strip() for answer in answers]
        rows.append([len(rows), question, '', *options, *([''] * (6-len(options))), qi.LETTERS[fractions.index(100.0)]])
    return table_from_rows(rows, filename, 'Moodle XML single-answer question bank')


def read_pdf_questions(data, filename):
    from pypdf import PdfReader
    if len(data) > 50 * 1024 * 1024:
        raise ValueError('PDF imports are limited to 50 MB.')
    reader = PdfReader(io.BytesIO(data))
    if reader.is_encrypted:
        raise ValueError('Password-protected PDFs are not supported. Upload an unlocked document.')
    if not 1 <= len(reader.pages) <= 200:
        raise ValueError('PDF imports are limited to 200 pages.')
    pages = []
    has_graphics = False
    for index, page in enumerate(reader.pages, 1):
        contents = page.get_contents()
        if contents is not None and len(contents.get_data()) > 10 * 1024 * 1024:
            raise ValueError(f'PDF page {index} is too complex to import safely. Use DOCX instead.')
        def inspect(operator, operands, cm, tm):
            nonlocal has_graphics
            if operator in {b'Do', b'BI', b're', b'm', b'l', b'c'}:
                has_graphics = True
        text = page.extract_text(visitor_operand_before=inspect) or ''
        if not text.strip():
            if contents is not None:
                raise ValueError(f'PDF page {index} has no readable text (it may be scanned). Run OCR first or upload the original document.')
            continue
        pages.append(text)
    if has_graphics:
        raise ValueError('This PDF contains images or drawn graphics that cannot be preserved by text import. Upload the original DOCX to retain the complete questions.')
    if not pages:
        raise ValueError('No readable text found in the PDF. Scanned PDFs require OCR before upload.')
    result = read_extracted_text('\n\n'.join(pages), filename)
    return require_review(result, filename, 'PDF text extraction: verify every question, answer option, and explicit answer against the original. Layout and visual answer formatting are not preserved.')


def read_additional_source(data, filename):
    suffix = Path(filename).suffix.lower()
    if suffix in WORD_VARIANTS:
        return read_word_variant(data, filename)
    if suffix in CONVERTED_DOCUMENTS:
        return convert_document(data, filename)
    if suffix in EXCEL_VARIANTS:
        return qi.read_excel_questions(data, filename)
    readers = {'.csv': read_delimited, '.tsv': read_delimited, '.json': read_json_questions,
               '.jsonl': read_json_questions, '.html': read_html_questions, '.htm': read_html_questions,
               '.md': read_markdown_questions, '.markdown': read_markdown_questions,
               '.ods': read_ods_questions, '.xml': read_moodle_xml, '.pdf': read_pdf_questions}
    if suffix not in readers:
        raise ValueError('Unsupported source type: ' + (suffix or 'unknown'))
    return readers[suffix](data, filename)
