import logging
import os
from typing import Iterable, List, Optional

import chardet
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

logger = logging.getLogger(__name__)


def detect_encoding(file_path: str, num_bytes: int = 10000) -> Optional[str]:
    """Detect file encoding using chardet with fallback for MacCyrillic."""
    with open(file_path, 'rb') as f:
        rawdata = f.read(num_bytes)
    result = chardet.detect(rawdata)
    encoding = result.get('encoding')
    if encoding and encoding.lower() == "maccyrillic":
        encoding = "cp1251"
    logger.debug("%s encoding detected as %s", file_path, encoding)
    return encoding


def read_lines_auto(
    file_path: str,
    default_encoding: str = 'utf-8',
    force_encoding: Optional[str] = None,
) -> List[str]:
    """Read lines from file using detected or forced encoding."""
    encoding = force_encoding if force_encoding else detect_encoding(file_path) or default_encoding
    source = "forced" if force_encoding else "auto-detected"
    logger.debug("Reading %s with %s encoding %s", file_path, source, encoding)
    try:
        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
            lines = [line.strip() for line in f.readlines()]
        logger.debug("Read %d lines from %s", len(lines), file_path)
        return lines
    except Exception as exc:
        logger.error("Failed reading %s with %s: %s", file_path, encoding, exc)
        with open(file_path, 'r', encoding=default_encoding, errors='replace') as f:
            lines = [line.strip() for line in f.readlines()]
        logger.debug(
            "Fallback read %d lines from %s using %s", len(lines), file_path, default_encoding
        )
        return lines


def iter_block_items(parent) -> Iterable:
    """Yield Paragraph and Table items from a Word document."""
    parent_elm = parent.element.body if hasattr(parent, 'element') else parent
    for child in parent_elm:
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)


def export_to_word(
    eng_folder: str,
    rus_folder: str,
    output_path: str,
    file_extension: Optional[str] = None,
    rus_force_encoding: Optional[str] = None,
) -> None:
    """Create Word document from pairs of subtitle files.

    ``file_extension`` may be ``"txt"`` or ``"srt"``. If ``None`` the
    extension will be auto-detected by looking for files in the english folder.
    """
    doc = Document()

    def gather(ext: str) -> List[str]:
        return sorted(
            [f for f in os.listdir(eng_folder) if f.lower().endswith(f".{ext}")]
        )

    if file_extension is None:
        for ext in ("txt", "srt"):
            files = gather(ext)
            if files:
                file_extension = ext
                break
        else:
            raise FileNotFoundError("No .txt or .srt files in english folder")
    else:
        file_extension = file_extension.lower().lstrip(".")
        files = gather(file_extension)
        if not files:
            raise FileNotFoundError(
                f"No .{file_extension} files in english folder"
            )

    # create a single two column table and populate it with all files
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'

    # first row stores file type so we can reconstruct later
    row = table.add_row()
    row.cells[0].text = f"Тип файлов: {file_extension}"
    row.cells[1].text = ""

    for filename in files:
        eng_file_path = os.path.join(eng_folder, filename)
        rus_file_path = os.path.join(rus_folder, filename)
        if not os.path.exists(rus_file_path):
            logger.warning("Missing russian file for %s", filename)
            continue
        eng_lines = read_lines_auto(eng_file_path, default_encoding='utf-8')
        rus_lines = read_lines_auto(
            rus_file_path, default_encoding='cp1251', force_encoding=rus_force_encoding
        )

        # row marking the beginning of a new file
        marker_row = table.add_row()
        file_cell = marker_row.cells[0]
        file_cell.text = ""
        run = file_cell.paragraphs[0].add_run(f"Файл: {filename}")
        run.bold = True
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        shading = parse_xml(
            r'<w:shd {} w:fill="CCFFCC"/>'.format(nsdecls("w"))
        )
        file_cell._tc.get_or_add_tcPr().append(shading)
        marker_row.cells[1].text = ""

        num_rows = max(len(eng_lines), len(rus_lines))
        for i in range(num_rows):
            data_row = table.add_row()
            data_row.cells[0].text = eng_lines[i] if i < len(eng_lines) else ''
            data_row.cells[1].text = rus_lines[i] if i < len(rus_lines) else ''
    doc.save(output_path)
    logger.info("Word document saved: %s", output_path)


def import_from_word(
    word_path: str, eng_output_folder: str, rus_output_folder: str
) -> None:
    """Split Word document back into english/russian txt or srt files."""
    doc = Document(word_path)

    # assume the document contains a single two column table
    if not doc.tables:
        raise ValueError("Word document does not contain a table")

    table = doc.tables[0]
    file_format = "txt"
    file_data = {}
    current_filename: Optional[str] = None

    for idx, row in enumerate(table.rows):
        first = row.cells[0].text.strip()
        second = row.cells[1].text.strip()

        if idx == 0 and first.startswith("Тип файлов:"):
            file_format = first.split(":", 1)[1].strip().lower()
            continue

        if first.startswith("Файл:"):
            current_filename = first.replace("Файл:", "").strip()
            file_data[current_filename] = ([], [])
            logger.debug("Processing section for %s", current_filename)
            continue

        if current_filename:
            eng_lines, rus_lines = file_data[current_filename]
            eng_lines.append(first)
            rus_lines.append(second)

    os.makedirs(eng_output_folder, exist_ok=True)
    os.makedirs(rus_output_folder, exist_ok=True)
    ext = "srt" if file_format.lower() == "srt" else "txt"
    for filename, (eng_lines, rus_lines) in file_data.items():
        filename_with_ext = (
            filename
            if filename.lower().endswith(f".{ext}")
            else f"{filename}.{ext}"
        )
        eng_file_path = os.path.join(eng_output_folder, filename_with_ext)
        rus_file_path = os.path.join(rus_output_folder, filename_with_ext)
        with open(eng_file_path, "w", encoding="utf-8") as f:
            f.write("\n".join(eng_lines))
        with open(rus_file_path, "w", encoding="cp1251") as f:
            f.write("\n".join(rus_lines))
        logger.debug("Saved %s and %s", eng_file_path, rus_file_path)
    logger.info("Finished importing from %s", word_path)
