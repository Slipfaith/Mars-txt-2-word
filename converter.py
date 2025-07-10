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
    rus_force_encoding: Optional[str] = None,
) -> None:
    """Create Word document from pairs of txt files."""
    doc = Document()
    files = sorted([f for f in os.listdir(eng_folder) if f.lower().endswith('.txt')])
    if not files:
        raise FileNotFoundError("No .txt files in english folder")

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
        doc.add_paragraph(f'Файл: {filename}')
        num_rows = max(len(eng_lines), len(rus_lines))
        table = doc.add_table(rows=num_rows, cols=2)
        table.style = 'Table Grid'
        for i in range(num_rows):
            table.cell(i, 0).text = eng_lines[i] if i < len(eng_lines) else ''
            table.cell(i, 1).text = rus_lines[i] if i < len(rus_lines) else ''
        doc.add_paragraph()
    doc.save(output_path)
    logger.info("Word document saved: %s", output_path)


def import_from_word(word_path: str, eng_output_folder: str, rus_output_folder: str) -> None:
    """Split Word document back into english/russian txt files."""
    doc = Document(word_path)
    current_filename: Optional[str] = None
    file_data = {}

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text.startswith("Файл:"):
                current_filename = text.replace("Файл:", "").strip()
                file_data[current_filename] = ([], [])
                logger.debug("Processing section for %s", current_filename)
        elif isinstance(block, Table) and current_filename:
            eng_lines, rus_lines = file_data[current_filename]
            for row in block.rows:
                eng_lines.append(row.cells[0].text.strip())
                rus_lines.append(row.cells[1].text.strip())
            current_filename = None

    os.makedirs(eng_output_folder, exist_ok=True)
    os.makedirs(rus_output_folder, exist_ok=True)
    for filename, (eng_lines, rus_lines) in file_data.items():
        filename_txt = filename if filename.lower().endswith('.txt') else f"{filename}.txt"
        eng_file_path = os.path.join(eng_output_folder, filename_txt)
        rus_file_path = os.path.join(rus_output_folder, filename_txt)
        with open(eng_file_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(eng_lines))
        with open(rus_file_path, 'w', encoding='cp1251') as f:
            f.write("\n".join(rus_lines))
        logger.debug("Saved %s and %s", eng_file_path, rus_file_path)
    logger.info("Finished importing from %s", word_path)
