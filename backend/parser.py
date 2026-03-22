"""
parser.py — Extract structured content from uploaded documents
Supports: .rtf, .txt, .docx, .pdf
"""

import re
import os


def _rtf_to_text(rtf_path: str) -> str:
    with open(rtf_path, 'rb') as f:
        raw = f.read().decode('latin-1', errors='ignore')

    # ── Step 0: Normalise line endings ──────────────────────
    raw = raw.replace('\r\n', '\n').replace('\r', '')

    # ── Step 1: Remove hex image data ────────────────────────
    raw = re.sub(r'[0-9a-f]{40,}', ' ', raw)
    raw = re.sub(r'[0-9A-F]{40,}', ' ', raw)

    # ── Step 2: Collapse font-switching sequences that split words ─
    # RTF inserts \n + \hich\afN\dbch\afN\loch\fN mid-word.
    # Must remove BEFORE splitting on newlines.
    raw = re.sub(r'\n\\hich\\af\d+\\dbch\\af\d+\\loch\\f\d+[ ]*', '', raw)
    raw = re.sub(r'\\hich\\af\d+\\dbch\\af\d+\\loch\\f\d+[ ]*', '', raw)
    raw = re.sub(r'\n\\hich\\f\d+[ ]*', '', raw)
    raw = re.sub(r'\\hich\\f\d+[ ]*', '', raw)
    raw = re.sub(r'\\loch\\f\d+[ ]*', '', raw)
    raw = re.sub(r'\\dbch\\af\d+[ ]*', '', raw)
    raw = re.sub(r'\\hich\\af\d+[ ]*', '', raw)
    raw = re.sub(r'\\loch\\af\d+[ ]*', '', raw)

        # ── Step 3: Paragraph breaks → newlines ──────────────────
    raw = raw.replace('\\par\r\n', '\n').replace('\\par\n', '\n')
    raw = re.sub(r'\\par\b', '\n', raw)
    raw = re.sub(r'\\pard\b', '\n', raw)
    raw = re.sub(r'\\line\b', '\n', raw)
    raw = re.sub(r'\\tab\b', ' ', raw)

    # ── Step 4: Remove all RTF control words ─────────────────
    raw = re.sub(r'\\[a-zA-Z]+\-?[0-9]*\s?', '', raw)
    raw = re.sub(r'\\[^\s\na-zA-Z]', '', raw)

    # ── Step 5: Remove braces ─────────────────────────────────
    raw = raw.replace('{', '').replace('}', '')

    # ── Step 6: Clean special char artifacts ─────────────────
    raw = raw.replace("'b7", '').replace("'3f", '→').replace("'B7", '')
    raw = raw.replace('\x00', '').replace('\x0c', '')
    # Remove leftover hex escapes like 'XX
    raw = re.sub(r"'[0-9a-fA-F]{2}", '', raw)

    # ── Step 7: Filter bad lines ──────────────────────────────
    lines = raw.split('\n')
    clean = []
    for line in lines:
        line = line.strip()
        if re.match(r'^[0-9a-fA-F\s]{20,}$', line):
            continue
        if re.search(r'(Times New Roman|Calibri|Arial|Symbol|Mangal)', line) and ';' in line and len(line) < 120:
            continue
        if re.search(r'shapeType|fFlipH|fLine0|wzName|dhgt|fFilled|fHidden|pictureGray|fLock', line):
            continue
        if re.match(r'^(Normal|heading|toc|index|caption|footer|header|List|Title|Closing)\s*[0-9]*\s*[;,]', line):
            continue
        if re.match(r'^[0-9]{8,}', line):
            continue
        if line and len(line) > 2:
            clean.append(line)

    text = '\n'.join(clean)

    # ── Step 8: Find start of real content ───────────────────
    lines_list = text.split('\n')
    start_line = 0
    for i, line in enumerate(lines_list):
        stripped = line.strip()
        # A real title line: has letters, no semicolons, no junk chars,
        # looks like natural language (spaces between words)
        if (len(stripped) > 10
                and re.search(r'[A-Za-z]{3}', stripped)
                and not re.search(r';', stripped)
                and not re.match(r'^[0-9a-fA-F\s]+$', stripped)
                and 'http' not in stripped
                and not re.match(r'^[0-9]+\s*$', stripped)
                and ' ' in stripped          # must have spaces (a real sentence)
                and len(stripped.split()) >= 3 # at least 3 words
                and stripped[0].isupper()):   # starts with capital
            start_line = i
            break

    text = '\n'.join(lines_list[start_line:])

    # ── Step 9: Remove trailing style tables ─────────────────
    for marker in ['footnote text;', 'Normal Table;', 'Default Paragraph Font']:
        idx = text.find(marker)
        if idx > 0:
            text = text[:idx]

    return text.strip()


def _clean_point(text: str) -> str:
    """Remove leading bullet chars, dots, numbers from a point."""
    text = re.sub(r'^[•\-*\.]+\s*', '', text).strip()
    text = re.sub(r"^'[0-9a-f]{2}", '', text).strip()
    # Remove leftover '3f' or 'b7' artifacts
    text = text.replace('3f', '').replace('b7', '')
    return text.strip()


def _is_good_point(text: str) -> bool:
    if not text or len(text) < 3:
        return False
    if re.match(r'^[0-9a-fA-F\s]{15,}$', text):
        return False
    if re.match(r'^[\d\s\.\-;:→]+$', text):
        return False
    if not re.search(r'[a-zA-Z]{2}', text):
        return False
    if re.search(r'shapeType|fFlipH|wzName|fHidden', text):
        return False
    return True


def _merge_broken_lines(raw_lines: list) -> list:
    """
    Merge lines that are broken mid-sentence.
    A line is a continuation if the previous line doesn't end a sentence
    AND the current line starts with lowercase.
    """
    if not raw_lines:
        return []

    merged = [raw_lines[0]]
    for line in raw_lines[1:]:
        prev = merged[-1].rstrip()
        curr = line.strip()

        prev_ends_sentence = bool(re.search(r'[.!?]$', prev))
        curr_starts_lower  = bool(curr and curr[0].islower())

        if not prev_ends_sentence and curr_starts_lower:
            merged[-1] = prev + ' ' + curr
        else:
            merged.append(line)

    return merged


def _parse_sections(text: str) -> dict:
    lines = [l.strip() for l in text.split('\n') if l.strip()]

    if not lines:
        return {"title": "Untitled", "sections": []}

    title = re.sub(r'^\d+[\.\)]\s*', '', lines[0]).strip()
    if not title or len(title) < 3:
        title = "Untitled Document"

    section_pattern = re.compile(r'^(\d+)[\.\)]\s+(.+)$')

    sections      = []
    current_head  = None
    current_raw   = []

    for line in lines[1:]:
        m = section_pattern.match(line)
        if m:
            if current_head is not None:
                merged = _merge_broken_lines(current_raw)
                pts = [_clean_point(p) for p in merged]
                pts = [p for p in pts if _is_good_point(p)]
                sections.append({"heading": current_head, "points": pts})
            current_head = m.group(2).strip()
            current_raw  = []
        else:
            if current_head is not None:
                clean = re.sub(r'^[•\-*b7]+\s*', '', line).strip()
                clean = re.sub(r'^[0-9]+\.\s+', '', clean)
                if clean and len(clean) > 1:
                    current_raw.append(clean)

    if current_head is not None:
        merged = _merge_broken_lines(current_raw)
        pts = [_clean_point(p) for p in merged]
        pts = [p for p in pts if _is_good_point(p)]
        sections.append({"heading": current_head, "points": pts})

    if not sections:
        sections = _fallback_sections(lines[1:])

    return {"title": title, "sections": sections}


def _fallback_sections(lines: list) -> list:
    sections = []
    current_head   = None
    current_points = []
    for line in lines:
        is_heading = (
            (line == line.upper() and len(line.split()) <= 6 and len(line) > 3)
            or (line.istitle() and len(line.split()) <= 5)
        )
        if is_heading and len(line) > 3:
            if current_head:
                sections.append({"heading": current_head, "points": current_points})
            current_head   = line
            current_points = []
        else:
            if current_head and _is_good_point(line):
                current_points.append(line)
    if current_head:
        sections.append({"heading": current_head, "points": current_points})
    if not sections:
        sections = [{"heading": "Content", "points": [l for l in lines if _is_good_point(l)]}]
    return sections


def extract_document_content(file_path: str) -> dict:
    ext = os.path.splitext(file_path)[1].lower()

    if ext == '.rtf':
        text = _rtf_to_text(file_path)
    elif ext == '.txt':
        with open(file_path, 'r', errors='ignore') as f:
            text = f.read()
    elif ext == '.docx':
        try:
            from docx import Document
            doc = Document(file_path)
            text = '\n'.join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception as e:
            raise ValueError(f"Cannot read .docx: {e}")
    elif ext == '.pdf':
        try:
            from pypdf import PdfReader
            reader = PdfReader(file_path)
            text = '\n'.join(page.extract_text() or '' for page in reader.pages)
        except Exception as e:
            raise ValueError(f"Cannot read .pdf: {e}")
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    result = _parse_sections(text)
    print(f"  Parsed: title='{result['title']}', sections={len(result['sections'])}")
    return result
