"""
Распознавание PDF-журнала: сначала пробуем текстовый слой (pdfplumber),
если его нет — рендерим страницы в картинки и прогоняем через Tesseract OCR.
"""
import io
import re
from typing import List, Dict

import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes


NAME_RX = re.compile(
    r"([А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі\-]+(?:\s+[А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі\-]+){1,3})"
)
MARK_RX = re.compile(r"^(1|2|3|4|5|10|А|С|Н)$", re.IGNORECASE)


def _extract_from_text(text: str) -> List[Dict]:
    """Извлекает строки вида 'ФИО 4 5 3 4 А 5 4 4'."""
    out = []
    seen = set()
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if re.search(r"учител|мұғалім|teacher|тоқсан|четверть|quarter", line, re.I):
            continue
        if re.fullmatch(r"[\d\s\.\-\/№#]+", line):
            continue
        m = NAME_RX.search(line)
        if not m:
            continue
        name = m.group(1).strip()
        if len(name.replace(" ", "")) < 4:
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)

        tail = line[m.end():]
        tokens = re.split(r"[\s,;|\t]+", tail)
        marks = [t.upper() for t in tokens if MARK_RX.match(t)]
        out.append({"name": name, "marks": marks, "raw": line})
    return out


def parse_pdf(file_bytes: bytes, use_ocr: bool = True) -> List[Dict]:
    """Главная функция: читает PDF и возвращает список учеников."""
    # 1) пробуем текстовый слой
    text_chunks = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            text_chunks.append(t)
    full_text = "\n".join(text_chunks)
    students = _extract_from_text(full_text)

    if students:
        return students

    # 2) если пусто и разрешён OCR — распознаём как скан
    if not use_ocr:
        return []

    images = convert_from_bytes(file_bytes, dpi=300)
    ocr_text = []
    for i, img in enumerate(images, 1):
        try:
            txt = pytesseract.image_to_string(img, lang="rus+kaz+eng")
        except pytesseract.TesseractError:
            txt = pytesseract.image_to_string(img, lang="rus+eng")
        ocr_text.append(txt)

    return _extract_from_text("\n".join(ocr_text))