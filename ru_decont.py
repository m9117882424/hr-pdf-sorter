from __future__ import annotations

import argparse
import csv
import datetime as dt
import re
import subprocess
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
import pdfplumber
import pytesseract
from openpyxl import load_workbook
from PIL import Image, ImageFilter, ImageOps

AH_START = 34  # AH
OUTPUT_HEADERS = ["Дата", "ИНН / Кимлик", "Фамилия / Наименование", "Имя", "Сумма", "Сумма 2"]

DEFAULT_TESSERACT = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
if DEFAULT_TESSERACT.exists():
    pytesseract.pytesseract.tesseract_cmd = str(DEFAULT_TESSERACT)


@dataclass
class PdfRecord:
    filename: str
    text_source: str = ""
    date: Optional[str] = None
    date_source: str = ""
    inn_kimlik: Optional[str] = None
    kimlik_source: str = ""
    surname_title: Optional[str] = None
    name: Optional[str] = None
    amounts: list[str] | None = None
    raw_top_date_ocr: str = ""
    raw_kimlik_ocr: str = ""
    page_count: int = 0
    payment_pages: list[int] | None = None
    skipped_pages: list[int] | None = None

    def __post_init__(self) -> None:
        self.amounts = self.amounts or []
        self.payment_pages = self.payment_pages or []
        self.skipped_pages = self.skipped_pages or []

    @property
    def full_name(self) -> str:
        return normalize_name(f"{self.surname_title or ''} {self.name or ''}")


def normalize_spaces(text) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def normalize_name(text) -> str:
    text = str(text or "").upper().replace("Ё", "Е")
    text = re.sub(r"[^А-ЯA-Z0-9 ]+", " ", text)
    return normalize_spaces(text)


def normalize_date(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime("%d.%m.%Y")
    text = str(value).strip()
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", text)
    if m:
        yyyy, mm, dd = m.groups()
        return f"{dd}.{mm}.{yyyy}"
    m = re.search(r"(\d{2})[./-](\d{2})[./-](\d{4})", text)
    if m:
        dd, mm, yyyy = m.groups()
        return f"{dd}.{mm}.{yyyy}"
    return normalize_spaces(text)


def normalize_id(value) -> str:
    return re.sub(r"\D+", "", str(value or ""))


def normalize_amount(text: str) -> str:
    text = re.sub(r"\s+", "", str(text or "").strip())
    m = re.search(r"\d{1,3}(?:[.\s]\d{3})*,\d{2}|\d+,\d{2}|\d{1,3}(?:[,\s]\d{3})*\.\d{2}|\d+\.\d{2}", text)
    if not m:
        return text
    amount = m.group(0).replace(" ", "")
    if "," in amount and "." in amount:
        amount = amount.replace(".", "") if amount.rfind(",") > amount.rfind(".") else amount.replace(",", "")
    elif "," in amount:
        amount = amount.replace(".", "")
    return amount


def amount_to_excel_number(text: str):
    if text is None:
        return None
    text = re.sub(r"\s+", "", str(text).strip())
    if not text:
        return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def check_tesseract() -> None:
    try:
        subprocess.run([pytesseract.pytesseract.tesseract_cmd, "--version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
    except Exception as exc:
        raise RuntimeError("Tesseract OCR не найден. Установите Tesseract для Windows и русский язык или проверьте путь C:\\Program Files\\Tesseract-OCR\\tesseract.exe") from exc


def preprocess_for_ocr(img: Image.Image, scale: int = 2, threshold: bool = True) -> Image.Image:
    img = img.convert("L")
    img = ImageOps.autocontrast(img)
    if scale > 1:
        img = img.resize((img.width * scale, img.height * scale), Image.Resampling.LANCZOS)
    img = img.filter(ImageFilter.SHARPEN)
    if threshold:
        img = img.point(lambda p: 255 if p > 175 else 0)
    return img


def ocr_image(img: Image.Image, *, lang: str = "rus+eng", psm: int = 6, whitelist: str | None = None, scale: int = 2, threshold: bool = True) -> str:
    config = f"--oem 3 --psm {psm}"
    if whitelist:
        safe = str(whitelist).replace(chr(34), "").replace(chr(39), "").replace(" ", "")
        config += f" -c tessedit_char_whitelist={safe}"
    prepared = preprocess_for_ocr(img, scale=scale, threshold=threshold)
    try:
        return pytesseract.image_to_string(prepared, lang=lang, config=config)
    except (pytesseract.TesseractError, ValueError):
        return pytesseract.image_to_string(prepared, lang="eng", config=config)


def render_pdf_pages(pdf_path: Path, dpi: int = 300) -> list[Image.Image]:
    pages: list[Image.Image] = []
    doc = fitz.open(str(pdf_path))
    try:
        matrix = fitz.Matrix(dpi / 72, dpi / 72)
        for page in doc:
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            pages.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
    finally:
        doc.close()
    return pages


def crop_rel(img: Image.Image, box: tuple[float, float, float, float]) -> Image.Image:
    w, h = img.size
    l, t, r, b = box
    return img.crop((int(w * l), int(h * t), int(w * r), int(h * b)))


def dark_ratio(img: Image.Image, box: tuple[float, float, float, float]) -> float:
    g = crop_rel(img, box).convert("L")
    hist = g.histogram()
    dark = sum(hist[:180])
    total = max(1, g.width * g.height)
    return dark / total


def looks_like_payment_page(img: Image.Image) -> bool:
    if dark_ratio(img, (0.04, 0.11, 0.96, 0.84)) >= 0.018:
        return True
    probe = crop_rel(img, (0.04, 0.05, 0.96, 0.62))
    txt = normalize_name(ocr_image(probe, lang="rus+eng", psm=6, scale=1, threshold=False))
    hits = sum(1 for key in ("КВИТАНЦИЯ", "СВЕДЕНИЯ", "ПЛАТЕЖ", "ИНН", "КИМЛИК", "ZIRAAT") if key in txt)
    return hits >= 2


def extract_top_right_date_from_page_image(img: Image.Image) -> tuple[str, str]:
    crops = [(0.58, 0.075, 0.91, 0.185), (0.61, 0.095, 0.82, 0.175), (0.63, 0.115, 0.80, 0.170), (0.55, 0.065, 0.92, 0.170)]
    raw: list[str] = []
    for box in crops:
        for psm in (6, 7, 11):
            txt = ocr_image(crop_rel(img, box), lang="eng+rus", psm=psm, whitelist="0123456789./- ", scale=2)
            raw.append(txt)
            m = re.search(r"\b(\d{2})[./-](\d{2})[./-](\d{4})\b", txt)
            if m:
                return f"{m.group(1)}.{m.group(2)}.{m.group(3)}", "\n".join(raw)
    return "", "\n".join(raw)


def extract_kimlik_from_page_image(img: Image.Image) -> tuple[str, str]:
    crops = [(0.10, 0.245, 0.42, 0.345), (0.14, 0.265, 0.36, 0.335), (0.05, 0.235, 0.55, 0.360), (0.10, 0.175, 0.42, 0.295), (0.03, 0.170, 0.55, 0.330)]
    raw: list[str] = []
    for box in crops:
        for psm in (6, 11):
            txt = ocr_image(crop_rel(img, box), lang="eng+rus", psm=psm, whitelist="0123456789 ", scale=2)
            raw.append(txt)
            found = re.findall(r"\b\d{10,12}\b", txt)
            if found:
                return found[0], "\n".join(raw)
    return "", "\n".join(raw)


def extract_amounts_from_page_image(img: Image.Image) -> tuple[list[str], str]:
    crops = [(0.60, 0.58, 0.76, 0.78), (0.58, 0.55, 0.78, 0.83), (0.30, 0.47, 0.86, 0.86)]
    raw: list[str] = []
    found: list[str] = []
    for box in crops:
        for psm in (6, 11):
            txt = ocr_image(crop_rel(img, box), lang="eng+rus", psm=psm, whitelist="0123456789., ", scale=2)
            raw.append(txt)
            for m in re.finditer(r"\b\d{1,3}(?:[ .]\d{3})*,\d{2}\b|\b\d+,\d{2}\b", txt):
                value = normalize_amount(m.group(0))
                num = amount_to_excel_number(value)
                if num is not None and num >= 100:
                    found.append(value)
        if found:
            break
    result: list[str] = []
    for value in found:
        if value not in result:
            result.append(value)
    return result, "\n".join(raw)


def extract_doc_fields_by_ocr(pdf_path: Path) -> dict:
    pages = render_pdf_pages(pdf_path, dpi=300)
    payment_indexes = [i for i, img in enumerate(pages) if looks_like_payment_page(img)]
    if not payment_indexes:
        payment_indexes = list(range(len(pages)))

    first_date = ""
    first_kimlik = ""
    date_raw = ""
    kimlik_raw = ""
    amounts: list[str] = []
    text_parts: list[str] = []

    for i in payment_indexes:
        img = pages[i]
        if not first_date:
            first_date, date_raw = extract_top_right_date_from_page_image(img)
        if not first_kimlik:
            first_kimlik, kimlik_raw = extract_kimlik_from_page_image(img)
        page_amounts, _ = extract_amounts_from_page_image(img)
        for value in page_amounts:
            if value not in amounts:
                amounts.append(value)
        text_parts.append(ocr_image(img, lang="rus+eng", psm=6, scale=2, threshold=True))

    payment_pages = [i + 1 for i in payment_indexes]
    skipped_pages = [i + 1 for i in range(len(pages)) if i not in payment_indexes]
    return {
        "text": "\n".join(text_parts),
        "date": first_date,
        "inn_kimlik": first_kimlik,
        "amounts": amounts,
        "date_raw": date_raw,
        "kimlik_raw": kimlik_raw,
        "page_count": len(pages),
        "payment_pages": payment_pages,
        "skipped_pages": skipped_pages,
    }


def extract_text_pages_from_pdf(pdf_path: Path) -> list[str]:
    chunks: list[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            chunks.append(page.extract_text() or "")
    return chunks


def extract_field(patterns: list[str], text: str, flags=re.IGNORECASE | re.MULTILINE) -> Optional[str]:
    for pattern in patterns:
        m = re.search(pattern, text, flags)
        if m:
            return normalize_spaces(m.group(1))
    return None


def extract_date_from_pdf_text(text: str) -> Optional[str]:
    date = extract_field([
        r"Дата\s+№\s*документа.*?\n.*?(\d{2}[./-]\d{2}[./-]\d{4})",
        r"Дата\s+Номер.*?\n.*?(\d{2}[./-]\d{2}[./-]\d{4})",
        r"Дата\s+оплаты.*?\n\s*02\s+(\d{2}[./-]\d{2}[./-]\d{4})",
    ], text, flags=re.IGNORECASE | re.DOTALL)
    return normalize_date(date) if date else None


def extract_kimlik_from_pdf_text(text: str) -> Optional[str]:
    value = extract_field([
        r"ИНН\s*/\s*Кимлик.*?\n\s*033202\s+(\d{10,12})",
        r"ИНН\s*/\s*Кимлик.*?\n\s*\d{6}\s+(\d{10,12})",
        r"Идентификационный\s+номер\s+налогоплательщика.*?(\d{10,12})",
        r"Идентификационный\s+номер\s+Т\.?Р\.?/ИНН\s*[:\-]?\s*(\d{10,12})",
        r"ИНН\s*[:\-]?\s*(\d{10,12})",
        r"Кимлик\s*[:\-]?\s*(\d{10,12})",
    ], text, flags=re.IGNORECASE | re.DOTALL)
    return normalize_id(value) if value else None


def extract_name_from_filename(filename: str) -> tuple[str, str]:
    stem = Path(filename).stem
    stem = re.sub(r"(?i)[_\s-]*(RU|RUS|TR)$", "", stem)
    stem = re.sub(r"^ПП\s+от\s+\d{2}[.]\d{2}[.]\d{4}\s+", "", stem, flags=re.IGNORECASE)
    stem = re.sub(r"\d{5,}$", "", stem)
    parts = normalize_name(stem).split()
    if len(parts) >= 2:
        return parts[0], " ".join(parts[1:])
    return (parts[0], "") if parts else ("", "")


def extract_all_amounts(text: str) -> list[str]:
    found: list[str] = []
    patterns = [
        r"Общее\s+количество\s+платежей\s*:\s*\d+\s+([\d\s.]+,\d{2})",
        r"ИТОГО\s*\d*\s+([\d\s.]+,\d{2})",
        r"Итог\s*\d*\s+([\d\s.]+,\d{2})",
        r"\b20\d{2}\s+\d+\s+([\d\s.]+,\d{2})",
        r"(?:Сумма|Размер)\s*[:\-]?\s*([\d\s.]+,\d{2})",
    ]
    for pattern in patterns:
        for m in re.finditer(pattern, text, re.IGNORECASE):
            found.append(normalize_amount(m.group(1)))
    for m in re.finditer(r"\b\d{1,3}(?:[ .]\d{3})*,\d{2}\b|\b\d+,\d{2}\b", text):
        value = normalize_amount(m.group(0))
        num = amount_to_excel_number(value)
        if num is not None and num >= 100:
            found.append(value)
    result: list[str] = []
    for x in found:
        if x and x not in result:
            result.append(x)
    return result


def parse_pdf(pdf_path: Path, use_ocr: bool = True) -> PdfRecord:
    text_pages = extract_text_pages_from_pdf(pdf_path)
    text = "\n".join(text_pages)
    text_source = "PDF_TEXT"
    ocr_fields: dict = {}

    date = extract_date_from_pdf_text(text)
    kimlik = extract_kimlik_from_pdf_text(text)
    amounts = extract_all_amounts(text)
    needs_ocr = use_ocr and (len(normalize_spaces(text)) < 80 or not date or not kimlik or not amounts)

    if needs_ocr:
        ocr_fields = extract_doc_fields_by_ocr(pdf_path)
        text = ocr_fields["text"] or text
        text_source = "OCR" if len(normalize_spaces("\n".join(text_pages))) < 80 else "PDF_TEXT+OCR"
        date = ocr_fields.get("date") or date
        kimlik = ocr_fields.get("inn_kimlik") or kimlik
        amounts = ocr_fields.get("amounts") or amounts or extract_all_amounts(text)

    surname_title = extract_field([
        r"Фамилия\s*/\s*Наименование\s*[:\-]?\s*([^\n:]+)",
        r"Фамилия\s*/\s*Назв\.?\s*имя\s*[:\-]?\s*([^\n:]+)",
        r"Фамилия\s*/\s*Должность\s*[:\-]?\s*([^\n:]+)",
    ], text, flags=re.IGNORECASE)
    name = extract_field([r"(?:^|\n)\s*Имя\s*[:\-]?\s*([^\n:]+)"], text, flags=re.IGNORECASE)

    file_surname, file_name = extract_name_from_filename(pdf_path.name)
    if not surname_title or len(normalize_name(surname_title)) < 2:
        surname_title = file_surname
    if not name or len(normalize_name(name)) < 2:
        name = file_name

    page_count = ocr_fields.get("page_count", len(text_pages))
    return PdfRecord(
        filename=pdf_path.name,
        text_source=text_source,
        date=normalize_date(date) if date else None,
        date_source="OCR_TOP_RIGHT_DATE_CELL" if ocr_fields.get("date") else ("PDF_TEXT_OR_FULL_OCR_REGEX" if date else ""),
        inn_kimlik=normalize_id(kimlik) if kimlik else None,
        kimlik_source="OCR_INN_KIMLIK_CELL" if ocr_fields.get("inn_kimlik") else ("PDF_TEXT_OR_FULL_OCR_REGEX" if kimlik else ""),
        surname_title=normalize_spaces(surname_title) if surname_title else None,
        name=normalize_spaces(name) if name else None,
        amounts=amounts,
        raw_top_date_ocr=ocr_fields.get("date_raw", ""),
        raw_kimlik_ocr=ocr_fields.get("kimlik_raw", ""),
        page_count=page_count,
        payment_pages=ocr_fields.get("payment_pages", list(range(1, page_count + 1))),
        skipped_pages=ocr_fields.get("skipped_pages", []),
    )


def find_header_row(ws, required_headers: list[str], max_scan_rows: int = 20) -> int:
    required = {normalize_spaces(x).lower() for x in required_headers}
    for row in range(1, max_scan_rows + 1):
        values = {normalize_spaces(ws.cell(row=row, column=col).value).lower() for col in range(1, ws.max_column + 1)}
        if required.issubset(values):
            return row
    raise ValueError(f"Не найдены заголовки: {', '.join(required_headers)}")


def map_headers(ws, header_row: int) -> dict[str, int]:
    return {normalize_spaces(ws.cell(row=header_row, column=col).value): col for col in range(1, ws.max_column + 1) if ws.cell(row=header_row, column=col).value is not None}


def build_excel_indexes(ws, header_row: int, headers: dict[str, int]) -> dict:
    fio_col = headers.get("ФИО рус")
    date_col = headers.get("Дата оплаты")
    kimlik_col = headers.get("Y.T.C № Кимлики")
    if not fio_col or not date_col or not kimlik_col:
        raise ValueError("В Excel не найдены обязательные столбцы: ФИО рус, Дата оплаты, Y.T.C № Кимлики")
    exact = {}
    by_name_date = defaultdict(list)
    by_name_id = defaultdict(list)
    by_name = defaultdict(list)
    for row in range(header_row + 1, ws.max_row + 1):
        fio = normalize_name(ws.cell(row=row, column=fio_col).value)
        pay_date = normalize_date(ws.cell(row=row, column=date_col).value)
        kimlik = normalize_id(ws.cell(row=row, column=kimlik_col).value)
        if not fio:
            continue
        exact[(fio, pay_date, kimlik)] = row
        by_name_date[(fio, pay_date)].append(row)
        by_name_id[(fio, kimlik)].append(row)
        by_name[fio].append(row)
    return {"exact": exact, "by_name_date": by_name_date, "by_name_id": by_name_id, "by_name": by_name}


def find_excel_row(indexes: dict, rec: PdfRecord) -> tuple[Optional[int], str]:
    fio = rec.full_name
    date = normalize_date(rec.date or "")
    kimlik = normalize_id(rec.inn_kimlik or "")
    if fio and date and kimlik and indexes["exact"].get((fio, date, kimlik)):
        return indexes["exact"][(fio, date, kimlik)], "ФИО + дата из документа + кимлик"
    if fio and date and len(indexes["by_name_date"].get((fio, date), [])) == 1:
        return indexes["by_name_date"][(fio, date)][0], "ФИО + дата из документа"
    if fio and kimlik and len(indexes["by_name_id"].get((fio, kimlik), [])) == 1:
        return indexes["by_name_id"][(fio, kimlik)][0], "ФИО + кимлик"
    if fio and len(indexes["by_name"].get(fio, [])) == 1:
        return indexes["by_name"][fio][0], "только ФИО, сотрудник уникален в Excel"
    return None, "Нет безопасного совпадения"


def ensure_output_headers(ws, header_row: int) -> None:
    for i, header in enumerate(OUTPUT_HEADERS):
        ws.cell(row=header_row, column=AH_START + i).value = header


def write_record_to_row(ws, row: int, record: PdfRecord) -> None:
    values = [record.date, record.inn_kimlik, record.surname_title, record.name]
    for i, value in enumerate(values):
        ws.cell(row=row, column=AH_START + i).value = value
    for offset, amount in enumerate(record.amounts[:2], start=4):
        cell = ws.cell(row=row, column=AH_START + offset)
        cell.value = amount_to_excel_number(amount)
        if cell.value is not None:
            cell.number_format = "#,##0.00"


def save_csv(path: Path, rows: list[dict]) -> None:
    if not rows:
        rows = [{"status": "empty"}]
    fieldnames: list[str] = []
    for row in rows:
        for key in row.keys():
            if key not in fieldnames:
                fieldnames.append(key)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def unique_pdf_files(pdf_dir: Path) -> list[Path]:
    result: list[Path] = []
    seen: set[str] = set()
    for path in sorted(pdf_dir.rglob("*")):
        if path.is_file() and path.suffix.lower() == ".pdf":
            key = str(path.resolve()).lower()
            if key not in seen:
                seen.add(key)
                result.append(path)
    return result


def process(excel_path: Path, pdf_dir: Path, sheet_name: Optional[str] = None, use_ocr: bool = True) -> None:
    if use_ocr:
        check_tesseract()
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel не найден: {excel_path}")
    if not pdf_dir.exists() or not pdf_dir.is_dir():
        raise FileNotFoundError(f"Папка с PDF не найдена: {pdf_dir}")
    pdf_files = unique_pdf_files(pdf_dir)
    if not pdf_files:
        raise FileNotFoundError(f"В папке нет PDF-файлов: {pdf_dir}")

    print(f"Найдено PDF-файлов: {len(pdf_files)}", flush=True)
    wb = load_workbook(excel_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    header_row = find_header_row(ws, ["ФИО рус", "Дата оплаты", "Y.T.C № Кимлики"])
    headers = map_headers(ws, header_row)
    indexes = build_excel_indexes(ws, header_row, headers)
    ensure_output_headers(ws, header_row)

    parsed_rows: list[dict] = []
    unmatched_rows: list[dict] = []
    debug_rows: list[dict] = []

    for idx, pdf_path in enumerate(pdf_files, start=1):
        print(f"[{idx}/{len(pdf_files)}] Читаю PDF: {pdf_path.name}", flush=True)
        rec = parse_pdf(pdf_path, use_ocr=use_ocr)
        row, match_mode = find_excel_row(indexes, rec)
        row_info = {
            "pdf_file": rec.filename,
            "text_source": rec.text_source,
            "Страниц PDF": rec.page_count,
            "Полезные страницы": ",".join(map(str, rec.payment_pages or [])),
            "Пропущенные страницы": ",".join(map(str, rec.skipped_pages or [])),
            "Дата": rec.date or "",
            "date_source": rec.date_source,
            "ИНН / Кимлик": rec.inn_kimlik or "",
            "kimlik_source": rec.kimlik_source,
            "Фамилия / Наименование": rec.surname_title or "",
            "Имя": rec.name or "",
            "Сумма": rec.amounts[0] if len(rec.amounts) >= 1 else "",
            "Сумма 2": rec.amounts[1] if len(rec.amounts) >= 2 else "",
            "Все суммы": ";".join(rec.amounts or []),
            "Кол-во сумм": len(rec.amounts or []),
            "match_key_name": rec.full_name,
            "match_key_date": normalize_date(rec.date or ""),
            "match_key_id": normalize_id(rec.inn_kimlik or ""),
            "match_mode": match_mode,
            "excel_row": row or "",
        }
        parsed_rows.append(row_info)
        debug_rows.append({"pdf_file": rec.filename, "raw_top_date_ocr": rec.raw_top_date_ocr, "raw_kimlik_ocr": rec.raw_kimlik_ocr, "payment_pages": row_info["Полезные страницы"], "skipped_pages": row_info["Пропущенные страницы"]})
        if row:
            write_record_to_row(ws, row, rec)
        else:
            unmatched_rows.append(row_info)

    output_excel = excel_path.with_name(f"{excel_path.stem}_RU_decont_filled.xlsx")
    save_csv(excel_path.with_name("RU_decont_parsed_report.csv"), parsed_rows)
    save_csv(excel_path.with_name("RU_decont_unmatched_report.csv"), unmatched_rows)
    save_csv(excel_path.with_name("RU_decont_ocr_debug_report.csv"), debug_rows)
    wb.save(output_excel)

    print("Готово.")
    print(f"PDF файлов обработано: {len(pdf_files)}")
    print(f"Совпадений записано в Excel: {len(pdf_files) - len(unmatched_rows)}")
    print(f"Несопоставленных PDF: {len(unmatched_rows)}")
    print(f"Результат сохранён: {output_excel}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Заполняет Excel данными из русских PDF-платёжек с OCR ячеек")
    parser.add_argument("--excel", required=True, help="Путь к исходному Excel-файлу")
    parser.add_argument("--pdf-dir", default="pdf_result", help="Папка с PDF-файлами")
    parser.add_argument("--sheet", default=None, help="Имя листа Excel, если нужен не активный лист")
    parser.add_argument("--no-ocr", action="store_true", help="Отключить OCR")
    args = parser.parse_args()
    process(Path(args.excel), Path(args.pdf_dir), args.sheet, use_ocr=not args.no_ocr)


if __name__ == "__main__":
    main()
