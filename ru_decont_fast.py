from __future__ import annotations

import argparse
from concurrent.futures import ProcessPoolExecutor, as_completed
from multiprocessing import freeze_support
from pathlib import Path

from openpyxl import load_workbook

import ru_decont as rd


def parse_task(pdf_path_text: str, use_ocr: bool):
    pdf_path = Path(pdf_path_text)
    return pdf_path_text, rd.parse_pdf(pdf_path, use_ocr=use_ocr)


def row_info(rec, row, match_mode: str) -> dict:
    return {
        "pdf_file": rec.filename,
        "text_source": rec.text_source,
        "pages": rec.page_count,
        "payment_pages": ",".join(map(str, rec.payment_pages or [])),
        "skipped_pages": ",".join(map(str, rec.skipped_pages or [])),
        "Дата": rec.date or "",
        "date_source": rec.date_source,
        "ИНН / Кимлик": rec.inn_kimlik or "",
        "kimlik_source": rec.kimlik_source,
        "Фамилия / Наименование": rec.surname_title or "",
        "Имя": rec.name or "",
        "Сумма": rec.amounts[0] if len(rec.amounts) >= 1 else "",
        "Сумма 2": rec.amounts[1] if len(rec.amounts) >= 2 else "",
        "all_amounts": ";".join(rec.amounts or []),
        "amount_count": len(rec.amounts or []),
        "match_key_name": rec.full_name,
        "match_key_date": rd.normalize_date(rec.date or ""),
        "match_key_id": rd.normalize_id(rec.inn_kimlik or ""),
        "match_mode": match_mode,
        "excel_row": row or "",
    }


def parse_pdfs(pdf_files: list[Path], use_ocr: bool, workers: int):
    total = len(pdf_files)
    if workers <= 1 or total <= 1:
        result = {}
        for idx, pdf_path in enumerate(pdf_files, start=1):
            print(f"[{idx}/{total}] Parse: {pdf_path.name}", flush=True)
            result[idx] = (pdf_path, rd.parse_pdf(pdf_path, use_ocr=use_ocr))
        return result

    workers = max(1, min(workers, total))
    print(f"Parallel PDF workers: {workers}", flush=True)
    result = {}
    with ProcessPoolExecutor(max_workers=workers) as executor:
        future_map = {
            executor.submit(parse_task, str(pdf_path), use_ocr): (idx, pdf_path)
            for idx, pdf_path in enumerate(pdf_files, start=1)
        }
        for future in as_completed(future_map):
            idx, pdf_path = future_map[future]
            try:
                _, rec = future.result()
            except Exception as exc:
                raise RuntimeError(f"Failed to parse PDF: {pdf_path}") from exc
            result[idx] = (pdf_path, rec)
            print(f"[{idx}/{total}] Done: {pdf_path.name}", flush=True)
    return result


def process(excel_path: Path, pdf_dir: Path, sheet_name=None, use_ocr: bool = True, workers: int = 3) -> None:
    if use_ocr:
        rd.check_tesseract()
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel not found: {excel_path}")
    if not pdf_dir.exists() or not pdf_dir.is_dir():
        raise FileNotFoundError(f"PDF folder not found: {pdf_dir}")

    pdf_files = rd.unique_pdf_files(pdf_dir)
    if not pdf_files:
        raise FileNotFoundError(f"No PDF files in folder: {pdf_dir}")

    print(f"PDF files found: {len(pdf_files)}", flush=True)
    wb = load_workbook(excel_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    header_row = rd.find_header_row(ws, ["ФИО рус", "Дата оплаты", "Y.T.C № Кимлики"])
    headers = rd.map_headers(ws, header_row)
    indexes = rd.build_excel_indexes(ws, header_row, headers)
    rd.ensure_output_headers(ws, header_row)

    parsed = parse_pdfs(pdf_files, use_ocr=use_ocr, workers=workers)
    parsed_rows = []
    unmatched_rows = []
    debug_rows = []

    for idx in range(1, len(pdf_files) + 1):
        _pdf_path, rec = parsed[idx]
        row, match_mode = rd.find_excel_row(indexes, rec)
        info = row_info(rec, row, match_mode)
        parsed_rows.append(info)
        debug_rows.append({
            "pdf_file": rec.filename,
            "raw_top_date_ocr": rec.raw_top_date_ocr,
            "raw_kimlik_ocr": rec.raw_kimlik_ocr,
            "raw_name_ocr": rec.raw_name_ocr,
            "payment_pages": info["payment_pages"],
            "skipped_pages": info["skipped_pages"],
        })
        if row:
            rd.write_record_to_row(ws, row, rec)
        else:
            unmatched_rows.append(info)

    output_excel = excel_path.with_name(f"{excel_path.stem}_RU_decont_filled.xlsx")
    rd.save_csv(excel_path.with_name("RU_decont_parsed_report.csv"), parsed_rows)
    rd.save_csv(excel_path.with_name("RU_decont_unmatched_report.csv"), unmatched_rows)
    rd.save_csv(excel_path.with_name("RU_decont_ocr_debug_report.csv"), debug_rows)
    wb.save(output_excel)

    print("Done.")
    print(f"Processed PDF files: {len(pdf_files)}")
    print(f"Matched rows: {len(pdf_files) - len(unmatched_rows)}")
    print(f"Unmatched PDF files: {len(unmatched_rows)}")
    print(f"Output Excel: {output_excel}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Parallel RU decont PDF to Excel matcher")
    parser.add_argument("--excel", required=True)
    parser.add_argument("--pdf-dir", default="pdf_result")
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--no-ocr", action="store_true")
    parser.add_argument("--workers", type=int, default=3)
    args = parser.parse_args()
    process(Path(args.excel), Path(args.pdf_dir), args.sheet, use_ocr=not args.no_ocr, workers=max(1, args.workers))


if __name__ == "__main__":
    freeze_support()
    main()
