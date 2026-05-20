#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HR PDF Sorter.

Moves payment PDF files into matching employee folders.

Default mode is dry-run. Use --move to actually move files.
"""

from __future__ import annotations

import argparse
import csv
import re
import shutil
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Literal

DATE_RE = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")
SERVICE_WORDS_RE = re.compile(
    r"\b(?:пп|от|rus|рус|pdf|payment|platezh|платеж|платежка)\b",
    flags=re.IGNORECASE,
)

Status = Literal[
    "DRY_RUN",
    "MOVED",
    "SKIPPED_NOT_FOUND",
    "SKIPPED_AMBIGUOUS",
    "ERROR",
]


@dataclass(frozen=True)
class MatchResult:
    folder: Path | None
    score: float
    extracted_name: str
    second_folder: Path | None = None
    second_score: float = 0.0
    ambiguous: bool = False


@dataclass(frozen=True)
class ProcessResult:
    pdf: Path
    status: Status
    extracted_name: str
    score: float
    target_folder: Path | None
    destination: Path | None
    message: str


def normalize(text: str) -> str:
    """Normalize text for stable comparison."""
    text = text.lower().replace("ё", "е")
    text = text.replace("_", " ").replace("-", " ")
    text = DATE_RE.sub(" ", text)
    text = SERVICE_WORDS_RE.sub(" ", text)
    text = re.sub(r"[^a-zа-я0-9\s]", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_name_from_pdf(pdf_path: Path) -> str:
    """Extract expected employee name from PDF filename."""
    name = pdf_path.stem
    name = name.replace("_", " ").replace("-", " ")
    name = DATE_RE.sub(" ", name)
    name = SERVICE_WORDS_RE.sub(" ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def token_score(folder_norm: str, pdf_norm: str) -> float:
    """Score by common tokens relative to the folder name."""
    folder_tokens = set(folder_norm.split())
    pdf_tokens = set(pdf_norm.split())

    if not folder_tokens or not pdf_tokens:
        return 0.0

    common = folder_tokens & pdf_tokens
    return len(common) / len(folder_tokens)


def combined_score(folder_norm: str, pdf_norm: str) -> float:
    """Combine strict, substring, token and fuzzy checks."""
    if not folder_norm or not pdf_norm:
        return 0.0

    if folder_norm == pdf_norm:
        return 1.0

    if folder_norm in pdf_norm or pdf_norm in folder_norm:
        return 0.98

    seq = similarity(folder_norm, pdf_norm)
    tok = token_score(folder_norm, pdf_norm)

    return max(seq, tok)


def find_best_folder(
    pdf_path: Path,
    folders: list[Path],
    min_score: float,
    ambiguity_gap: float,
) -> MatchResult:
    extracted_name = extract_name_from_pdf(pdf_path)
    pdf_norm = normalize(extracted_name)

    candidates: list[tuple[Path, float]] = []

    for folder in folders:
        folder_norm = normalize(folder.name)
        score = combined_score(folder_norm, pdf_norm)
        candidates.append((folder, score))

    candidates.sort(key=lambda item: item[1], reverse=True)

    if not candidates:
        return MatchResult(None, 0.0, extracted_name)

    best_folder, best_score = candidates[0]
    second_folder: Path | None = None
    second_score = 0.0

    if len(candidates) > 1:
        second_folder, second_score = candidates[1]

    if best_score < min_score:
        return MatchResult(
            None,
            best_score,
            extracted_name,
            second_folder=second_folder,
            second_score=second_score,
        )

    if second_folder is not None and abs(best_score - second_score) < ambiguity_gap:
        return MatchResult(
            best_folder,
            best_score,
            extracted_name,
            second_folder=second_folder,
            second_score=second_score,
            ambiguous=True,
        )

    return MatchResult(
        best_folder,
        best_score,
        extracted_name,
        second_folder=second_folder,
        second_score=second_score,
    )


def unique_destination_path(dest_folder: Path, filename: str) -> Path:
    """Return non-existing destination path by adding numeric suffix when needed."""
    dest = dest_folder / filename

    if not dest.exists():
        return dest

    source_name = Path(filename)
    stem = source_name.stem
    suffix = source_name.suffix

    counter = 2
    while True:
        candidate = dest_folder / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def collect_folders(base_dir: Path, recursive_folders: bool) -> list[Path]:
    if recursive_folders:
        return [p for p in base_dir.rglob("*") if p.is_dir()]
    return [p for p in base_dir.iterdir() if p.is_dir()]


def collect_pdfs(base_dir: Path, recursive_pdfs: bool) -> list[Path]:
    pattern = "**/*.pdf" if recursive_pdfs else "*.pdf"
    return sorted([p for p in base_dir.glob(pattern) if p.is_file()])


def process_pdf(
    pdf: Path,
    folders: list[Path],
    dry_run: bool,
    min_score: float,
    ambiguity_gap: float,
) -> ProcessResult:
    match = find_best_folder(
        pdf_path=pdf,
        folders=folders,
        min_score=min_score,
        ambiguity_gap=ambiguity_gap,
    )

    if match.ambiguous:
        second_name = match.second_folder.name if match.second_folder else ""
        return ProcessResult(
            pdf=pdf,
            status="SKIPPED_AMBIGUOUS",
            extracted_name=match.extracted_name,
            score=match.score,
            target_folder=match.folder,
            destination=None,
            message=f"Ambiguous match. Second candidate: {second_name} ({match.second_score:.3f})",
        )

    if match.folder is None:
        return ProcessResult(
            pdf=pdf,
            status="SKIPPED_NOT_FOUND",
            extracted_name=match.extracted_name,
            score=match.score,
            target_folder=None,
            destination=None,
            message="Suitable folder was not found",
        )

    destination = unique_destination_path(match.folder, pdf.name)

    if dry_run:
        return ProcessResult(
            pdf=pdf,
            status="DRY_RUN",
            extracted_name=match.extracted_name,
            score=match.score,
            target_folder=match.folder,
            destination=destination,
            message="Dry-run only. File was not moved",
        )

    try:
        shutil.move(str(pdf), str(destination))
    except Exception as exc:  # noqa: BLE001 - CLI tool should report all file-operation errors
        return ProcessResult(
            pdf=pdf,
            status="ERROR",
            extracted_name=match.extracted_name,
            score=match.score,
            target_folder=match.folder,
            destination=destination,
            message=str(exc),
        )

    return ProcessResult(
        pdf=pdf,
        status="MOVED",
        extracted_name=match.extracted_name,
        score=match.score,
        target_folder=match.folder,
        destination=destination,
        message="Moved successfully",
    )


def write_report(report_path: Path, results: list[ProcessResult]) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)

    with report_path.open("w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.DictWriter(
            csvfile,
            fieldnames=[
                "pdf",
                "status",
                "extracted_name",
                "score",
                "target_folder",
                "destination",
                "message",
            ],
            delimiter=";",
        )
        writer.writeheader()

        for result in results:
            writer.writerow(
                {
                    "pdf": str(result.pdf),
                    "status": result.status,
                    "extracted_name": result.extracted_name,
                    "score": f"{result.score:.3f}",
                    "target_folder": str(result.target_folder) if result.target_folder else "",
                    "destination": str(result.destination) if result.destination else "",
                    "message": result.message,
                }
            )


def print_result(index: int, total: int, result: ProcessResult) -> None:
    print(f"[{index}/{total}] {result.pdf.name}")
    print(f"  Status: {result.status}")
    print(f"  Extracted name: {result.extracted_name}")
    print(f"  Score: {result.score:.3f}")

    if result.target_folder:
        print(f"  Folder: {result.target_folder.name}")

    if result.destination:
        print(f"  Destination: {result.destination}")

    print(f"  Message: {result.message}")
    print()


def process(
    base_dir: Path,
    dry_run: bool,
    min_score: float,
    ambiguity_gap: float,
    recursive_pdfs: bool,
    recursive_folders: bool,
    report_path: Path | None,
) -> int:
    if not base_dir.exists():
        print(f"[ERROR] Directory not found: {base_dir}")
        return 2

    folders = collect_folders(base_dir, recursive_folders=recursive_folders)
    pdf_files = collect_pdfs(base_dir, recursive_pdfs=recursive_pdfs)

    print(f"Working directory: {base_dir}")
    print(f"Folders found: {len(folders)}")
    print(f"PDF files found: {len(pdf_files)}")
    print(f"Mode: {'DRY-RUN' if dry_run else 'MOVE'}")
    print("-" * 80)

    results: list[ProcessResult] = []

    for index, pdf in enumerate(pdf_files, start=1):
        result = process_pdf(
            pdf=pdf,
            folders=folders,
            dry_run=dry_run,
            min_score=min_score,
            ambiguity_gap=ambiguity_gap,
        )
        results.append(result)
        print_result(index, len(pdf_files), result)

    totals = {
        "DRY_RUN": 0,
        "MOVED": 0,
        "SKIPPED_NOT_FOUND": 0,
        "SKIPPED_AMBIGUOUS": 0,
        "ERROR": 0,
    }

    for result in results:
        totals[result.status] += 1

    print("=" * 80)
    print("SUMMARY")
    print(f"Dry-run matches: {totals['DRY_RUN']}")
    print(f"Moved: {totals['MOVED']}")
    print(f"Skipped, not found: {totals['SKIPPED_NOT_FOUND']}")
    print(f"Skipped, ambiguous: {totals['SKIPPED_AMBIGUOUS']}")
    print(f"Errors: {totals['ERROR']}")

    if report_path:
        write_report(report_path, results)
        print(f"Report: {report_path}")

    if dry_run:
        print("\nDry-run completed. Use --move to actually move files.")

    return 1 if totals["ERROR"] else 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Move payment PDF files into matching employee folders."
    )

    parser.add_argument(
        "--dir",
        default=".",
        help="Working directory with employee folders and PDF files. Default: current directory.",
    )
    parser.add_argument(
        "--move",
        action="store_true",
        help="Actually move files. Without this flag only dry-run is performed.",
    )
    parser.add_argument(
        "--min-score",
        type=float,
        default=0.78,
        help="Minimum match score. Default: 0.78.",
    )
    parser.add_argument(
        "--ambiguity-gap",
        type=float,
        default=0.05,
        help="If best and second scores differ by less than this value, file is skipped. Default: 0.05.",
    )
    parser.add_argument(
        "--recursive-pdfs",
        action="store_true",
        help="Search PDF files recursively.",
    )
    parser.add_argument(
        "--recursive-folders",
        action="store_true",
        help="Search target folders recursively.",
    )
    parser.add_argument(
        "--report",
        default=None,
        help="Optional CSV report path, for example reports/result.csv.",
    )

    return parser.parse_args()


def main() -> int:
    args = parse_args()

    return process(
        base_dir=Path(args.dir).resolve(),
        dry_run=not args.move,
        min_score=args.min_score,
        ambiguity_gap=args.ambiguity_gap,
        recursive_pdfs=args.recursive_pdfs,
        recursive_folders=args.recursive_folders,
        report_path=Path(args.report).resolve() if args.report else None,
    )


if __name__ == "__main__":
    raise SystemExit(main())
