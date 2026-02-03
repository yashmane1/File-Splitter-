#!/usr/bin/env python3
"""Split large CSV or Excel files into smaller files by row count."""

from __future__ import annotations

import argparse
import csv
import os
from pathlib import Path
from typing import Iterable, List, Optional


SUPPORTED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"}


def _chunked(iterable: Iterable[List[str]], size: int) -> Iterable[List[List[str]]]:
    chunk: List[List[str]] = []
    for row in iterable:
        chunk.append(row)
        if len(chunk) >= size:
            yield chunk
            chunk = []
    if chunk:
        yield chunk


def split_csv(path: Path, rows_per_file: int, output_dir: Path, base_name: str) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    with path.open(newline="", encoding="utf-8") as handle:
        reader = csv.reader(handle)
        try:
            header = next(reader)
        except StopIteration:
            print(f"{path} is empty. Nothing to split.")
            return

        for index, rows in enumerate(_chunked(reader, rows_per_file), start=1):
            output_path = output_dir / f"{base_name}_part{index}.csv"
            with output_path.open("w", newline="", encoding="utf-8") as output_handle:
                writer = csv.writer(output_handle)
                writer.writerow(header)
                writer.writerows(rows)
            print(f"Wrote {output_path}")


def _load_openpyxl():
    try:
        import openpyxl  # type: ignore
    except ImportError as exc:
        raise SystemExit(
            "openpyxl is required for Excel splitting. "
            "Install it with: pip install openpyxl"
        ) from exc
    return openpyxl


def split_excel(
    path: Path,
    rows_per_file: int,
    output_dir: Path,
    base_name: str,
    sheet_name: Optional[str] = None,
) -> None:
    openpyxl = _load_openpyxl()
    output_dir.mkdir(parents=True, exist_ok=True)

    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook[sheet_name] if sheet_name else workbook.active

    rows_iter = sheet.iter_rows(values_only=True)
    try:
        header = next(rows_iter)
    except StopIteration:
        print(f"{path} is empty. Nothing to split.")
        return

    chunk_index = 0
    buffer: List[List[object]] = []
    for row in rows_iter:
        buffer.append(list(row))
        if len(buffer) >= rows_per_file:
            chunk_index += 1
            _write_excel_chunk(openpyxl, header, buffer, output_dir, base_name, chunk_index)
            buffer = []

    if buffer:
        chunk_index += 1
        _write_excel_chunk(openpyxl, header, buffer, output_dir, base_name, chunk_index)



def _write_excel_chunk(
    openpyxl_module,
    header: Iterable[object],
    rows: List[List[object]],
    output_dir: Path,
    base_name: str,
    index: int,
) -> None:
    workbook = openpyxl_module.Workbook()
    sheet = workbook.active
    sheet.append(list(header))
    for row in rows:
        sheet.append(row)
    output_path = output_dir / f"{base_name}_part{index}.xlsx"
    workbook.save(output_path)
    print(f"Wrote {output_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Split CSV or Excel files into smaller files by row count."
    )
    parser.add_argument("input_file", type=Path, help="Path to CSV or Excel file.")
    parser.add_argument(
        "rows_per_file",
        type=int,
        help="Number of rows (excluding header) per output file.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Directory to store split files (default: ./output).",
    )
    parser.add_argument(
        "--sheet",
        type=str,
        default=None,
        help="Excel sheet name (default: active sheet).",
    )
    parser.add_argument(
        "--base-name",
        type=str,
        default=None,
        help="Base name for output files (default: input file stem).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path: Path = args.input_file

    if not input_path.exists():
        raise SystemExit(f"Input file does not exist: {input_path}")
    if args.rows_per_file <= 0:
        raise SystemExit("rows_per_file must be greater than 0")

    base_name = args.base_name or input_path.stem
    extension = input_path.suffix.lower()

    if extension == ".csv":
        split_csv(input_path, args.rows_per_file, args.output_dir, base_name)
    elif extension in SUPPORTED_EXCEL_EXTENSIONS:
        split_excel(
            input_path,
            args.rows_per_file,
            args.output_dir,
            base_name,
            sheet_name=args.sheet,
        )
    else:
        raise SystemExit(
            f"Unsupported file type: {extension}. Use CSV or Excel formats."
        )


if __name__ == "__main__":
    main()
