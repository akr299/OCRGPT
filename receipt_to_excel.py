#!/usr/bin/env python3
"""
Batch process receipt images:
1) OCR each image with pytesseract
2) Ask OpenAI to extract structured fields
3) Validate JSON schema and types
4) Append results into an existing Excel template (expense.xlsx)

Usage:
    python receipt_to_excel.py --input-folder ./receipts --excel-template ./expense.xlsx

Environment:
    export OPENAI_API_KEY="..."

Notes:
- Tesseract engine must be installed on your machine.
- If needed, set --tesseract-cmd to the tesseract executable path.
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
import pytesseract
from openai import OpenAI
from openpyxl import load_workbook

# Expected output schema from the model
REQUIRED_SCHEMA = {
    "store": str,
    "date": str,      # YYYY-MM-DD
    "total": int,
    "tax8": int,
    "tax10": int,
    "payment": str,
    "category": str,
}

SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}


@dataclass
class ProcessResult:
    """Processing result for a single image file."""

    file_name: str
    ok: bool
    data: Dict[str, Any] | None = None
    error: str | None = None


def extract_json_from_text(text: str) -> Dict[str, Any]:
    """
    Extract JSON object from model output.

    Handles either pure JSON response or JSON inside markdown code fences.
    """
    cleaned = text.strip()

    # Remove markdown code fences, e.g. ```json ... ```
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\s*```$", "", cleaned)

    # Try direct parse first
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # Fallback: find first JSON object in text
    match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
    if not match:
        raise ValueError("Model response did not contain a JSON object.")

    return json.loads(match.group(0))


def validate_record(record: Dict[str, Any]) -> Dict[str, Any]:
    """
    Validate required keys and expected types.

    Also validates date format YYYY-MM-DD and normalizes integer-like strings.
    Returns normalized record if valid.
    """
    missing = [k for k in REQUIRED_SCHEMA if k not in record]
    if missing:
        raise ValueError(f"Missing keys: {missing}")

    normalized: Dict[str, Any] = {}

    for key, expected_type in REQUIRED_SCHEMA.items():
        value = record[key]

        if expected_type is int:
            # Accept numeric strings (e.g., "14619", "14,619", "¥14,619") and normalize to int
            if isinstance(value, int):
                normalized[key] = value
            elif isinstance(value, str):
                digits = re.sub(r"[^0-9-]", "", value)
                if digits == "" or digits == "-":
                    raise ValueError(f"Key '{key}' is not a valid integer string: {value!r}")
                normalized[key] = int(digits)
            else:
                raise ValueError(f"Key '{key}' expected int, got {type(value).__name__}")
        elif expected_type is str:
            if not isinstance(value, str):
                raise ValueError(f"Key '{key}' expected str, got {type(value).__name__}")
            normalized[key] = value.strip()
        else:
            raise ValueError(f"Unsupported schema type for key '{key}'")

    # Validate date format strictly
    try:
        datetime.strptime(normalized["date"], "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError(f"Invalid date format for 'date': {normalized['date']!r}; expected YYYY-MM-DD") from exc

    return normalized


def ocr_image(image_path: Path, lang: str = "jpn+eng") -> str:
    """Run OCR on one image file and return extracted text."""
    return pytesseract.image_to_string(str(image_path), lang=lang)


def call_openai_extract(client: OpenAI, model: str, ocr_text: str) -> Dict[str, Any]:
    """
    Ask OpenAI to extract receipt fields as strict JSON object.

    Prompt is explicit to improve consistency and reduce extra text.
    """
    system_prompt = (
        "You extract structured receipt data. "
        "Return ONLY one JSON object with exactly these keys: "
        "store (string), date (YYYY-MM-DD string), total (integer), tax8 (integer), "
        "tax10 (integer), payment (string), category (string). "
        "Do not include markdown. Do not include explanations."
    )

    user_prompt = (
        "Extract data from the OCR text below.\n"
        "Rules:\n"
        "- date must be formatted as YYYY-MM-DD.\n"
        "- total, tax8, tax10 must be integers in yen without commas/symbols.\n"
        "- if tax8 or tax10 is not found, set it to 0.\n"
        "- payment examples: cash, credit_card, ic_card, quickpay, mobile_pay.\n"
        "- category should be a concise spending category (e.g., food, transport, office, shopping).\n\n"
        f"OCR text:\n{ocr_text}"
    )

    response = client.chat.completions.create(
        model=model,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )

    content = response.choices[0].message.content or ""
    parsed = extract_json_from_text(content)
    return validate_record(parsed)


def append_to_excel(excel_path: Path, records: List[Dict[str, Any]], sheet_name: str | None = None) -> Tuple[str, int]:
    """
    Append records into existing Excel file.

    If first row contains headers matching all required keys, append values by header order.
    Otherwise append values in fixed schema order.
    Returns tuple: (used_sheet_name, appended_count)
    """
    wb = load_workbook(excel_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    required_keys = list(REQUIRED_SCHEMA.keys())

    # Try to detect header row and map by names
    header_row = [ws.cell(row=1, column=i).value for i in range(1, ws.max_column + 1)]
    header_map = {str(v).strip(): idx + 1 for idx, v in enumerate(header_row) if v is not None}
    has_all_headers = all(key in header_map for key in required_keys)

    appended = 0
    for rec in records:
        if has_all_headers:
            next_row = ws.max_row + 1
            for key in required_keys:
                ws.cell(row=next_row, column=header_map[key], value=rec[key])
        else:
            ws.append([rec[key] for key in required_keys])
        appended += 1

    wb.save(excel_path)
    return ws.title, appended


def collect_image_files(input_folder: Path) -> List[Path]:
    """Collect supported image files from input folder (non-recursive)."""
    return sorted(
        p for p in input_folder.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def process_folder(
    input_folder: Path,
    excel_template: Path,
    model: str,
    ocr_lang: str,
    tesseract_cmd: str | None,
    sheet_name: str | None,
) -> None:
    """End-to-end pipeline for OCR -> OpenAI extraction -> Excel append."""
    if not input_folder.exists() or not input_folder.is_dir():
        raise FileNotFoundError(f"Input folder not found: {input_folder}")

    if not excel_template.exists():
        raise FileNotFoundError(f"Excel template not found: {excel_template}")

    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    image_files = collect_image_files(input_folder)
    if not image_files:
        print(f"No image files found in: {input_folder}")
        return

    client = OpenAI()

    results: List[ProcessResult] = []
    valid_records: List[Dict[str, Any]] = []

    for image_path in image_files:
        try:
            text = ocr_image(image_path, lang=ocr_lang)
            record = call_openai_extract(client=client, model=model, ocr_text=text)
            results.append(ProcessResult(file_name=image_path.name, ok=True, data=record))
            valid_records.append(record)
            print(f"[OK] {image_path.name}: {record}")
        except Exception as exc:
            results.append(ProcessResult(file_name=image_path.name, ok=False, error=str(exc)))
            print(f"[NG] {image_path.name}: {exc}")

    if valid_records:
        used_sheet, n = append_to_excel(excel_template, valid_records, sheet_name=sheet_name)
        print(f"\nAppended {n} rows to '{excel_template}' (sheet: '{used_sheet}').")
    else:
        print("\nNo valid records to append.")

    # pandas usage: create a compact summary report
    df = pd.DataFrame([r.data for r in results if r.ok and r.data])
    ok_count = sum(r.ok for r in results)
    ng_count = len(results) - ok_count

    print("\n=== Summary ===")
    print(f"Total images: {len(results)}")
    print(f"Succeeded  : {ok_count}")
    print(f"Failed     : {ng_count}")

    if not df.empty:
        print("\nBy category:")
        print(df.groupby("category")["total"].agg(["count", "sum"]).reset_index())

    if ng_count > 0:
        print("\nFailed files:")
        for r in results:
            if not r.ok:
                print(f"- {r.file_name}: {r.error}")


def parse_args() -> argparse.Namespace:
    """Parse command-line options."""
    parser = argparse.ArgumentParser(description="OCR receipts and append structured data to an Excel template.")
    parser.add_argument("--input-folder", required=True, help="Folder containing receipt images")
    parser.add_argument("--excel-template", default="expense.xlsx", help="Path to existing Excel file")
    parser.add_argument("--model", default="gpt-4.1-mini", help="OpenAI model name")
    parser.add_argument("--ocr-lang", default="jpn+eng", help="Tesseract language(s), e.g. 'jpn+eng'")
    parser.add_argument("--tesseract-cmd", default=None, help="Optional full path to tesseract executable")
    parser.add_argument("--sheet-name", default=None, help="Optional sheet name in Excel workbook")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    process_folder(
        input_folder=Path(args.input_folder),
        excel_template=Path(args.excel_template),
        model=args.model,
        ocr_lang=args.ocr_lang,
        tesseract_cmd=args.tesseract_cmd,
        sheet_name=args.sheet_name,
    )


if __name__ == "__main__":
    main()
