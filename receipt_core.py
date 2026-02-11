from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

import pytesseract
from openai import OpenAI
from openai import AuthenticationError, RateLimitError
from openpyxl import load_workbook

SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}

JAPANESE_ACCOUNTING_CATEGORIES = [
    "旅費交通費",
    "会議費",
    "接待交際費",
    "消耗品費",
    "通信費",
    "水道光熱費",
    "広告宣伝費",
    "新聞図書費",
    "地代家賃",
    "外注工賃",
    "支払手数料",
    "租税公課",
    "雑費",
]


@dataclass
class ReceiptRecord:
    date: str
    store: str
    total: int
    category: str
    payment: str
    note: str

    def as_excel_row(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ProcessFailure:
    file_name: str
    reason: str


def _extract_json_from_text(text: str) -> Dict[str, Any]:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\s*```$", "", cleaned)

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
    if not match:
        raise ValueError("OpenAI response did not contain a JSON object")
    return json.loads(match.group(0))


def _normalize_record(raw: Dict[str, Any], file_name: str) -> ReceiptRecord:
    required = ["date", "store", "total", "category", "payment"]
    missing = [key for key in required if key not in raw]
    if missing:
        raise ValueError(f"Missing keys: {missing}")

    date_value = str(raw["date"]).strip().replace("/", "-")
    try:
        date_value = datetime.strptime(date_value, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError as exc:
        raise ValueError(f"Invalid date format: {date_value!r} (expected YYYY-MM-DD)") from exc

    total_raw = raw["total"]
    if isinstance(total_raw, int):
        total_value = total_raw
    else:
        digits = re.sub(r"[^0-9-]", "", str(total_raw))
        if digits in {"", "-"}:
            raise ValueError(f"Invalid total value: {total_raw!r}")
        total_value = int(digits)

    store_value = str(raw["store"]).strip()
    category_value = str(raw["category"]).strip() or "雑費"
    payment_value = str(raw["payment"]).strip() or "不明"
    note_value = str(raw.get("note", "")).strip() or f"OCR元: {file_name}"

    return ReceiptRecord(
        date=date_value,
        store=store_value,
        total=total_value,
        category=category_value,
        payment=payment_value,
        note=note_value,
    )


def collect_image_files(input_folder: Path) -> List[Path]:
    return sorted(
        path
        for path in input_folder.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    )


class ReceiptProcessor:
    def __init__(
        self,
        api_key: str,
        model: str = "gpt-4.1-mini",
        ocr_lang: str = "jpn+eng",
        tesseract_cmd: str | None = None,
    ):
        self.api_key = api_key
        self.model = model
        self.ocr_lang = ocr_lang
        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    def _ocr_image(self, image_path: Path) -> str:
        text = pytesseract.image_to_string(str(image_path), lang=self.ocr_lang)
        if not text.strip():
            raise ValueError("OCR結果が空です")
        return text

    def _call_openai(self, client: OpenAI, ocr_text: str, file_name: str) -> ReceiptRecord:
        system_prompt = (
            "You extract data from Japanese receipts. "
            "Return ONLY one JSON object with keys: date, store, total, category, payment, note."
        )
        user_prompt = (
            "OCR text から経費精算向けデータを抽出してください。\n"
            "Rules:\n"
            "- date は YYYY-MM-DD\n"
            "- total は整数(円)\n"
            "- category は日本語の会計カテゴリ\n"
            "- payment は支払方法\n"
            "- note は短い補足(なければ空文字)\n"
            f"OCR:\n{ocr_text}"
        )

        response = client.chat.completions.create(
            model=self.model,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        )
        content = response.choices[0].message.content or ""
        raw_record = _extract_json_from_text(content)
        return _normalize_record(raw_record, file_name=file_name)

    def process_folder(self, input_folder: Path) -> Tuple[List[Tuple[str, ReceiptRecord]], List[ProcessFailure]]:
        if not input_folder.exists() or not input_folder.is_dir():
            raise FileNotFoundError(f"画像フォルダが見つかりません: {input_folder}")

        image_files = collect_image_files(input_folder)
        if not image_files:
            raise FileNotFoundError("対応画像ファイルがありません。")

        if not self.api_key.strip():
            raise RuntimeError("OpenAI APIキーが設定されていません。設定から再登録してください。")

        client = OpenAI(api_key=self.api_key)

        records: List[Tuple[str, ReceiptRecord]] = []
        failures: List[ProcessFailure] = []

        for image_path in image_files:
            try:
                ocr_text = self._ocr_image(image_path)
                record = self._call_openai(client, ocr_text, file_name=image_path.name)
                if not record.note:
                    record.note = f"OCR元: {image_path.name}"
                records.append((image_path.name, record))
            except AuthenticationError:
                raise RuntimeError("OpenAI APIキー認証に失敗しました。設定からAPIキーを再登録してください。")
            except RateLimitError:
                raise RuntimeError("OpenAIの利用上限に達しました。課金状況を確認し、必要ならAPIキーを再設定してください。")
            except Exception as exc:
                failures.append(ProcessFailure(file_name=image_path.name, reason=str(exc)))

        return records, failures


def append_records_to_excel(excel_path: Path, records: Sequence[ReceiptRecord], sheet_name: str | None = None) -> int:
    if not excel_path.exists():
        raise FileNotFoundError(f"Excelファイルが見つかりません: {excel_path}")

    if not records:
        raise ValueError("保存対象データがありません")

    wb = load_workbook(excel_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    required_columns = ["date", "store", "total", "category", "payment", "note"]
    header_row = [ws.cell(row=1, column=i).value for i in range(1, ws.max_column + 1)]
    header_map = {str(v).strip(): idx + 1 for idx, v in enumerate(header_row) if v is not None}
    has_header = all(col in header_map for col in required_columns)

    appended = 0
    for record in records:
        row_dict = record.as_excel_row()
        if has_header:
            row_idx = ws.max_row + 1
            for col in required_columns:
                ws.cell(row=row_idx, column=header_map[col], value=row_dict[col])
        else:
            ws.append([row_dict[col] for col in required_columns])
        appended += 1

    wb.save(excel_path)
    return appended
