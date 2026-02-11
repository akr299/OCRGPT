# OCRGPT

Batch receipt OCR + OpenAI extraction script.

## Script
- `receipt_to_excel.py`

## What it does
1. Reads receipt images from an input folder.
2. Runs OCR with `pytesseract`.
3. Sends OCR text to OpenAI and requests strict JSON fields:
   - `store` (string)
   - `date` (`YYYY-MM-DD`)
   - `total` (integer)
   - `tax8` (integer)
   - `tax10` (integer)
   - `payment` (string)
   - `category` (string)
4. Validates keys and types.
5. Appends rows to an existing `expense.xlsx` template.
6. Prints a processing summary.

## Requirements
```bash
pip install pytesseract openai pandas openpyxl
```
Also install Tesseract OCR engine on your OS.

## Usage
```bash
export OPENAI_API_KEY="your_api_key"
python receipt_to_excel.py \
  --input-folder ./receipts \
  --excel-template ./expense.xlsx
```

Optional flags:
- `--model gpt-4.1-mini`
- `--ocr-lang jpn+eng`
- `--tesseract-cmd /usr/bin/tesseract`
- `--sheet-name Sheet1`
