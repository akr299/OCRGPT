# OCRGPT

Tkinter desktop application for receipt OCR + OpenAI extraction + Excel append.

## Features
- Select a folder of receipt images (`png/jpg/jpeg/bmp/tif/tiff/webp`)
- Run OCR with `pytesseract`
- Extract structured JSON with OpenAI:
  - `store` (string)
  - `date` (`YYYY-MM-DD`)
  - `total` (integer)
  - `tax8` (integer)
  - `tax10` (integer)
  - `payment` (string)
  - `category` (string)
- Validate fields and append to an existing Excel file
- Progress indicator + scrollable logs

## Requirements
```bash
pip install pytesseract openai openpyxl pandas
```

### External dependency (required)
Install **Tesseract OCR engine** on your OS and ensure `tesseract` is available in PATH.
If it is not in PATH, set the path in the GUI field: `Tesseract Path (optional)`.

## Run the GUI
```bash
export OPENAI_API_KEY="your_api_key"
python receipt_to_excel.py
```

## Build standalone executable (PyInstaller)
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name ReceiptOCRApp receipt_to_excel.py
```

Generated executable is typically in:
- `dist/ReceiptOCRApp` (macOS/Linux)
- `dist/ReceiptOCRApp.exe` (Windows)

## Notes for distribution
- End users still need Tesseract OCR installed on their machine.
- Provide an existing Excel template file (e.g., `expense.xlsx`) before running.
