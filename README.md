# OCRGPT

領収書画像を OCR + OpenAI で構造化し、**GUIで手修正してからExcel出力できる** Tkinter アプリです。

## 改善ポイント（本バージョン）
- 和暦/西暦の日付入力を受け付け、内部的に `YYYY/MM/DD` に正規化
- OCR失敗や解析失敗を `is_error` / `error_reason` で管理
- GUIテーブルでセルをダブルクリックして直接編集（空白でも可）
- エラー行は赤背景で可視化
- 元画像リンク（`file:///...`）を Excel に出力
- OCR結果 → データモデル → GUI編集 → Excel出力 を分離

## 主な機能
1. 指定フォルダ内の領収書画像を一括読み込み
2. `pytesseract` で OCR 実行
3. OpenAI API で項目抽出（未定義キーも保持）
4. 結果をテーブル表示（1行=1領収書）
5. セル編集 or フォーム編集で値を手修正
6. 正常/エラーを同じExcelに保存（フィルタしやすい形式）

## 必要環境
- Python 3.10+
- Tesseract OCR（OS側にインストール）
- OpenAI APIキー

依存ライブラリ:
```bash
pip install pytesseract openai openpyxl cryptography
```

## 使い方（GUI）
```bash
python receipt_to_excel.py
```

画面操作:
1. **画像フォルダ**を選択
2. **保存先Excel**を選択（新規作成可）
3. `OCR + OpenAI実行` を押下
4. 抽出テーブルを確認
   - ダブルクリックで各セル編集
   - 下部フォームから「日付/店名/金額/その他」をまとめて修正
   - 日付は和暦・西暦どちらも入力可能
5. 必要に応じて `元画像を開く` で画像確認
6. `Save to Excel` を押下

## Excel出力列（例）
- `file_name`
- `date`
- `store`
- `total`
- `category`
- `payment`
- `note`
- `is_error`
- `error_reason`
- `source_image_link`
- その他OCRで取得できた追加項目

出力シートには自動フィルタが設定されます。

## 実行ファイル化（PyInstaller）
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name ocrgpt_gui receipt_to_excel.py
```

## ファイル構成
- `receipt_to_excel.py` : Tkinter GUI（一覧編集・エラー可視化・画像オープン）
- `receipt_core.py` : OCR / OpenAI / データ正規化 / Excel出力
