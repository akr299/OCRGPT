# OCRGPT

領収書画像をOCR + OpenAIで構造化し、**保存前にカテゴリを手修正できる** Tkinter GUI アプリです。

## 主な機能
1. 指定フォルダ内の領収書画像を一括読み込み
2. `pytesseract` で OCR 実行
3. OpenAI API で以下の項目を抽出
   - `date`
   - `store`
   - `total`
   - `category`
   - `payment`
   - `note`
4. 結果をテーブル表示（1行=1領収書）
5. 行選択 + ドロップダウンでカテゴリを手修正
6. 修正済みデータを既存 Excel に追記保存（上書きしない）

## 必要環境
- Python 3.10+
- Tesseract OCR（OS側にインストール）
- OpenAI APIキー

依存ライブラリ:
```bash
pip install pytesseract openai openpyxl
```

環境変数:
```bash
export OPENAI_API_KEY="your_api_key"
```

## 使い方（GUI）
```bash
python receipt_to_excel.py
```

画面操作:
1. **画像フォルダ**を選択
2. **保存先Excel**（既存 `.xlsx`）を選択
3. `OCR + OpenAI実行` を押下
4. テーブルの行をクリックし、下部の「カテゴリ修正」コンボボックスでカテゴリ変更
   - 選択後すぐ `ReceiptRecord` に反映されます
5. `Save to Excel` を押下して追記保存

## エラー時の挙動
- APIキー未設定: エラーダイアログ表示
- OCR結果が空/画像処理失敗: 対象ファイルを警告ダイアログで表示
- 日付・金額など不正データ: 対象ファイルを警告ダイアログで表示

## 実行ファイル化（PyInstaller）
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name ocrgpt_gui receipt_to_excel.py
```

生成物:
- `dist/ocrgpt_gui`（Windowsなら `ocrgpt_gui.exe`）

> 配布先PCにも Tesseract OCR のインストールが必要です。

## ファイル構成
- `receipt_to_excel.py` : Tkinter GUI
- `receipt_core.py` : OCR / OpenAI / Excel追記のコアロジック
