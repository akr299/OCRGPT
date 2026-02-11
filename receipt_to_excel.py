#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import platform
import subprocess
import threading
import tkinter as tk
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from queue import Empty, Queue
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Dict, List

from api_key_manager import ApiKeyManager
from receipt_core import (
    EXCEL_COLUMNS,
    JAPANESE_ACCOUNTING_CATEGORIES,
    ReceiptProcessor,
    ReceiptRecord,
    append_records_to_excel,
    calculate_expense_amount,
    infer_tax_category,
    normalize_date_value,
)


@dataclass
class AppPathsConfig:
    input_folder: str = ""
    excel_path: str = ""


class AppConfigStore:
    """入力/出力パスをOS標準の設定ディレクトリへ保存する。"""

    def __init__(self, config_dir: Path | None = None):
        self.api_key_manager = ApiKeyManager(config_dir=config_dir)
        self.config_dir = self.api_key_manager.config_dir
        self.config_file = self.config_dir / "app_paths.json"

    def load(self) -> AppPathsConfig:
        if not self.config_file.exists():
            return AppPathsConfig()
        try:
            payload = json.loads(self.config_file.read_text(encoding="utf-8"))
            return AppPathsConfig(
                input_folder=str(payload.get("input_folder", "")),
                excel_path=str(payload.get("excel_path", "")),
            )
        except Exception:
            return AppPathsConfig()

    def save(self, config: AppPathsConfig) -> None:
        self.config_dir.mkdir(parents=True, exist_ok=True)
        payload = {"input_folder": config.input_folder, "excel_path": config.excel_path}
        self.config_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


class ReceiptApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OCRGPT 領収書OCR")
        self.root.geometry("1240x720")

        self.records: List[ReceiptRecord] = []
        self.tree_columns: List[str] = []
        self.selected_item_id: str | None = None
        self.result_queue: Queue = Queue()

        self.config_store = AppConfigStore()
        self.api_key_manager = self.config_store.api_key_manager
        self.api_key: str | None = None

        paths = self.config_store.load()
        self.input_folder_var = tk.StringVar(value=paths.input_folder)
        self.excel_path_var = tk.StringVar(value=paths.excel_path)
        self.model_var = tk.StringVar(value="gpt-4.1-mini")
        self.ocr_lang_var = tk.StringVar(value="jpn+eng")
        self.sheet_name_var = tk.StringVar()
        self.api_key_status_var = tk.StringVar(value="API Key: 未設定")
        self.status_var = tk.StringVar(value="準備完了")

        self.form_vars: Dict[str, tk.StringVar] = {
            "日付": tk.StringVar(),
            "勘定科目": tk.StringVar(),
            "内容": tk.StringVar(),
            "支払先": tk.StringVar(),
            "支払方法": tk.StringVar(),
            "金額(税込)": tk.StringVar(),
            "税区分": tk.StringVar(),
            "事業利用割合(%)": tk.StringVar(),
            "経費計上額": tk.StringVar(),
            "備考": tk.StringVar(),
            "is_error": tk.StringVar(),
            "error_reason": tk.StringVar(),
        }

        self._build_ui()
        self._load_saved_api_key_or_prompt()
        self.root.after(200, self._poll_worker)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        config_frame = ttk.LabelFrame(self.root, text="設定")
        config_frame.pack(fill="x", padx=12, pady=8)

        ttk.Label(config_frame, text="画像フォルダ").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(config_frame, textvariable=self.input_folder_var, width=85).grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(config_frame, text="参照", command=self._pick_input_folder).grid(row=0, column=2, padx=4, pady=4)

        ttk.Label(config_frame, text="保存先Excel").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(config_frame, textvariable=self.excel_path_var, width=85).grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(config_frame, text="参照", command=self._pick_excel_path).grid(row=1, column=2, padx=4, pady=4)

        ttk.Label(config_frame, text="OpenAIモデル").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(config_frame, textvariable=self.model_var, width=30).grid(row=2, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(config_frame, text="OCR言語").grid(row=2, column=1, sticky="e", padx=4, pady=4)
        ttk.Entry(config_frame, textvariable=self.ocr_lang_var, width=12).grid(row=2, column=2, sticky="w", padx=4, pady=4)

        ttk.Label(config_frame, text="Sheet名(任意)").grid(row=3, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(config_frame, textvariable=self.sheet_name_var, width=30).grid(row=3, column=1, sticky="w", padx=4, pady=4)

        config_frame.columnconfigure(1, weight=1)

        action_frame = ttk.Frame(self.root)
        action_frame.pack(fill="x", padx=12, pady=6)

        self.process_button = ttk.Button(action_frame, text="OCR + OpenAI実行", command=self._start_processing)
        self.process_button.pack(side="left", padx=4)

        self.save_button = ttk.Button(action_frame, text="Save to Excel", command=self._save_to_excel, state="disabled")
        self.save_button.pack(side="left", padx=4)

        ttk.Button(action_frame, text="元画像を開く", command=self._open_selected_image).pack(side="left", padx=4)
        self.settings_button = ttk.Button(action_frame, text="設定 / APIキー", command=self._open_settings_dialog)
        self.settings_button.pack(side="left", padx=4)

        ttk.Label(action_frame, textvariable=self.api_key_status_var).pack(side="left", padx=16)
        ttk.Label(action_frame, textvariable=self.status_var).pack(side="left", padx=20)

        table_frame = ttk.LabelFrame(self.root, text="抽出結果（ダブルクリックでセル編集、エラー行は赤表示）")
        table_frame.pack(fill="both", expand=True, padx=12, pady=8)

        self.tree = ttk.Treeview(table_frame, show="headings", height=14)
        self.tree.tag_configure("error", background="#ffb3b3")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.pack(side="top", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")
        x_scroll.pack(side="bottom", fill="x")

        self.tree.bind("<<TreeviewSelect>>", self._on_row_selected)
        self.tree.bind("<Double-1>", self._on_double_click_cell)

        edit_frame = ttk.LabelFrame(self.root, text="選択レコードの手動修正")
        edit_frame.pack(fill="x", padx=12, pady=8)

        self._make_form_entry(edit_frame, "日付", "日付", 0, hint="和暦/西暦どちらでも可")
        self._make_form_entry(edit_frame, "支払先", "支払先", 1)
        self._make_form_entry(edit_frame, "勘定科目", "勘定科目", 2, hint="例: " + " / ".join(JAPANESE_ACCOUNTING_CATEGORIES[:5]))

        self._make_form_entry(edit_frame, "内容", "内容", 0, col_offset=3)
        self._make_form_entry(edit_frame, "支払方法", "支払方法", 1, col_offset=3)
        self._make_form_entry(edit_frame, "金額(税込)", "金額(税込)", 2, col_offset=3)

        self._make_form_entry(edit_frame, "税区分", "税区分", 0, col_offset=6, hint="8% / 10%")
        self._make_form_entry(edit_frame, "事業利用割合(%)", "事業利用割合(%)", 1, col_offset=6)
        self._make_form_entry(edit_frame, "経費計上額", "経費計上額", 2, col_offset=6)
        self._make_form_entry(edit_frame, "備考", "備考", 0, col_offset=9)
        self._make_form_entry(edit_frame, "is_error", "is_error", 1, col_offset=9)
        self._make_form_entry(edit_frame, "エラー理由", "error_reason", 2, col_offset=9)

        ttk.Button(edit_frame, text="選択行へ反映", command=self._apply_form_to_selected).grid(row=3, column=11, padx=8, pady=4, sticky="e")

        ttk.Label(edit_frame, text="※ 空白入力でも保存可能。人が判断して修正する前提です。", foreground="#444").grid(
            row=3, column=0, columnspan=12, sticky="w", padx=8, pady=4
        )

    def _make_form_entry(self, parent: ttk.LabelFrame, label: str, key: str, row: int, col_offset: int = 0, hint: str = "") -> None:
        c = col_offset
        ttk.Label(parent, text=label).grid(row=row, column=c, sticky="w", padx=8, pady=4)
        entry = ttk.Entry(parent, textvariable=self.form_vars[key], width=22)
        entry.grid(row=row, column=c + 1, sticky="w", padx=8, pady=4)
        if hint:
            ttk.Label(parent, text=hint, foreground="#666").grid(row=row, column=c + 2, sticky="w", padx=2)

    def _update_api_key_status(self) -> None:
        self.api_key_status_var.set("API Key: 設定済み" if self.api_key else "API Key: 未設定")

    def _prompt_api_key(self, required: bool = False) -> bool:
        while True:
            title = "初期設定" if required else "APIキー設定"
            prompt = "OpenAI APIキーを入力してください" + ("（必須）" if required else "")
            value = simpledialog.askstring(title, prompt, show="*", parent=self.root)

            if value is None:
                if required:
                    retry = messagebox.askretrycancel("APIキー必須", "APIキーが未設定です。入力を続けますか？")
                    if retry:
                        continue
                    self.root.destroy()
                    return False
                return False

            value = value.strip()
            if not value:
                if required:
                    messagebox.showwarning("入力不足", "APIキーを入力してください。")
                    continue
                return False

            self.api_key_manager.save_api_key(value)
            self.api_key = value
            self._update_api_key_status()
            messagebox.showinfo("保存完了", "APIキーを保存しました。")
            return True

    def _load_saved_api_key_or_prompt(self) -> None:
        self.api_key = self.api_key_manager.load_api_key()
        self._update_api_key_status()
        if not self.api_key:
            self._prompt_api_key(required=True)

    def _open_settings_dialog(self) -> None:
        settings = tk.Toplevel(self.root)
        settings.title("設定 / APIキー")
        settings.geometry("500x200")
        settings.transient(self.root)
        settings.grab_set()

        ttk.Label(settings, text="OpenAI APIキー").pack(anchor="w", padx=12, pady=(12, 4))
        key_var = tk.StringVar(value=self.api_key or "")
        ttk.Entry(settings, textvariable=key_var, width=65, show="*").pack(anchor="w", padx=12)

        def save_key() -> None:
            key = key_var.get().strip()
            if not key:
                messagebox.showwarning("入力不足", "APIキーを入力してください。", parent=settings)
                return
            self.api_key_manager.save_api_key(key)
            self.api_key = key
            self._update_api_key_status()
            messagebox.showinfo("保存完了", "APIキーを保存しました。", parent=settings)

        def delete_key() -> None:
            self.api_key_manager.delete_api_key()
            self.api_key = None
            self._update_api_key_status()
            messagebox.showinfo("削除完了", "保存済みAPIキーを削除しました。", parent=settings)
            settings.destroy()

        button_row = ttk.Frame(settings)
        button_row.pack(fill="x", padx=12, pady=(16, 12))
        ttk.Button(button_row, text="保存", command=save_key).pack(side="left", padx=4)
        ttk.Button(button_row, text="削除", command=delete_key).pack(side="left", padx=4)
        ttk.Button(button_row, text="閉じる", command=settings.destroy).pack(side="right", padx=4)

    def _pick_input_folder(self) -> None:
        folder = filedialog.askdirectory(title="領収書画像フォルダを選択")
        if folder:
            self.input_folder_var.set(folder)

    def _pick_excel_path(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存先Excelファイルを選択",
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")],
        )
        if path:
            self.excel_path_var.set(path)

    def _set_processing_state(self, running: bool) -> None:
        self.process_button.config(state="disabled" if running else "normal")
        self.save_button.config(state="disabled" if running or not self.records else "normal")

    def _start_processing(self) -> None:
        if not self.api_key:
            messagebox.showerror("APIキー未設定", "設定ボタンからOpenAI APIキーを登録してください。")
            return

        input_folder = self.input_folder_var.get().strip()
        if not input_folder:
            messagebox.showerror("入力不足", "画像フォルダを指定してください。")
            return

        self._set_processing_state(True)
        self.status_var.set("OCR + OpenAI 処理中...")
        self.tree.delete(*self.tree.get_children())
        self.records.clear()

        threading.Thread(target=self._worker_process, daemon=True).start()

    def _worker_process(self) -> None:
        try:
            processor = ReceiptProcessor(
                api_key=self.api_key,
                model=self.model_var.get().strip() or "gpt-4.1-mini",
                ocr_lang=self.ocr_lang_var.get().strip() or "jpn+eng",
            )
            records, failures = processor.process_folder(Path(self.input_folder_var.get().strip()))
            self.result_queue.put(("ok", records, failures))
        except Exception as exc:
            self.result_queue.put(("error", str(exc)))

    def _poll_worker(self) -> None:
        try:
            result = self.result_queue.get_nowait()
            if result[0] == "ok":
                _, records, failures = result
                self.records = records
                self._render_table()
                self._set_processing_state(False)
                error_count = len([r for r in records if r.is_error])
                self.status_var.set(f"処理完了: 全 {len(records)}件 / エラー {error_count}件")
                if failures:
                    detail = "\n".join(f"{f.file_name}: {f.reason}" for f in failures)
                    messagebox.showwarning("一部失敗", f"以下のファイルでOCR/解析に失敗しました:\n\n{detail}")
            else:
                _, error = result
                self._set_processing_state(False)
                self.status_var.set("エラー")
                messagebox.showerror("処理エラー", str(error))
        except Empty:
            pass
        finally:
            self.root.after(200, self._poll_worker)

    def _derive_columns(self) -> List[str]:
        base = ["file_name", *EXCEL_COLUMNS, "source_image_link"]
        extras = sorted({k for r in self.records for k in r.fields.keys() if k not in base})
        return base + extras

    def _render_table(self) -> None:
        self.tree.delete(*self.tree.get_children())

        self.tree_columns = self._derive_columns()
        self.tree["columns"] = self.tree_columns

        for col in self.tree_columns:
            self.tree.heading(col, text=col)
            width = 140 if col not in {"備考", "error_reason", "source_image_link"} else 260
            self.tree.column(col, width=width, anchor="w")

        for index, record in enumerate(self.records):
            row = []
            for col in self.tree_columns:
                if col == "file_name":
                    row.append(record.file_name)
                elif col == "is_error":
                    row.append("1" if record.is_error else "0")
                elif col == "error_reason":
                    row.append(record.error_reason)
                elif col == "source_image_link":
                    row.append(record.source_image_link())
                else:
                    row.append(record.get(col, ""))

            tags = ("error",) if record.is_error else ()
            self.tree.insert("", "end", iid=str(index), values=row, tags=tags)

    def _on_row_selected(self, _event: object) -> None:
        selected = self.tree.selection()
        if not selected:
            self.selected_item_id = None
            return

        self.selected_item_id = selected[0]
        rec = self.records[int(self.selected_item_id)]
        for key in self.form_vars:
            if key == "error_reason":
                self.form_vars[key].set(rec.error_reason)
            elif key == "is_error":
                self.form_vars[key].set("1" if rec.is_error else "0")
            else:
                self.form_vars[key].set(str(rec.get(key, "")))

    def _apply_record_rules(self, rec: ReceiptRecord) -> None:
        normalized_date, date_error = normalize_date_value(rec.get("日付", ""))
        rec.set("日付", normalized_date)

        tax_value, tax_error = infer_tax_category(rec.get("税区分", ""))
        rec.set("税区分", tax_value)

        if not rec.get("経費計上額", ""):
            rec.set("経費計上額", calculate_expense_amount(rec.get("金額(税込)", ""), rec.get("事業利用割合(%)", "")))

        reasons = []
        if not rec.get("支払先", ""):
            reasons.append("store_not_found")
        if not rec.get("金額(税込)", ""):
            reasons.append("amount_not_found")
        if date_error:
            reasons.append(date_error)
        if tax_error:
            reasons.append(tax_error)

        manual_reasons = [item for item in self.form_vars["error_reason"].get().split("|") if item.strip()] if self.selected_item_id else []
        merged = []
        for reason in [*reasons, *manual_reasons]:
            if reason and reason not in merged:
                merged.append(reason)

        rec.error_reason = "|".join(merged)
        manual_is_error = self.form_vars["is_error"].get().strip().lower() in {"1", "true", "t", "yes", "y"}
        rec.is_error = manual_is_error or bool(rec.error_reason)

    def _apply_form_to_selected(self) -> None:
        if self.selected_item_id is None:
            return

        rec = self.records[int(self.selected_item_id)]
        for key in EXCEL_COLUMNS:
            if key in {"is_error", "error_reason"}:
                continue
            rec.set(key, self.form_vars[key].get().strip())

        self._apply_record_rules(rec)

        self._render_table()
        self.tree.selection_set(self.selected_item_id)

    def _on_double_click_cell(self, event: tk.Event) -> None:
        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item_id or not column_id:
            return

        col_idx = int(column_id.replace("#", "")) - 1
        if col_idx < 0 or col_idx >= len(self.tree_columns):
            return

        key = self.tree_columns[col_idx]
        if key in {"file_name", "source_image_link"}:
            return

        bbox = self.tree.bbox(item_id, column_id)
        if not bbox:
            return

        x, y, width, height = bbox
        current_value = self.tree.item(item_id, "values")[col_idx]
        editor = ttk.Entry(self.tree)
        editor.place(x=x, y=y, width=width, height=height)
        editor.insert(0, current_value)
        editor.focus()

        def commit_edit(_event: object | None = None) -> None:
            new_value = editor.get().strip()
            editor.destroy()

            row_idx = int(item_id)
            rec = self.records[row_idx]

            if key == "error_reason":
                rec.error_reason = new_value
            elif key == "is_error":
                rec.is_error = new_value.lower() in {"1", "true", "t", "yes", "y"}
            else:
                rec.set(key, new_value)

            self.selected_item_id = item_id
            self.form_vars["error_reason"].set(rec.error_reason)
            self.form_vars["is_error"].set("1" if rec.is_error else "0")
            self._apply_record_rules(rec)

            self._render_table()
            self.tree.selection_set(item_id)

        editor.bind("<Return>", commit_edit)
        editor.bind("<FocusOut>", commit_edit)
        editor.bind("<Escape>", lambda _e: editor.destroy())

    def _open_selected_image(self) -> None:
        if self.selected_item_id is None:
            messagebox.showinfo("行未選択", "画像を開く行を選択してください。")
            return

        rec = self.records[int(self.selected_item_id)]
        image_path = Path(rec.source_image_path)
        if not image_path.exists():
            messagebox.showwarning("画像なし", f"元画像が見つかりません: {image_path}")
            return

        try:
            if platform.system() == "Windows":
                os.startfile(str(image_path))  # type: ignore[attr-defined]
            elif platform.system() == "Darwin":
                subprocess.run(["open", str(image_path)], check=False)
            else:
                subprocess.run(["xdg-open", str(image_path)], check=False)
        except Exception:
            webbrowser.open(rec.source_image_link())

    def _save_to_excel(self) -> None:
        excel_path = self.excel_path_var.get().strip()
        if not excel_path:
            messagebox.showerror("入力不足", "保存先Excelを指定してください。")
            return

        if not self.records:
            messagebox.showerror("保存不可", "保存対象データがありません。")
            return

        try:
            sheet_name = self.sheet_name_var.get().strip() or None
            appended = append_records_to_excel(Path(excel_path), self.records, sheet_name=sheet_name)
            messagebox.showinfo("保存完了", f"{appended}件のレコードをExcelに出力しました。")
            self.status_var.set(f"Excel保存完了: {appended}件")
        except Exception as exc:
            messagebox.showerror("保存エラー", str(exc))

    def _on_close(self) -> None:
        self.config_store.save(
            AppPathsConfig(
                input_folder=self.input_folder_var.get().strip(),
                excel_path=self.excel_path_var.get().strip(),
            )
        )
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    ReceiptApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
