#!/usr/bin/env python3
from __future__ import annotations

import os
import platform
import subprocess
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from queue import Empty, Queue
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Dict, List

from api_key_manager import ApiKeyManager
from receipt_core import (
    JAPANESE_ACCOUNTING_CATEGORIES,
    ReceiptProcessor,
    ReceiptRecord,
    append_records_to_excel,
    normalize_date_value,
)


class ReceiptApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OCRGPT 領収書OCR")
        self.root.geometry("1240x720")

        self.records: List[ReceiptRecord] = []
        self.tree_columns: List[str] = []
        self.selected_item_id: str | None = None
        self.result_queue: Queue = Queue()

        self.api_key_manager = ApiKeyManager()
        self.api_key: str | None = None

        self.input_folder_var = tk.StringVar()
        self.excel_path_var = tk.StringVar()
        self.model_var = tk.StringVar(value="gpt-4.1-mini")
        self.ocr_lang_var = tk.StringVar(value="jpn+eng")
        self.sheet_name_var = tk.StringVar()
        self.api_key_status_var = tk.StringVar(value="API Key: 未設定")
        self.status_var = tk.StringVar(value="準備完了")

        self.form_vars: Dict[str, tk.StringVar] = {
            "date": tk.StringVar(),
            "store": tk.StringVar(),
            "total": tk.StringVar(),
            "category": tk.StringVar(),
            "payment": tk.StringVar(),
            "note": tk.StringVar(),
            "error_reason": tk.StringVar(),
        }

        self._build_ui()
        self._load_saved_api_key_or_prompt()
        self.root.after(200, self._poll_worker)

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
        self.tree.tag_configure("error", background="#ffd9d9")

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

        self._make_form_entry(edit_frame, "日付", "date", 0, hint="和暦/西暦どちらでも可")
        self._make_form_entry(edit_frame, "店名", "store", 1)
        self._make_form_entry(edit_frame, "金額", "total", 2)

        ttk.Label(edit_frame, text="カテゴリ").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        self.category_combo = ttk.Combobox(
            edit_frame,
            textvariable=self.form_vars["category"],
            values=JAPANESE_ACCOUNTING_CATEGORIES,
            state="normal",
            width=32,
        )
        self.category_combo.grid(row=1, column=1, sticky="w", padx=8, pady=4)

        self._make_form_entry(edit_frame, "支払", "payment", 1, col_offset=2)
        self._make_form_entry(edit_frame, "メモ", "note", 2, col_offset=2)
        self._make_form_entry(edit_frame, "エラー理由", "error_reason", 0, col_offset=2)

        ttk.Button(edit_frame, text="選択行へ反映", command=self._apply_form_to_selected).grid(row=3, column=5, padx=8, pady=4, sticky="e")

        ttk.Label(edit_frame, text="※ 空白入力でも保存可能。必須チェックは行いません。").grid(
            row=3, column=0, columnspan=4, sticky="w", padx=8, pady=4
        )

    def _make_form_entry(self, parent: ttk.LabelFrame, label: str, key: str, row: int, col_offset: int = 0, hint: str = "") -> None:
        c = col_offset * 2
        ttk.Label(parent, text=label).grid(row=row, column=c, sticky="w", padx=8, pady=4)
        entry = ttk.Entry(parent, textvariable=self.form_vars[key], width=32)
        entry.grid(row=row, column=c + 1, sticky="w", padx=8, pady=4)
        if hint:
            ttk.Label(parent, text=hint, foreground="#666").grid(row=row, column=c + 2, sticky="w", padx=8)

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

            self.api_key_manager.save_key(value)
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
            self.api_key_manager.save_key(key)
            self.api_key = key
            self._update_api_key_status()
            messagebox.showinfo("保存完了", "APIキーを保存しました。", parent=settings)

        def delete_key() -> None:
            self.api_key_manager.delete_key()
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
        base = ["file_name", "date", "store", "total", "category", "payment", "note", "is_error", "error_reason", "source_image_link"]
        extras = sorted({k for r in self.records for k in r.fields.keys() if k not in base})
        return base + extras

    def _render_table(self) -> None:
        self.tree.delete(*self.tree.get_children())

        self.tree_columns = self._derive_columns()
        self.tree["columns"] = self.tree_columns

        for col in self.tree_columns:
            self.tree.heading(col, text=col)
            width = 140 if col not in {"note", "error_reason", "source_image_link"} else 260
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
            else:
                self.form_vars[key].set(str(rec.get(key, "")))

    def _apply_form_to_selected(self) -> None:
        if self.selected_item_id is None:
            return

        rec = self.records[int(self.selected_item_id)]
        rec.set("date", self.form_vars["date"].get().strip())
        rec.set("store", self.form_vars["store"].get().strip())
        rec.set("total", self.form_vars["total"].get().strip())
        rec.set("category", self.form_vars["category"].get().strip())
        rec.set("payment", self.form_vars["payment"].get().strip())
        rec.set("note", self.form_vars["note"].get().strip())
        rec.error_reason = self.form_vars["error_reason"].get().strip()
        rec.is_error = bool(rec.error_reason)

        normalized_date, date_error = normalize_date_value(rec.get("date", ""))
        rec.set("date", normalized_date)
        if date_error and rec.get("date", ""):
            rec.mark_error(date_error)

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
        if key in {"file_name", "is_error", "source_image_link"}:
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
                rec.is_error = bool(new_value)
            elif key == "date":
                normalized, date_error = normalize_date_value(new_value)
                rec.set("date", normalized)
                if date_error and normalized:
                    rec.mark_error(date_error)
            else:
                rec.set(key, new_value)

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


def main() -> None:
    root = tk.Tk()
    ReceiptApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
