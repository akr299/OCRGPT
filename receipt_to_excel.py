#!/usr/bin/env python3
from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from queue import Empty, Queue
from tkinter import filedialog, messagebox, ttk
from typing import List, Tuple

from receipt_core import (
    JAPANESE_ACCOUNTING_CATEGORIES,
    ReceiptProcessor,
    ReceiptRecord,
    append_records_to_excel,
)


class ReceiptApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OCRGPT 領収書OCR")
        self.root.geometry("1050x650")

        self.records: List[Tuple[str, ReceiptRecord]] = []
        self.result_queue: Queue = Queue()

        self.input_folder_var = tk.StringVar()
        self.excel_path_var = tk.StringVar()
        self.model_var = tk.StringVar(value="gpt-4.1-mini")
        self.ocr_lang_var = tk.StringVar(value="jpn+eng")
        self.sheet_name_var = tk.StringVar()

        self.selected_item_id: str | None = None
        self._build_ui()
        self.root.after(200, self._poll_worker)

    def _build_ui(self) -> None:
        config_frame = ttk.LabelFrame(self.root, text="設定")
        config_frame.pack(fill="x", padx=12, pady=8)

        ttk.Label(config_frame, text="画像フォルダ").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(config_frame, textvariable=self.input_folder_var, width=70).grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(config_frame, text="参照", command=self._pick_input_folder).grid(row=0, column=2, padx=4, pady=4)

        ttk.Label(config_frame, text="保存先Excel").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(config_frame, textvariable=self.excel_path_var, width=70).grid(row=1, column=1, sticky="ew", padx=4, pady=4)
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

        self.status_var = tk.StringVar(value="準備完了")
        ttk.Label(action_frame, textvariable=self.status_var).pack(side="left", padx=20)

        table_frame = ttk.LabelFrame(self.root, text="抽出結果（カテゴリは下で修正できます）")
        table_frame.pack(fill="both", expand=True, padx=12, pady=8)

        columns = ("file", "date", "store", "total", "category", "payment", "note")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=16)
        for col, label, width in [
            ("file", "画像", 180),
            ("date", "日付", 100),
            ("store", "店舗", 220),
            ("total", "金額", 90),
            ("category", "カテゴリ", 130),
            ("payment", "支払", 110),
            ("note", "メモ", 220),
        ]:
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width, anchor="w")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self._on_row_selected)

        edit_frame = ttk.LabelFrame(self.root, text="カテゴリ修正")
        edit_frame.pack(fill="x", padx=12, pady=8)

        ttk.Label(edit_frame, text="選択行のカテゴリ").pack(side="left", padx=8)
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(
            edit_frame,
            textvariable=self.category_var,
            values=JAPANESE_ACCOUNTING_CATEGORIES,
            state="readonly",
            width=24,
        )
        self.category_combo.pack(side="left", padx=8)
        self.category_combo.bind("<<ComboboxSelected>>", self._update_selected_category)

    def _pick_input_folder(self) -> None:
        folder = filedialog.askdirectory(title="領収書画像フォルダを選択")
        if folder:
            self.input_folder_var.set(folder)

    def _pick_excel_path(self) -> None:
        path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            filetypes=[("Excel file", "*.xlsx *.xlsm")],
        )
        if path:
            self.excel_path_var.set(path)

    def _set_processing_state(self, running: bool) -> None:
        if running:
            self.process_button.config(state="disabled")
            self.save_button.config(state="disabled")
        else:
            self.process_button.config(state="normal")
            self.save_button.config(state="normal" if self.records else "disabled")

    def _start_processing(self) -> None:
        input_folder = self.input_folder_var.get().strip()
        if not input_folder:
            messagebox.showerror("入力不足", "画像フォルダを指定してください。")
            return

        self._set_processing_state(True)
        self.status_var.set("OCR + OpenAI 処理中...")
        self.tree.delete(*self.tree.get_children())
        self.records.clear()

        thread = threading.Thread(target=self._worker_process, daemon=True)
        thread.start()

    def _worker_process(self) -> None:
        try:
            processor = ReceiptProcessor(
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
                self.status_var.set(f"処理完了: 成功 {len(records)}件 / 失敗 {len(failures)}件")
                if failures:
                    detail = "\n".join(f"{f.file_name}: {f.reason}" for f in failures)
                    messagebox.showwarning("一部失敗", f"以下のファイルで処理に失敗しました:\n\n{detail}")
                if not records:
                    messagebox.showerror("処理結果", "有効なレコードを作成できませんでした。")
            else:
                _, error = result
                self._set_processing_state(False)
                self.status_var.set("エラー")
                messagebox.showerror("処理エラー", str(error))
        except Empty:
            pass
        finally:
            self.root.after(200, self._poll_worker)

    def _render_table(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for index, (file_name, rec) in enumerate(self.records):
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(file_name, rec.date, rec.store, rec.total, rec.category, rec.payment, rec.note),
            )

    def _on_row_selected(self, _event: object) -> None:
        selected = self.tree.selection()
        if not selected:
            self.selected_item_id = None
            return

        item_id = selected[0]
        self.selected_item_id = item_id
        row_idx = int(item_id)
        self.category_var.set(self.records[row_idx][1].category)

    def _update_selected_category(self, _event: object) -> None:
        if self.selected_item_id is None:
            return

        row_idx = int(self.selected_item_id)
        file_name, record = self.records[row_idx]
        record.category = self.category_var.get()
        self.records[row_idx] = (file_name, record)

        current_values = list(self.tree.item(self.selected_item_id, "values"))
        current_values[4] = record.category
        self.tree.item(self.selected_item_id, values=current_values)

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
            appended = append_records_to_excel(
                excel_path=Path(excel_path),
                records=[r for _, r in self.records],
                sheet_name=sheet_name,
            )
            messagebox.showinfo("保存完了", f"{appended}件のレコードをExcelに追記しました。")
            self.status_var.set(f"Excel保存完了: {appended}件追記")
        except Exception as exc:
            messagebox.showerror("保存エラー", str(exc))


def main() -> None:
    root = tk.Tk()
    ReceiptApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
