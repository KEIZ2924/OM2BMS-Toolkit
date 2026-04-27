from __future__ import annotations

import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from queue import Queue, Empty

from om2bms.gui.constants import APP_TITLE, TABLEGEN_EXPORT_FIELDS
from om2bms.gui.config_store import (
    load_default_bms_dir,
    save_default_bms_dir,
    load_default_json_dir,
    save_default_json_dir,
)
from om2bms.gui.exporters.csv_exporter import export_dict_rows_to_csv
from om2bms.gui.workers.tablegen_worker import TableGenWorker


class TableGenTab:
    def __init__(self, app, parent: ttk.Frame) -> None:
        self.app = app
        self.parent = parent

        self.input_var = tk.StringVar()
        self.score_json_var = tk.StringVar()
        self.auto_append_var = tk.BooleanVar(value=False)
        self.use_custom_level_var = tk.BooleanVar(value=False)
        self.custom_level_var = tk.StringVar()

        self.queue: Queue[tuple[str, object]] = Queue()
        self.worker_thread: threading.Thread | None = None
        self.export_rows: list[dict] = []

        self._build()
        self._load_defaults()
        self.parent.after(150, self._process_queue)

    def _build(self) -> None:
        self.parent.columnconfigure(0, weight=1)
        self.parent.rowconfigure(0, weight=1)

        root = ttk.Frame(self.parent, style="App.TFrame", padding=12)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)
        root.rowconfigure(4, weight=1)

        input_box = ttk.LabelFrame(root, text="BMS 输入", padding=10)
        input_box.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        input_box.columnconfigure(0, weight=1)

        ttk.Entry(input_box, textvariable=self.input_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(input_box, text="选择文件", command=self._select_file).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(input_box, text="选择文件夹", command=self._select_folder).grid(row=0, column=2)

        json_box = ttk.LabelFrame(root, text="目标 score.json", padding=10)
        json_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        json_box.columnconfigure(0, weight=1)

        ttk.Entry(json_box, textvariable=self.score_json_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(json_box, text="选择 score.json", command=self._select_score_json).grid(row=0, column=1)

        option_box = ttk.LabelFrame(root, text="选项", padding=10)
        option_box.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        option_box.columnconfigure(3, weight=1)

        ttk.Checkbutton(
            option_box,
            text="未收录时自动追加到 score.json",
            variable=self.auto_append_var,
        ).grid(row=0, column=0, sticky="w", padx=(0, 16))

        ttk.Checkbutton(
            option_box,
            text="使用自定义 Level",
            variable=self.use_custom_level_var,
            command=self._toggle_custom_level,
        ).grid(row=0, column=1, sticky="w", padx=(0, 8))

        self.custom_level_entry = ttk.Entry(
            option_box,
            textvariable=self.custom_level_var,
            width=16,
            state="disabled",
        )
        self.custom_level_entry.grid(row=0, column=2, sticky="w", padx=(0, 16))

        self.start_btn = ttk.Button(option_box, text="开始处理", command=self._start)
        self.start_btn.grid(row=0, column=3, sticky="e", padx=(0, 8))

        self.export_btn = ttk.Button(option_box, text="导出表格", command=self._export_results_table)
        self.export_btn.grid(row=0, column=4, sticky="e")

        self.progress = ttk.Progressbar(option_box, mode="indeterminate")
        self.progress.grid(row=1, column=0, columnspan=5, sticky="ew", pady=(8, 0))

        table_box = ttk.LabelFrame(root, text="处理结果", padding=8)
        table_box.grid(row=3, column=0, sticky="nsew", pady=(0, 8))
        table_box.columnconfigure(0, weight=1)
        table_box.rowconfigure(0, weight=1)

        columns = ("Chart", "Title", "Artist", "Level", "Status", "Appended")
        self.tree = ttk.Treeview(table_box, columns=columns, show="headings")

        headers = {
            "Chart": "Chart",
            "Title": "Title",
            "Artist": "Artist",
            "Level": "Level",
            "Status": "Status",
            "Appended": "Appended",
        }

        for col in columns:
            self.tree.heading(col, text=headers[col])
            self.tree.column(col, anchor="center", width=150)

        self.tree.grid(row=0, column=0, sticky="nsew")

        tree_scroll = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=tree_scroll.set)

        log_box = ttk.LabelFrame(root, text="日志", padding=8)
        log_box.grid(row=4, column=0, sticky="nsew")
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)

        self.log = ScrolledText(log_box, height=8, wrap="word")
        self.log.grid(row=0, column=0, sticky="nsew")

        self.app.register_theme_widget(self.log)

    def _load_defaults(self) -> None:
        default_bms = load_default_bms_dir()
        default_json = load_default_json_dir()

        if default_bms:
            self.input_var.set(default_bms)

        if default_json:
            self.score_json_var.set(default_json)

    def _toggle_custom_level(self) -> None:
        if self.use_custom_level_var.get():
            self.custom_level_entry.configure(state="normal")
        else:
            self.custom_level_entry.configure(state="disabled")

    def _select_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 BMS 谱面文件",
            filetypes=[
                ("BMS files", "*.bms *.bme *.bml *.pms"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.input_var.set(path)
            save_default_bms_dir(path)

    def _select_folder(self) -> None:
        path = filedialog.askdirectory(title="选择 BMS 文件夹")
        if path:
            self.input_var.set(path)
            save_default_bms_dir(path)

    def _select_score_json(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 score.json",
            filetypes=[
                ("JSON files", "*.json"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.score_json_var.set(path)
            save_default_json_dir(path)

    def _start(self) -> None:
        if self.worker_thread is not None and self.worker_thread.is_alive():
            return

        input_path = self.input_var.get().strip()
        score_json_path = self.score_json_var.get().strip()

        if not input_path:
            messagebox.showinfo(APP_TITLE, "请选择 BMS 文件或文件夹路径")
            return

        if not score_json_path:
            messagebox.showinfo(APP_TITLE, "请选择目标 score.json")
            return

        self.tree.delete(*self.tree.get_children())
        self.export_rows.clear()
        self.log.delete("1.0", "end")

        self.progress.start(10)
        self.start_btn.configure(state="disabled")

        self._log(f"开始处理：{input_path}")

        worker = TableGenWorker(self.queue)

        self.worker_thread = threading.Thread(
            target=worker.run,
            args=(
                input_path,
                score_json_path,
                self.auto_append_var.get(),
                self.use_custom_level_var.get(),
                self.custom_level_var.get().strip(),
            ),
            daemon=True,
        )

        self.worker_thread.start()

    def _process_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()

                if kind == "log":
                    self._log(str(payload))

                elif kind == "log_json":
                    self._log("生成的新条目：")
                    self._log(json.dumps(payload, ensure_ascii=False, indent=2))

                elif kind == "result":
                    row = payload

                    self.tree.insert(
                        "",
                        "end",
                        values=[
                            row.get("Chart", ""),
                            row.get("Title", ""),
                            row.get("Artist", ""),
                            row.get("Level", ""),
                            row.get("Status", ""),
                            row.get("Appended", ""),
                        ],
                    )

                    self.export_rows.append(row)

                elif kind == "done":
                    self._log(str(payload))
                    self.progress.stop()
                    self.start_btn.configure(state="normal")
                    messagebox.showinfo(APP_TITLE, str(payload))

                elif kind == "error":
                    self._log("错误：" + str(payload))
                    self.progress.stop()
                    self.start_btn.configure(state="normal")
                    messagebox.showerror(APP_TITLE, str(payload))

        except Empty:
            pass

        self.parent.after(150, self._process_queue)

    def _export_results_table(self) -> None:
        if not self.export_rows:
            messagebox.showinfo(APP_TITLE, "当前没有可导出的处理结果。")
            return

        export_path = filedialog.asksaveasfilename(
            title="导出处理结果",
            defaultextension=".csv",
            initialfile="TableGen_Results.csv",
            filetypes=[("CSV 表格", "*.csv"), ("所有文件", "*.*")],
        )

        if not export_path:
            return

        try:
            export_dict_rows_to_csv(
                export_path,
                self.export_rows,
                TABLEGEN_EXPORT_FIELDS,
            )
            self._log(f"已导出结果表格：{export_path}")
            messagebox.showinfo(APP_TITLE, f"结果表格已导出到：\n{export_path}")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"导出失败：{exc}")

    def _log(self, text: str) -> None:
        self.log.insert("end", text + "\n")
        self.log.see("end")
