from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from queue import Queue, Empty

from om2bms.gui.constants import APP_TITLE, ANALYZER_EXPORT_FIELDS
from om2bms.gui.exporters.csv_exporter import export_dict_rows_to_csv
from om2bms.gui.workers.analyzer_worker import AnalyzerWorker


class AnalyzerTab:
    def __init__(self, app, parent: ttk.Frame) -> None:
        self.app = app
        self.parent = parent

        self.input_var = tk.StringVar()
        self.queue: Queue[tuple[str, object]] = Queue()
        self.worker_thread: threading.Thread | None = None
        self.export_rows: list[dict] = []

        self._build()
        self.parent.after(150, self._process_queue)

    def _build(self) -> None:
        self.parent.columnconfigure(0, weight=1)
        self.parent.rowconfigure(0, weight=1)

        root = ttk.Frame(self.parent, style="App.TFrame", padding=12)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)
        root.rowconfigure(3, weight=1)

        input_box = ttk.LabelFrame(root, text="输入 BMS 文件或文件夹", padding=10)
        input_box.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        input_box.columnconfigure(0, weight=1)

        ttk.Entry(input_box, textvariable=self.input_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))

        ttk.Button(input_box, text="选择文件", command=self._select_file).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(input_box, text="选择文件夹", command=self._select_folder).grid(row=0, column=2)

        action_box = ttk.Frame(root, style="App.TFrame")
        action_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        action_box.columnconfigure(0, weight=1)

        self.start_btn = ttk.Button(action_box, text="开始分析", command=self._start)
        self.start_btn.grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.export_btn = ttk.Button(action_box, text="导出表格", command=self._export_results_table)
        self.export_btn.grid(row=0, column=1, sticky="w")

        self.progress = ttk.Progressbar(action_box, mode="indeterminate")
        self.progress.grid(row=0, column=2, sticky="ew", padx=(12, 0))
        action_box.columnconfigure(2, weight=1)

        table_box = ttk.LabelFrame(root, text="分析结果", padding=8)
        table_box.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        table_box.columnconfigure(0, weight=1)
        table_box.rowconfigure(0, weight=1)

        columns = ("Chart", "Difficulty", "Level", "Raw", "Source")
        self.tree = ttk.Treeview(table_box, columns=columns, show="headings")

        headers = {
            "Chart": "Chart",
            "Difficulty": "Difficulty",
            "Level": "Level",
            "Raw": "Raw",
            "Source": "Source",
        }

        for col in columns:
            self.tree.heading(col, text=headers[col])
            self.tree.column(col, anchor="center", width=150)

        self.tree.grid(row=0, column=0, sticky="nsew")

        tree_scroll = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=tree_scroll.set)

        log_box = ttk.LabelFrame(root, text="日志", padding=8)
        log_box.grid(row=3, column=0, sticky="nsew")
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)

        self.log = ScrolledText(log_box, height=8, wrap="word")
        self.log.grid(row=0, column=0, sticky="nsew")

        self.app.register_theme_widget(self.log)

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

    def _select_folder(self) -> None:
        path = filedialog.askdirectory(title="选择 BMS 文件夹")
        if path:
            self.input_var.set(path)

    def _start(self) -> None:
        if self.worker_thread is not None and self.worker_thread.is_alive():
            return

        input_path = self.input_var.get().strip()

        if not input_path:
            messagebox.showinfo(APP_TITLE, "请选择 BMS 文件或文件夹路径")
            return

        self.tree.delete(*self.tree.get_children())
        self.export_rows.clear()
        self.log.delete("1.0", "end")

        self.progress.start(10)
        self.start_btn.configure(state="disabled")

        worker = AnalyzerWorker(self.queue)

        self.worker_thread = threading.Thread(
            target=worker.run,
            args=(input_path,),
            daemon=True,
        )
        self.worker_thread.start()

    def _process_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()

                if kind == "log":
                    self._log(str(payload))

                elif kind == "result":
                    row = payload

                    self.tree.insert(
                        "",
                        "end",
                        values=[
                            row.get("Chart", ""),
                            row.get("Difficulty", ""),
                            row.get("Level", ""),
                            row.get("Raw", ""),
                            row.get("Source", ""),
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
            messagebox.showinfo(APP_TITLE, "当前没有可导出的分析结果。")
            return

        export_path = filedialog.asksaveasfilename(
            title="导出分析结果",
            defaultextension=".csv",
            initialfile="Analyzer_Results.csv",
            filetypes=[("CSV 表格", "*.csv"), ("所有文件", "*.*")],
        )

        if not export_path:
            return

        try:
            export_dict_rows_to_csv(
                export_path,
                self.export_rows,
                ANALYZER_EXPORT_FIELDS,
            )
            self._log(f"已导出结果表格：{export_path}")
            messagebox.showinfo(APP_TITLE, f"结果表格已导出到：\n{export_path}")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"导出失败：{exc}")

    def _log(self, text: str) -> None:
        self.log.insert("end", text + "\n")
        self.log.see("end")
