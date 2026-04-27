from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from om2bms.gui.constants import (
    ANALYSIS_MODE_LABELS,
    JUDGE_OPTIONS,
)


class ConverterTab:
    def __init__(self, app, parent: ttk.Frame) -> None:
        self.app = app
        self.parent = parent
        self._build()

    def _set_choice(self, variable: tk.StringVar, value: str, buttons: list[ttk.Button]) -> None:
        variable.set(value)
        self._refresh_choice_group(variable, buttons)

    def _refresh_choice_group(self, variable: tk.StringVar, buttons: list[ttk.Button]) -> None:
        current = variable.get()

        for button in buttons:
            value = getattr(button, "_choice_value", None)
            if value == current:
                button.configure(style="SelectedChoice.TButton")
            else:
                button.configure(style="Choice.TButton")

    def _register_choice_group(
        self,
        parent,
        variable: tk.StringVar,
        choices: list[tuple[str, str]],
        buttons: list[ttk.Button],
        command_after=None,
    ) -> None:
        for index, (value, text) in enumerate(choices):
            def command(v=value):
                variable.set(v)
                self._refresh_choice_group(variable, buttons)
                if command_after is not None:
                    command_after()

            btn = ttk.Button(
                parent,
                text=text,
                style="Choice.TButton",
                command=command,
            )
            btn._choice_value = value
            btn.grid(row=0, column=index, padx=(0, 8), sticky="ew")

            parent.columnconfigure(index, weight=1)
            buttons.append(btn)

        self._refresh_choice_group(variable, buttons)

    def _toggle_boolean_choice(self, variable: tk.BooleanVar, buttons: list[ttk.Button]) -> None:
        variable.set(not variable.get())
        self._refresh_boolean_group(variable, buttons)

    def _refresh_boolean_group(self, variable: tk.BooleanVar, buttons: list[ttk.Button]) -> None:
        current = variable.get()

        for button in buttons:
            value = getattr(button, "_choice_value", None)
            if value == current:
                button.configure(style="SelectedChoice.TButton")
            else:
                button.configure(style="Choice.TButton")

    def _register_boolean_group(
        self,
        parent,
        variable: tk.BooleanVar,
        buttons: list[ttk.Button],
        true_text: str = "开启",
        false_text: str = "关闭",
    ) -> None:
        choices = [
            (True, true_text),
            (False, false_text),
        ]

        for index, (value, text) in enumerate(choices):
            btn = ttk.Button(
                parent,
                text=text,
                style="Choice.TButton",
                command=lambda v=value: (
                    variable.set(v),
                    self._refresh_boolean_group(variable, buttons),
                ),
            )
            btn._choice_value = value
            btn.grid(row=0, column=index, padx=(0, 8), sticky="ew")

            parent.columnconfigure(index, weight=1)
            buttons.append(btn)

        self._refresh_boolean_group(variable, buttons)

    def _build(self) -> None:
        app = self.app
        parent = self.parent

        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        outer = ttk.Frame(parent, style="App.TFrame", padding=12)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=0)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(0, weight=1)

        # 左侧滚动区域
        left_canvas = tk.Canvas(outer, highlightthickness=0)
        left_scroll = ttk.Scrollbar(outer, orient="vertical", command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_scroll.set)

        left_canvas.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        left_scroll.grid(row=0, column=0, sticky="nse")

        left = ttk.Frame(left_canvas, style="App.TFrame")
        left_window = left_canvas.create_window((0, 0), window=left, anchor="nw")

        def _on_left_configure(event):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))

        def _on_canvas_configure(event):
            left_canvas.itemconfigure(left_window, width=event.width)

        left.bind("<Configure>", _on_left_configure)
        left_canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel(event):
            left_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_mousewheel(event):
            left_canvas.unbind_all("<MouseWheel>")

        left_canvas.bind("<Enter>", _bind_mousewheel)
        left_canvas.bind("<Leave>", _unbind_mousewheel)
        left.bind("<Enter>", _bind_mousewheel)
        left.bind("<Leave>", _unbind_mousewheel)

        app.register_theme_canvas(left_canvas)

        # 右侧日志区域
        right = ttk.Frame(outer, style="App.TFrame")
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        # ================= 模式 =================
        mode_box = ttk.LabelFrame(left, text="处理模式", padding=10)
        mode_box.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        mode_box.columnconfigure(0, weight=1)

        mode_choices_frame = ttk.Frame(mode_box, style="App.TFrame")
        mode_choices_frame.grid(row=0, column=0, sticky="ew")

        self._register_choice_group(
            mode_choices_frame,
            app.mode_var,
            [
                ("single", "单文件转换"),
                ("batch", "批量转换"),
                ("zip", "ZIP 改 OSZ"),
            ],
            app.mode_buttons,
            command_after=app._sync_mode_widgets,
        )

        ttk.Label(
            mode_box,
            textvariable=app.mode_hint_var,
            style="Subtitle.TLabel",
            wraplength=420,
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        # ================= 输入 =================
        input_box = ttk.LabelFrame(left, text="输入", padding=10)
        input_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        input_box.columnconfigure(1, weight=1)

        ttk.Label(input_box, textvariable=app.input_label_var).grid(row=0, column=0, sticky="w")
        ttk.Entry(input_box, textvariable=app.input_var).grid(row=1, column=0, columnspan=3, sticky="ew", pady=(4, 6))

        app.input_file_button = ttk.Button(
            input_box,
            text="选择文件",
            command=app._select_input_file,
        )
        app.input_file_button.grid(row=2, column=0, sticky="ew", padx=(0, 6))

        app.input_folder_button = ttk.Button(
            input_box,
            text="选择文件夹",
            command=app._select_input_folder,
        )
        app.input_folder_button.grid(row=2, column=1, sticky="ew", padx=(0, 6))

        # ================= 输出 =================
        app.output_frame = ttk.LabelFrame(left, text="输出", padding=10)
        app.output_frame.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        app.output_frame.columnconfigure(0, weight=1)

        ttk.Entry(app.output_frame, textvariable=app.output_var).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        app.output_button = ttk.Button(
            app.output_frame,
            text="选择输出目录",
            command=app._select_output,
        )
        app.output_button.grid(row=0, column=1, sticky="ew")

        ttk.Button(
            app.output_frame,
            text="打开输出目录",
            command=app._open_output_folder,
        ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        # ================= 选项 =================
        option_box = ttk.LabelFrame(left, text="转换选项", padding=10)
        option_box.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        option_box.columnconfigure(1, weight=1)

        ttk.Label(option_box, text="Hitsound").grid(row=0, column=0, sticky="w", pady=(0, 4))
        hitsound_frame = ttk.Frame(option_box, style="App.TFrame")
        hitsound_frame.grid(row=0, column=1, sticky="ew", pady=(0, 4))
        self._register_boolean_group(hitsound_frame, app.hitsound_var, [], "开启", "关闭")

        ttk.Label(option_box, text="背景图片").grid(row=1, column=0, sticky="w", pady=(0, 4))
        bg_frame = ttk.Frame(option_box, style="App.TFrame")
        bg_frame.grid(row=1, column=1, sticky="ew", pady=(0, 4))
        self._register_boolean_group(bg_frame, app.bg_var, [], "保留", "移除")

        ttk.Label(option_box, text="Offset").grid(row=2, column=0, sticky="w", pady=(0, 4))
        app.offset_entry = ttk.Entry(option_box, textvariable=app.offset_var)
        app.offset_entry.grid(row=2, column=1, sticky="ew", pady=(0, 4))

        ttk.Label(option_box, text="判定").grid(row=3, column=0, sticky="w", pady=(0, 4))
        judge_frame = ttk.Frame(option_box, style="App.TFrame")
        judge_frame.grid(row=3, column=1, sticky="ew", pady=(0, 4))

        self._register_choice_group(
            judge_frame,
            app.judge_var,
            [(key, key) for key in JUDGE_OPTIONS.keys()],
            app.judge_buttons,
        )

        ttk.Label(option_box, text="T/N").grid(row=4, column=0, sticky="w", pady=(0, 4))
        app.tn_entry = ttk.Entry(option_box, textvariable=app.tn_var)
        app.tn_entry.grid(row=4, column=1, sticky="ew", pady=(0, 4))

        # ================= 难度分析 =================
        analysis_box = ttk.LabelFrame(left, text="难度分析", padding=10)
        analysis_box.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        analysis_box.columnconfigure(0, weight=1)

        analysis_mode_frame = ttk.Frame(analysis_box, style="App.TFrame")
        analysis_mode_frame.grid(row=0, column=0, sticky="ew")

        self._register_choice_group(
            analysis_mode_frame,
            app.analysis_mode_var,
            [(key, label) for key, label in ANALYSIS_MODE_LABELS.items()],
            app.analysis_mode_buttons,
            command_after=app._sync_mode_widgets,
        )

        app.analysis_target_entry = ttk.Entry(
            analysis_box,
            textvariable=app.analysis_target_var,
        )
        app.analysis_target_entry.grid(row=1, column=0, sticky="ew", pady=(8, 0))

        # ================= 执行按钮 =================
        action_box = ttk.Frame(left, style="App.TFrame")
        action_box.grid(row=5, column=0, sticky="ew", pady=(0, 8))
        action_box.columnconfigure(0, weight=1)
        action_box.columnconfigure(1, weight=1)

        app.start_button = ttk.Button(
            action_box,
            text="开始执行",
            command=app._start,
        )
        app.start_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        app.export_button = ttk.Button(
            action_box,
            text="导出结果",
            command=app._export_results_table,
        )
        app.export_button.grid(row=0, column=1, sticky="ew")

        app.progress_bar = ttk.Progressbar(left, mode="indeterminate")
        app.progress_bar.grid(row=6, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(left, textvariable=app.status_var, style="Subtitle.TLabel").grid(row=7, column=0, sticky="w")

        # 控件运行时禁用列表
        for widget in [
            app.input_file_button,
            app.input_folder_button,
            app.output_button,
            app.offset_entry,
            app.analysis_target_entry,
        ]:
            if widget is not None:
                app.controls_to_toggle.append(widget)

        # ================= 日志 =================
        ttk.Label(right, text="运行日志", style="Title.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))

        app.log_box = ScrolledText(right, height=20, wrap="word")
        app.log_box.grid(row=1, column=0, sticky="nsew")

        app.register_theme_widget(app.log_box)

        log_actions = ttk.Frame(right, style="App.TFrame")
        log_actions.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        log_actions.columnconfigure(0, weight=1)

        ttk.Button(log_actions, text="清空日志", command=app._clear_log).grid(row=0, column=0, sticky="e")
