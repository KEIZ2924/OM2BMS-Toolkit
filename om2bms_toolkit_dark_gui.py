from __future__ import annotations

import csv
import ctypes
import multiprocessing
import os
import sys
import threading
import zipfile

from datetime import datetime
from pathlib import Path
from queue import Empty, Queue
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import tkinter as tk
import tkinter.font as tkfont
import json

from om2bms.pipeline.types import ConversionOptions, ConversionResult, DifficultyAnalysisMode
from om2bms.analysis.service import DifficultyAnalyzerService
from om2bms.services.conversion_service import ConversionService
from om2bms.table_generator.bms2json import append_missing_entry_if_needed
from om2bms.table_generator.bms2json import append_score_entry_to_json
from om2bms.table_generator.bms2json import match_or_build_missing_entry

APP_TITLE = "OMSTOOLKIT"
DEFAULT_OUTPUT_DIRNAME = "output"
JUDGE_OPTIONS = {
    "EASY": 3,
    "NORMAL": 2,
    "HARD": 1,
    "VERYHARD": 0,
}
ANALYSIS_MODE_LABELS = {
    DifficultyAnalysisMode.OFF.value: "关闭分析",
    DifficultyAnalysisMode.SINGLE.value: "分析单个输出",
    DifficultyAnalysisMode.ALL.value: "分析全部输出",
}
ANALYSIS_MODE_VALUE_BY_LABEL = {
    label: value for value, label in ANALYSIS_MODE_LABELS.items()}
SOURCE_EXPORT_SKIP_DIRS = {"__pycache__",
                           ".git", ".pytest_cache", ".mypy_cache"}
SOURCE_EXPORT_SKIP_FILES = {"default_bms_dir.txt"}
SOURCE_EXPORT_SKIP_SUFFIXES = {".exe", ".osz", ".pyc", ".pyo", ".zip"}
EXPORT_COLUMNS = [
    ("difficulty_table", "难度表"),
    ("analysis_label", "分析标签"),
    ("difficulty_display", "难度显示"),
    ("estimated_difficulty", "估计难度"),
    ("raw_score", "原始分数"),
    ("source_difficulty_label", "源难度标签"),
    ("output_file_name", "输出文件"),
    ("output_path", "输出路径"),
    ("conversion_status", "转换状态"),
    ("conversion_error", "转换错误"),
    ("analysis_enabled", "分析启用"),
    ("analysis_status", "分析状态"),
    ("analysis_source", "分析来源"),
    ("runtime_provider", "推理后端"),
    ("analysis_error", "分析错误"),
    ("analysis_selection_error", "分析选择错误"),
    ("archive_name", "压缩包"),
    ("output_directory", "输出目录"),
    ("chart_id", "chartId"),
    ("chart_index", "序号"),
    ("source_chart_name", "源谱面文件"),
    ("source_osu_path", "源谱面相对路径"),
]
def apply_theme(root: tk.Tk, dark: bool = False) -> None:
    style = ttk.Style(root)

    # 选择兼容性较好的基础主题
    for theme_name in ("clam", "vista", "xpnative", "default"):
        if theme_name in style.theme_names():
            style.theme_use(theme_name)
            break

    if dark:
        colors = {
            # ===== 基础背景 =====
            "bg": "#1e1f22",
            "card": "#2b2d31",
            "card_alt": "#25272c",
            "canvas_bg": "#1e1f22",

            # ===== 文本层级 =====
            "fg": "#c7ccd4",
            "sub_fg": "#a9b1bb",
            "hint_fg": "#8b95a1",
            "disabled_fg": "#6f7682",

            # ===== 边框 =====
            "border": "#3a3f4b",
            "border_soft": "#323743",

            # ===== 输入框 / 文本 =====
            "input_bg": "#262a31",   # 比 card 略深一点，更自然；如需完全一致可改成 #2b2d31
            "input_fg": "#c7ccd4",
            "text_bg": "#1b1d22",
            "text_fg": "#c7ccd4",

            # ===== 按钮 =====
            "button_bg": "#353b45",
            "button_hover": "#414856",
            "button_active": "#4a5364",

            # ===== 强调色 =====
            "accent": "#5b8cff",
            "accent_hover": "#6b97ff",
            "accent_pressed": "#4b78e6",

            # ===== 表格 =====
            "tree_bg": "#23262d",
            "tree_fg": "#c7ccd4",
            "tree_head_bg": "#313743",
            "tree_head_fg": "#c3c9d1",

            # ===== 选中 =====
            "select_bg": "#5b8cff",
            "select_fg": "#ffffff",

            # ===== 进度 / 滚动条 =====
            "trough": "#20232a",
        }
    else:
        colors = {
            # ===== 基础背景 =====
            "bg": "#f4f7fb",
            "card": "#ffffff",
            "card_alt": "#f8fafc",
            "canvas_bg": "#f4f7fb",

            # ===== 文本层级 =====
            "fg": "#1f2328",
            "sub_fg": "#4b5563",
            "hint_fg": "#667085",
            "disabled_fg": "#98a2b3",

            # ===== 边框 =====
            "border": "#d7deea",
            "border_soft": "#e5eaf3",

            # ===== 输入框 / 文本 =====
            "input_bg": "#ffffff",
            "input_fg": "#1f2328",
            "text_bg": "#ffffff",
            "text_fg": "#1f2328",

            # ===== 按钮 =====
            "button_bg": "#edf2f9",
            "button_hover": "#e1e8f5",
            "button_active": "#d6e0f0",

            # ===== 强调色 =====
            "accent": "#2f6feb",
            "accent_hover": "#4380f5",
            "accent_pressed": "#1f5ed8",

            # ===== 表格 =====
            "tree_bg": "#ffffff",
            "tree_fg": "#1f2328",
            "tree_head_bg": "#edf2f9",
            "tree_head_fg": "#344054",

            # ===== 选中 =====
            "select_bg": "#2f6feb",
            "select_fg": "#ffffff",

            # ===== 进度 / 滚动条 =====
            "trough": "#e9eef6",
        }

    bg = colors["bg"]
    card = colors["card"]
    fg = colors["fg"]
    sub_fg = colors["sub_fg"]
    hint_fg = colors["hint_fg"]
    disabled_fg = colors["disabled_fg"]
    border = colors["border"]
    border_soft = colors["border_soft"]
    input_bg = colors["input_bg"]
    input_fg = colors["input_fg"]
    button_bg = colors["button_bg"]
    button_hover = colors["button_hover"]
    button_active = colors["button_active"]
    accent = colors["accent"]
    accent_hover = colors["accent_hover"]
    accent_pressed = colors["accent_pressed"]
    tree_bg = colors["tree_bg"]
    tree_fg = colors["tree_fg"]
    tree_head_bg = colors["tree_head_bg"]
    tree_head_fg = colors["tree_head_fg"]
    trough = colors["trough"]
    select_bg = colors["select_bg"]
    select_fg = colors["select_fg"]

    try:
        root.configure(bg=bg)
    except Exception:
        pass

    # =========================
    # 字体
    # =========================
    try:
        default_font = tkfont.nametofont("TkDefaultFont")
        text_font = tkfont.nametofont("TkTextFont")
        fixed_font = tkfont.nametofont("TkFixedFont")
        heading_font = tkfont.nametofont("TkHeadingFont")
        menu_font = tkfont.nametofont("TkMenuFont")

        default_font.configure(family="Microsoft YaHei", size=10)
        text_font.configure(family="Microsoft YaHei", size=10)
        fixed_font.configure(family="Consolas", size=10)
        heading_font.configure(family="Microsoft YaHei", size=10, weight="bold")
        menu_font.configure(family="Microsoft YaHei", size=10)
    except Exception:
        pass

    # =========================
    # 基础
    # =========================
    style.configure(".", background=bg, foreground=fg)
    style.configure("TFrame", background=bg)
    style.configure("App.TFrame", background=bg)
    style.configure("CardInner.TFrame", background=card)

    # =========================
    # Label
    # =========================
    style.configure("TLabel", background=bg, foreground=fg)

    style.configure(
        "Card.TLabel",
        background=card,
        foreground=fg,
    )

    style.configure(
        "Title.TLabel",
        background=bg,
        foreground="#d2d7df" if dark else "#111827",
        font=("Microsoft YaHei", 16, "bold"),
    )

    style.configure(
        "Subtitle.TLabel",
        background=bg,
        foreground=sub_fg,
        font=("Microsoft YaHei", 10),
    )

    style.configure(
        "SectionLabel.TLabel",
        background=card,
        foreground=fg,
        font=("Microsoft YaHei", 10, "bold"),
    )

    style.configure(
        "Hint.TLabel",
        background=card,
        foreground=hint_fg,
        font=("Microsoft YaHei", 9),
    )

    # =========================
    # LabelFrame
    # =========================
    style.configure(
        "TLabelframe",
        background=card,
        bordercolor=border,
        relief="solid",
        borderwidth=1,
    )
    style.configure(
        "TLabelframe.Label",
        background=card,
        foreground=sub_fg,
        font=("Microsoft YaHei", 10, "bold"),
    )

    style.configure(
        "Card.TLabelframe",
        background=card,
        bordercolor=border,
        relief="solid",
        borderwidth=1,
    )
    style.configure(
        "Card.TLabelframe.Label",
        background=card,
        foreground=sub_fg,
        font=("Microsoft YaHei", 10, "bold"),
    )

    # =========================
    # Button
    # =========================
    style.configure(
        "TButton",
        background=button_bg,
        foreground=fg,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        borderwidth=1,
        relief="solid",
        focusthickness=0,
        focuscolor=accent,
        padding=(12, 7),
    )
    style.map(
        "TButton",
        background=[
            ("disabled", border_soft),
            ("pressed", button_active),
            ("active", button_hover),
        ],
        foreground=[
            ("disabled", disabled_fg),
        ],
        bordercolor=[
            ("disabled", border),
            ("focus", accent),
            ("active", accent),
        ],
        lightcolor=[
            ("pressed", button_active),
            ("active", button_hover),
        ],
        darkcolor=[
            ("pressed", button_active),
            ("active", button_hover),
        ],
    )

    style.configure(
        "Accent.TButton",
        background=accent,
        foreground="#ffffff",
        bordercolor=accent,
        lightcolor=accent,
        darkcolor=accent,
        borderwidth=1,
        relief="solid",
        padding=(12, 7),
    )
    style.map(
        "Accent.TButton",
        background=[
            ("disabled", border_soft),
            ("pressed", accent_pressed),
            ("active", accent_hover),
        ],
        foreground=[
            ("disabled", "#d0d5dd"),
        ],
        bordercolor=[
            ("focus", accent),
            ("active", accent_hover),
        ],
    )

    style.configure(
        "Choice.TButton",
        background=button_bg,
        foreground=fg,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        borderwidth=1,
        relief="solid",
        padding=(14, 10),
    )
    style.map(
        "Choice.TButton",
        background=[
            ("disabled", border_soft),
            ("active", button_hover),
            ("pressed", button_active),
        ],
        foreground=[
            ("disabled", disabled_fg),
        ],
        bordercolor=[
            ("focus", accent),
            ("active", accent),
        ],
    )

    style.configure(
        "ChoiceSelected.TButton",
        background=accent,
        foreground="#ffffff",
        bordercolor=accent,
        lightcolor=accent,
        darkcolor=accent,
        borderwidth=1,
        relief="solid",
        padding=(14, 10),
    )
    style.map(
        "ChoiceSelected.TButton",
        background=[
            ("disabled", border_soft),
            ("active", accent_hover),
            ("pressed", accent_pressed),
        ],
        foreground=[
            ("disabled", "#d0d5dd"),
        ],
        bordercolor=[
            ("focus", accent),
            ("active", accent_hover),
        ],
    )

    # =========================
    # Entry / Combobox / Spinbox
    # =========================
    style.configure(
        "TEntry",
        fieldbackground=input_bg,
        background=input_bg,
        foreground=input_fg,
        insertcolor=input_fg,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        borderwidth=1,
        relief="solid",
        padding=6,
    )
    style.map(
        "TEntry",
        bordercolor=[("focus", accent)],
        lightcolor=[("focus", accent)],
        darkcolor=[("focus", accent)],
    )

    style.configure(
        "TCombobox",
        fieldbackground=input_bg,
        background=input_bg,
        foreground=input_fg,
        arrowcolor=fg,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        borderwidth=1,
        relief="solid",
        padding=4,
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", input_bg)],
        bordercolor=[("focus", accent)],
        lightcolor=[("focus", accent)],
        darkcolor=[("focus", accent)],
    )

    style.configure(
        "TSpinbox",
        fieldbackground=input_bg,
        background=input_bg,
        foreground=input_fg,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        arrowsize=12,
        borderwidth=1,
        relief="solid",
        padding=4,
    )

    # =========================
    # Check / Radio
    # =========================
    style.configure(
        "TCheckbutton",
        background=card,
        foreground=fg,
        font=("Microsoft YaHei", 10),
        indicatorcolor=input_bg,
        indicatormargin=6,
        padding=(2, 2),
    )
    style.map(
        "TCheckbutton",
        foreground=[
            ("disabled", disabled_fg),
            ("active", fg),
            ("selected", fg),
        ],
        background=[
            ("active", card),
        ],
        indicatorcolor=[
            ("selected", accent),
            ("active", input_bg),
            ("!selected", input_bg),
        ],
    )

    style.configure(
        "TRadiobutton",
        background=card,
        foreground=fg,
        font=("Microsoft YaHei", 10),
        indicatorcolor=input_bg,
        indicatormargin=6,
        padding=(2, 2),
    )
    style.map(
        "TRadiobutton",
        foreground=[
            ("disabled", disabled_fg),
            ("active", fg),
            ("selected", fg),
        ],
        background=[
            ("active", card),
        ],
        indicatorcolor=[
            ("selected", accent),
            ("active", input_bg),
            ("!selected", input_bg),
        ],
    )

    # =========================
    # Notebook
    # =========================
    style.configure(
        "TNotebook",
        background=bg,
        borderwidth=0,
        tabmargins=(0, 0, 0, 0),
    )
    style.configure(
        "TNotebook.Tab",
        background=button_bg,
        foreground=sub_fg if dark else fg,
        padding=(16, 9),
        borderwidth=0,
    )
    style.map(
        "TNotebook.Tab",
        background=[
            ("selected", card),
            ("active", button_hover),
        ],
        foreground=[
            ("selected", fg),
            ("active", fg),
        ],
    )

    # =========================
    # Progressbar
    # =========================
    style.configure(
        "TProgressbar",
        troughcolor=trough,
        background=accent,
        lightcolor=accent,
        darkcolor=accent,
        bordercolor=border,
        thickness=10,
    )

    # =========================
    # Scrollbar
    # =========================
    style.configure(
        "TScrollbar",
        background=button_bg,
        troughcolor=bg,
        bordercolor=border_soft,
        arrowcolor=fg,
        relief="flat",
    )
    style.map(
        "TScrollbar",
        background=[
            ("active", button_hover),
            ("pressed", button_active),
        ]
    )

    # =========================
    # Treeview
    # =========================
    style.configure(
        "Treeview",
        background=tree_bg,
        fieldbackground=tree_bg,
        foreground=tree_fg,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        rowheight=32,
        relief="solid",
    )
    style.map(
        "Treeview",
        background=[("selected", select_bg)],
        foreground=[("selected", select_fg)],
    )

    style.configure(
        "Treeview.Heading",
        background=tree_head_bg,
        foreground=tree_head_fg,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        relief="solid",
        borderwidth=1,
        font=("Microsoft YaHei", 10, "bold"),
    )
    style.map(
        "Treeview.Heading",
        background=[("active", button_hover)],
        foreground=[("active", fg)],
    )

    # =========================
    # 其他
    # =========================
    style.configure("TSeparator", background=border_soft)
    style.configure("TSizegrip", background=bg)
    style.configure("TPanedwindow", background=bg)
    style.configure("Sash", background=border)

    # 给外部 Text / Canvas / 其他自定义控件使用
    root._theme_colors = colors




def enable_high_dpi() -> None:
    if sys.platform != "win32":
        return

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
CONFIG_FILE = APP_DIR / "default_bms_dir.txt"


def load_default_output_dir() -> str:
    if CONFIG_FILE.exists():
        saved = CONFIG_FILE.read_text(encoding="utf-8").strip()
        if saved:
            return saved
    default_dir = APP_DIR / DEFAULT_OUTPUT_DIRNAME
    default_dir.mkdir(parents=True, exist_ok=True)
    return str(default_dir)


def save_default_output_dir(path: str) -> None:
    CONFIG_FILE.write_text(path.strip(), encoding="utf-8")


def configure_ui(root: tk.Tk) -> None:
    apply_theme(root, dark=False)



def rename_zip_to_osz(path: str) -> int:
    source = Path(path)
    if source.is_file():
        if source.suffix.lower() != ".zip":
            return 0
        target = source.with_suffix(".osz")
        if target.exists():
            return 0
        source.rename(target)
        return 1

    if not source.is_dir():
        return 0

    count = 0
    for zip_file in sorted(source.glob("*.zip")):
        target = zip_file.with_suffix(".osz")
        if target.exists():
            continue
        zip_file.rename(target)
        count += 1
    return count


class ConverterTab:
    def __init__(self, app, parent):
        self.app = app
        self.parent = parent
        self.parent.update_idletasks()
        self._build()

    def _set_choice(self, var, value, group_name: str) -> None:
        var.set(value)
        self._refresh_choice_group(group_name)

        if group_name in ("mode", "analysis_mode"):
            self.app._sync_mode_widgets()

    def _refresh_choice_group(self, group_name: str) -> None:
        group = getattr(self, "_choice_groups", {}).get(group_name)
        if not group:
            return

        var = group["var"]
        buttons = group["buttons"]
        current = var.get()

        for value, btn in buttons.items():
            btn.configure(
                style="ChoiceSelected.TButton" if value == current else "Choice.TButton"
            )

    def _register_choice_group(self, group_name: str, var, buttons: dict) -> None:
        if not hasattr(self, "_choice_groups"):
            self._choice_groups = {}

        self._choice_groups[group_name] = {
            "var": var,
            "buttons": buttons,
        }
        self._refresh_choice_group(group_name)
    def _toggle_boolean_choice(self, var, group_name: str) -> None:
        var.set(not bool(var.get()))
        self._refresh_boolean_group(group_name)

    def _refresh_boolean_group(self, group_name: str) -> None:
        group = getattr(self, "_boolean_choice_groups", {}).get(group_name)
        if not group:
            return

        var = group["var"]
        button = group["button"]

        button.configure(
            style="ChoiceSelected.TButton" if bool(var.get()) else "Choice.TButton"
        )

    def _register_boolean_group(self, group_name: str, var, button) -> None:
        if not hasattr(self, "_boolean_choice_groups"):
            self._boolean_choice_groups = {}

        self._boolean_choice_groups[group_name] = {
            "var": var,
            "button": button,
        }
        self._refresh_boolean_group(group_name)


    def _build(self):
        app = self.app

        # ================= 根容器 =================
        main = ttk.Frame(self.parent, style="App.TFrame", padding=12)
        main.grid(row=0, column=0, sticky="nsew")

        self.parent.columnconfigure(0, weight=1)
        self.parent.rowconfigure(0, weight=1)

        main.columnconfigure(0, weight=3, minsize=360)
        main.columnconfigure(1, weight=5, minsize=420)
        main.rowconfigure(1, weight=1)

        # ================= HEADER =================
        header = ttk.Frame(main, style="App.TFrame")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text=APP_TITLE, style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )

        ttk.Label(
            header,
            text="支持 .osz / .zip 输入与 .osz 批量处理，并附带难度分析功能。",
            style="Subtitle.TLabel",
            wraplength=1400,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        # ================= 左侧：可滚动配置区 =================
        left_outer = ttk.Frame(main, style="App.TFrame")
        left_outer.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        left_outer.columnconfigure(0, weight=1)
        left_outer.rowconfigure(0, weight=1)

        left_canvas = tk.Canvas(left_outer, highlightthickness=0, bd=0)
        left_canvas.grid(row=0, column=0, sticky="nsew")

        left_scrollbar = ttk.Scrollbar(
            left_outer, orient="vertical", command=left_canvas.yview
        )
        left_scrollbar.grid(row=0, column=1, sticky="ns")

        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        left = ttk.Frame(left_canvas, style="App.TFrame")
        left.columnconfigure(0, weight=1)

        left_window = left_canvas.create_window((0, 0), window=left, anchor="nw")

        def _update_left_scrollregion(event=None):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))

        def _resize_left_inner(event):
            left_canvas.itemconfigure(left_window, width=event.width)

        left.bind("<Configure>", _update_left_scrollregion)
        left_canvas.bind("<Configure>", _resize_left_inner)

        def _on_mousewheel(event):
            left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        left_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # ===== 转换模式 =====
        mode_box = ttk.LabelFrame(
            left, text="转换模式", style="Card.TLabelframe", padding=10
        )
        mode_box.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        mode_box.columnconfigure(0, weight=1)

        mode_items = [
            ("single", "转换单个压缩包"),
            ("batch", "批量转换文件夹中的 .osz"),
            ("zip", "把 .zip 改名成 .osz"),
        ]
        self.mode_choice_buttons = {}

        for i, (v, t) in enumerate(mode_items):
            btn = ttk.Button(
                mode_box,
                text=t,
                style="Choice.TButton",
                command=lambda value=v: self._set_choice(app.mode_var, value, "mode"),
            )
            btn.grid(row=i, column=0, sticky="ew", pady=4)
            self.mode_choice_buttons[v] = btn

        self._register_choice_group("mode", app.mode_var, self.mode_choice_buttons)

        ttk.Label(
            mode_box,
            textvariable=app.mode_hint_var,
            style="Hint.TLabel",
            wraplength=520,
            justify="left",
        ).grid(row=3, column=0, sticky="ew", pady=(8, 0))

        # ===== 文件路径 =====
        path_box = ttk.LabelFrame(
            left, text="文件路径", style="Card.TLabelframe", padding=10
        )
        path_box.grid(row=1, column=0, sticky="ew", pady=8)
        path_box.columnconfigure(0, weight=1)

        ttk.Label(path_box, textvariable=app.input_label_var).grid(
            row=0, column=0, sticky="w"
        )

        input_entry = ttk.Entry(path_box, textvariable=app.input_var)
        input_entry.grid(row=1, column=0, sticky="ew", pady=(4, 4))

        input_btn_row = ttk.Frame(path_box, style="App.TFrame")
        input_btn_row.grid(row=2, column=0, sticky="w", pady=(0, 6))

        app.input_file_button = ttk.Button(
            input_btn_row, text="选择文件", command=app._select_input_file
        )
        app.input_file_button.grid(row=0, column=0, padx=(0, 6))

        app.input_folder_button = ttk.Button(
            input_btn_row, text="选择文件夹", command=app._select_input_folder
        )
        app.input_folder_button.grid(row=0, column=1)

        app.output_frame = ttk.Frame(path_box, style="App.TFrame")
        app.output_frame.grid(row=3, column=0, sticky="ew")
        app.output_frame.columnconfigure(0, weight=1)

        ttk.Label(app.output_frame, text="输出文件夹").grid(row=0, column=0, sticky="w")

        output_entry = ttk.Entry(app.output_frame, textvariable=app.output_var)
        output_entry.grid(row=1, column=0, sticky="ew", pady=(4, 4))

        app.output_button = ttk.Button(
            app.output_frame, text="选择输出文件夹", command=app._select_output
        )
        app.output_button.grid(row=2, column=0, sticky="w")

        # ===== 转换选项 =====
        option_box = ttk.LabelFrame(
            left, text="转换选项", style="Card.TLabelframe", padding=10
        )
        option_box.grid(row=2, column=0, sticky="ew", pady=8)
        option_box.columnconfigure(0, weight=2, minsize=150)
        option_box.columnconfigure(1, weight=3, minsize=220)
        option_box.columnconfigure(2, weight=1, minsize=90)

        ttk.Label(option_box, text="音效", style="SectionLabel.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Label(option_box, text="背景 / 判定", style="SectionLabel.TLabel").grid(
            row=0, column=1, sticky="w", pady=(0, 4)
        )
        ttk.Label(option_box, text="参数", style="SectionLabel.TLabel").grid(
            row=0, column=2, sticky="w", pady=(0, 4)
        )
        app.hitsound_var.set(False)

        hitsound_check = ttk.Button(
            option_box,
            text="包含击打音效",
            style="Choice.TButton",
            command=lambda: self._toggle_boolean_choice(app.hitsound_var, "hitsound"),
        )
        hitsound_check.grid(row=1, column=0, sticky="ew", pady=(0, 10), padx=(0, 10))

        bg_check = ttk.Button(
            option_box,
            text="处理背景图片",
            style="Choice.TButton",
            command=lambda: self._toggle_boolean_choice(app.bg_var, "bg"),
        )
        bg_check.grid(row=1, column=1, sticky="ew", pady=(0, 10), padx=(0, 10))

        self._register_boolean_group("hitsound", app.hitsound_var, hitsound_check)
        self._register_boolean_group("bg", app.bg_var, bg_check)


        ttk.Label(option_box, text="偏移量 (ms)", style="SectionLabel.TLabel").grid(
            row=2, column=0, sticky="w"
        )
        ttk.Label(
            option_box,
            text="正值表示整体延后，负值表示整体提前。",
            style="Hint.TLabel",
            wraplength=220,
            justify="left",
        ).grid(row=3, column=0, sticky="ew", pady=(2, 6))

        ttk.Label(option_box, text="判定难度", style="SectionLabel.TLabel").grid(
            row=2, column=1, sticky="w"
        )
        ttk.Label(
            option_box,
            text="选择转换后使用的默认判定级别。",
            style="Hint.TLabel",
            wraplength=500,
            justify="left",
        ).grid(row=3, column=1, sticky="ew", pady=(2, 6))

        ttk.Label(option_box, text="T/N", style="SectionLabel.TLabel").grid(
            row=2, column=2, sticky="w"
        )
        ttk.Label(
            option_box,
            text="默认T/N=0.2 TOTAL>=300",
            style="Hint.TLabel",
            wraplength=220,
            justify="left",
        ).grid(row=3, column=2, sticky="ew", pady=(2, 6))

        app.offset_entry = ttk.Entry(option_box, textvariable=app.offset_var)
        app.offset_entry.grid(row=4, column=0, sticky="ew", padx=(0, 10), pady=(2, 0))

        judge_frame = ttk.Frame(option_box, style="App.TFrame")
        judge_frame.grid(row=4, column=1, sticky="ew", padx=(0, 10), pady=(2, 0))
        judge_frame.columnconfigure(0, weight=1)
        judge_frame.columnconfigure(1, weight=1)

        self.judge_choice_buttons = {}
        for i, label in enumerate(JUDGE_OPTIONS.keys()):
            btn = ttk.Button(
                judge_frame,
                text=label,
                style="Choice.TButton",
                command=lambda value=label: self._set_choice(app.judge_var, value, "judge"),
            )
            btn.grid(row=i // 2, column=i % 2, sticky="ew", padx=4, pady=4)
            self.judge_choice_buttons[label] = btn

        self._register_choice_group("judge", app.judge_var, self.judge_choice_buttons)

        app.tn_var = tk.DoubleVar(value=0.2)
        app.tn_entry = ttk.Entry(option_box, textvariable=app.tn_var, width=8)
        app.tn_entry.grid(row=4, column=2, sticky="ew", pady=(2, 0))

        # ===== 难度分析 =====
        analysis_box = ttk.LabelFrame(
            left, text="难度分析", style="Card.TLabelframe", padding=10
        )
        analysis_box.grid(row=3, column=0, sticky="ew", pady=8)
        analysis_box.columnconfigure(0, weight=2, minsize=180)
        analysis_box.columnconfigure(1, weight=2, minsize=180)

        ttk.Label(analysis_box, text="分析模式", style="SectionLabel.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(analysis_box, text="目标", style="SectionLabel.TLabel").grid(
            row=0, column=1, sticky="w"
        )

        ttk.Label(
            analysis_box,
            text="选择是否启用难度分析功能。",
            style="Hint.TLabel",
            wraplength=320,
            justify="left",
        ).grid(row=1, column=0, sticky="ew", pady=(2, 6))

        ttk.Label(
            analysis_box,
            text="可填写目标等级、目标值或相关参数。",
            style="Hint.TLabel",
            wraplength=440,
            justify="left",
        ).grid(row=1, column=1, sticky="ew", pady=(2, 6))

        analysis_frame = ttk.Frame(analysis_box, style="App.TFrame")
        analysis_frame.grid(row=2, column=0, sticky="ew", pady=(0, 0))
        analysis_frame.columnconfigure(0, weight=1)

        app.analysis_mode_var = tk.StringVar(value=DifficultyAnalysisMode.OFF.value)

        self.analysis_choice_buttons = {}
        for i, (v, t) in enumerate(ANALYSIS_MODE_LABELS.items()):
            btn = ttk.Button(
                analysis_frame,
                text=t,
                style="Choice.TButton",
                command=lambda value=v: self._set_choice(
                    app.analysis_mode_var, value, "analysis_mode"
                ),
            )
            btn.grid(row=i, column=0, sticky="ew", pady=4)
            self.analysis_choice_buttons[v] = btn

        self._register_choice_group(
            "analysis_mode",
            app.analysis_mode_var,
            self.analysis_choice_buttons,
        )

        app.analysis_target_entry = ttk.Entry(
            analysis_box, textvariable=app.analysis_target_var
        )
        app.analysis_target_entry.grid(
            row=2, column=1, sticky="ew", padx=(8, 0), pady=(0, 0)
        )

        # ===== 操作区 =====
        action_box = ttk.LabelFrame(
            left, text="开始执行", style="Card.TLabelframe", padding=10
        )
        action_box.grid(row=4, column=0, sticky="ew", pady=(8, 8))

        for i in range(4):
            action_box.columnconfigure(i, weight=1)

        app.start_button = ttk.Button(
            action_box, text="开始执行", command=app._start, style="Accent.TButton"
        )
        app.start_button.grid(row=0, column=0, sticky="ew")

        ttk.Button(
            action_box,
            text="打开输出文件夹",
            command=app._open_output_folder,
        ).grid(row=0, column=1, sticky="ew", padx=5)

        app.export_button = ttk.Button(
            action_box,
            text="导出表格",
            command=app._export_results_table,
            state="disabled",
        )
        app.export_button.grid(row=0, column=2, sticky="ew", padx=5)

        ttk.Button(
            action_box,
            text="清空日志",
            command=app._clear_log,
        ).grid(row=0, column=3, sticky="ew", padx=5)

        app.progress_bar = ttk.Progressbar(action_box, mode="indeterminate")
        app.progress_bar.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(8, 0))

        # ================= 右侧日志区 =================
        right = ttk.Frame(main, style="App.TFrame")
        right.grid(row=1, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        log_box = ttk.LabelFrame(right, text="运行日志", style="Card.TLabelframe", padding=8)
        log_box.grid(row=0, column=0, sticky="nsew")
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)

        app.log_box = ScrolledText(
            log_box,
            wrap="word",
            font=("Consolas", 10),
            padx=8,
            pady=8,
        )
        app.log_box.grid(row=0, column=0, sticky="nsew")

        # ================= 控件注册 =================
        app.mode_buttons = list(self.mode_choice_buttons.values())
        app.judge_buttons = list(self.judge_choice_buttons.values())
        app.analysis_mode_buttons = list(self.analysis_choice_buttons.values())

        app.controls_to_toggle = [
            *app.mode_buttons,
            *app.judge_buttons,
            *app.analysis_mode_buttons,
            hitsound_check,
            bg_check,
            input_entry,
            output_entry,
            app.offset_entry,
            app.tn_entry,
            app.input_file_button,
            app.input_folder_button,
            app.output_button,
            app.analysis_target_entry,
        ]

        app.register_theme_canvas(left_canvas)
        app.register_theme_widget(app.log_box)

        app._append_log("Get Ready.")
        app._sync_export_button()
        app._sync_mode_widgets()


class AnalyzerTab:
    def __init__(self, app, parent):
        self.app = app
        self.parent = parent
        self.queue = Queue()
        self.worker_thread = None
        self.export_rows: list[dict] = []  # 用于导出
        self.export_default_name = "Analyzer_Results.csv"
        self._build()
        self.parent.after(150, self._process_queue)

    # ===========================
    # 构建界面
    # ===========================
    def _build(self):
        app = self.app

        main = ttk.Frame(self.parent, style="App.TFrame", padding=12)
        main.grid(row=0, column=0, sticky="nsew")

        self.parent.columnconfigure(0, weight=1)
        self.parent.rowconfigure(0, weight=1)

        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)
        main.rowconfigure(3, weight=1)

        # ===== 输入区 =====
        input_box = ttk.LabelFrame(main, text="输入", style="Card.TLabelframe", padding=10)
        input_box.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        input_box.columnconfigure(0, weight=1)

        self.input_var = tk.StringVar()

        input_entry = ttk.Entry(input_box, textvariable=self.input_var)
        input_entry.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        btn_row = ttk.Frame(input_box, style="App.TFrame")
        btn_row.grid(row=1, column=0, sticky="w")

        self.select_file_btn = ttk.Button(
            btn_row, text="选择文件", command=self._select_file
        )
        self.select_file_btn.grid(row=0, column=0)

        self.select_folder_btn = ttk.Button(
            btn_row, text="选择文件夹", command=self._select_folder
        )
        self.select_folder_btn.grid(row=0, column=1, padx=(6, 0))

        # ===== 控制区 =====
        control_box = ttk.LabelFrame(main, text="控制", style="Card.TLabelframe", padding=10)
        control_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        control_box.columnconfigure(2, weight=1)

        self.start_btn = ttk.Button(
            control_box, text="开始分析", command=self._start
        )
        self.start_btn.grid(row=0, column=0, padx=(0, 6), pady=(0, 6), sticky="w")

        self.export_btn = ttk.Button(
            control_box, text="导出分析表格", command=self._export_results_table
        )
        self.export_btn.grid(row=0, column=1, padx=(0, 6), pady=(0, 6), sticky="w")

        self.progress = ttk.Progressbar(control_box, mode="indeterminate")
        self.progress.grid(row=1, column=0, columnspan=3, sticky="ew")

        # ===== 结果表格 =====
        table_box = ttk.LabelFrame(main, text="结果", style="Card.TLabelframe", padding=10)
        table_box.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        table_box.columnconfigure(0, weight=1)
        table_box.rowconfigure(0, weight=1)

        columns = ("Chart", "Difficulty", "Level")
        self.tree = ttk.Treeview(table_box, columns=columns, show="headings")

        for col in columns:
            self.tree.heading(col, text=col.capitalize())
            self.tree.column(col, anchor="center", width=150)

        y_scrollbar = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scrollbar.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scrollbar.grid(row=0, column=1, sticky="ns")

        # ===== 日志 =====
        log_box = ttk.LabelFrame(main, text="日志", style="Card.TLabelframe", padding=10)
        log_box.grid(row=3, column=0, sticky="nsew")
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)

        self.log = ScrolledText(log_box, height=8, wrap="word")
        self.log.grid(row=0, column=0, sticky="nsew")

        # 注册到全局主题系统，确保深浅色切换时一起生效
        app.register_theme_widget(self.log)


    # ===========================
    # 文件选择
    # ===========================
    def _select_file(self):
        path = filedialog.askopenfilename(
            title="选择BMS文件",
            filetypes=[("BMS 文件", "*.bms *.bme *.bml *.pms"), ("所有文件", "*.*")]
        )
        if path:
            self.input_var.set(path)

    def _select_folder(self):
        path = filedialog.askdirectory(title="选择BMS文件夹")
        if path:
            self.input_var.set(path)

    # ===========================
    # 开始分析
    # ===========================
    def _start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return

        path = self.input_var.get().strip()
        if not path:
            messagebox.showinfo(APP_TITLE, "请选择文件或文件夹路径")
            return

        self.progress.start(10)
        self.start_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.export_rows.clear()

        self.worker_thread = threading.Thread(
            target=self._run, args=(path,), daemon=True
        )
        self.worker_thread.start()

    # ===========================
    # 后台执行分析
    # ===========================
    def _run(self, path):
        service = DifficultyAnalyzerService()
        try:
            p = Path(path)
            if p.is_file():
                self._analyze_one(service, p)
            else:
                files = list(p.rglob("*.bms")) + list(p.rglob("*.bme")) \
                    + list(p.rglob("*.bml")) + list(p.rglob("*.pms"))
                if not files:
                    self.queue.put(("log", "未找到任何可分析文件"))
                    return

                total = len(files)
                for i, f in enumerate(files, 1):
                    self.queue.put(("log", f"[{i}/{total}] 分析 {f.name}..."))
                    self._analyze_one(service, f)

            self.queue.put(("done", "分析完成"))
        except Exception as e:
            self.queue.put(("error", str(e)))

    # ===========================
    # 单文件分析
    # ===========================
    def _analyze_one(self, service, path: Path):
        try:
            result = service.analyze_path(path)
            file_stem = Path(path).stem  # 自动去除 .bms / .bme / .bml / .pms 等后缀

            # 获取难度值与小数部分
            diff_value = float(result.estimated_difficulty or 0)
            diff_int = int(diff_value)
            diff_frac = diff_value - diff_int
            addon = ""
            if 0.25 <= diff_frac < 0.50:
                addon = "+"
            elif 0.50 <= diff_frac < 0.75:
                addon = "-"
            difficulty_display = f"{result.label}{addon}"
            row = {
                "Chart": file_stem,
                "Difficulty": f"{result.estimated_difficulty:.2f}",
                "Level": difficulty_display or "-",
                "raw": f"{result.raw_score:.4f}",
                "source": result.source,
            }

            # 添加到 Treeview
            self.queue.put(("result", row))
        except Exception as e:
            self.queue.put(("log", f"失败: {path.name}: {e}"))

    # ===========================
    # UI 队列更新
    # ===========================
    def _process_queue(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()

                if kind == "log":
                    self._log(payload)

                elif kind == "result":
                    self.tree.insert("", "end", values=[
                        payload["Chart"],
                        payload["Difficulty"],
                        payload["Level"],
                        payload["raw"],
                        payload["source"],
                    ])
                    self.export_rows.append(payload)  # ✅ 保存导出数据

                elif kind == "done":
                    self._log(payload)
                    self.progress.stop()
                    self.start_btn.config(state="normal")
                    self.export_btn.config(state="normal")

                elif kind == "error":
                    self._log("错误：" + payload)
                    self.progress.stop()
                    self.start_btn.config(state="normal")
                    self.export_btn.config(state="normal")

        except Empty:
            pass

        self.parent.after(150, self._process_queue)

    # ===========================
    # 日志输出
    # ===========================
    def _log(self, text: str):
        self.log.insert("end", text + "\n")
        self.log.see("end")
    def _export_results_table(self) -> None:
        """导出当前 GUI 表格内容到 CSV（仅导出 row 中定义的 5 列）"""
        if not self.export_rows:
            messagebox.showinfo(APP_TITLE, "当前还没有可导出的分析结果。")
            return

        export_path = filedialog.asksaveasfilename(
            title="导出分析结果",
            defaultextension=".csv",
            initialfile=self.export_default_name,
            filetypes=[("CSV 表格", "*.csv"), ("所有文件", "*.*")]
        )
        if not export_path:
            return

        # ✅ 只包含 row 里的五个字段
        fieldnames = ["Chart", "Difficulty", "Level", "raw", "source"]

        try:
            with open(export_path, "w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=fieldnames)
                writer.writeheader()
                for row in self.export_rows:
                    # 按当前 row 字段导出
                    writer.writerow({k: row.get(k, "") for k in fieldnames})

            self._log(f"已导出结果表格：{export_path}")
            messagebox.showinfo(APP_TITLE, f"结果表格已导出到：\n{export_path}")

        except Exception as e:
            messagebox.showerror(APP_TITLE, f"导出失败：{e}")


class TableGenTab:
    def __init__(self, app, parent):
        self.app = app
        self.parent = parent
        self.queue: Queue = Queue()
        self.worker_thread: threading.Thread | None = None
        self.export_rows: list[dict] = []

        self.input_var = tk.StringVar()
        self.score_json_var = tk.StringVar()
        self.auto_append_var = tk.BooleanVar(value=False)
        self.use_custom_level_var = tk.BooleanVar(value=False)
        self.custom_level_var = tk.StringVar()

        self._build()
        self._load_default_bms_dir()
        self._process_queue()

    # ===========================
    # 按钮式布尔开关
    # ===========================
    def _toggle_boolean_choice(self, var, group_name: str) -> None:
        var.set(not bool(var.get()))
        self._refresh_boolean_group(group_name)

        if group_name == "custom_level":
            self._toggle_custom_level()

    def _refresh_boolean_group(self, group_name: str) -> None:
        group = getattr(self, "_boolean_choice_groups", {}).get(group_name)
        if not group:
            return

        var = group["var"]
        button = group["button"]

        button.configure(
            style="ChoiceSelected.TButton" if bool(var.get()) else "Choice.TButton"
        )

    def _register_boolean_group(self, group_name: str, var, button) -> None:
        if not hasattr(self, "_boolean_choice_groups"):
            self._boolean_choice_groups = {}

        self._boolean_choice_groups[group_name] = {
            "var": var,
            "button": button,
        }
        self._refresh_boolean_group(group_name)

    # ===========================
    # UI 构建
    # ===========================
    def _build(self):
        app = self.app

        main = ttk.Frame(self.parent, style="App.TFrame", padding=12)
        main.grid(row=0, column=0, sticky="nsew")

        self.parent.columnconfigure(0, weight=1)
        self.parent.rowconfigure(0, weight=1)

        main.columnconfigure(0, weight=1)
        main.rowconfigure(3, weight=3)
        main.rowconfigure(4, weight=2)

        # ===== 顶部标题 =====
        header = ttk.Frame(main, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="表格生成器", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text="根据 BMS 文件与 score.json 匹配或生成谱面条目，可选择自动追加缺失项目。",
            style="Subtitle.TLabel",
            wraplength=1200,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        # ===== 输入区 =====
        input_box = ttk.LabelFrame(
            main, text="输入配置", style="Card.TLabelframe", padding=10
        )
        input_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        input_box.columnconfigure(0, weight=1)

        ttk.Label(input_box, text="BMS 文件 / 文件夹路径", style="SectionLabel.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            input_box,
            text="支持单个谱面文件或整个文件夹批量处理。",
            style="Hint.TLabel",
            wraplength=900,
            justify="left",
        ).grid(row=1, column=0, sticky="ew", pady=(2, 6))

        self.input_entry = ttk.Entry(input_box, textvariable=self.input_var)
        self.input_entry.grid(row=2, column=0, sticky="ew", pady=(0, 8))

        btn_row_1 = ttk.Frame(input_box, style="CardInner.TFrame")
        btn_row_1.grid(row=3, column=0, sticky="w", pady=(0, 10))

        self.select_file_btn = ttk.Button(
            btn_row_1, text="选择文件", command=self._select_file
        )
        self.select_file_btn.grid(row=0, column=0)

        self.select_folder_btn = ttk.Button(
            btn_row_1, text="选择文件夹", command=self._select_folder
        )
        self.select_folder_btn.grid(row=0, column=1, padx=(6, 0))

        ttk.Label(input_box, text="目标 score.json 路径", style="SectionLabel.TLabel").grid(
            row=4, column=0, sticky="w"
        )
        ttk.Label(
            input_box,
            text="用于匹配现有条目，或在需要时生成并追加新的谱面信息。",
            style="Hint.TLabel",
            wraplength=900,
            justify="left",
        ).grid(row=5, column=0, sticky="ew", pady=(2, 6))

        self.score_json_entry = ttk.Entry(input_box, textvariable=self.score_json_var)
        self.score_json_entry.grid(row=6, column=0, sticky="ew", pady=(0, 8))

        btn_row_2 = ttk.Frame(input_box, style="CardInner.TFrame")
        btn_row_2.grid(row=7, column=0, sticky="w", pady=(0, 10))

        self.select_score_json_btn = ttk.Button(
            btn_row_2, text="选择 score.json", command=self._select_score_json
        )
        self.select_score_json_btn.grid(row=0, column=0)

        self.auto_append_btn = ttk.Button(
            btn_row_2,
            text="自动追加到难易度",
            style="Choice.TButton",
            command=lambda: self._toggle_boolean_choice(self.auto_append_var, "auto_append"),
        )
        self.auto_append_btn.grid(row=0, column=1, padx=(10, 0), sticky="w")

        self._register_boolean_group("auto_append", self.auto_append_var, self.auto_append_btn)

        ttk.Label(input_box, text="Level 设置", style="SectionLabel.TLabel").grid(
            row=8, column=0, sticky="w"
        )
        ttk.Label(
            input_box,
            text="可直接填写自定义 Level；开启后将不自动分析等级。",
            style="Hint.TLabel",
            wraplength=900,
            justify="left",
        ).grid(row=9, column=0, sticky="ew", pady=(2, 6))

        level_row = ttk.Frame(input_box, style="CardInner.TFrame")
        level_row.grid(row=10, column=0, sticky="ew")
        level_row.columnconfigure(1, weight=1)

        self.custom_level_btn = ttk.Button(
            level_row,
            text="启用自定义 Level",
            style="Choice.TButton",
            command=lambda: self._toggle_boolean_choice(self.use_custom_level_var, "custom_level"),
        )
        self.custom_level_btn.grid(row=0, column=0, sticky="w", padx=(0, 10))

        self.custom_level_entry = ttk.Entry(
            level_row,
            textvariable=self.custom_level_var,
            state="disabled",
        )
        self.custom_level_entry.grid(row=0, column=1, sticky="ew")

        self._register_boolean_group(
            "custom_level",
            self.use_custom_level_var,
            self.custom_level_btn,
        )

        # ===== 控制区 =====
        control_box = ttk.LabelFrame(
            main, text="操作控制", style="Card.TLabelframe", padding=10
        )
        control_box.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        control_box.columnconfigure(0, weight=1)
        control_box.columnconfigure(1, weight=4)

        ttk.Label(
            control_box,
            text="开始后将扫描谱面、匹配 score.json，并根据设置决定是否生成缺失条目。",
            style="Hint.TLabel",
            wraplength=900,
            justify="left",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        self.start_btn = ttk.Button(
            control_box,
            text="开始生成",
            command=self._start,
            style="Accent.TButton",
        )
        self.start_btn.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(0, 6))

        self.progress = ttk.Progressbar(control_box, mode="indeterminate")
        self.progress.grid(row=1, column=1, sticky="ew")

        # ===== 结果表格 =====
        table_box = ttk.LabelFrame(
            main, text="生成结果", style="Card.TLabelframe", padding=10
        )
        table_box.grid(row=3, column=0, sticky="nsew", pady=(0, 8))
        table_box.columnconfigure(0, weight=1)
        table_box.rowconfigure(0, weight=1)

        columns = ("Chart", "Title", "Artist", "Level", "Status", "Appended")
        self.tree = ttk.Treeview(table_box, columns=columns, show="headings")

        self.tree.heading("Chart", text="谱面文件")
        self.tree.heading("Title", text="标题")
        self.tree.heading("Artist", text="作者")
        self.tree.heading("Level", text="等级")
        self.tree.heading("Status", text="状态")
        self.tree.heading("Appended", text="已追加")

        self.tree.column("Chart", anchor="w", width=220)
        self.tree.column("Title", anchor="w", width=260)
        self.tree.column("Artist", anchor="w", width=180)
        self.tree.column("Level", anchor="center", width=100)
        self.tree.column("Status", anchor="center", width=120)
        self.tree.column("Appended", anchor="center", width=80)

        y_scrollbar = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        x_scrollbar = ttk.Scrollbar(table_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar.grid(row=1, column=0, sticky="ew")

        # ===== 日志 =====
        log_box = ttk.LabelFrame(
            main, text="运行日志", style="Card.TLabelframe", padding=10
        )
        log_box.grid(row=4, column=0, sticky="nsew")
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)

        self.log = ScrolledText(
            log_box,
            height=10,
            wrap="word",
            font=("Consolas", 10),
            padx=8,
            pady=8,
            relief="flat",
            bd=0,
        )
        self.log.grid(row=0, column=0, sticky="nsew")

        app.register_theme_widget(self.log)

        self._log("Get Ready.")

    # ===========================
    # 文件选择
    # ===========================
    def _toggle_custom_level(self):
        if self.use_custom_level_var.get():
            self.custom_level_entry.config(state="normal")
        else:
            self.custom_level_entry.config(state="disabled")

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="选择BMS文件",
            filetypes=[("BMS 文件", "*.bms *.bme *.bml *.pms"), ("所有文件", "*.*")]
        )
        if path:
            self.input_var.set(path)

    def _select_folder(self):
        config_path = Path("default_bms_dir.txt")
        current_path = self.input_var.get().strip()

        initialdir = current_path if current_path else str(Path.cwd())
        if not Path(initialdir).exists():
            initialdir = str(Path.cwd())

        path = filedialog.askdirectory(
            title="选择文件夹",
            initialdir=initialdir,
        )

        if path:
            self.input_var.set(path)
            try:
                config_path.write_text(path, encoding="utf-8")
            except OSError:
                pass

    def _load_default_bms_dir(self):
        config_path = Path("default_bms_dir.txt")
        default_dir = Path.cwd()

        if not config_path.exists():
            try:
                config_path.write_text(str(default_dir), encoding="utf-8")
            except OSError:
                pass

        try:
            saved_path = config_path.read_text(encoding="utf-8-sig").strip()
        except OSError:
            saved_path = ""

        if not saved_path:
            saved_path = str(default_dir)

        p = Path(saved_path)
        if not p.exists():
            p = default_dir

        self.input_var.set(str(p))

    def _select_score_json(self):
        config_path = Path("default_json_dir.txt")

        initialdir = None
        initialfile = None

        if config_path.exists():
            try:
                saved_path = config_path.read_text(encoding="utf-8-sig").strip()
                if saved_path:
                    p = Path(saved_path)

                    if p.suffix.lower() == ".json":
                        initialdir = str(p.parent)
                        initialfile = p.name
                    else:
                        initialdir = str(p)
                        initialfile = "score.json"
            except OSError:
                pass

        path = filedialog.askopenfilename(
            title="选择 score.json",
            initialdir=initialdir,
            initialfile=initialfile,
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")]
        )

        if path:
            self.score_json_var.set(path)

    # ===========================
    # 开始执行
    # ===========================
    def _start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return

        input_path = self.input_var.get().strip()
        score_json_path = self.score_json_var.get().strip()

        if not input_path:
            messagebox.showinfo(APP_TITLE, "请选择 BMS 文件或文件夹路径")
            return

        if not score_json_path:
            messagebox.showinfo(APP_TITLE, "请选择目标 score.json")
            return

        self.progress.start(10)
        self.start_btn.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.export_rows.clear()
        self._log(f"开始处理：{input_path}")

        self.worker_thread = threading.Thread(
            target=self._run,
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

    # ===========================
    # 后台执行
    # ===========================
    def _run(
        self,
        input_path: str,
        score_json_path: str,
        auto_append: bool,
        use_custom_level: bool,
        custom_level: str,
    ):
        try:
            p = Path(input_path)

            if not p.exists():
                self.queue.put(("error", f"路径不存在：{p}"))
                return

            if p.is_file():
                files = [p]
            else:
                files = self._collect_chart_files(p)

            if not files:
                self.queue.put(("log", "未找到任何可处理的谱面文件"))
                self.queue.put(("done", "处理结束"))
                return

            total = len(files)
            self.queue.put(("log", f"共找到 {total} 个谱面文件"))

            for i, chart_path in enumerate(files, 1):
                self.queue.put(("log", f"[{i}/{total}] 处理 {chart_path.name} ..."))
                self._process_one(
                    chart_path,
                    score_json_path,
                    auto_append,
                    use_custom_level,
                    custom_level,
                )

            self.queue.put(("done", "处理完成"))

        except Exception as e:
            self.queue.put(("error", str(e)))

    # ===========================
    # 收集谱面文件
    # ===========================
    def _collect_chart_files(self, folder: Path) -> list[Path]:
        exts = {".bms", ".bme", ".bml", ".pms"}
        files: list[Path] = []

        for p in folder.rglob("*"):
            if p.is_file() and p.suffix.lower() in exts:
                files.append(p)

        files.sort()
        return files

    # ===========================
    # 单文件处理
    # ===========================
    def _process_one(
        self,
        chart_path: Path,
        score_json_path: str,
        auto_append: bool,
        use_custom_level: bool,
        custom_level: str,
    ):
        try:
            if auto_append:
                result = append_missing_entry_if_needed(
                    chart_path,
                    score_json_path,
                    use_custom_level=use_custom_level,
                    custom_level=custom_level,
                )
            else:
                result = match_or_build_missing_entry(
                    chart_path,
                    score_json_path,
                    use_custom_level=use_custom_level,
                    custom_level=custom_level,
                )

            song_info = result.get("song_info") or {}
            score_entry = result.get("score_entry") or {}
            generated_entry = result.get("generated_score_entry") or {}

            title = (
                score_entry.get("title")
                or generated_entry.get("title")
                or song_info.get("title")
                or chart_path.stem
            )
            artist = (
                score_entry.get("artist")
                or generated_entry.get("artist")
                or song_info.get("artist")
                or ""
            )
            level = (
                score_entry.get("level")
                or generated_entry.get("level")
                or song_info.get("level")
                or custom_level
                or ""
            )

            matched = bool(result.get("matched"))
            sha256_verified = result.get("sha256_verified")
            appended = result.get("appended", False)

            if matched:
                if sha256_verified is True:
                    status = "已匹配"
                    self.queue.put(("log", f"匹配成功：{chart_path.name}"))
                elif sha256_verified is False:
                    status = "SHA256不一致"
                    self.queue.put(("log", f"SHA256 不一致：{chart_path.name}"))
                else:
                    status = "已匹配"
                    self.queue.put(("log", f"匹配成功，JSON 未提供 SHA256，已跳过校验：{chart_path.name}"))
            else:
                status = "未收录"
                self.queue.put(("log", f"未收录：{chart_path.name}"))

            if not matched and generated_entry:
                self.queue.put(("log_json", generated_entry))

            if appended:
                self.queue.put(("log", f"已追加到 score.json：{chart_path.name}"))

            row = {
                "Chart": chart_path.name,
                "Title": title,
                "Artist": artist,
                "Level": level,
                "Status": status,
                "Appended": "是" if appended else "否",
            }

            self.queue.put(("result", row))

        except Exception as e:
            self.queue.put(("log", f"失败: {chart_path.name}: {e}"))

    # ===========================
    # UI 队列更新
    # ===========================
    def _process_queue(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()

                if kind == "log":
                    self._log(payload)

                elif kind == "log_json":
                    self._log("生成的新条目：")
                    self._log(json.dumps(payload, ensure_ascii=False, indent=2))

                elif kind == "result":
                    self.tree.insert(
                        "",
                        "end",
                        values=[
                            payload["Chart"],
                            payload["Title"],
                            payload["Artist"],
                            payload["Level"],
                            payload["Status"],
                            payload["Appended"],
                        ],
                    )
                    self.export_rows.append(payload)

                elif kind == "done":
                    self._log(payload)
                    self.progress.stop()
                    self.start_btn.config(state="normal")

                elif kind == "error":
                    self._log("错误：" + payload)
                    self.progress.stop()
                    self.start_btn.config(state="normal")

        except Empty:
            pass

        self.parent.after(150, self._process_queue)

    # ===========================
    # 日志输出
    # ===========================
    def _log(self, text: str):
        self.log.insert("end", text + "\n")
        self.log.see("end")



class Om2BmsGuiApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        configure_ui(self.root)
        self.root.title(APP_TITLE)
        self.root.geometry("1920x1080")
        self.root.minsize(1020, 760)
        self.root.resizable(True, True)
        self.root.configure(bg="#f3f6fb")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.conversion_service = ConversionService()

        self.mode_var = tk.StringVar(value="single")
        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar(value=load_default_output_dir())
        self.hitsound_var = tk.BooleanVar(value=True)
        self.bg_var = tk.BooleanVar(value=True)
        self.offset_var = tk.StringVar(value="0")
        self.judge_var = tk.StringVar(value="EASY")
        self.analysis_mode_var = tk.StringVar(
            value=DifficultyAnalysisMode.OFF.value)
        self.analysis_target_var = tk.StringVar()

        self.status_var = tk.StringVar(value="就绪")
        self.mode_hint_var = tk.StringVar(value="适合处理单个 .osz 或 .zip 谱面包。")
        self.input_label_var = tk.StringVar(value="输入 .osz 或 .zip 文件")

        self.queue: Queue[tuple[str, object]] = Queue()
        self.worker_thread: threading.Thread | None = None
        self.export_rows: list[dict[str, object]] = []
        self.export_default_name = "om2bms-results.csv"
        self.open_source_notice_shown = False

        self.mode_buttons: list[ttk.Radiobutton] = []
        self.controls_to_toggle: list[tk.Widget] = []

        self.output_frame: ttk.Frame | None = None
        self.input_file_button: ttk.Button | None = None
        self.input_folder_button: ttk.Button | None = None
        self.output_button: ttk.Button | None = None
        self.offset_entry: ttk.Entry | None = None
        self.judge_buttons: list[ttk.Radiobutton] = []
        self.analysis_mode_buttons: list[ttk.Radiobutton] = []
        self.analysis_target_entry: ttk.Entry | None = None
        self.start_button: ttk.Button | None = None
        self.source_button: ttk.Button | None = None
        self.source_export_button: ttk.Button | None = None
        self.export_button: ttk.Button | None = None
        self.progress_bar: ttk.Progressbar | None = None
        self.log_box: ScrolledText | None = None
        self.is_dark_theme = False
        self.theme_button_var = tk.StringVar(value="深色模式")
        self._theme_registered_widgets = []
        self._theme_registered_canvases = []


        self._build_ui()
        self._sync_mode_widgets()
        self.root.after(150, self._process_queue)
        # self.root.after(350, self._show_open_source_warning)

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, style="App.TFrame", padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        # ================= 顶部工具栏 =================
        top_bar = ttk.Frame(container, style="App.TFrame")
        top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top_bar.columnconfigure(0, weight=1)
        top_bar.columnconfigure(1, weight=0)

        ttk.Label(
            top_bar,
            text=APP_TITLE,
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")

        ttk.Button(
            top_bar,
            textvariable=self.theme_button_var,
            command=self.toggle_theme,
        ).grid(row=0, column=1, sticky="e")

        # ================= Notebook =================
        notebook = ttk.Notebook(container)
        notebook.grid(row=1, column=0, sticky="nsew")

        tab_converter = ttk.Frame(notebook, style="App.TFrame")
        tab_analyzer = ttk.Frame(notebook, style="App.TFrame")
        tab_tablegen = ttk.Frame(notebook, style="App.TFrame")

        notebook.add(tab_converter, text="转谱工具")
        notebook.add(tab_analyzer, text="BMS难度分析")
        notebook.add(tab_tablegen, text="难易度生成工具")

        # 把各个 Tab 实例保存下来也更方便后续扩展
        self.converter_tab = ConverterTab(self, tab_converter)
        self.analyzer_tab = AnalyzerTab(self, tab_analyzer)
        self.tablegen_tab = TableGenTab(self,tab_tablegen)

        # 构建完成后统一应用一次当前主题
        self.apply_current_theme()

        self._sync_export_button()

    def register_theme_widget(self, widget) -> None:
        if widget is not None:
            self._theme_registered_widgets.append(widget)

    def register_theme_canvas(self, canvas) -> None:
        if canvas is not None:
            self._theme_registered_canvases.append(canvas)

    def apply_current_theme(self) -> None:
        apply_theme(self.root, dark=self.is_dark_theme)

        colors = getattr(self.root, "_theme_colors", {})

        text_bg = colors.get("text_bg", "#ffffff")
        text_fg = colors.get("text_fg", "#1f2328")
        select_bg = colors.get("select_bg", "#2f6feb")
        select_fg = colors.get("select_fg", "#ffffff")
        canvas_bg = colors.get("canvas_bg", "#f4f7fb")

        for widget in self._theme_registered_widgets:
            try:
                widget.configure(
                    background=text_bg,
                    foreground=text_fg,
                    insertbackground=text_fg,
                    selectbackground=select_bg,
                    selectforeground=select_fg,
                )
            except Exception:
                try:
                    widget.configure(
                        bg=text_bg,
                        fg=text_fg,
                        insertbackground=text_fg,
                        selectbackground=select_bg,
                        selectforeground=select_fg,
                    )
                except Exception:
                    pass

        for canvas in self._theme_registered_canvases:
            try:
                canvas.configure(background=canvas_bg)
            except Exception:
                try:
                    canvas.configure(bg=canvas_bg)
                except Exception:
                    pass

        self.theme_button_var.set("浅色模式" if self.is_dark_theme else "深色模式")

    def toggle_theme(self) -> None:
        self.is_dark_theme = not self.is_dark_theme
        self.apply_current_theme()




    def _sync_mode_widgets(self) -> None:
        mode = self.mode_var.get()
        analysis_mode = self._get_analysis_mode_value()

        if mode == "single":
            self.input_label_var.set("输入 .osz 或 .zip 文件")
            self.mode_hint_var.set("适合处理单个 .osz 或 .zip 谱面包。")
            file_state = "normal"
            folder_state = "disabled"
        elif mode == "batch":
            self.input_label_var.set("输入包含 .osz 的文件夹")
            self.mode_hint_var.set("会遍历所选文件夹里的全部 .osz，并对每个压缩包分别执行转换和分析。")
            file_state = "disabled"
            folder_state = "normal"
        else:
            self.input_label_var.set("输入 .zip 文件或包含 .zip 的文件夹")
            self.mode_hint_var.set("只改扩展名，不触发 BMS 转换，也不会执行难度分析。")
            file_state = "normal"
            folder_state = "normal"

        if self.input_file_button is not None:
            self.input_file_button.configure(state=file_state)
        if self.input_folder_button is not None:
            self.input_folder_button.configure(state=folder_state)

        show_output = mode != "zip"
        if self.output_frame is not None:
            if show_output:
                self.output_frame.grid()
            else:
                self.output_frame.grid_remove()

        analysis_mode_state = "disabled" if mode == "zip" else "normal"
        for button in self.analysis_mode_buttons:
            button.configure(state=analysis_mode_state)
        if self.analysis_target_entry is not None:
            target_state = "normal" if (
                mode != "zip" and analysis_mode == DifficultyAnalysisMode.SINGLE.value) else "disabled"
            self.analysis_target_entry.configure(state=target_state)

    def _get_analysis_mode_value(self) -> str:
        raw_value = self.analysis_mode_var.get().strip()
        if raw_value in ANALYSIS_MODE_LABELS:
            return raw_value
        return ANALYSIS_MODE_VALUE_BY_LABEL.get(raw_value, DifficultyAnalysisMode.OFF.value)

    def _select_input_file(self) -> None:
        mode = self.mode_var.get()
        if mode == "single":
            path = filedialog.askopenfilename(
                title="选择 .osz 或 .zip 文件",
                filetypes=[("谱面压缩包", "*.osz *.zip"), ("所有文件", "*.*")],
            )
        else:
            path = filedialog.askopenfilename(
                title="选择 .zip 文件",
                filetypes=[("ZIP 文件", "*.zip"), ("所有文件", "*.*")],
            )

        if path:
            self.input_var.set(path)

    def _select_input_folder(self) -> None:
        path = filedialog.askdirectory(title="选择文件夹")
        if path:
            self.input_var.set(path)

    def _select_output(self) -> None:
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.output_var.set(path)
            save_default_output_dir(path)

    def _open_output_folder(self) -> None:
        path = self.output_var.get().strip()
        if not path:
            messagebox.showwarning(APP_TITLE, "请先选择输出文件夹。")
            return

        Path(path).mkdir(parents=True, exist_ok=True)
        os.startfile(path)

    def _clear_log(self) -> None:
        if self.log_box is None:
            return
        self.log_box.delete("1.0", "end")
        self._append_log("日志已清空。")

    def _append_log(self, text: str) -> None:
        if self.log_box is None:
            return
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")

    def _sync_export_button(self) -> None:
        if self.export_button is None:
            return

        is_running = self.worker_thread is not None and self.worker_thread.is_alive()
        if is_running:
            self.export_button.configure(state="disabled")
        else:
            self.export_button.configure(
                state="normal" if self.export_rows else "disabled")

    def _build_export_default_name(self, input_path: str) -> str:
        candidate = Path(input_path)
        if candidate.is_file():
            base_name = candidate.stem
        else:
            base_name = candidate.name or "om2bms-results"

        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        return f"{base_name}-conversion-results-{timestamp}.csv"

    def _record_export_rows(self, result: ConversionResult) -> None:
        archive_name = Path(
            result.output_directory).name if result.output_directory else ""

        if not result.charts:
            self.export_rows.append(
                {
                    "archive_name": archive_name,
                    "output_directory": result.output_directory or "",
                    "chart_id": "",
                    "chart_index": "",
                    "source_chart_name": "",
                    "source_osu_path": "",
                    "source_difficulty_label": "",
                    "output_file_name": "",
                    "output_path": "",
                    "conversion_status": "failed",
                    "conversion_error": result.conversion_error or "",
                    "analysis_enabled": "FALSE",
                    "analysis_status": "",
                    "estimated_difficulty": "",
                    "raw_score": "",
                    "difficulty_table": "",
                    "analysis_label": "",
                    "difficulty_display": "",
                    "analysis_source": "",
                    "runtime_provider": "",
                    "analysis_error": "",
                    "analysis_selection_error": result.analysis_error or "",
                }
            )
            return

        for chart in result.charts:
            analysis = result.analysis_result_for(chart.chart_id)
            self.export_rows.append(
                {
                    "archive_name": archive_name,
                    "output_directory": result.output_directory or "",
                    "chart_id": chart.chart_id,
                    "chart_index": chart.chart_index,
                    "source_chart_name": chart.source_chart_name,
                    "source_osu_path": chart.source_osu_path,
                    "source_difficulty_label": chart.difficulty_label or "",
                    "output_file_name": chart.output_file_name or "",
                    "output_path": chart.output_path or "",
                    "conversion_status": chart.conversion_status,
                    "conversion_error": chart.conversion_error or result.conversion_error or "",
                    "analysis_enabled": (
                        "TRUE" if analysis is not None and analysis.enabled else "FALSE"
                    ),
                    "analysis_status": analysis.status if analysis is not None else "",
                    "estimated_difficulty": (
                        f"{analysis.estimated_difficulty:.4f}"
                        if analysis is not None and analysis.estimated_difficulty is not None
                        else ""
                    ),
                    "raw_score": (
                        f"{analysis.raw_score:.6f}"
                        if analysis is not None and analysis.raw_score is not None
                        else ""
                    ),
                    "difficulty_table": analysis.difficulty_table if analysis is not None else "",
                    "analysis_label": analysis.difficulty_label if analysis is not None else "",
                    "difficulty_display": analysis.difficulty_display if analysis is not None else "",
                    "analysis_source": analysis.analysis_source if analysis is not None else "",
                    "runtime_provider": analysis.runtime_provider if analysis is not None else "",
                    "analysis_error": analysis.error if analysis is not None else "",
                    "analysis_selection_error": result.analysis_error or "",
                }
            )

    def _export_results_table(self) -> None:
        if not self.export_rows:
            messagebox.showinfo(APP_TITLE, "当前还没有可导出的转换结果。")
            return

        export_path = filedialog.asksaveasfilename(
            title="导出结果表格",
            defaultextension=".csv",
            initialfile=self.export_default_name,
            filetypes=[("CSV 表格", "*.csv"), ("所有文件", "*.*")],
        )
        if not export_path:
            return

        fieldnames = [label for _, label in EXPORT_COLUMNS]
        with open(export_path, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            for row in self.export_rows:
                writer.writerow({label: row.get(key, "")
                                for key, label in EXPORT_COLUMNS})

        self._append_log(f"已导出结果表格：{export_path}")
        messagebox.showinfo(APP_TITLE, f"结果表格已导出到：\n{export_path}")

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        for widget in self.controls_to_toggle:
            widget.configure(state=state)

        if self.start_button is not None:
            self.start_button.configure(
                state="disabled" if running else "normal")

        if self.progress_bar is not None:
            if running:
                self.progress_bar.start(12)
            else:
                self.progress_bar.stop()

        if not running:
            self._sync_mode_widgets()
        self._sync_export_button()

    def _build_conversion_options(self) -> ConversionOptions:
        mode = self.mode_var.get()

        try:
            offset = int(self.offset_var.get().strip())
        except ValueError as exc:
            raise ValueError("偏移量必须是整数。") from exc

        analysis_mode = self._get_analysis_mode_value()
        enable_analysis = mode != "zip" and analysis_mode != DifficultyAnalysisMode.OFF.value
        try:
            tn_value = float(self.tn_var.get())
        except (ValueError, tk.TclError):
            raise ValueError("T/N 必须是数字。")

        if not (0 < tn_value < 10):
            raise ValueError("T/N 值必须大于 0 且小于 10。")
        judge_label = self.judge_var.get().strip()
        if judge_label not in JUDGE_OPTIONS:
            raise ValueError("请从下拉框选择有效的判定难度。")

        return ConversionOptions(
            hitsound=self.hitsound_var.get(),
            bg=self.bg_var.get(),
            offset=offset,
            judge=JUDGE_OPTIONS[judge_label],
            enable_difficulty_analysis=enable_analysis,
            difficulty_analysis_mode=analysis_mode,
            difficulty_target_id=self.analysis_target_var.get().strip() or None,
            tn_value=tn_value,
        )

    def _validate_inputs(self) -> tuple[str, str, ConversionOptions]:
        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip()
        mode = self.mode_var.get()

        if not input_path:
            raise ValueError("请先选择输入文件或文件夹。")
        if not Path(input_path).exists():
            raise FileNotFoundError(f"输入路径不存在：{input_path}")

        options = self._build_conversion_options()

        if mode != "zip":
            if not output_path:
                raise ValueError("请先选择输出文件夹。")
            Path(output_path).mkdir(parents=True, exist_ok=True)
            save_default_output_dir(output_path)

        if options.enable_difficulty_analysis and options.resolved_analysis_mode() == DifficultyAnalysisMode.SINGLE:
            if not options.difficulty_target_id:
                raise ValueError("单目标分析模式必须填写目标选择器。")

        return input_path, output_path, options

    def _start(self) -> None:
        if self.worker_thread is not None and self.worker_thread.is_alive():
            return

        try:
            input_path, output_path, options = self._validate_inputs()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return

        self._set_running(True)
        self.export_rows.clear()
        self.export_default_name = self._build_export_default_name(input_path)
        self._sync_export_button()
        self.status_var.set("运行中")
        self._append_log("")
        self._append_log(f"模式：{self.mode_var.get()}")
        self._append_log(f"输入：{input_path}")
        if self.mode_var.get() != "zip":
            self._append_log(f"输出根目录：{output_path}")
            self._append_log(
                f"分析模式：{ANALYSIS_MODE_LABELS[options.resolved_analysis_mode().value]}")
            if options.difficulty_target_id:
                self._append_log(f"分析目标选择器：{options.difficulty_target_id}")

        self.worker_thread = threading.Thread(
            target=self._run_task,
            args=(self.mode_var.get(), input_path, output_path, options),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_task(self, mode: str, input_path: str, output_path: str, options: ConversionOptions) -> None:
        try:
            
            if mode == "zip":
                count = rename_zip_to_osz(input_path)
                self.queue.put(("done", f"已把 {count} 个文件从 .zip 改名为 .osz。"))
                return

            if mode == "single":
                self.queue.put(("log", "开始处理单个压缩包"))
                result = self.conversion_service.convert_osz(
                    input_path, output_path, options)
                self.queue.put(("result", result))
                message = "转换完成。" if result.conversion_success else "转换失败。"
                self.queue.put(("done", message))
                return

            archives = sorted(Path(input_path).glob("*.osz"))
            if not archives:
                raise FileNotFoundError("所选文件夹中没有找到 .osz 文件。")

            total = len(archives)
            success_count = 0
            for index, archive in enumerate(archives, start=1):
                self.queue.put(("status", f"批量任务 {index}/{total}"))
                self.queue.put(("log", f"开始处理 {archive.name}"))
                result = self.conversion_service.convert_osz(
                    str(archive), output_path, options)
                self.queue.put(("result", result))
                if result.conversion_success:
                    success_count += 1

            self.queue.put(
                ("done", f"批量处理完成：成功 {success_count}/{total} 个压缩包。"))
        except Exception as exc:
            self.queue.put(("error", str(exc)))

    def _log_conversion_result(self, result: ConversionResult) -> None:
        if result.output_directory:
            self._append_log(f"输出目录：{result.output_directory}")
        if result.conversion_error:
            self._append_log(f"转换错误：{result.conversion_error}")
        if result.analysis_error:
            self._append_log(f"分析选择错误：{result.analysis_error}")

        if result.charts:
            self._append_log("转换产物：")
            for chart in result.charts:
                if chart.conversion_status == "success":
                    self._append_log(
                        f"  [{chart.chart_index}] chartId={chart.chart_id} "
                        f"file={chart.output_file_name or '-'} "
                        f"difficulty={chart.difficulty_label or '-'}"
                    )
                else:
                    self._append_log(
                        f"  [{chart.chart_index}] FAILED source={chart.source_chart_name} error={chart.conversion_error}"
                    )

        if result.analysis_results:
            self._append_log("分析结果：")
            for analysis in result.analysis_results:
                if analysis.status == "success":
                    self._append_log(
                        f"  chartId={analysis.chart_id} status=success "
                        f"label={analysis.difficulty_label} "
                        f"provider={analysis.runtime_provider or '-'} "
                        f"display={analysis.difficulty_display} "
                        f"raw={analysis.raw_score:.6f}"
                    )
                elif analysis.status == "failed":
                    self._append_log(
                        f"  chartId={analysis.chart_id} status=failed error={analysis.error}"
                    )
                else:
                    self._append_log(
                        f"  chartId={analysis.chart_id} status=skipped")

    def _process_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "log":
                    self._append_log(str(payload))
                elif kind == "status":
                    self.status_var.set(str(payload))
                elif kind == "result":
                    self._record_export_rows(payload)
                    self._log_conversion_result(payload)
                    self._sync_export_button()
                elif kind == "done":
                    self._append_log(str(payload))
                    self.status_var.set("完成")
                    self._set_running(False)
                    messagebox.showinfo(APP_TITLE, str(payload))
                elif kind == "error":
                    self._append_log(f"错误：{payload}")
                    self.status_var.set("出错")
                    self._set_running(False)
                    messagebox.showerror(APP_TITLE, str(payload))
        except Empty:
            pass
        finally:
            self.root.after(150, self._process_queue)

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    enable_high_dpi()
    multiprocessing.freeze_support()
    Om2BmsGuiApp().run()


if __name__ == "__main__":
    main()
