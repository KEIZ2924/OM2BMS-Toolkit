from __future__ import annotations

import tkinter as tk
from tkinter import ttk


LIGHT_COLORS = {
    "bg": "#f3f6fb",
    "panel_bg": "#ffffff",
    "text_bg": "#ffffff",
    "text_fg": "#1f2328",
    "fg": "#1f2328",
    "muted_fg": "#57606a",
    "select_bg": "#2f6feb",
    "select_fg": "#ffffff",
    "canvas_bg": "#f4f7fb",
    "border": "#d0d7de",
    "button_bg": "#ffffff",
    "button_fg": "#1f2328",
}


DARK_COLORS = {
    "bg": "#0d1117",
    "panel_bg": "#161b22",
    "text_bg": "#0d1117",
    "text_fg": "#e6edf3",
    "fg": "#e6edf3",
    "muted_fg": "#8b949e",
    "select_bg": "#1f6feb",
    "select_fg": "#ffffff",
    "canvas_bg": "#0d1117",
    "border": "#30363d",
    "button_bg": "#21262d",
    "button_fg": "#e6edf3",
}


def configure_ui(root: tk.Tk) -> None:
    """
    初始化 ttk 样式。
    """
    style = ttk.Style(root)

    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure("App.TFrame", background="#f3f6fb")
    style.configure("Card.TFrame", background="#ffffff", relief="flat")
    style.configure("Title.TLabel", background="#f3f6fb", foreground="#1f2328", font=("Microsoft YaHei UI", 18, "bold"))
    style.configure("Subtitle.TLabel", background="#f3f6fb", foreground="#57606a", font=("Microsoft YaHei UI", 10))
    style.configure("TLabel", background="#f3f6fb", foreground="#1f2328")
    style.configure("TButton", padding=(10, 6))
    style.configure("Choice.TButton", padding=(10, 6))
    style.configure("SelectedChoice.TButton", padding=(10, 6))
    style.configure("TLabelframe", background="#f3f6fb")
    style.configure("TLabelframe.Label", background="#f3f6fb", foreground="#1f2328")
    style.configure("Treeview", rowheight=26)
    style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))


def apply_theme(root: tk.Tk, dark: bool = False) -> None:
    """
    应用浅色/深色主题。
    同时把颜色保存到 root._theme_colors，供 app 里的 Text/Canvas 注册组件使用。
    """
    colors = DARK_COLORS if dark else LIGHT_COLORS
    root._theme_colors = colors

    root.configure(bg=colors["bg"])

    style = ttk.Style(root)

    style.configure("App.TFrame", background=colors["bg"])
    style.configure("Card.TFrame", background=colors["panel_bg"])
    style.configure("Title.TLabel", background=colors["bg"], foreground=colors["fg"])
    style.configure("Subtitle.TLabel", background=colors["bg"], foreground=colors["muted_fg"])
    style.configure("TLabel", background=colors["bg"], foreground=colors["fg"])

    style.configure(
        "TLabelframe",
        background=colors["bg"],
        foreground=colors["fg"],
        bordercolor=colors["border"],
    )
    style.configure(
        "TLabelframe.Label",
        background=colors["bg"],
        foreground=colors["fg"],
    )

    style.configure(
        "TButton",
        background=colors["button_bg"],
        foreground=colors["button_fg"],
    )

    style.configure(
        "Choice.TButton",
        background=colors["button_bg"],
        foreground=colors["button_fg"],
    )

    style.configure(
        "SelectedChoice.TButton",
        background=colors["select_bg"],
        foreground=colors["select_fg"],
    )

    style.map(
        "SelectedChoice.TButton",
        background=[("active", colors["select_bg"]), ("pressed", colors["select_bg"])],
        foreground=[("active", colors["select_fg"]), ("pressed", colors["select_fg"])],
    )

    style.configure(
        "Treeview",
        background=colors["text_bg"],
        foreground=colors["text_fg"],
        fieldbackground=colors["text_bg"],
        bordercolor=colors["border"],
    )

    style.configure(
        "Treeview.Heading",
        background=colors["button_bg"],
        foreground=colors["button_fg"],
    )

    style.map(
        "Treeview",
        background=[("selected", colors["select_bg"])],
        foreground=[("selected", colors["select_fg"])],
    )
