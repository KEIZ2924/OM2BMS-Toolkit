import ctypes
import sys


def enable_high_dpi() -> None:
    """
    Windows 高 DPI 适配。
    非 Windows 平台自动跳过。
    """
    if sys.platform != "win32":
        return

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
