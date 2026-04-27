import multiprocessing

from om2bms.gui.app import Om2BmsGuiApp
from om2bms.gui.ui_utils import enable_high_dpi


def main() -> None:
    enable_high_dpi()
    multiprocessing.freeze_support()
    Om2BmsGuiApp().run()


if __name__ == "__main__":
    main()
