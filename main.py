import sys
from core.subprocess_generation import main as subprocess_main


def _run_subprocess_generation_if_requested(argv: list[str]) -> int | None:
    if len(argv) == 4 and argv[1] == "--subprocess-generation":
        return subprocess_main(["core.subprocess_generation", argv[2], argv[3]])
    return None

def main():
    maybe_code = _run_subprocess_generation_if_requested(sys.argv)
    if maybe_code is not None:
        sys.exit(maybe_code)

    from PyQt5.QtWidgets import QApplication
    from gui.main_window import MainWindow

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
