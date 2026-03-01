import os
import sys
import logging
from logging.handlers import RotatingFileHandler

from PySide6 import QtWidgets

from .app import AppController
from .config import AppConfig


def _enable_windows_dpi_awareness() -> None:
    """Робимо процес Per-Monitor DPI Aware (V2), щоб координати оверлею не їхали.

    Без цього MSS повертає фізичні пікселі, а Qt може працювати в "логічних".
    Результат: бокси зміщуються, особливо на 125%/150% scaling.
    """
    if os.name != 'nt':
        return

    try:
        import ctypes

        user32 = ctypes.windll.user32
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = (HANDLE)-4
        user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        # Якщо не вийшло, просто живемо далі (буде як було)
        return


def run() -> int:
    _enable_windows_dpi_awareness()

    # Логи в файл + консоль, щоб можна було нарешті зрозуміти, чому "нічого не працює".
    try:
        cfg = AppConfig()
        log_dir = os.path.join(cfg.config_dir(), "logs")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, "win-screen-translator.log")

        root = logging.getLogger()
        if not root.handlers:
            root.setLevel(logging.INFO)
            fmt = logging.Formatter(
                "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S"
            )

            fh = RotatingFileHandler(log_path, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
            fh.setFormatter(fmt)
            root.addHandler(fh)

            sh = logging.StreamHandler(sys.stdout)
            sh.setFormatter(fmt)
            root.addHandler(sh)
    except Exception:
        pass

    app = QtWidgets.QApplication(sys.argv)
    # не показуємо головне вікно, тільки tray + overlay
    controller = AppController(app)
    return app.exec()


if __name__ == '__main__':
    raise SystemExit(run())
