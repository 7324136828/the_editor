from __future__ import annotations

import os
import sys
from pathlib import Path

WINDOWS_FONT_FILES = (
    "segoeui.ttf",
    "segoeuib.ttf",
    "segoeuii.ttf",
    "segoeuiz.ttf",
    "calibri.ttf",
    "calibrib.ttf",
    "calibrii.ttf",
    "calibriz.ttf",
)

APP_STYLESHEET = """
QMainWindow, QDialog { background: #f3f4f6; }
QToolButton {
    border: 1px solid transparent;
    border-radius: 3px;
    padding: 2px 4px;
    background: transparent;
}
QToolButton:hover { background: #dce6f5; border-color: #b8cbe8; }
QToolButton:pressed, QToolButton:checked { background: #c4d8f2; border-color: #8fb4e0; }
QToolButton:disabled { color: #9aa3ad; }
QComboBox, QLineEdit, QSpinBox, QDoubleSpinBox {
    background: #ffffff;
    border: 1px solid #c8cdd4;
    border-radius: 3px;
    padding: 1px 4px;
    selection-background-color: #185abd;
}
QComboBox:focus, QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus {
    border-color: #185abd;
}
QComboBox QAbstractItemView {
    background: #ffffff;
    border: 1px solid #c8cdd4;
    selection-background-color: #dce6f5;
    selection-color: #17365d;
}
QPushButton {
    background: #ffffff;
    border: 1px solid #c8cdd4;
    border-radius: 3px;
    padding: 4px 12px;
}
QPushButton:hover { background: #e8eefb; border-color: #8fb4e0; }
QPushButton:pressed { background: #c4d8f2; }
QPushButton:disabled { color: #9aa3ad; background: #f0f1f3; }
QPushButton:default { background: #185abd; color: #ffffff; border-color: #185abd; }
QPushButton:default:hover { background: #2a6ac4; }
QPushButton:default:pressed { background: #14489c; }
QTabWidget::pane { border: 1px solid #d4d9df; background: #ffffff; }
QTabBar::tab {
    background: #e8ebef;
    border: 1px solid #d4d9df;
    border-bottom: none;
    padding: 4px 14px;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
}
QTabBar::tab:selected { background: #ffffff; color: #185abd; }
QTabBar::tab:hover:!selected { background: #f0f3f7; }
QDockWidget { color: #17365d; }
QListWidget { background: #ffffff; border: 1px solid #d4d9df; }
QMenuBar { background: #f3f4f6; }
QMenu { background: #ffffff; border: 1px solid #d4d9df; }
QMenu::item:selected { background: #dce6f5; }
QStatusBar { background: #eceff3; }
QToolTip { background: #ffffff; color: #17365d; border: 1px solid #c8cdd4; }
"""


def _application_icon():
    from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor
    from PySide6.QtCore import Qt

    pixmap = QPixmap(64, 64)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setBrush(QColor("#185abd"))
    painter.setPen(Qt.PenStyle.NoPen)
    painter.drawRoundedRect(2, 2, 60, 60, 12, 12)
    painter.setPen(QColor("white"))
    font = painter.font()
    font.setPointSize(30)
    font.setBold(True)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, "F")
    painter.end()
    return QIcon(pixmap)


def _windows_font_dir() -> Path:
    windir = os.environ.get("WINDIR") or os.environ.get("SystemRoot") or r"C:\Windows"
    candidate = Path(windir) / "Fonts"
    if candidate.is_dir():
        return candidate
    return Path("C:/Windows/Fonts")


def register_platform_fonts() -> list[str]:
    if sys.platform != "win32":
        return []
    from PySide6.QtGui import QFontDatabase

    font_dir = _windows_font_dir()
    try:
        available = {entry.name.lower() for entry in font_dir.iterdir()}
    except OSError:
        return []
    loaded = []
    for name in WINDOWS_FONT_FILES:
        if name not in available:
            continue
        if QFontDatabase.addApplicationFont(str(font_dir / name)) >= 0:
            loaded.append(name)
    return loaded


def configure_application(app) -> None:
    from PySide6.QtGui import QFont
    from PySide6.QtWidgets import QStyleFactory

    app.setOrganizationName("Folio")
    app.setApplicationName("Folio")
    if sys.platform == "win32" and app.platformName() == "offscreen":
        register_platform_fonts()
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setFont(QFont("Segoe UI", 10))
    app.setStyleSheet(APP_STYLESHEET)
    app.setWindowIcon(_application_icon())


def main(argv: list[str] | None = None) -> int:
    from PySide6.QtWidgets import QApplication

    args = list(sys.argv if argv is None else argv)
    app = QApplication(args)
    configure_application(app)

    from .window import FolioWindow

    window = FolioWindow()
    filename = next((a for a in args[1:] if not a.startswith("-")), None)
    if filename:
        window.open_path(Path(filename))
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
