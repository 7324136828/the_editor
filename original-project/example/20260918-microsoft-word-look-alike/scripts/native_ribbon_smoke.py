from __future__ import annotations

import os
from pathlib import Path
import sys
import tempfile


def main():
    workspace = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(workspace))
    previous_platform = os.environ.get("QT_QPA_PLATFORM")
    previous_data = os.environ.get("FOLIO_DATA_DIR")
    os.environ["QT_QPA_PLATFORM"] = "windows"
    try:
        with tempfile.TemporaryDirectory(prefix="folio-ribbon-smoke-") as data_dir:
            os.environ["FOLIO_DATA_DIR"] = data_dir
            from PySide6.QtGui import QFont, QRawFont, QTextCursor
            from PySide6.QtWidgets import QApplication
            from folio.app import configure_application
            from folio.document_features import insert_node
            from folio.models import DocumentState
            from folio.references import insert_citation, insert_note, registry
            from folio.storage import LocalStore
            from folio.window import FolioWindow

            app = QApplication([])
            configure_application(app)
            if app.platformName() != "windows":
                raise RuntimeError("Expected native Windows platform.")
            for family in ("Segoe UI", "Calibri"):
                font = QRawFont.fromFont(QFont(family, 11))
                if not font.isValid() or any(index == 0 for index in font.glyphIndexesForString("Folio123")):
                    raise RuntimeError("Missing native glyphs: " + family)
            store = LocalStore()
            window = None
            try:
                if store.root.resolve() != Path(data_dir).resolve():
                    raise RuntimeError("Smoke requires a temporary store.")
                store.set_setting("proofing", False)
                window = FolioWindow(store=store, restore=False)
                window.load_state(DocumentState(title="Research and review", html=(
                    "<h1>Research and review</h1><p>Folio keeps references and review history in this local document.</p>"
                    "<h2>Methods</h2><p>Prepare your sources, add notes and captions, then review the published pages.</p>")), clean=True)
                window.editor.moveCursor(QTextCursor.MoveOperation.End)
                window.editor.insertPlainText(" ")
                insert_note(window.editor, window.state, "This numbered footnote is printed beneath the page body.")
                window.editor.moveCursor(QTextCursor.MoveOperation.End)
                window.editor.insertPlainText(" ")
                registry(window.state)["sources"]["demo"] = {
                    "type": "book", "author": "Example, Alex", "title": "Local research methods",
                    "year": "2026", "publisher": "Example Press"}
                insert_citation(window.editor, window.state, "demo")
                cursor = window.editor.textCursor()
                cursor.movePosition(QTextCursor.MoveOperation.End)
                cursor.insertBlock()
                window.editor.setTextCursor(cursor)
                insert_node(window.editor, window.state,
                            {"kind": "caption", "label": "Figure", "text": "A reproducible local workflow"}, "Figure")
                window._refresh_references()
                window.resize(1440, 960)
                window.show()
                output = workspace / "artifacts"
                output.mkdir(exist_ok=True)
                for name in ("References", "Mailings", "Review", "View", "Help"):
                    index = next(index for index in range(window.ribbon.count()) if window.ribbon.tabText(index) == name)
                    window.ribbon.setCurrentIndex(index)
                    app.processEvents()
                    path = output / f"folio-native-{name.lower()}.png"
                    if not window.grab().save(str(path)) or path.stat().st_size < 1000:
                        raise RuntimeError("Screenshot failed: " + name)
                view = window._open_view("Read Mode")
                app.processEvents()
                if not view.grab().save(str(output / "folio-native-read-mode.png")):
                    raise RuntimeError("Read Mode screenshot failed.")
                print("Native ribbon smoke passed; all five ribbons and Read Mode captured with native fonts.")
            finally:
                if window is not None:
                    window._dirty = False
                    window.close()
                app.processEvents()
                for widget in app.topLevelWidgets():
                    widget.close()
                app.processEvents()
                store.close()
                if any(widget.isVisible() for widget in app.topLevelWidgets()):
                    raise RuntimeError("A smoke window remained open.")
                app.quit()
    finally:
        for key, value in (("QT_QPA_PLATFORM", previous_platform), ("FOLIO_DATA_DIR", previous_data)):
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


if __name__ == "__main__":
    main()
