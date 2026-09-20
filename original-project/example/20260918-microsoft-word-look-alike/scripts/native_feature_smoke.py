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
        with tempfile.TemporaryDirectory(prefix="folio-native-features-") as data_dir:
            os.environ["FOLIO_DATA_DIR"] = data_dir
            from PySide6.QtGui import QFont, QRawFont, QTextCursor
            from PySide6.QtWidgets import QApplication
            from folio.app import configure_application
            from folio.design import DesignSettings, apply_design
            from folio.document_features import insert_node, uid
            from folio.models import DocumentState
            from folio.storage import LocalStore
            from folio.structured_media import render_equation
            from folio.window import FolioWindow

            app = QApplication([])
            configure_application(app)
            if app.platformName() != "windows":
                raise RuntimeError("Native smoke must use the Windows platform plugin.")
            for family, text in (("Segoe UI", "Folio 123"), ("Calibri", "Native Windows fonts"),
                                 ("Cambria Math", "x² + √y = ∑n")):
                raw = QRawFont.fromFont(QFont(family, 12))
                if not raw.isValid():
                    raise RuntimeError(f"Native font unavailable: {family}")
                for character, glyph in zip(text, raw.glyphIndexesForString(text)):
                    if not character.isspace() and glyph == 0:
                        raise RuntimeError(f"Native font {family} lacks {character!r}")
            store = LocalStore()
            window = None
            try:
                if store.root.resolve() != Path(data_dir).resolve():
                    raise RuntimeError("Native smoke must use its temporary local store.")
                store.set_setting("proofing", False)
                settings = DesignSettings(theme="Woodland", page_border="box",
                                          watermark_kind="text", watermark_text="LOCAL DRAFT")
                state = DocumentState(title="Native feature smoke", html=(
                    "<h1>Native Windows rendering</h1>"
                    "<p>A local document with semantic styles, a rendered equation, and a floating text container.</p>"
                    "<p>Structured equation:</p><p></p>"))
                state.design = settings.to_dict()
                window = FolioWindow(store=store, restore=False)
                window.load_state(state, clean=True)
                apply_design(window.editor.document(), settings)
                cursor = window.editor.textCursor()
                cursor.movePosition(QTextCursor.MoveOperation.End)
                window.editor.setTextCursor(cursor)
                source = r"\frac{x^{2}+1}{\sqrt{y}} = \sum_{n=1}^{5}n"
                insert_node(window.editor, state, {"kind": "equation", "syntax": "tex", "source": source},
                            image=render_equation(source))
                state.features["objects"].append({
                    "id": uid(), "kind": "text", "name": "Local caption", "text": "Floating objects stay on disk with the document.",
                    "x": 145, "y": 430, "width": 310, "height": 95, "page": 0, "z": 0,
                    "wrap": "square", "visible": True, "fill": "#e4f1e8", "stroke": "#166534",
                })
                window._floating_inspector()
                window.floating_dock.widget().select_ids({state.features["objects"][0]["id"]})
                window.ribbon.setCurrentIndex(2)
                window.resize(1440, 960)
                window.show()
                app.processEvents()
                target = workspace / "artifacts" / "folio-native-features.png"
                target.parent.mkdir(exist_ok=True)
                if not window.grab().save(str(target)) or target.stat().st_size < 1000:
                    raise RuntimeError("The native screenshot was not saved.")
                print(f"Native feature smoke passed; Segoe UI, Calibri, and Cambria Math glyphs verified. Screenshot: {target}")
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
                    raise RuntimeError("A native smoke window remained visible.")
                app.quit()
    finally:
        if previous_platform is None:
            os.environ.pop("QT_QPA_PLATFORM", None)
        else:
            os.environ["QT_QPA_PLATFORM"] = previous_platform
        if previous_data is None:
            os.environ.pop("FOLIO_DATA_DIR", None)
        else:
            os.environ["FOLIO_DATA_DIR"] = previous_data


if __name__ == "__main__":
    main()
