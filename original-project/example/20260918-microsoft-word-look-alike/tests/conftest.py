import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pytest


@pytest.fixture(scope="session")
def qapp():
    from PySide6.QtWidgets import QApplication
    from folio.app import configure_application

    app = QApplication.instance()
    if app is None:
        app = QApplication([])
        configure_application(app)
    yield app
    app.processEvents()


@pytest.fixture
def store(tmp_path):
    from folio.storage import LocalStore

    local = LocalStore(tmp_path / "folio-data")
    yield local
    local.close()


@pytest.fixture
def editor(qapp):
    from folio.editor import RichEditor
    from folio.models import PageSettings

    widget = RichEditor()
    widget.apply_page_settings(PageSettings())
    yield widget
    widget.deleteLater()
