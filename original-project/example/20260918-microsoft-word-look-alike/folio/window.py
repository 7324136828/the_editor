from __future__ import annotations

import copy
import html as html_module
import time
from pathlib import Path

from PySide6.QtCore import (
    QObject,
    QPoint,
    QRectF,
    QRunnable,
    QSignalBlocker,
    Qt,
    QThreadPool,
    QTimer,
    Signal,
)
from PySide6.QtGui import (
    QAction,
    QColor,
    QFont,
    QKeySequence,
    QPainter,
    QPen,
    QShortcut,
    QTextBlockFormat,
    QTextCharFormat,
    QTextCursor,
    QTextFormat,
    QSyntaxHighlighter,
)
from PySide6.QtPrintSupport import (
    QPrintDialog,
    QPrinter,
    QPrintPreviewDialog,
)
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDockWidget,
    QFileDialog,
    QFontComboBox,
    QFrame,
    QGraphicsDropShadowEffect,
    QGraphicsProxyWidget,
    QGraphicsRectItem,
    QGraphicsScene,
    QGraphicsView,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSlider,
    QStatusBar,
    QTabWidget,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from . import authoring, dialogs
from .docx_io import DocxError, atomic_write, read_docx, write_docx
from .editor import RichEditor, STYLE_ORDER, mm_to_px, pt_to_px
from .icons import icon, mark
from .models import (DocumentState, PageSettings, TocRegion, ValidationError,
                     sanitize_ranges)
from .proofing import GRAMMAR, STYLE, WORD, ProofingEngine
from .publishing import Publication, page_layout_for, write_pdf
from .review import ReviewError, ReviewTracker
from .storage import LocalStore
from .templates import make_template
from .feature_commands import FeatureCommands
from .view_commands import ViewCommands
from .reference_commands import ReferenceCommands
from .mailing_commands import MailingCommands
from .review_commands import ReviewCommands
from . import document_features
from .design import DesignSettings, paint_page_decoration

AUTOSAVE_DEBOUNCE_MS = 1500
AUTOSAVE_SAFETY_MS = 30_000
VERSION_INTERVAL_S = 300
PROOF_DEBOUNCE_MS = 600


class SpellHighlighter(QSyntaxHighlighter):
    def __init__(self, document):
        super().__init__(document)
        self._issues: list = []
        self._enabled = False

    def set_issues(self, issues: list) -> None:
        self._issues = issues
        with QSignalBlocker(self.document()):
            self.rehighlight()

    def set_enabled(self, enabled: bool) -> None:
        self._enabled = enabled
        with QSignalBlocker(self.document()):
            self.rehighlight()

    def highlightBlock(self, text: str) -> None:
        if not self._enabled or not self._issues:
            return
        block = self.currentBlock()
        start = block.position()
        end = start + block.length()
        for issue in self._issues:
            issue_end = issue.start + issue.length
            if issue.start >= end or issue_end <= start:
                continue
            fmt = QTextCharFormat()
            fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SpellCheckUnderline)
            if issue.kind == WORD:
                fmt.setUnderlineColor(QColor("#d64545"))
            elif issue.kind == GRAMMAR:
                fmt.setUnderlineColor(QColor("#2e8540"))
            else:
                fmt.setUnderlineColor(QColor("#b26a00"))
            local_start = max(0, issue.start - start)
            local_len = min(issue_end, end) - max(issue.start, start)
            self.setFormat(local_start, local_len, fmt)


class _ProofEmitter(QObject):
    finished = Signal(int, object)


class _ProofJob(QRunnable):
    def __init__(self, engine: ProofingEngine, text: str,
                 generation: int, emitter: _ProofEmitter):
        super().__init__()
        self.engine = engine
        self.text = text
        self.generation = generation
        self.emitter = emitter

    def run(self) -> None:
        try:
            issues = self.engine.check(self.text)
        except Exception:
            issues = None
        self.emitter.finished.emit(self.generation, issues)


class RulerWidget(QWidget):
    def __init__(self, canvas: "DocumentCanvas"):
        super().__init__(canvas)
        self.canvas = canvas
        self.setFixedHeight(24)

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#f5f7fa"))
        painter.setPen(QColor("#9aa7b8"))
        metrics = self.canvas.page_metrics()
        if metrics is None:
            painter.end()
            return
        page_w_px, origin_x, zoom = metrics
        px_per_mm = mm_to_px(1.0) * zoom
        if px_per_mm < 2:
            painter.end()
            return
        mm = 0
        x = origin_x
        while x < self.width() + px_per_mm:
            if x >= -1:
                tick = 10 if mm % 10 == 0 else (7 if mm % 5 == 0 else 4)
                painter.drawLine(int(x), self.height() - tick, int(x), self.height())
                if mm % 10 == 0 and mm > 0:
                    painter.drawText(int(x) + 2, 12, str(mm // 10))
            mm += 1
            x += px_per_mm
        painter.setPen(QColor("#c7d0dc"))
        painter.drawLine(0, self.height() - 1, self.width(), self.height() - 1)
        painter.end()


class DocumentPaper(QGraphicsRectItem):
    def __init__(self, editor):
        super().__init__()
        self.editor = editor

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)
        state = getattr(self.editor.document(), "folio_state", None)
        if state is None:
            return
        settings = DesignSettings.from_dict(state.design)
        rect = self.rect()
        page_h = mm_to_px(state.page.size_mm()[1])
        for index in range(int(rect.height() / page_h) + 1):
            painter.save()
            painter.translate(rect.left(), rect.top() + index * page_h)
            paint_page_decoration(painter, QRectF(0, 0, rect.width(), page_h), settings)
            document_features.paint_objects(painter, state.features["objects"], index, "behind")
            painter.restore()


class DocumentCanvas(QGraphicsView):
    def __init__(self, editor: RichEditor, parent=None):
        super().__init__(parent)
        self.editor = editor
        self._scene = QGraphicsScene(self)
        self._paper = DocumentPaper(editor)
        editor.setStyleSheet("QTextEdit { background: transparent; border: none; }")
        editor.viewport().setAutoFillBackground(False)
        self._paper.setBrush(QColor("white"))
        self._paper.setPen(Qt.PenStyle.NoPen)
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(24)
        shadow.setColor(QColor(0, 0, 0, 70))
        shadow.setOffset(0, 4)
        self._paper.setGraphicsEffect(shadow)
        self._scene.addItem(self._paper)
        self._proxy: QGraphicsProxyWidget = self._scene.addWidget(editor)
        self._proxy.setPos(48, 48)
        self.setScene(self._scene)
        self.setBackgroundBrush(QColor("#edf0f4"))
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self._zoom = 1.0
        self.gridlines = False
        self._page_w = 794.0
        self._page_h = 1123.0
        editor.document().contentsChanged.connect(self.refresh_layout)
        editor.cursorPositionChanged.connect(self.scroll_to_cursor)
        self.refresh_layout()

    def set_page_metrics(self, page_w_px: float, page_h_px: float) -> None:
        self._page_w = page_w_px
        self._page_h = page_h_px
        self.editor.setFixedWidth(int(page_w_px))
        self.refresh_layout()

    def page_metrics(self):
        x = self.mapFromScene(self._proxy.pos()).x()
        return self._page_w, x, self._zoom

    def refresh_layout(self) -> None:
        doc = self.editor.document()
        page = doc.pageSize()
        self.editor._sync_height()
        content_h = max(page.height(), self.editor.height() + 2)
        self._paper.setRect(QRectF(self._proxy.pos().x(), self._proxy.pos().y(),
                                   self._page_w, content_h))
        margin = 48
        self._scene.setSceneRect(0, 0, self._page_w + margin * 2,
                                 content_h + margin * 2)
        self.viewport().update()

    def drawForeground(self, painter: QPainter, rect: QRectF) -> None:
        super().drawForeground(painter, rect)
        page_h = self._page_h
        if page_h <= 0:
            return
        origin = self._proxy.pos()
        content_h = self._paper.rect().height()
        state = getattr(self.editor.document(), "folio_state", None)
        if state:
            settings = DesignSettings.from_dict(state.design)
            for page in range(int(content_h / page_h) + 1):
                painter.save()
                painter.translate(origin.x(), origin.y() + page * page_h)
                document_features.paint_objects(painter, state.features["objects"], page, "front")
                paint_page_decoration(painter, QRectF(0, 0, self._page_w, page_h), settings, "front")
                painter.restore()
        if self.gridlines:
            painter.save()
            painter.setClipRect(self._paper.rect().intersected(rect))
            painter.setPen(QPen(QColor(100, 125, 155, 65), 0.5))
            step = mm_to_px(5)
            left = max(origin.x(), rect.left())
            top = max(origin.y(), rect.top())
            right = min(origin.x() + self._page_w, rect.right())
            bottom = min(origin.y() + content_h, rect.bottom())
            x = origin.x() + int((left - origin.x()) / step) * step
            while x <= right:
                painter.drawLine(QPoint(int(x), int(top)), QPoint(int(x), int(bottom)))
                x += step
            y_grid = origin.y() + int((top - origin.y()) / step) * step
            while y_grid <= bottom:
                painter.drawLine(QPoint(int(left), int(y_grid)), QPoint(int(right), int(y_grid)))
                y_grid += step
            painter.restore()
        pen = QPen(QColor("#c7d0dc"))
        pen.setStyle(Qt.PenStyle.DashLine)
        painter.setPen(pen)
        y = origin.y() + page_h
        while y < origin.y() + content_h - 4:
            painter.drawLine(QPoint(int(origin.x() + 8), int(y)),
                             QPoint(int(origin.x() + self._page_w - 8), int(y)))
            y += page_h

    def set_zoom(self, factor: float) -> None:
        factor = max(0.1, min(5.0, factor))
        self._zoom = factor
        self.resetTransform()
        self.scale(factor, factor)
        self.viewport().update()

    def zoom(self) -> float:
        return self._zoom

    def zoom_fit_width(self) -> None:
        available = self.viewport().width() - 120
        if available > 0:
            self.set_zoom(available / self._page_w)

    def scroll_to_cursor(self) -> None:
        rect = self.editor.cursorRect()
        if rect.isNull():
            return
        scene_rect = self._proxy.mapToScene(rect).boundingRect()
        self.ensureVisible(scene_rect, 40, 40)


class FolioWindow(ReferenceCommands, MailingCommands, ReviewCommands, ViewCommands, FeatureCommands, QMainWindow):
    def __init__(self, store: LocalStore | None = None, restore: bool = True):
        super().__init__()
        self.store = store if store is not None else LocalStore()
        self.state = DocumentState()
        self._path: Path | None = None
        self._dirty = False
        self._saved_mtime_ns: int | None = None
        self._title = "Untitled"
        self._author = str(self.store.get_setting("author", "You"))
        self._last_version_time = time.monotonic()
        self._proof_generation = 0
        self._proof_enabled = bool(self.store.get_setting("proofing", True))
        paste_mode = str(self.store.get_setting(
            "paste_mode", "keep_formatting"))
        self._paste_mode = (paste_mode if paste_mode in (
            "keep_formatting", "text", "image") else "keep_formatting")
        self._ignored_words: set = set()
        self._focus = False
        self._pre_focus = (True, True)
        self._show_comments = True
        self._show_changes = True
        self._secondary_windows = []

        self.setWindowTitle("Folio")
        screen = QApplication.primaryScreen()
        available = (screen.availableGeometry()
                     if screen is not None else None)
        win_w, win_h = 1440, 960
        if available is not None and available.width() > 0:
            win_w = min(win_w, max(800, available.width() - 40))
            win_h = min(win_h, max(600, available.height() - 60))
        self.resize(win_w, win_h)
        self.editor = RichEditor()
        self._initialize_features()
        self.editor.apply_page_settings(self.state.page)
        self.tracker = ReviewTracker(self.editor.document(), self.state)
        self.tracker.author = self._author
        self.proof_engine = ProofingEngine(self.store)
        self.highlighter = SpellHighlighter(self.editor.document())
        self._proof_emitter = _ProofEmitter()
        self._proof_emitter.finished.connect(self._proof_results)
        self._proof_issues: list = []
        self._pool = QThreadPool(self)

        self._build_chrome()
        self._build_ribbon()
        chrome = QWidget()
        chrome_layout = QVBoxLayout(chrome)
        chrome_layout.setContentsMargins(0, 0, 0, 0)
        chrome_layout.setSpacing(0)
        chrome_layout.addWidget(self._header)
        chrome_layout.addWidget(self.ribbon)
        self.setMenuWidget(chrome)
        self._build_docks()
        self._build_statusbar()

        self.canvas = DocumentCanvas(self.editor)
        self.canvas.set_page_metrics(
            mm_to_px(self.state.page.size_mm()[0]),
            mm_to_px(self.state.page.size_mm()[1]),
        )
        center = QWidget()
        center_layout = QVBoxLayout(center)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(0)
        self.ruler = RulerWidget(self.canvas)
        center_layout.addWidget(self.ruler, 0)
        center_layout.addWidget(self.canvas, 1)
        self.setCentralWidget(center)
        self.ribbon.setCurrentIndex(1)
        self.canvas.horizontalScrollBar().valueChanged.connect(
            lambda _v: self.ruler.update())

        self._autosave_timer = QTimer(self)
        self._autosave_timer.setSingleShot(True)
        self._autosave_timer.setInterval(AUTOSAVE_DEBOUNCE_MS)
        self._autosave_timer.timeout.connect(self._autosave)
        self._safety_timer = QTimer(self)
        self._safety_timer.setInterval(AUTOSAVE_SAFETY_MS)
        self._safety_timer.timeout.connect(self._autosave)
        self._safety_timer.start()
        self._outline_timer = QTimer(self)
        self._outline_timer.setSingleShot(True)
        self._outline_timer.setInterval(500)
        self._outline_timer.timeout.connect(self._refresh_outline)
        self._proof_timer = QTimer(self)
        self._proof_timer.setSingleShot(True)
        self._proof_timer.setInterval(PROOF_DEBOUNCE_MS)
        self._proof_timer.timeout.connect(self._run_proofing)
        self._counts_timer = QTimer(self)
        self._counts_timer.setSingleShot(True)
        self._counts_timer.setInterval(400)
        self._counts_timer.timeout.connect(self._refresh_counts)
        self._initialize_references()

        self.editor.document().contentsChanged.connect(self._document_changed)
        self.tracker.changed.connect(self._review_changed)
        self.editor.cursorPositionChanged.connect(self._sync_toolbar)
        self.editor.currentCharFormatChanged.connect(self._sync_toolbar)
        self.setup_editor_context_menu()
        self._initialize_review_services()

        for sequence, slot in (
            (QKeySequence.StandardKey.Save, self._save),
            (QKeySequence.StandardKey.Open, self._open_document),
            (QKeySequence.StandardKey.New, self._new_document),
            (QKeySequence.StandardKey.Print, self._print),
            (QKeySequence.StandardKey.Find, self._find),
            (QKeySequence.StandardKey.Replace, self._replace),
        ):
            QShortcut(sequence, self, slot)
        QShortcut(QKeySequence(Qt.Key.Key_Escape), self, self._exit_focus)

        if self._proof_enabled:
            self.highlighter.set_enabled(True)
            self._proof_timer.start()

        if restore:
            self._offer_recovery()
        if self.state.html == "":
            self.load_state(make_template("welcome"), None, clean=True)
        self._sync_toolbar()


    def _build_chrome(self) -> None:
        header = QWidget()
        header.setFixedHeight(46)
        header.setStyleSheet("background:#185abd;")
        row = QHBoxLayout(header)
        row.setContentsMargins(14, 0, 14, 0)
        row.setSpacing(10)
        logo = QLabel()
        logo.setPixmap(mark(20))
        name = QLabel("Folio")
        name.setStyleSheet("color:white;font-weight:600;font-size:15px;")
        row.addWidget(logo)
        row.addWidget(name)
        row.addSpacing(16)
        for icon_name, tip, slot in (
            ("save", "Save (Ctrl+S)", self._save),
            ("undo", "Undo", self.editor.undo),
            ("redo", "Redo", self.editor.redo),
        ):
            button = QToolButton()
            button.setIcon(icon(icon_name, "#ffffff"))
            button.setToolTip(tip)
            button.setStyleSheet("QToolButton{border:none;padding:4px;}"
                                 "QToolButton:hover{background:#2f6fd0;}")
            if icon_name == "save":
                button.clicked.connect(slot)
            else:
                button.clicked.connect(lambda _checked=False, action=slot: action() if self._can_edit() else None)
            row.addWidget(button)
        row.addStretch(1)
        self.title_edit = QLabel()
        self.title_edit.setStyleSheet("color:white;font-size:13px;")
        row.addWidget(self.title_edit)
        row.addStretch(1)
        self._exit_focus_btn = QToolButton()
        self._exit_focus_btn.setText("Exit focus")
        self._exit_focus_btn.setToolTip(
            "Leave focus mode (Esc)")
        self._exit_focus_btn.setStyleSheet(
            "QToolButton{color:white;border:1px solid #7aa7e8;"
            "border-radius:3px;padding:3px 10px;}"
            "QToolButton:hover{background:#2f6fd0;}")
        self._exit_focus_btn.setVisible(False)
        self._exit_focus_btn.clicked.connect(self._exit_focus)
        row.addWidget(self._exit_focus_btn)
        self._header = header
        self.permission_combo = QComboBox()
        self.permission_combo.setObjectName("permission_mode")
        self.permission_combo.addItems(["Editing", "Reviewing", "Viewing"])
        self.permission_combo.currentTextChanged.connect(self._permission_mode)
        self.permission_combo.setStyleSheet("background:white;color:#263244;padding:3px;")
        row.addWidget(self.permission_combo)

    def _tool(self, icon_name: str, text: str, slot=None, tip: str = "",
              *, text_mode: bool = False, checkable: bool = False) -> QToolButton:
        button = QToolButton()
        button.setIcon(icon(icon_name))
        if text_mode:
            button.setText(text)
            button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        else:
            button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
        button.setToolTip(tip or text)
        button.setCheckable(checkable)
        if slot is not None:
            if checkable:
                button.toggled.connect(slot)
            else:
                button.clicked.connect(slot)
        return button

    @staticmethod
    def _vsep() -> QFrame:
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.VLine)
        sep.setStyleSheet("color:#d5dce5;")
        return sep

    def _group(self, title: str) -> tuple[QFrame, QHBoxLayout]:
        frame = QFrame()
        frame.setStyleSheet("QFrame{border:none;}")
        outer = QVBoxLayout(frame)
        outer.setContentsMargins(6, 4, 6, 2)
        outer.setSpacing(2)
        inner = QHBoxLayout()
        inner.setSpacing(2)
        outer.addLayout(inner)
        label = QLabel(title)
        label.setStyleSheet("color:#7d8aa0;font-size:10px;")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(label)
        return frame, inner

    def _build_ribbon(self) -> None:
        self.ribbon = QTabWidget()
        self.ribbon.setDocumentMode(True)
        self.ribbon.setStyleSheet(
            "QTabWidget::pane{border-top:1px solid #d5dce5;background:white;}"
            "QTabBar::tab{background:#eef1f6;color:#33415c;padding:6px 18px;"
            "border-top-left-radius:6px;border-top-right-radius:6px;}"
            "QTabBar::tab:selected{background:white;color:#185abd;"
            "border-bottom:2px solid #185abd;}"
        )
        self._build_file_tab()
        self._build_home_tab()
        self._build_insert_tab()
        self._build_design_tab()
        self._build_layout_tab()
        self._build_references_tab()
        self._build_mailings_tab()
        self._build_review_tab()
        self._build_view_tab()
        self._build_help_tab()
        self.ribbon.setFixedHeight(124)

    @staticmethod
    def _extra_row(group: QFrame) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(2)
        group.layout().insertLayout(group.layout().count() - 1, row)
        return row

    def _wrap(self, groups: list[QFrame]) -> QWidget:
        inner = QWidget()
        row = QHBoxLayout(inner)
        row.setContentsMargins(8, 2, 8, 4)
        row.setSpacing(4)
        for i, frame in enumerate(groups):
            if i:
                row.addWidget(self._vsep())
            row.addWidget(frame)
        row.addStretch(1)
        scroll = QScrollArea()
        scroll.setWidget(inner)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        return scroll

    def _build_file_tab(self) -> None:
        groups = []
        g, l = self._group("Document")
        l.addWidget(self._tool("new", "New", self._new_document,
                               "New from template", text_mode=True))
        l.addWidget(self._tool("open", "Open", self._open_document,
                               "Open (Ctrl+O)", text_mode=True))
        l.addWidget(self._tool("save", "Save", self._save,
                               "Save (Ctrl+S)", text_mode=True))
        l.addWidget(self._tool("save", "Save As", self._save_as,
                               "Save As (DOCX)", text_mode=True))
        l.addWidget(self._tool("review", "Close", self._close_document,
                               "Close document", text_mode=True))
        groups.append(g)
        g, l = self._group("Output")
        l.addWidget(self._tool("pdf", "PDF", self._export_pdf,
                               "Export as PDF", text_mode=True))
        l.addWidget(self._tool("preview", "Preview", self._print_preview,
                               "Print preview", text_mode=True))
        l.addWidget(self._tool("print", "Print", self._print,
                               "Print (Ctrl+P)", text_mode=True))
        l.addWidget(self._tool("copy", "Export HTML", self._export_html,
                               "Export as HTML", text_mode=True))
        groups.append(g)
        g, l = self._group("More")
        l.addWidget(self._tool("history", "History", self._show_history,
                               "Version history", text_mode=True))
        l.addWidget(self._tool("save", "Checkpoint", self._checkpoint,
                               "Save a manual checkpoint to history",
                               text_mode=True))
        l.addWidget(self._tool("about", "About", self._about,
                               "About Folio", text_mode=True))
        groups.append(g)
        self.ribbon.addTab(self._wrap(groups), "File")

    def _build_home_tab(self) -> None:
        groups = []
        g, l = self._group("Clipboard")
        l.addWidget(self._tool("cut", "Cut", self.editor.cut, "Cut"))
        l.addWidget(self._tool("copy", "Copy", self.editor.copy, "Copy"))
        l.addWidget(self._tool("paste", "Paste", self._paste, "Paste"))
        l.addWidget(self._tool("paste", "Special", self._paste_special,
                               "Paste Special", text_mode=True))
        l2 = self._extra_row(g)
        l2.addWidget(self._tool("undo", "Undo", self.editor.undo, "Undo"))
        l2.addWidget(self._tool("redo", "Redo", self.editor.redo, "Redo"))
        l2.addWidget(self._tool("pagesetup", "Default paste",
                                self._set_default_paste,
                                "Choose the default paste format",
                                text_mode=True))
        painter_menu = QMenu(self)
        painter_menu.addAction("Sample Format", self._sample_format)
        painter_menu.addAction("Apply Format", self._apply_format)
        l2.addWidget(self._menu_tool("Format Painter", painter_menu))
        groups.append(g)
        g, l = self._group("Font")
        self.font_combo = QFontComboBox()
        self.font_combo.setCurrentFont(QFont("Calibri"))
        self.font_combo.setFixedWidth(150)
        self.font_combo.currentFontChanged.connect(
            lambda f: self.editor.set_font_family(f.family()))
        l.addWidget(self.font_combo)
        self.size_combo = QComboBox()
        self.size_combo.setEditable(True)
        self.size_combo.setFixedWidth(56)
        for size in (8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72):
            self.size_combo.addItem(str(size))
        self.size_combo.setCurrentText("11")
        self.size_combo.activated.connect(self._font_size_chosen)
        self.size_combo.lineEdit().editingFinished.connect(
            self._font_size_edited)
        l.addWidget(self.size_combo)
        l.addWidget(self._tool("heading", "Grow font", lambda: self.editor.step_font_size(1)))
        l.addWidget(self._tool("heading", "Shrink font", lambda: self.editor.step_font_size(-1)))
        case_menu = QMenu(self)
        for label, mode in (("Sentence case", "sentence"), ("lowercase", "lower"),
                            ("UPPERCASE", "upper"), ("Title Case", "title")):
            case_menu.addAction(label, lambda _checked=False, value=mode: self.editor.change_case(value))
        l.addWidget(self._menu_tool("Aa", case_menu))
        l = self._extra_row(g)
        self.bold_button = self._tool("bold", "Bold", self.editor.toggle_bold,
                                      "Bold (Ctrl+B)", checkable=True)
        self.italic_button = self._tool("italic", "Italic",
                                        self.editor.toggle_italic,
                                        "Italic (Ctrl+I)", checkable=True)
        self.underline_button = self._tool(
            "underline", "Underline", self.editor.toggle_underline,
            "Underline (Ctrl+U)", checkable=True)
        self.strike_button = self._tool("strike", "Strike",
                                        self.editor.toggle_strikethrough,
                                        "Strikethrough", checkable=True)
        for b in (self.bold_button, self.italic_button,
                  self.underline_button, self.strike_button):
            l.addWidget(b)
        l.addWidget(self._tool("field", "Subscript", self.editor.toggle_subscript, text_mode=True))
        l.addWidget(self._tool("field", "Superscript", self.editor.toggle_superscript, text_mode=True))
        l.addWidget(self._tool("color", "Text color", self._pick_text_color,
                               "Text color"))
        l.addWidget(self._tool("highlight", "Highlight",
                               self._pick_highlight, "Highlight color"))
        l.addWidget(self._tool("clear", "No HL",
                               lambda: self.editor.set_highlight(None),
                               "Clear highlight"))
        l.addWidget(self._tool("clear", "Clear", self._clear_format,
                               "Clear formatting"))
        l.addWidget(self._tool("shape", "Shadow",
                               self.editor.toggle_text_shadow,
                               "Toggle text shadow"))
        groups.append(g)
        g, l = self._group("Style")
        self.style_combo = QComboBox()
        self.style_combo.addItems(STYLE_ORDER)
        self.style_combo.activated.connect(
            lambda _i: self.editor.apply_style(self.style_combo.currentText()))
        l.addWidget(self.style_combo)
        l = self._extra_row(g)
        for name in ("Normal", "Title", "Heading 1", "Heading 2", "Quote"):
            b = QToolButton()
            b.setText(name.replace("Heading ", "H"))
            b.setToolTip(f"Apply {name}")
            b.setStyleSheet("padding:3px 6px;")
            b.clicked.connect(lambda _c, n=name: self.editor.apply_style(n))
            l.addWidget(b)
        groups.append(g)
        g, l = self._group("Paragraph")
        self.align_buttons = {}
        for name, flag in (("align-left", Qt.AlignmentFlag.AlignLeft),
                           ("align-center", Qt.AlignmentFlag.AlignHCenter),
                           ("align-right", Qt.AlignmentFlag.AlignRight),
                           ("align-justify", Qt.AlignmentFlag.AlignJustify)):
            b = self._tool(name, name.replace("align-", "Align "),
                           None, name.replace("align-", "Align "),
                           checkable=True)
            b.clicked.connect(lambda _c, f=flag: self._set_alignment(f))
            self.align_buttons[flag] = b
            l.addWidget(b)
        l.addWidget(self._tool("bullets", "Bullets",
                               lambda: self.editor.set_list(False), "Bullets"))
        l.addWidget(self._tool("numbering", "Numbering",
                               lambda: self.editor.set_list(True), "Numbering"))
        l.addWidget(self._tool("outdent", "Outdent",
                               lambda: self.editor.indent_paragraph(-1),
                               "Decrease indent"))
        l.addWidget(self._tool("indent", "Indent",
                               lambda: self.editor.indent_paragraph(1),
                               "Increase indent"))
        l = self._extra_row(g)
        self.spacing_combo = QComboBox()
        for label, value in (("1.0", 1.0), ("1.15", 1.15),
                             ("1.5", 1.5), ("2.0", 2.0)):
            self.spacing_combo.addItem(label, value)
        self.spacing_combo.setCurrentIndex(1)
        self.spacing_combo.activated.connect(
            lambda _i: self.editor.set_line_spacing(
                float(self.spacing_combo.currentData())))
        l.addWidget(self.spacing_combo)
        l.addWidget(self._tool("pagesetup", "Para", self._paragraph_dialog,
                               "Paragraph settings"))
        l.addWidget(self._tool("numbering", "Multilevel", self.editor.set_multilevel_list, text_mode=True))
        adornment_menu = QMenu(self)
        adornment_menu.addAction("Shading…", self._paragraph_shading)
        adornment_menu.addAction("Clear Shading", lambda: self.editor.set_paragraph_shading(None))
        for border in ("none", "top", "bottom", "left", "right", "box"):
            adornment_menu.addAction(border.title() + " Border", lambda _checked=False, value=border: self.editor.set_paragraph_border(value))
        l.addWidget(self._menu_tool("Borders & Shading", adornment_menu))
        groups.append(g)
        g, l = self._group("Editing")
        l.addWidget(self._tool("find", "Find", self._find,
                               "Find (Ctrl+F)", text_mode=True))
        l.addWidget(self._tool("field", "Transcription", self._dictation, text_mode=True))
        l2 = self._extra_row(g)
        l2.addWidget(self._tool("find", "Replace", self._replace,
                                "Replace (Ctrl+H)", text_mode=True))
        l2.addWidget(self._tool("nav", "Select All", self.editor.selectAll,
                                "Select all content", text_mode=True))
        l2.addWidget(self._tool("image", "Selection Pane",
                                self._show_object_pane,
                                "List document objects", text_mode=True))
        groups.append(g)
        self.ribbon.addTab(self._wrap(groups), "Home")

    def _build_insert_tab(self) -> None:
        groups = []
        g, l = self._group("Pages")
        l.addWidget(self._tool("pagesetup", "Cover Page",
                               self._insert_cover_page,
                               "Insert a designed cover page", text_mode=True))
        l.addWidget(self._tool("break", "Blank Page", self._blank_page, text_mode=True))
        groups.append(g)
        g, l = self._group("Illustrations")
        l.addWidget(self._tool("image", "Image", self._insert_image,
                               "Insert image", text_mode=True))
        l.addWidget(self._tool("chart", "Chart", self._insert_chart,
                               "Insert chart", text_mode=True))
        l.addWidget(self._tool("image", "Screenshot",
                               self._insert_screenshot,
                               "Capture the primary screen", text_mode=True))
        l.addWidget(self._tool("shape", "Shape", self._insert_shape,
                               "Insert shape", text_mode=True))
        l.addWidget(self._tool("diagram", "SmartArt", self._insert_diagram,
                               "Insert a structured diagram", text_mode=True))
        l.addWidget(self._tool("shape", "3D Model", self._model_insert, text_mode=True))
        groups.append(g)
        g, l = self._group("Table")
        l.addWidget(self._tool("table", "Table", self._insert_table,
                               "Insert table", text_mode=True))
        self._table_tools = []
        for icon_name, text, slot in (
            ("table", "Dimensions", self._table_dimensions),
            ("add-row", "+Row", lambda: self._table_op("add_row")),
            ("add-col", "+Col", lambda: self._table_op("add_col")),
            ("del-row", "-Row", lambda: self._table_op("del_row")),
            ("del-col", "-Col", lambda: self._table_op("del_col")),
            ("merge-cells", "Merge", lambda: self._table_op("merge")),
        ):
            b = self._tool(icon_name, text, slot, text)
            self._table_tools.append(b)
            l.addWidget(b)
        groups.append(g)
        g, l = self._group("Insert")
        l.addWidget(self._tool("link", "Link", self._insert_link,
                               "Insert hyperlink", text_mode=True))
        l.addWidget(self._tool("link", "Bookmark", self._bookmark, text_mode=True))
        l.addWidget(self._tool("link", "Cross-reference", self._cross_reference, text_mode=True))
        l.addWidget(self._tool("break", "Break", self.editor.insert_page_break,
                               "Page break (Ctrl+Enter)", text_mode=True))
        l.addWidget(self._tool("rule", "Rule", self.editor.insert_horizontal_rule,
                               "Horizontal rule", text_mode=True))
        l.addWidget(self._tool("date", "Date", self._insert_date,
                               "Insert current date", text_mode=True))
        l.addWidget(self._tool("field", "Field", self._insert_field,
                               "Insert {{merge field}}", text_mode=True))
        l.addWidget(self._tool("merge", "Merge", self._mail_merge,
                               "Mail merge (CSV/XLSX)", text_mode=True))
        groups.append(g)
        g, l = self._group("Media & Text")
        for label, slot in (("Online Video", self._online_video),
                            ("Stock Media", self._stock_media),
                            ("WordArt", lambda: self._text_insert("WordArt"))):
            l.addWidget(self._tool("image", label, slot, text_mode=True))
        l = self._extra_row(g)
        for label, slot in (("Symbol", lambda: self._text_insert("Symbol")),
                            ("Equation", self._equation_editor),
                            ("Signature Fields", self._signature_fields)):
            l.addWidget(self._tool("field", label, slot, text_mode=True))
        groups.append(g)
        g, l = self._group("Objects")
        for label, slot in (("Text Box", self._float_insert),
                            ("Floating Image", lambda: self._float_insert(True)),
                            ("Drop Cap", self._drop_cap)):
            l.addWidget(self._tool("shape", label, slot, text_mode=True))
        l = self._extra_row(g)
        for label, slot in (("Embed File", self._attachment_insert),
                            ("Edit / Sign / Extract", self._edit_content_node)):
            l.addWidget(self._tool("field", label, slot, text_mode=True))
        groups.append(g)
        g, l = self._group("Header & Footer")
        page_menu = QMenu(self)
        for label, position in (("Top of Page", "top"), ("Bottom of Page", "bottom"),
                                ("Remove Page Numbers", "remove")):
            page_menu.addAction(label, lambda _checked=False, value=position: self._page_numbers(value))
        l.addWidget(self._menu_tool("Page Numbers", page_menu))
        groups.append(g)
        self.ribbon.addTab(self._wrap(groups), "Insert")

    def _build_design_tab(self) -> None:
        groups = []
        g, l = self._group("Document Formatting")
        for name in ("Title", "Subtitle", "Heading 1", "Heading 2", "Normal"):
            l.addWidget(self._tool(
                "heading", name,
                lambda _checked=False, value=name: self.editor.apply_style(value),
                f"Apply {name}", text_mode=True))
        groups.append(g)
        g, l = self._group("Page")
        l.addWidget(self._tool("pagesetup", "Page Setup", self._page_setup,
                               "Paper, margins, headers, and layout",
                               text_mode=True))
        groups.append(g)
        g, l = self._group("Themes, Colors & Fonts")
        l.addWidget(self._tool("color", "Document Theme", self._semantic_design,
                               "Apply colors, fonts, and a style set to the document",
                               text_mode=True))
        l.addWidget(self._tool("color", "Page Background", self._semantic_design, text_mode=True))
        groups.append(g)
        self.ribbon.addTab(self._wrap(groups), "Design")

    def _build_layout_tab(self) -> None:
        groups = []
        g, l = self._group("Page")
        l.addWidget(self._tool("pagesetup", "Setup", self._page_setup,
                               "Page setup", text_mode=True))
        l.addWidget(self._tool("portrait", "Portrait",
                               lambda: self._quick_orientation(False),
                               "Portrait", text_mode=True))
        l.addWidget(self._tool("landscape", "Landscape",
                               lambda: self._quick_orientation(True),
                               "Landscape", text_mode=True))
        groups.append(g)
        g, l = self._group("Margins")
        self.margins_combo = QComboBox()
        self.margins_combo.addItem("Normal (25.4 mm)")
        self.margins_combo.addItem("Narrow (12.7 mm)")
        self.margins_combo.addItem("Wide (38.1 mm)")
        self.margins_combo.addItem("Custom")
        self.margins_combo.activated.connect(self._margins_preset)
        l.addWidget(self.margins_combo)
        self.columns_combo = QComboBox()
        self.columns_combo.addItems(["1 column", "2 columns", "3 columns"])
        self.columns_combo.activated.connect(
            lambda i: self._quick_columns(i + 1))
        l.addWidget(self.columns_combo)
        groups.append(g)
        g, l = self._group("Preview")
        l.addWidget(self._tool("preview", "Preview", self._print_preview,
                               "Publishing preview", text_mode=True))
        groups.append(g)
        g, l = self._group("Flow")
        self.breaks_combo = QComboBox()
        self.breaks_combo.addItems(["Breaks", "Page", "Column", "Section: Next Page",
                                   "Section: Continuous", "Section: Even Page", "Section: Odd Page"])
        self.breaks_combo.activated.connect(self._insert_break)
        l.addWidget(self.breaks_combo)
        self.line_numbers_combo = QComboBox()
        self.line_numbers_combo.addItem("Line Numbers", None)
        self.line_numbers_combo.addItem("None", "none")
        self.line_numbers_combo.addItem("Continuous", "continuous")
        self.line_numbers_combo.addItem("Restart Each Page", "restart_page")
        self.line_numbers_combo.addItem("Restart Each Section",
                                        "restart_section")
        self.line_numbers_combo.addItem("Line Numbering Options", "options")
        self.line_numbers_combo.activated.connect(self._line_numbers_chosen)
        l.addWidget(self.line_numbers_combo)
        self.hyphenation_combo = QComboBox()
        self.hyphenation_combo.addItem("Hyphenation", None)
        self.hyphenation_combo.addItem("None", "none")
        self.hyphenation_combo.addItem("Automatic", "automatic")
        self.hyphenation_combo.addItem("Manual", "manual")
        self.hyphenation_combo.addItem("Hyphenation Options", "options")
        self.hyphenation_combo.activated.connect(self._hyphenation_chosen)
        l.addWidget(self.hyphenation_combo)
        groups.append(g)
        g, l = self._group("Size & Arrange")
        size_menu = QMenu(self)
        for paper in ("A4", "Letter", "Legal"):
            size_menu.addAction(paper, lambda _checked=False, value=paper: self._paper_size(value))
        size_menu.addAction("More Paper Sizes…", self._page_setup)
        l.addWidget(self._menu_tool("Size", size_menu))
        l = self._extra_row(g)
        l.addWidget(self._tool("image", "Selection Pane", self._show_object_pane,
                               text_mode=True))
        section_menu = QMenu(self)
        for label, kind in (("Next Page", "next"), ("Continuous", "continuous"),
                            ("Even Page", "even"), ("Odd Page", "odd")):
            section_menu.addAction(label, lambda _checked=False, value=kind: self._section_break(value))
        l.addWidget(self._menu_tool("Section Break", section_menu))
        groups.append(g)
        self.ribbon.addTab(self._wrap(groups), "Layout")

    def _build_references_tab(self) -> None:
        return ReferenceCommands._build_references_tab(self)

    def _build_mailings_tab(self) -> None:
        return MailingCommands._build_mailings_tab(self)

    def _build_review_tab(self) -> None:
        groups = []
        g, l = self._group("Proofing")
        self.proof_button = self._tool(
            "spell", "Proofing", self._toggle_proofing,
            "Offline spelling + basic grammar/style heuristics",
            checkable=True)
        with QSignalBlocker(self.proof_button):
            self.proof_button.setChecked(self._proof_enabled)
        l.addWidget(self.proof_button)
        l.addWidget(self._tool("spell", "Check now", lambda: self._run_proofing(),
                               "Run checks now", text_mode=True))
        groups.append(g)
        g, l = self._group("Changes")
        self.track_button = self._tool(
            "track", "Track", self._toggle_tracking,
            "Track text edits; enable Track formatting to include formatting edits",
            checkable=True)
        with QSignalBlocker(self.track_button):
            self.track_button.setChecked(self.state.track_changes)
        l.addWidget(self.track_button)
        l.addWidget(self._tool("review", "Accept", self._accept_selected,
                               "Accept selected change", text_mode=True))
        l.addWidget(self._tool("clear", "Reject", self._reject_selected,
                               "Reject selected change", text_mode=True))
        groups.append(g)
        g, l = self._group("Comments")
        l.addWidget(self._tool("comment", "Comment", self._new_comment,
                               "New comment on selection", text_mode=True))
        l.addWidget(self._tool("author", "Author", self._author_dialog,
                               "Set author name", text_mode=True))
        groups.append(g)
        g, l = self._group("Reference & Statistics")
        l.addWidget(self._tool("numbering", "Word Count", self._word_count, text_mode=True))
        l = self._extra_row(g)
        l.addWidget(self._tool("find", "Thesaurus", self._thesaurus, text_mode=True))
        groups.append(g)
        g, l = self._group("Tracking & Compare")
        tracking = QMenu(self)
        tracking.addAction("Track Changes On", lambda: self._tracking_mode(True))
        tracking.addAction("Track Changes Off", lambda: self._tracking_mode(False))
        tracking.addAction("Reviewing Pane", self._reviewing_pane)
        l.addWidget(self._menu_tool("Tracking", tracking))
        l.addWidget(self._tool("review", "Compare", self._compare_documents, text_mode=True))
        l = self._extra_row(g)
        markup = QMenu(self)
        for label, kind in (("Comments", "comments"), ("Insertions and Deletions", "changes")):
            action = markup.addAction(label)
            action.setCheckable(True)
            action.setChecked(True)
            action.toggled.connect(lambda checked, value=kind: self._markup(value, checked))
        l.addWidget(self._menu_tool("Show Markup", markup))
        l.addWidget(self._tool("review", "Reviewing Pane", self._reviewing_pane, text_mode=True))
        groups.append(g)
        groups.extend(self._review_extra_groups())
        self.ribbon.addTab(self._wrap(groups), "Review")

    def _build_view_tab(self) -> None:
        return ViewCommands._build_view_tab(self)

    def _build_help_tab(self) -> None:
        return ViewCommands._build_help_tab(self)

    def _build_docks(self) -> None:
        self.nav_dock = QDockWidget("Navigation", self)
        self.nav_dock.setObjectName("nav_dock")
        self.nav_dock.setMinimumWidth(240)
        self.nav_dock.setMaximumWidth(300)
        self.outline = QListWidget()
        self.outline.setWordWrap(True)
        self.outline.itemClicked.connect(self._outline_jump)
        self.pages_list = QListWidget()
        self.pages_list.itemClicked.connect(self._page_jump)
        self.nav_results = QListWidget()
        self.objects_list = QListWidget()
        self.objects_list.itemClicked.connect(self._object_jump)
        self.navigation_tabs = QTabWidget()
        self.navigation_tabs.tabBar().setExpanding(True)
        self.navigation_tabs.tabBar().setUsesScrollButtons(False)
        self.navigation_tabs.addTab(self.outline, "Headings")
        self.navigation_tabs.addTab(self.pages_list, "Pages")
        self.navigation_tabs.addTab(self.nav_results, "Results")
        self.navigation_tabs.addTab(self.objects_list, "Objects")
        self.nav_dock.setWidget(self.navigation_tabs)
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.nav_dock)

        self.review_dock = QDockWidget("Review", self)
        self.review_dock.setObjectName("review_dock")
        self.review_dock.setMinimumWidth(280)
        self.review_dock.setMaximumWidth(330)
        review_tabs = QTabWidget()
        comments_page = QWidget()
        cl = QVBoxLayout(comments_page)
        self.comments_list = QListWidget()
        self.comments_list.setWordWrap(True)
        self.comments_list.itemClicked.connect(self._comment_jump)
        cl.addWidget(self.comments_list)
        cbtn = QHBoxLayout()
        resolve = QPushButton("Resolve")
        resolve.clicked.connect(self._resolve_selected_comment)
        newc = QPushButton("New")
        newc.clicked.connect(self._new_comment)
        cbtn.addWidget(resolve)
        cbtn.addWidget(newc)
        cl.addLayout(cbtn)
        review_tabs.addTab(comments_page, "Comments")
        changes_page = QWidget()
        gl = QVBoxLayout(changes_page)
        self.changes_list = QListWidget()
        self.changes_list.setWordWrap(True)
        self.changes_list.itemClicked.connect(self._change_jump)
        gl.addWidget(self.changes_list)
        gbtn = QGridLayout()
        for index, (label, slot) in enumerate(
                (("Accept", self._accept_selected),
                 ("Reject", self._reject_selected),
                 ("Accept all", self._accept_all),
                 ("Reject all", self._reject_all))):
            b = QPushButton(label)
            b.clicked.connect(slot)
            gbtn.addWidget(b, index // 2, index % 2)
        gl.addLayout(gbtn)
        review_tabs.addTab(changes_page, "Changes")
        proof_page = QWidget()
        pl = QVBoxLayout(proof_page)
        self.proof_list = QListWidget()
        self.proof_list.setWordWrap(True)
        self.proof_list.itemClicked.connect(self._proof_jump)
        pl.addWidget(self.proof_list)
        pbtn = QHBoxLayout()
        fix = QPushButton("Apply fix")
        fix.clicked.connect(self._proof_apply)
        pbtn.addWidget(fix)
        pl.addLayout(pbtn)
        review_tabs.addTab(proof_page, "Proofing")
        self.review_dock.setWidget(review_tabs)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea,
                           self.review_dock)
        self._review_tabs = review_tabs
        self.resizeDocks([self.nav_dock], [280], Qt.Orientation.Horizontal)
        self.resizeDocks([self.review_dock], [300], Qt.Orientation.Horizontal)
        self.review_dock.setVisible(False)

    def _build_statusbar(self) -> None:
        bar = QStatusBar()
        self.counts_label = QLabel("0 words")
        bar.addWidget(self.counts_label)
        self.page_label = QLabel("Page 1")
        bar.addWidget(self.page_label)
        self.proof_label = QLabel("Proofing on" if self._proof_enabled
                                  else "Proofing off")
        bar.addWidget(self.proof_label)
        bar.addPermanentWidget(QLabel("Zoom"))
        self.zoom_slider = QSlider(Qt.Orientation.Horizontal)
        self.zoom_slider.setRange(10, 500)
        self.zoom_slider.setFixedWidth(140)
        self.zoom_label = QLabel("100%")
        self.zoom_slider.valueChanged.connect(
            lambda v: self.zoom_label.setText(f"{v}%"))
        self.zoom_slider.valueChanged.connect(
            lambda v: getattr(self, "canvas", None)
            and self.canvas.set_zoom(v / 100.0))
        bar.addPermanentWidget(self.zoom_slider)
        bar.addPermanentWidget(self.zoom_label)
        self.zoom_slider.setValue(100)
        self.save_state_label = QLabel("")
        bar.addPermanentWidget(self.save_state_label)
        self.setStatusBar(bar)


    def _set_dirty(self, dirty: bool) -> None:
        self._dirty = dirty
        self._refresh_title()

    def _refresh_title(self) -> None:
        marker = " •" if self._dirty else ""
        self.title_edit.setText(f"{self._title}{marker}")
        self.setWindowTitle(f"Folio — {self._title}{marker}")

    def _document_changed(self) -> None:
        self._set_dirty(True)
        self._proof_generation += 1
        self._proof_issues = []
        self.highlighter.set_issues([])
        self._autosave_timer.start()
        self._outline_timer.start()
        self._counts_timer.start()
        if self._proof_enabled:
            self._proof_timer.start()
        self._update_table_tools()

    def _review_changed(self) -> None:
        self._set_dirty(True)
        self._autosave_timer.start()
        self._refresh_review_panes()
        self._apply_review_highlights()
        self._counts_timer.start()

    def _update_table_tools(self) -> None:
        in_table = self.editor.current_table() is not None
        for b in self._table_tools:
            b.setEnabled(in_table)


    def load_state(self, state: DocumentState, path: Path | None = None,
                   clean: bool = True) -> None:
        self._close_document_views()
        if hasattr(self, "thesaurus_dock"):
            self.thesaurus_dock.hide()
            self._synonym_cursor = QTextCursor()
        self.state = state
        self.editor.document().folio_state = state
        self.editor.folio_state = state
        self.editor.document().design_settings = DesignSettings.from_dict(state.design)
        self.editor.design_settings = self.editor.document().design_settings
        with self.tracker.loading():
            self.editor.document().setHtml(state.html or "")
            self.editor.apply_page_settings(state.page)
        sanitize_ranges(state,
                        self.editor.document().characterCount() - 1)
        self.tracker.reset(state)
        self.tracker.author = self._author
        if hasattr(self, "floating_dock"):
            self.floating_dock.widget().set_objects(state.features["objects"])
        self._path = Path(path) if path else None
        try:
            self._saved_mtime_ns = (self._path.stat().st_mtime_ns
                                    if self._path and self._path.exists()
                                    else None)
        except OSError:
            self._saved_mtime_ns = None
        self._title = state.title
        with QSignalBlocker(self.track_button):
            self.track_button.setChecked(state.track_changes)
        self._ignored_words.clear()
        self._proof_generation += 1
        self._proof_issues = []
        self.highlighter.set_issues([])
        self._set_dirty(not clean)
        if clean:
            self.editor.document().setModified(False)
        self._sync_all_panes()
        self._sync_page_controls()
        self.review_dock.setVisible(
            bool(state.comments or state.revisions))
        self.canvas.set_page_metrics(
            mm_to_px(state.page.size_mm()[0]),
            mm_to_px(state.page.size_mm()[1]),
        )
        if self._proof_enabled:
            self._proof_timer.start()

        self._references_loaded()
        self._review_state_loaded()

    def _sync_page_controls(self) -> None:
        page = self._active_page_settings()
        with QSignalBlocker(self.columns_combo):
            self.columns_combo.setCurrentIndex(
                max(0, min(2, page.columns - 1)))
        margins = (page.top_mm, page.right_mm,
                   page.bottom_mm, page.left_mm)
        presets = ((25.4,) * 4, (12.7,) * 4, (25.4, 38.1, 25.4, 38.1))
        index = 3
        for i, preset in enumerate(presets):
            if all(abs(m - v) < 0.05 for m, v in zip(margins, preset)):
                index = i
                break
        with QSignalBlocker(self.margins_combo):
            self.margins_combo.setCurrentIndex(index)

    def current_state(self) -> DocumentState:
        self._refresh_references()
        self._refresh_fields()
        self.state.html = self.editor.document().toHtml()
        self.state.title = self._title
        return self.state

    def _check_external_change(self, path: Path) -> str:
        if (self._path is None or path != self._path
                or self._saved_mtime_ns is None or not path.exists()):
            return "overwrite"
        try:
            current = path.stat().st_mtime_ns
        except OSError:
            return "overwrite"
        if current == self._saved_mtime_ns:
            return "overwrite"
        box = QMessageBox(self)
        box.setWindowTitle("File changed on disk")
        box.setText(f"{path.name} was modified by another program since it "
                    "was opened.")
        overwrite = box.addButton("Overwrite", QMessageBox.ButtonRole.AcceptRole)
        copy_btn = box.addButton("Save a Copy",
                                 QMessageBox.ButtonRole.ActionRole)
        box.addButton(QMessageBox.StandardButton.Cancel)
        box.exec()
        clicked = box.clickedButton()
        if clicked is overwrite:
            return "overwrite"
        if clicked is copy_btn:
            return "copy"
        return "cancel"

    def save_to(self, path: Path) -> bool:
        path = Path(path)
        choice = self._check_external_change(path)
        if choice == "copy":
            return self._save_as()
        if choice != "overwrite":
            return False
        self._refresh_toc_quietly()
        state = self.current_state()
        try:
            write_docx(path, state, self.editor.document())
        except Exception as exc:
            QMessageBox.warning(self, "Save", f"Could not save the file:\n{exc}")
            return False
        self._path = path
        try:
            self._saved_mtime_ns = path.stat().st_mtime_ns
        except OSError:
            self._saved_mtime_ns = None
        self._set_dirty(False)
        self.editor.document().setModified(False)
        try:
            self.store.save_recovery(state, str(path), dirty=False)
            self.store.mark_clean(state.id)
        except Exception:
            self.statusBar().showMessage(
                "Saved, but local recovery could not be updated.", 4000)
        try:
            self.store.add_version(state, "Saved")
            self._last_version_time = time.monotonic()
        except Exception:
            self.statusBar().showMessage(
                "Saved, but version history is unavailable.", 4000)
        self.statusBar().showMessage(f"Saved to {path}", 3000)
        self.save_state_label.setText("Saved")
        return True

    def _save(self) -> bool:
        if self._path is not None:
            return self.save_to(self._path)
        return self._save_as()

    def _save_as(self) -> bool:
        path, _ = QFileDialog.getSaveFileName(
            self, "Save As", str(self._path or f"{self._title}.docx"),
            "Word document (*.docx)")
        if not path:
            return False
        if not path.lower().endswith(".docx"):
            path += ".docx"
        return self.save_to(Path(path))

    def _maybe_save(self) -> str:
        if not self._dirty:
            return "proceed"
        ret = QMessageBox.question(
            self, "Unsaved changes",
            f"Save changes to {self._title}?",
            QMessageBox.StandardButton.Save
            | QMessageBox.StandardButton.Discard
            | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save,
        )
        if ret == QMessageBox.StandardButton.Save:
            return "save"
        if ret == QMessageBox.StandardButton.Discard:
            return "discard"
        return "cancel"

    def _discard_recovery(self, doc_id) -> None:
        try:
            self.store.mark_clean(doc_id)
        except Exception:
            pass

    def _proceed_or_cancel(self) -> bool:
        choice = self._maybe_save()
        if choice == "cancel":
            return False
        if choice == "save" and not self._save():
            return False
        return True


    def _new_document(self) -> None:
        dialog = dialogs.TemplateGalleryDialog(self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        choice = dialog.choice()
        if not choice:
            return
        if not self._proceed_or_cancel():
            return
        self._discard_recovery(self.state.id)
        self.load_state(make_template(choice), None, clean=True)

    def _open_document(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Open", "",
            "Documents (*.docx *.html *.htm *.txt);;"
            "Word documents (*.docx);;HTML (*.html *.htm);;Text (*.txt)")
        if not path:
            return
        if not self._proceed_or_cancel():
            return
        prior_id = self.state.id
        if self.open_path(Path(path)):
            self._discard_recovery(prior_id)

    def open_path(self, path: Path) -> bool:
        suffix = path.suffix.lower()
        try:
            if suffix == ".docx":
                state, warnings = read_docx(path)
                self.load_state(state, path, clean=True)
                if warnings:
                    QMessageBox.information(self, "Open", "\n".join(warnings))
            elif suffix in (".html", ".htm"):
                try:
                    text = path.read_text(encoding="utf-8")
                except UnicodeDecodeError as exc:
                    raise DocxError(f"Could not decode the HTML file: {exc}")
                state = DocumentState(title=path.stem, html=text)
                self.load_state(state, None, clean=False)
            elif suffix == ".txt":
                try:
                    text = path.read_text(encoding="utf-8")
                except UnicodeDecodeError as exc:
                    raise DocxError(f"Could not decode the text file: {exc}")
                body = "".join(
                    f"<p>{html_module.escape(line)}</p>"
                    for line in text.splitlines()
                )
                state = DocumentState(
                    title=path.stem,
                    html=("<!DOCTYPE html><html><body>"
                          + body + "</body></html>"))
                self.load_state(state, None, clean=False)
            else:
                QMessageBox.warning(self, "Open",
                                    "Folio can open .docx, .html, and .txt files.")
                return False
        except (DocxError, OSError) as exc:
            QMessageBox.warning(self, "Open", str(exc))
            return False
        return True

    def _close_document(self) -> None:
        if not self._proceed_or_cancel():
            return
        self._discard_recovery(self.state.id)
        self.load_state(make_template("blank"), None, clean=True)

    def _export_pdf(self) -> None:
        self._refresh_references()
        path, _ = QFileDialog.getSaveFileName(
            self, "Export PDF", f"{self._title}.pdf",
            "PDF document (*.pdf)")
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            pages = write_pdf(Path(path), self.editor.document(),
                              self.state.page, self._title)
        except Exception as exc:
            QMessageBox.warning(self, "Export PDF",
                                f"Could not write the PDF:\n{exc}")
            return
        self.statusBar().showMessage(
            f"Exported {pages} page(s) to {path}", 4000)

    def _export_html(self) -> None:
        self._refresh_references()
        path, _ = QFileDialog.getSaveFileName(
            self, "Export HTML", f"{self._title}.html",
            "HTML document (*.html)")
        if not path:
            return
        try:
            atomic_write(Path(path),
                         self.editor.document().toHtml().encode("utf-8"))
        except Exception as exc:
            QMessageBox.warning(self, "Export HTML", str(exc))

    def _make_printer(self) -> QPrinter:
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setPageLayout(page_layout_for(self.state.page))
        printer.setFullPage(True)
        return printer

    def _print_preview(self) -> None:
        self._refresh_references()
        publication = Publication(self.editor.document(), self.state.page,
                                  self._title)
        printer = self._make_printer()
        preview = QPrintPreviewDialog(printer, self)
        preview.paintRequested.connect(
            lambda p: publication.paint(p))
        preview.exec()

    def _print(self) -> None:
        self._refresh_references()
        publication = Publication(self.editor.document(), self.state.page,
                                  self._title)
        printer = self._make_printer()
        dialog = QPrintDialog(printer, self)
        if dialog.exec() != QPrintDialog.DialogCode.Accepted:
            return
        try:
            publication.paint(printer)
        except Exception as exc:
            QMessageBox.warning(self, "Print", f"Printing failed:\n{exc}")

    def _checkpoint(self) -> None:
        try:
            self.store.add_version(self.current_state(), "Manual checkpoint")
            self._last_version_time = time.monotonic()
        except Exception as exc:
            QMessageBox.warning(self, "Checkpoint",
                                f"Could not save a checkpoint:\n{exc}")
            return
        self.statusBar().showMessage("Checkpoint saved to history.", 3000)

    def _show_history(self) -> None:
        try:
            versions = self.store.versions(self.state.id)
        except Exception as exc:
            QMessageBox.warning(self, "Version History",
                                f"History is unavailable:\n{exc}")
            return
        dialog = dialogs.VersionHistoryDialog(versions, self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        version_id = dialog.selected()
        if version_id is None:
            return
        ret = QMessageBox.question(
            self, "Restore version",
            "Restore this snapshot? A checkpoint of the current document "
            "will be kept in history.")
        if ret != QMessageBox.StandardButton.Yes:
            return
        try:
            self.store.add_version(self.current_state(), "Before restore")
            restored = self.store.load_version(version_id)
        except Exception as exc:
            QMessageBox.warning(self, "Version History", str(exc))
            return
        restored.id = self.state.id
        self.load_state(restored, self._path, clean=False)
        self.statusBar().showMessage(
            "Snapshot restored (not written to disk until you save).", 4000)

    def _about(self) -> None:
        dialogs.AboutDialog(self).exec()


    def _autosave(self) -> None:
        try:
            self.store.save_recovery(self.current_state(),
                                     str(self._path) if self._path else None,
                                     dirty=self._dirty)
            if self._dirty:
                self.save_state_label.setText("Recovery saved")
                if time.monotonic() - self._last_version_time >= VERSION_INTERVAL_S:
                    self.store.add_version(self.state, "Automatic checkpoint")
                    self._last_version_time = time.monotonic()
        except Exception:
            self.save_state_label.setText("Recovery unavailable")

    def _offer_recovery(self) -> None:
        try:
            entries = self.store.recoverable()
        except Exception:
            entries = []
        if not entries:
            return
        dialog = dialogs.RecoveryDialog(entries, self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        doc_id, discard = dialog.choice()
        if not doc_id:
            return
        if discard:
            try:
                self.store.discard_recovery(doc_id)
            except Exception:
                pass
            return
        try:
            loaded = self.store.load_recovery(doc_id)
        except Exception as exc:
            QMessageBox.warning(self, "Recovery", str(exc))
            return
        if loaded is None:
            return
        state, path = loaded
        self.load_state(state, Path(path) if path else None, clean=False)
        self.statusBar().showMessage("Recovered unsaved work.", 4000)

    def closeEvent(self, event) -> None:
        choice = self._maybe_save()
        if choice == "cancel":
            event.ignore()
            return
        if choice == "save" and not self._save():
            event.ignore()
            return
        self._discard_recovery(self.state.id)
        self._close_document_views()
        self._references_timer.stop()
        self._fields_timer.stop()
        self.speech_reader.stop()
        for timer in (self._autosave_timer, self._safety_timer,
                      self._outline_timer, self._proof_timer,
                      self._counts_timer):
            timer.stop()
        self._proof_generation += 1
        self._pool.waitForDone(3000)
        try:
            self.store.close()
        except Exception:
            pass
        event.accept()


    def _sync_toolbar(self, *_args) -> None:
        cursor = self.editor.textCursor()
        fmt = cursor.charFormat()
        with QSignalBlocker(self.font_combo):
            font = QFont(fmt.font().family() or "Calibri")
            self.font_combo.setCurrentFont(font)
        with QSignalBlocker(self.size_combo):
            size = fmt.fontPointSize()
            self.size_combo.setCurrentText(
                str(int(size)) if size > 0 else "11")
        with QSignalBlocker(self.bold_button):
            self.bold_button.setChecked(fmt.fontWeight() >= QFont.Weight.Bold)
        with QSignalBlocker(self.italic_button):
            self.italic_button.setChecked(fmt.fontItalic())
        with QSignalBlocker(self.underline_button):
            self.underline_button.setChecked(fmt.fontUnderline())
        with QSignalBlocker(self.strike_button):
            self.strike_button.setChecked(fmt.fontStrikeOut())
        with QSignalBlocker(self.style_combo):
            self.style_combo.setCurrentText(self.editor.current_style_name())
        alignment = self.editor.alignment() & Qt.AlignmentFlag.AlignHorizontal_Mask
        for flag, button in self.align_buttons.items():
            with QSignalBlocker(button):
                button.setChecked(alignment == flag or
                                  (flag == Qt.AlignmentFlag.AlignLeft and
                                   alignment == 0))
        self._update_table_tools()

    def _set_alignment(self, flag) -> None:
        self.editor.set_alignment(flag)

    def _font_size_chosen(self, _index=0) -> None:
        try:
            self.editor.set_font_size(float(self.size_combo.currentText()))
        except ValueError:
            self._sync_toolbar()

    def _font_size_edited(self) -> None:
        self._font_size_chosen()

    def _pick_text_color(self) -> None:
        from PySide6.QtWidgets import QColorDialog
        color = QColorDialog.getColor(parent=self, title="Text color")
        if color.isValid():
            self.editor.set_foreground(color)

    def _menu_tool(self, label: str, menu: QMenu) -> QToolButton:
        button = self._tool("pagesetup", label, text_mode=True)
        button.setMenu(menu)
        button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        return button

    def _online_video(self) -> None:
        dialog = authoring.VideoDialog(self)
        if dialog.exec() == dialogs.QDialog.DialogCode.Accepted:
            url = dialog.url.text().strip()
            title = dialog.title.text().strip() or authoring.video_provider(url) + " video"
            self.editor.insert_link(url, title)

    def _stock_media(self) -> None:
        dialog = authoring.StockMediaDialog(self)
        if (dialog.exec() == dialogs.QDialog.DialogCode.Accepted
                and dialog.selected_image is not None):
            self.editor.insert_image(dialog.selected_image)

    def _text_insert(self, kind: str) -> None:
        dialog = authoring.TextInsertDialog(kind, self)
        if dialog.exec() == dialogs.QDialog.DialogCode.Accepted:
            dialog.insert(self.editor)

    def _signature_fields(self) -> None:
        if not hasattr(self, "signature_dock"):
            self.signature_dock = QDockWidget("Signature Workflow", self)
            self.signature_dock.setObjectName("signature_dock")
            pane = QWidget()
            layout = QVBoxLayout(pane)
            note = QLabel("Add an interactive signature field, place the cursor in it, then use "
                          "Insert → Edit / Sign / Extract to sign with a typed name and explicit consent. "
                          "The local record stores the name, time, and a body digest. "
                          "It does not authenticate identity or use a signing certificate.")
            note.setWordWrap(True)
            layout.addWidget(note)
            add = QPushButton("Add Signature Fields…")
            add.clicked.connect(self._insert_signature_fields)
            layout.addWidget(add)
            requirements = QPushButton("Electronic Signature Requirements")
            requirements.clicked.connect(lambda: QMessageBox.information(
                self, "Electronic Signature Requirements",
                "Folio supports a local typed signature and explicit consent record. "
                "This is not a certificate-based digital signature, identity verification, "
                "or a tamper-proof audit service. Subsequent document edits are allowed. "
                "No data is submitted to any signing provider."))
            layout.addWidget(requirements)
            layout.addStretch()
            self.signature_dock.setWidget(pane)
            self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.signature_dock)
        self.signature_dock.show()

    def _insert_signature_fields(self) -> None:
        self._signature_node()

    def _page_numbers(self, position: str) -> None:
        if not self._can_edit():
            return
        page = self._active_page_settings()
        if position == "remove":
            for name in ("header", "footer"):
                value = getattr(page, name).replace("Page {page} of {pages}", "")
                value = value.replace("Page {CurrentPage} of {TotalPages}", "")
                for token in ("{page}", "{pages}", "{CurrentPage}", "{TotalPages}"):
                    value = value.replace(token, "")
                setattr(page, name, value.strip())
        else:
            name = "header" if position == "top" else "footer"
            value = getattr(page, name)
            if "{page}" not in value and "{CurrentPage}" not in value:
                setattr(page, name, (value + "    Page {page} of {pages}").strip())
        self._set_active_page(page)
        self.statusBar().showMessage("Page numbers updated in preview and exported output.", 5000)

    def _paper_size(self, paper: str) -> None:
        candidate = self._active_page_settings()
        candidate.paper = paper
        try:
            candidate.validate()
        except ValidationError as exc:
            QMessageBox.warning(self, "Paper Size", str(exc))
            return
        self._set_active_page(candidate)

    def _design_dialog(self) -> None:
        dialog = dialogs.QDialog(self)
        dialog.setWindowTitle("Document Theme")
        layout = QVBoxLayout(dialog)
        form = dialogs.QFormLayout()
        colors = QComboBox()
        colors.addItems(list(authoring.PALETTES))
        fonts = QComboBox()
        fonts.addItem("Calibri / Calibri", ("Calibri", "Calibri"))
        fonts.addItem("Cambria / Calibri", ("Cambria", "Calibri"))
        fonts.addItem("Georgia / Georgia", ("Georgia", "Georgia"))
        fonts.addItem("Segoe UI / Segoe UI", ("Segoe UI", "Segoe UI"))
        styles = QComboBox()
        styles.addItems(list(authoring.STYLE_SETS))
        form.addRow("Colors:", colors)
        form.addRow("Heading / body fonts:", fonts)
        form.addRow("Style set:", styles)
        layout.addLayout(form)
        note = QLabel("Applies to existing document text. Headings keep their levels; "
                      "images and links are preserved. Undo reverses the change.")
        note.setWordWrap(True)
        layout.addWidget(note)
        authoring.buttons(dialog, layout)
        if dialog.exec() == dialogs.QDialog.DialogCode.Accepted:
            authoring.apply_design(self.editor, colors.currentText(),
                                   fonts.currentData(), styles.currentText())

    def _word_count(self) -> None:
        stats = authoring.document_statistics(self.editor.document(), self.state.page, self._title)
        dialog = dialogs.QDialog(self)
        dialog.setWindowTitle("Word Count")
        layout = QVBoxLayout(dialog)
        form = dialogs.QFormLayout()
        for name, value in stats.items():
            form.addRow(name + ":", QLabel(f"{value:,}"))
        layout.addLayout(form)
        layout.addWidget(QLabel("Document body; pages and lines use publishing layout."))
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            import re
            count = len(re.findall(r"\b\w+\b", cursor.selectedText()))
            layout.addWidget(QLabel(f"Selection: {count:,} words"))
        close = QPushButton("Close")
        close.clicked.connect(dialog.accept)
        layout.addWidget(close)
        dialog.exec()

    def _thesaurus(self) -> None:
        if not hasattr(self, "thesaurus_dock"):
            self.thesaurus_dock = QDockWidget("Thesaurus", self)
            self.thesaurus_dock.setObjectName("thesaurus_dock")
            page = QWidget()
            layout = QVBoxLayout(page)
            self.thesaurus_query = QLineEdit()
            self.thesaurus_query.setPlaceholderText("English word")
            layout.addWidget(self.thesaurus_query)
            lookup = QPushButton("Look Up")
            lookup.clicked.connect(self._lookup_synonyms)
            self.thesaurus_query.returnPressed.connect(self._lookup_synonyms)
            layout.addWidget(lookup)
            self.synonym_list = QListWidget()
            layout.addWidget(self.synonym_list)
            replace = QPushButton("Replace Selected Word")
            replace.clicked.connect(self._replace_synonym)
            layout.addWidget(replace)
            note = QLabel("Small built-in English glossary. Suggestions depend on context.")
            note.setWordWrap(True)
            layout.addWidget(note)
            self.thesaurus_status = QLabel()
            self.thesaurus_status.setWordWrap(True)
            layout.addWidget(self.thesaurus_status)
            self.thesaurus_dock.setWidget(page)
            self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.thesaurus_dock)
        cursor = self.editor.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        self._synonym_cursor = QTextCursor(cursor)
        self.thesaurus_query.setText(cursor.selectedText())
        self._lookup_synonyms()
        self.thesaurus_dock.show()

    def _lookup_synonyms(self) -> None:
        self.synonym_list.clear()
        values = authoring.synonyms(self.thesaurus_query.text())
        self.synonym_list.addItems(values)
        self.thesaurus_status.setText("" if values else "No entry in the local glossary.")

    def _replace_synonym(self) -> None:
        if not self._can_edit():
            return
        item = self.synonym_list.currentItem()
        cursor = self._synonym_cursor
        if (item is None or not cursor.hasSelection()
                or cursor.selectedText().lower() != self.thesaurus_query.text().strip().lower()):
            self.thesaurus_status.setText("Select a word in the document and reopen Thesaurus to replace it.")
            return
        replacement = item.text()
        if cursor.selectedText().istitle():
            replacement = replacement.title()
        elif cursor.selectedText().isupper():
            replacement = replacement.upper()
        cursor.insertText(replacement)
        self.editor.setTextCursor(cursor)
        self.thesaurus_status.setText("Word replaced.")

    def _tracking_mode(self, enabled: bool) -> None:
        self.track_button.setChecked(enabled)

    def _reviewing_pane(self) -> None:
        self._show_review_pane()
        self._review_tabs.setCurrentIndex(1)

    def _markup(self, kind: str, visible: bool) -> None:
        if kind == "comments":
            self._show_comments = visible
        else:
            self._show_changes = visible
        self._refresh_review_panes()
        self._apply_review_highlights()

    def _compare_documents(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Compare Current Document with Original", "", "Word documents (*.docx)")
        if not path:
            return
        try:
            original, warnings = read_docx(Path(path))
        except (DocxError, OSError) as exc:
            QMessageBox.warning(self, "Compare", str(exc))
            return
        source = authoring.SafeDocument()
        source.setHtml(original.html)
        dialog = dialogs.QDialog(self)
        dialog.setWindowTitle(f"Compare — {Path(path).name} → {self._title}")
        dialog.resize(950, 700)
        layout = QVBoxLayout(dialog)
        if warnings:
            note = QLabel("\n".join(warnings))
            note.setWordWrap(True)
            layout.addWidget(note)
        view = QTextEdit()
        view.setReadOnly(True)
        view.setHtml(authoring.comparison_html(source.toPlainText(), self.editor.toPlainText()))
        layout.addWidget(view)
        close = QPushButton("Close")
        close.clicked.connect(dialog.accept)
        layout.addWidget(close)
        dialog.exec()

    def _immersive_reader(self, fullscreen=False, web=False) -> None:
        dialog = authoring.ReadingDialog(self.editor.document(), self, web=web)
        if fullscreen:
            dialog.showFullScreen()
        dialog.exec()

    def _gridlines(self, enabled: bool) -> None:
        self.canvas.gridlines = enabled
        self.canvas.viewport().update()

    def _zoom_dialog(self) -> None:
        dialog = authoring.ZoomDialog(round(self.canvas.zoom() * 100), self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        mode = dialog.mode.currentText()
        if mode == "Percent":
            self._set_zoom(dialog.percent.value() / 100)
        elif mode == "Page Width":
            self._zoom_fit()
        else:
            publication = Publication(self.editor.document(), self.state.page, self._title)
            columns = {"One Page": 1, "Two Pages": 2, "Multiple Pages": 3}[mode]
            preview = authoring.PagePreviewDialog(publication, columns, self)
            preview.exec()

    def _populate_windows(self, menu: QMenu) -> None:
        menu.clear()
        for window in QApplication.topLevelWidgets():
            if isinstance(window, FolioWindow) and window.isVisible():
                action = menu.addAction(window.windowTitle())
                action.setCheckable(True)
                action.setChecked(window is self)
                action.triggered.connect(lambda _checked=False, target=window: self._activate_window(target))
        menu.addSeparator()
        menu.addAction("New Document Window", self._new_window)

    @staticmethod
    def _activate_window(window) -> None:
        if window.isMinimized():
            window.showNormal()
        window.raise_()
        window.activateWindow()

    def _new_window(self) -> None:
        window = FolioWindow(store=LocalStore(self.store.root), restore=False)
        window.load_state(make_template("blank"), clean=True)
        window.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self._secondary_windows.append(window)
        window.destroyed.connect(lambda: self._secondary_windows.remove(window)
                                 if window in self._secondary_windows else None)
        window.show()

    def _object_jump(self, item: QListWidgetItem) -> None:
        self._outline_jump(item)
        cursor = self.editor.textCursor()
        if item.text().startswith("Picture"):
            cursor.movePosition(QTextCursor.MoveOperation.NextCharacter,
                                QTextCursor.MoveMode.KeepAnchor)
        else:
            table = cursor.currentTable()
            if table is not None:
                cursor.setPosition(table.firstPosition())
                cursor.setPosition(table.lastPosition(), QTextCursor.MoveMode.KeepAnchor)
        self.editor.setTextCursor(cursor)

    def _help_topics(self) -> None:
        self._search_help()

    def _pick_highlight(self) -> None:
        from PySide6.QtWidgets import QColorDialog
        color = QColorDialog.getColor(QColor("#ffff00"), self,
                                      "Highlight color")
        if color.isValid():
            self.editor.set_highlight(color)

    def _paste(self) -> None:
        self.editor.paste_special(self._paste_mode)

    def _paste_special(self) -> None:
        mime_data = QApplication.clipboard().mimeData()
        dialog = dialogs.PasteSpecialDialog(mime_data, self)
        if dialog.exec() == dialogs.QDialog.DialogCode.Accepted:
            self.editor.paste_special(dialog.mode())

    def _set_default_paste(self) -> None:
        labels = {
            "Keep source formatting": "keep_formatting",
            "Plain text": "text",
            "Picture when available": "image",
        }
        current_label = next(
            (label for label, value in labels.items()
             if value == self._paste_mode), "Keep source formatting")
        value, ok = QInputDialog.getItem(
            self, "Default Paste", "Paste clipboard content as:",
            list(labels), list(labels).index(current_label), False)
        if not ok:
            return
        self._paste_mode = labels[value]
        self.store.set_setting("paste_mode", self._paste_mode)

    def _clear_format(self) -> None:
        cursor = self.editor.textCursor()
        if not cursor.hasSelection():
            self.editor.set_highlight(None)
            return
        self.editor.clear_formatting()


    def _insert_image(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Insert image", "",
            "Images (*.png *.jpg *.jpeg *.svg *.bmp *.gif *.webp)")
        if not path:
            return
        try:
            if Path(path).suffix.lower() == ".svg":
                from .structured_media import render_svg
                self.editor.insert_image(render_svg(Path(path).read_bytes()))
            else:
                self.editor.insert_image_file(path)
        except (ValueError, OSError) as exc:
            QMessageBox.warning(self, "Insert image", str(exc))

    def _insert_cover_page(self) -> None:
        dialog = dialogs.CoverPageDialog(self, self._title)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        style, title, subtitle = dialog.values()
        self.editor.insert_cover_page(title, subtitle, style)

    def _insert_chart(self) -> None:
        dialog = dialogs.ChartDialog(self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        values = dialog.values()
        if values is None:
            return
        chart_type, rows = values
        self.editor.insert_image(dialogs.ChartDialog.render(chart_type, rows))

    def _insert_screenshot(self) -> None:
        screen = QApplication.primaryScreen()
        if screen is None:
            QMessageBox.warning(
                self, "Insert Screenshot", "No screen is available to capture.")
            return
        image = screen.grabWindow(0).toImage()
        if image.isNull():
            QMessageBox.warning(
                self, "Insert Screenshot", "The screen capture was empty.")
            return
        self.editor.insert_image(image)

    def _insert_shape(self) -> None:
        dialog = dialogs.ShapeDialog(self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        shape, color, label = dialog.values()
        image = dialogs.ShapeDialog.render(shape, color, label)
        self.editor.insert_image(image)

    def _insert_diagram(self) -> None:
        dialog = dialogs.DiagramDialog(self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        image = dialogs.DiagramDialog.render(dialog.values())
        document_features.insert_node(self.editor, self.state,
                                      {"kind": "diagram", "steps": dialog.values()}, image=image)

    def _insert_table(self) -> None:
        dialog = dialogs.TableDialog(self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        rows, cols = dialog.values()
        self.editor.insert_table(rows, cols)

    def _table_op(self, op: str) -> None:
        if op == "add_row":
            self.editor.table_insert_row(True)
        elif op == "add_col":
            self.editor.table_insert_column(True)
        elif op == "del_row":
            self.editor.table_delete_row()
        elif op == "del_col":
            self.editor.table_delete_column()
        elif op == "merge":
            if not self.editor.table_merge_cells():
                self.statusBar().showMessage(
                    "Select cells in a table to merge them.", 3000)

    def _insert_link(self) -> None:
        cursor = self.editor.textCursor()
        dialog = dialogs.LinkDialog(self, cursor.selectedText())
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        url, text = dialog.values()
        try:
            self.editor.insert_link(url, text)
        except ValueError as exc:
            QMessageBox.warning(self, "Insert link", str(exc))

    def _insert_date(self) -> None:
        self._date_field()

    def _insert_field(self) -> None:
        name, ok = QInputDialog.getText(
            self, "Merge Field", "Field name (inserted as {{Name}}):")
        if not (ok and name.strip()):
            return
        try:
            self.editor.insert_merge_field(name.strip())
        except ValueError as exc:
            QMessageBox.warning(self, "Merge Field", str(exc))

    def _mail_merge(self) -> None:
        return MailingCommands._mail_merge(self)

    def _page_setup(self) -> None:
        dialog = dialogs.PageSetupDialog(self._active_page_settings(), self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        settings = dialog.settings()
        if settings is None:
            return
        self._set_active_page(settings)

    def _quick_orientation(self, landscape: bool) -> None:
        candidate = self._active_page_settings()
        if candidate.landscape == landscape:
            return
        candidate.landscape = landscape
        try:
            candidate.validate()
        except ValidationError as exc:
            QMessageBox.warning(self, "Orientation", str(exc))
            return
        self._set_active_page(candidate)

    def _margins_preset(self, index: int) -> None:
        presets = ((25.4, 25.4, 25.4, 25.4),
                   (12.7, 12.7, 12.7, 12.7),
                   (25.4, 38.1, 25.4, 38.1))
        if index >= len(presets):
            self._page_setup()
            return
        top, right, bottom, left = presets[index]
        candidate = self._active_page_settings()
        candidate.top_mm, candidate.right_mm = top, right
        candidate.bottom_mm, candidate.left_mm = bottom, left
        try:
            candidate.validate()
        except ValidationError as exc:
            QMessageBox.warning(self, "Margins", str(exc))
            return
        self._set_active_page(candidate)

    def _quick_columns(self, columns: int) -> None:
        candidate = self._active_page_settings()
        candidate.columns = columns
        try:
            candidate.validate()
        except ValidationError as exc:
            QMessageBox.warning(self, "Columns", str(exc))
            self.columns_combo.setCurrentIndex(self.state.page.columns - 1)
            return
        self._set_active_page(candidate)
        self.statusBar().showMessage(
            "Columns are applied in the publishing preview and exported "
            "output; the editing canvas stays one column wide.", 6000)

    def _insert_break(self, index: int) -> None:
        try:
            if index == 1:
                self.editor.insert_page_break()
            elif index == 2:
                self.editor.insert_column_break()
                self.statusBar().showMessage(
                    "Column break inserted. It is shown in publishing preview "
                    "and DOCX output.", 5000)
            elif 3 <= index <= 6:
                self._section_break(("next", "continuous", "even", "odd")[index - 3])
        finally:
            with QSignalBlocker(self.breaks_combo):
                self.breaks_combo.setCurrentIndex(0)

    def _line_numbers_chosen(self, index: int) -> None:
        mode = self.line_numbers_combo.itemData(index)
        try:
            if mode == "options":
                self._page_setup()
                return
            if mode is None:
                return
            candidate = self._active_page_settings()
            candidate.line_numbering = str(mode)
            candidate.validate()
            self._set_active_page(candidate)
        finally:
            with QSignalBlocker(self.line_numbers_combo):
                self.line_numbers_combo.setCurrentIndex(0)

    def _hyphenation_chosen(self, index: int) -> None:
        mode = self.hyphenation_combo.itemData(index)
        try:
            if mode == "options":
                self._page_setup()
                return
            if mode == "manual":
                count = self.editor.apply_manual_hyphenation(
                    self.state.page.hyphenate_caps)
                QMessageBox.information(
                    self, "Hyphenation",
                    (f"Hyphenation is complete. {count} discretionary "
                     f"hyphen{' was' if count == 1 else 's were'} inserted."))
                return
            if mode is None:
                return
            candidate = self._active_page_settings()
            candidate.hyphenation = str(mode)
            candidate.validate()
            self._set_active_page(candidate)
            self.statusBar().showMessage(
                "Dictionary hyphenation enabled in preview/PDF and DOCX."
                if mode == "automatic" else "Automatic hyphenation disabled.",
                5000)
        finally:
            with QSignalBlocker(self.hyphenation_combo):
                self.hyphenation_combo.setCurrentIndex(0)

    def _paragraph_dialog(self) -> None:
        dialog = dialogs.ParagraphDialog(self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        values = dialog.values()
        if values:
            self.editor.set_paragraph(*values)


    def _find(self) -> None:
        if not hasattr(self, "_find_dialog") or self._find_dialog is None:
            self._find_dialog = dialogs.FindReplaceDialog(self.editor, self)
        self._find_dialog.replace_button.setEnabled(self._can_edit())
        self._find_dialog.replace_all_button.setEnabled(self._can_edit())
        self._find_dialog.show()
        self._find_dialog.raise_()
        self._find_dialog.find_edit.setFocus()

    def _replace(self) -> None:
        self._find()
        self._find_dialog.replace_edit.setFocus()


    def _toggle_nav(self) -> None:
        self.nav_dock.setVisible(not self.nav_dock.isVisible())

    def _show_object_pane(self) -> None:
        self._floating_inspector()
        self._refresh_outline()
        self.nav_dock.setVisible(True)
        self.navigation_tabs.setCurrentWidget(self.objects_list)

    def _toggle_review_pane(self) -> None:
        if self.review_dock.isVisible():
            self.review_dock.setVisible(False)
        else:
            self._show_review_pane()

    def _show_review_pane(self) -> None:
        if self.review_dock.isVisible():
            return
        self.review_dock.setVisible(True)
        QTimer.singleShot(0, self._ensure_page_fits)

    def _ensure_page_fits(self) -> None:
        try:
            page_w, _x, zoom = self.canvas.page_metrics()
            viewport = self.canvas.viewport().width()
        except RuntimeError:
            return
        needed = page_w * zoom + 120
        if 0 < viewport < needed:
            self._zoom_fit()

    def _toggle_focus(self) -> None:
        self._focus = not self._focus
        if self._focus:
            self._pre_focus = (self.nav_dock.isVisible(),
                               self.review_dock.isVisible())
            self.ribbon.setVisible(False)
            self.nav_dock.setVisible(False)
            self.review_dock.setVisible(False)
            self.statusBar().setVisible(False)
            self._exit_focus_btn.setVisible(True)
        else:
            self.ribbon.setVisible(True)
            self.statusBar().setVisible(True)
            self.nav_dock.setVisible(self._pre_focus[0])
            self.review_dock.setVisible(self._pre_focus[1])
            self._exit_focus_btn.setVisible(False)

    def _exit_focus(self) -> None:
        if self._focus:
            self._toggle_focus()

    def _set_zoom(self, factor: float) -> None:
        self.canvas.set_zoom(factor)
        percent = int(round(self.canvas.zoom() * 100))
        with QSignalBlocker(self.zoom_slider):
            self.zoom_slider.setValue(percent)
        self.zoom_label.setText(f"{percent}%")

    def _zoom_fit(self) -> None:
        self.canvas.zoom_fit_width()
        percent = int(round(self.canvas.zoom() * 100))
        with QSignalBlocker(self.zoom_slider):
            self.zoom_slider.setValue(percent)
        self.zoom_label.setText(f"{percent}%")

    def _refresh_outline(self) -> None:
        self.outline.clear()
        for pos, level, text in self.editor.headings():
            item = QListWidgetItem("    " * (level - 1) + (text or "(empty)"))
            item.setData(Qt.ItemDataRole.UserRole, pos)
            self.outline.addItem(item)
        self.pages_list.clear()
        try:
            publication = Publication(
                self.editor.document(), self.state.page, self._title)
            first_positions: dict[int, int] = {}
            block = self.editor.document().begin()
            while block.isValid():
                page = publication.page_for_position(block.position())
                first_positions.setdefault(page, block.position())
                block = block.next()
            for page in range(1, publication.page_count() + 1):
                item = QListWidgetItem(f"Page {page}")
                item.setData(Qt.ItemDataRole.UserRole,
                             first_positions.get(page, 0))
                self.pages_list.addItem(item)
        except Exception:
            item = QListWidgetItem("Page 1")
            item.setData(Qt.ItemDataRole.UserRole, 0)
            self.pages_list.addItem(item)
        self.objects_list.clear()
        image_number = 0
        table_number = 0
        seen_tables: set[int] = set()
        block = self.editor.document().begin()
        while block.isValid():
            table = QTextCursor(block).currentTable()
            table_key = table.firstPosition() if table is not None else -1
            if table is not None and table_key not in seen_tables:
                seen_tables.add(table_key)
                table_number += 1
                item = QListWidgetItem(f"Table {table_number}")
                item.setData(Qt.ItemDataRole.UserRole, block.position())
                self.objects_list.addItem(item)
            iterator = block.begin()
            while not iterator.atEnd():
                fragment = iterator.fragment()
                if (fragment.isValid()
                        and fragment.charFormat().isImageFormat()):
                    image_number += 1
                    item = QListWidgetItem(f"Picture {image_number}")
                    item.setData(Qt.ItemDataRole.UserRole,
                                 fragment.position())
                    self.objects_list.addItem(item)
                iterator += 1
            block = block.next()

    def _outline_jump(self, item: QListWidgetItem) -> None:
        pos = item.data(Qt.ItemDataRole.UserRole)
        cursor = self.editor.textCursor()
        cursor.setPosition(int(pos))
        self.editor.setTextCursor(cursor)
        self.editor.setFocus()
        self.canvas.scroll_to_cursor()

    def _page_jump(self, item: QListWidgetItem) -> None:
        self._outline_jump(item)

    def _collect_toc_entries(self):
        toc = self.state.toc
        entries = []
        for pos, level, text in self.editor.headings():
            if toc and toc.start <= pos < toc.start + toc.length:
                continue
            entries.append((pos, level, text))
        return entries

    def _toc_line_specs(self, entries_pages) -> list:
        title_block = QTextBlockFormat()
        title_block.setBottomMargin(pt_to_px(12))
        title_char = QTextCharFormat()
        title_char.setFontPointSize(16)
        title_char.setFontWeight(QFont.Weight.Bold)
        title_char.setForeground(QColor("#17365d"))
        specs = [("Contents", title_block, title_char)]
        for level, text, page in entries_pages:
            block_fmt = QTextBlockFormat()
            block_fmt.setLeftMargin((level - 1) * 24.0)
            block_fmt.setBottomMargin(pt_to_px(4))
            char_fmt = QTextCharFormat()
            char_fmt.setFontPointSize(11)
            char_fmt.setForeground(QColor("#263244"))
            specs.append((f"{text or '(untitled)'}  ·  p. {page}",
                          block_fmt, char_fmt))
        return specs

    def _apply_toc_lines(self, specs, region, join: bool) -> None:
        doc = self.editor.document()
        self.state.toc = None
        edit = QTextCursor(doc)
        with self.tracker.applying():
            if join:
                edit.joinPreviousEditBlock()
            else:
                edit.beginEditBlock()
            if region is not None:
                edit.setPosition(region.start)
                edit.setPosition(region.start + region.length,
                                 QTextCursor.MoveMode.KeepAnchor)
                edit.removeSelectedText()
                edit.setPosition(region.start)
            else:
                edit.setPosition(0)
                edit.insertBlock()
                edit.movePosition(QTextCursor.MoveOperation.PreviousBlock)
                if edit.block().text():
                    edit.movePosition(QTextCursor.MoveOperation.NextBlock)
                    edit.insertBlock()
                    edit.movePosition(QTextCursor.MoveOperation.PreviousBlock)
            start = edit.block().position()
            for text, block_fmt, char_fmt in specs:
                edit.setBlockFormat(block_fmt)
                edit.insertText(text, char_fmt)
                edit.insertBlock()
            edit.deletePreviousChar()
            end = edit.position()
            edit.endEditBlock()
        self.state.toc = TocRegion(start, max(0, end - start))

    def _update_toc(self) -> None:
        toc = self.state.toc
        if toc and not self.tracker.allows_edit(toc.start, toc.start + toc.length, formatting=True):
            return
        if not self._collect_toc_entries() and self.state.toc is None:
            QMessageBox.information(
                self, "Table of Contents",
                "Add some Heading 1-3 paragraphs first, then insert a "
                "table of contents.")
            return
        doc = self.editor.document()
        changed = False
        for attempt in range(3):
            publication = Publication(doc, self.state.page, self._title)
            entries_pages = [
                (level, text, publication.page_for_position(pos))
                for pos, level, text in self._collect_toc_entries()
            ]
            specs = self._toc_line_specs(entries_pages)
            plain = "\n".join(text for text, _b, _c in specs)
            region = self.state.toc
            if region is not None and region.length > 0:
                probe = QTextCursor(doc)
                probe.setPosition(region.start)
                probe.setPosition(region.start + region.length,
                                  QTextCursor.MoveMode.KeepAnchor)
                current = probe.selectedText().replace(" ", "\n")
                current = current.replace(" ", "")
                if current == plain:
                    break
            self._apply_toc_lines(specs, region, join=attempt > 0)
            changed = True
        if changed:
            self._set_dirty(True)
        self._refresh_outline()

    def _refresh_toc_quietly(self) -> None:
        if self.state.toc and self.state.toc.length > 0:
            try:
                self._update_toc()
            except Exception:
                pass


    def _toggle_proofing(self, checked: bool) -> None:
        self._proof_enabled = bool(checked)
        self.store.set_setting("proofing", self._proof_enabled)
        self.highlighter.set_enabled(self._proof_enabled)
        self.proof_label.setText(
            "Proofing on" if self._proof_enabled else "Proofing off")
        if self._proof_enabled:
            self._run_proofing()
        else:
            self.highlighter.set_issues([])
            self._proof_issues = []
            self._refresh_review_panes()

    def _toggle_tracking(self, checked: bool) -> None:
        if not self._can_edit():
            return
        if self.hooks.mode == "Reviewing":
            checked = True
        self.tracker.set_tracking(bool(checked))
        if checked:
            self._show_review_pane()
        self._set_dirty(True)

    def _new_comment(self) -> None:
        if not self._can_edit():
            return
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            start, end = cursor.selectionStart(), cursor.selectionEnd()
        else:
            start = cursor.block().position()
            end = start + max(0, cursor.block().length() - 1)
        text, ok = QInputDialog.getMultiLineText(
            self, "New Comment", "Comment:")
        if not ok or not text.strip():
            return
        comment = self.tracker.add_comment(start, end, text.strip(), self._author)
        self.hooks.dispatch_comment(comment)
        self._show_review_pane()

    def _author_dialog(self) -> None:
        dialog = dialogs.AuthorDialog(self._author, self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        value = dialog.value()
        if value:
            self._author = value
            self.tracker.author = value
            self.store.set_setting("author", value)

    def _selected_revision(self):
        item = self.changes_list.currentItem()
        if item is None:
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def _accept_selected(self) -> None:
        if not self._can_edit():
            return
        self._show_review_pane()
        rev_id = self._selected_revision()
        if not rev_id:
            self.statusBar().showMessage("Select a change first.", 3000)
            return
        try:
            self.tracker.accept_revision(rev_id)
        except ReviewError as exc:
            QMessageBox.warning(self, "Accept", str(exc))

    def _reject_selected(self) -> None:
        if not self._can_edit():
            return
        self._show_review_pane()
        rev_id = self._selected_revision()
        if not rev_id:
            self.statusBar().showMessage("Select a change first.", 3000)
            return
        try:
            self.tracker.reject_revision(rev_id)
        except ReviewError as exc:
            QMessageBox.warning(self, "Reject", str(exc))

    def _accept_all(self) -> None:
        if not self._can_edit():
            return
        self._show_review_pane()
        self.tracker.accept_all()

    def _reject_all(self) -> None:
        if not self._can_edit():
            return
        self._show_review_pane()
        try:
            self.tracker.reject_all()
        except ReviewError as exc:
            QMessageBox.warning(self, "Reject All", str(exc))

    def _resolve_selected_comment(self) -> None:
        if not self._can_edit():
            return
        item = self.comments_list.currentItem()
        if item is None:
            return
        self.tracker.resolve_comment(item.data(Qt.ItemDataRole.UserRole))

    def _comment_jump(self, item: QListWidgetItem) -> None:
        comment = self._find_comment(item.data(Qt.ItemDataRole.UserRole))
        if comment is None:
            return
        cursor = self.editor.textCursor()
        cursor.setPosition(comment.start)
        cursor.setPosition(comment.end, QTextCursor.MoveMode.KeepAnchor)
        self.editor.setTextCursor(cursor)
        self.canvas.scroll_to_cursor()
        self._apply_review_highlights(active_comment=comment.id)

    def _find_comment(self, comment_id: str):
        for comment in self.state.comments:
            if comment.id == comment_id:
                return comment
        return None

    def _change_jump(self, item: QListWidgetItem) -> None:
        rev_id = item.data(Qt.ItemDataRole.UserRole)
        for rev in self.state.revisions:
            if rev.id == rev_id:
                cursor = self.editor.textCursor()
                cursor.setPosition(rev.start)
                cursor.setPosition(rev.start + rev.length,
                                   QTextCursor.MoveMode.KeepAnchor)
                self.editor.setTextCursor(cursor)
                self.canvas.scroll_to_cursor()
                return

    def _refresh_review_panes(self) -> None:
        self.comments_list.clear()
        for comment in self.state.comments:
            if comment.resolved or not self._show_comments:
                continue
            quote = comment.quote.replace("\n", " ")
            if len(quote) > 40:
                quote = quote[:40] + "…"
            label = f"{comment.author}: {comment.text}  (on “{quote}”)"
            if comment.replies:
                label += f"  [{len(comment.replies)} replies]"
            if comment.orphaned:
                label += "  [orphaned]"
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, comment.id)
            self.comments_list.addItem(item)
        self.changes_list.clear()
        for rev in self.state.revisions:
            if rev.status != "pending" or not self._show_changes:
                continue
            if rev.kind == "format":
                label = f"Format “{rev.after_text}”"
            elif rev.before_text and rev.after_text:
                label = f"Replace “{rev.before_text}” with “{rev.after_text}”"
            elif rev.after_text:
                label = f"Insert “{rev.after_text}”"
            else:
                label = f"Delete “{rev.before_text}”"
            if rev.conflicted:
                label += "  [conflicted]"
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, rev.id)
            self.changes_list.addItem(item)
        self.proof_list.clear()
        for issue in self._proof_issues[:200]:
            item = QListWidgetItem(f"[{issue.kind}] {issue.message}")
            item.setData(Qt.ItemDataRole.UserRole,
                         (issue.start, issue.length))
            self.proof_list.addItem(item)

    def _apply_review_highlights(self, active_comment: str | None = None) -> None:
        selections = []
        doc = self.editor.document()
        length = doc.characterCount()
        for rev in self.state.revisions:
            if rev.status != "pending" or not rev.after_text or not self._show_changes:
                continue
            selection = QTextCursor(doc)
            selection.setPosition(min(rev.start, max(0, length - 1)))
            selection.setPosition(
                min(rev.start + rev.length, max(0, length - 1)),
                QTextCursor.MoveMode.KeepAnchor)
            fmt = QTextCharFormat()
            fmt.setBackground(QColor("#dcf5dd"))
            extra = QTextEdit.ExtraSelection()
            extra.cursor = selection
            extra.format = fmt
            selections.append(extra)
        for comment in self.state.comments:
            if comment.resolved or not self._show_comments:
                continue
            if active_comment and comment.id != active_comment:
                continue
            selection = QTextCursor(doc)
            selection.setPosition(min(comment.start, max(0, length - 1)))
            selection.setPosition(
                min(comment.end, max(0, length - 1)),
                QTextCursor.MoveMode.KeepAnchor)
            fmt = QTextCharFormat()
            fmt.setBackground(QColor("#fdf3c8"))
            extra = QTextEdit.ExtraSelection()
            extra.cursor = selection
            extra.format = fmt
            selections.append(extra)
        self.editor.setExtraSelections(selections)

    def _run_proofing(self) -> None:
        if not self._proof_enabled:
            return
        self._proof_generation += 1
        text = self.editor.document().toPlainText()
        job = _ProofJob(self.proof_engine, text, self._proof_generation,
                        self._proof_emitter)
        self._pool.start(job)

    def _proof_results(self, generation: int, issues) -> None:
        if generation != self._proof_generation or not self._proof_enabled:
            return
        if issues is None:
            self.proof_label.setText("Proofing unavailable")
            self.statusBar().showMessage(
                "The proofing check failed; try again.", 4000)
            return
        issues = [i for i in issues
                  if getattr(i, "word", "").lower() not in self._ignored_words]
        self._proof_issues = issues
        self.highlighter.set_issues(issues)
        self._refresh_review_panes()
        self.proof_label.setText(
            f"Proofing on · {len(issues)} issue(s)")

    def _proof_jump(self, item: QListWidgetItem) -> None:
        start, length = item.data(Qt.ItemDataRole.UserRole)
        cursor = self.editor.textCursor()
        cursor.setPosition(start)
        cursor.setPosition(start + length, QTextCursor.MoveMode.KeepAnchor)
        self.editor.setTextCursor(cursor)
        self.canvas.scroll_to_cursor()

    def _proof_apply(self) -> None:
        if not self._can_edit():
            return
        item = self.proof_list.currentItem()
        if item is None:
            return
        start, length = item.data(Qt.ItemDataRole.UserRole)
        for issue in self._proof_issues:
            if issue.start == start and issue.length == length:
                if issue.replacement is None:
                    self.statusBar().showMessage(
                        "No automatic fix; use the context menu for "
                        "spelling suggestions.", 3000)
                    return
                cursor = self.editor.textCursor()
                cursor.setPosition(issue.start)
                cursor.setPosition(issue.start + issue.length,
                                   QTextCursor.MoveMode.KeepAnchor)
                if cursor.selectedText() != issue.word:
                    self.statusBar().showMessage(
                        "That issue is stale; run the check again.", 3000)
                    return
                cursor.insertText(issue.replacement)
                self._proof_generation += 1
                self._proof_timer.start()
                return


    def _refresh_counts(self) -> None:
        text = self.editor.document().toPlainText()
        words = len(text.split())
        chars = len(text)
        self.counts_label.setText(f"{words} words · {chars} characters")
        try:
            pages = Publication(self.editor.document(), self.state.page, self._title).page_count()
        except Exception:
            pages = 1
        self.page_label.setText(
            f"{max(1, pages)} page(s)")

    def _sync_all_panes(self) -> None:
        self._refresh_outline()
        self._refresh_review_panes()
        self._apply_review_highlights()
        self._refresh_counts()
        self._update_table_tools()


    def setup_editor_context_menu(self) -> None:
        self.editor.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.editor.customContextMenuRequested.connect(self._editor_menu)

    def _editor_menu(self, point: QPoint) -> None:
        menu = self.editor.createStandardContextMenu(point)
        cursor = self.editor.cursorForPosition(point)
        pos = cursor.position()
        issue = next((i for i in self._proof_issues
                      if i.start <= pos < i.start + i.length), None)
        first = menu.actions()[0] if menu.actions() else None
        if issue is not None and issue.kind == WORD and self._proof_enabled:
            menu.insertSeparator(first)
            for suggestion in self.proof_engine.suggestions(issue.word, 5):
                action = QAction(suggestion, menu)
                action.triggered.connect(
                    lambda _c, s=suggestion, iss=issue:
                        self._apply_spelling(iss, s))
                menu.insertAction(first, action)
            ignore = QAction("Ignore", menu)
            ignore.triggered.connect(lambda: self._ignore_issue(issue))
            menu.insertAction(first, ignore)
            add = QAction("Add to dictionary", menu)
            add.triggered.connect(
                lambda: self._add_to_dictionary(issue.word))
            menu.insertAction(first, add)
        image_fmt = self.editor.selected_image_format()
        if image_fmt is not None:
            resize = QAction("Resize image…", menu)
            resize.triggered.connect(lambda: self._resize_image(image_fmt))
            menu.addAction(resize)
        menu.exec(self.editor.viewport().mapToGlobal(point))

    def _apply_spelling(self, issue, replacement: str) -> None:
        if not self._can_edit():
            return
        cursor = self.editor.textCursor()
        cursor.setPosition(issue.start)
        cursor.setPosition(issue.start + issue.length,
                           QTextCursor.MoveMode.KeepAnchor)
        if cursor.selectedText() != issue.word:
            self.statusBar().showMessage(
                "That issue is stale; run the check again.", 3000)
            return
        cursor.insertText(replacement)
        self._proof_generation += 1
        self._proof_timer.start()

    def _ignore_issue(self, issue) -> None:
        self._ignored_words.add(issue.word.lower())
        self._proof_issues = [i for i in self._proof_issues if i is not issue]
        self.highlighter.set_issues(self._proof_issues)
        self._refresh_review_panes()

    def _add_to_dictionary(self, word: str) -> None:
        self.proof_engine.add_to_dictionary(word)
        self._proof_generation += 1
        self._proof_timer.start()

    def _resize_image(self, image_fmt) -> None:
        if not self._can_edit():
            return
        dialog = dialogs.ImageSizeDialog(
            image_fmt.width(), self.editor.content_width_px(), self)
        if dialog.exec() != dialogs.QDialog.DialogCode.Accepted:
            return
        self.editor.resize_selected_image(dialog.values())
