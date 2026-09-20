from __future__ import annotations

from PySide6.QtCore import QSignalBlocker, Qt
from PySide6.QtWidgets import (QDialog, QDockWidget, QMenu, QTextBrowser,
                              QVBoxLayout)

from .view_tools import DocumentViewport
from .workspace_tools import (FeedbackDialog, HelpDialog, MacroDialog,
                              RELEASE_NOTES, TutorialDialog, document_properties)


class ViewCommands:
    def _build_view_tab(self):
        groups = []
        group, row = self._group("Document Views")
        for label, mode in (("Print Layout", "Print Layout"), ("Web Layout", "Web Layout"),
                            ("Read Mode", "Read Mode")):
            row.addWidget(self._tool("preview", label,
                                    lambda _checked=False, value=mode: self._open_view(value), text_mode=True))
        row = self._extra_row(group)
        for label in ("Outline Mode", "Draft Mode"):
            row.addWidget(self._tool("nav", label,
                                    lambda _checked=False, value=label: self._open_view(value), text_mode=True))
        row.addWidget(self._tool("preview", "Immersive Reader", self._immersive_reader, text_mode=True))
        groups.append(group)
        group, row = self._group("Show")
        row.addWidget(self._tool("nav", "Navigation", self._toggle_nav, text_mode=True))
        row.addWidget(self._tool("focus", "Focus", self._toggle_focus, text_mode=True))
        row = self._extra_row(group)
        ruler = self._tool("ruler", "Ruler", lambda checked: self.ruler.setVisible(checked),
                           checkable=True, text_mode=True)
        with QSignalBlocker(ruler):
            ruler.setChecked(True)
        row.addWidget(ruler)
        row.addWidget(self._tool("table", "Gridlines", self._gridlines, checkable=True, text_mode=True))
        groups.append(group)
        group, row = self._group("Zoom")
        row.addWidget(self._tool("zoom-in", "100%", lambda: self._set_zoom(1.0), text_mode=True))
        row.addWidget(self._tool("zoom-fit", "Page Width", self._zoom_fit, text_mode=True))
        row = self._extra_row(group)
        row.addWidget(self._tool("zoom-fit", "Single Page", lambda: self._fit_pages(False), text_mode=True))
        row.addWidget(self._tool("zoom-fit", "Multiple Pages", lambda: self._fit_pages(True), text_mode=True))
        row.addWidget(self._tool("zoom-in", "Zoom…", self._zoom_dialog, text_mode=True))
        groups.append(group)
        group, row = self._group("Windows")
        row.addWidget(self._tool("preview", "New View", lambda: self._open_view("Web Layout"), text_mode=True))
        row.addWidget(self._tool("preview", "Split", self._split_view, text_mode=True))
        row = self._extra_row(group)
        windows = QMenu(self)
        windows.aboutToShow.connect(lambda: self._view_windows_menu(windows))
        row.addWidget(self._menu_tool("Switch Windows", windows))
        row.addWidget(self._tool("history", "History", self._show_history, text_mode=True))
        groups.append(group)
        group, row = self._group("Automation & Properties")
        row.addWidget(self._tool("field", "Macros", self._macros, text_mode=True))
        row = self._extra_row(group)
        row.addWidget(self._tool("about", "Properties", self._properties, text_mode=True))
        groups.append(group)
        self.ribbon.addTab(self._wrap(groups), "View")

    def _open_view(self, mode="Print Layout"):
        view = DocumentViewport(self, mode, self)
        view.setWindowFlag(Qt.WindowType.Window, True)
        view.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        if not hasattr(self, "_document_views"):
            self._document_views = []
        self._document_views.append(view)
        view.destroyed.connect(lambda: self._document_views.remove(view)
                               if view in self._document_views else None)
        view.show()
        return view

    def _split_view(self):
        if hasattr(self, "_split_dock"):
            self._split_dock.setVisible(not self._split_dock.isVisible())
            return
        dock = QDockWidget("Document split", self)
        dock.setAllowedAreas(Qt.DockWidgetArea.BottomDockWidgetArea | Qt.DockWidgetArea.TopDockWidgetArea)
        dock.setWidget(DocumentViewport(self, "Web Layout", dock))
        self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, dock)
        self._split_dock = dock
        dock.show()

    def _fit_pages(self, multiple):
        if multiple:
            view = self._open_view()
            view.zoom.setCurrentText("Multiple Pages")
        else:
            width, height = self.state.page.size_mm()
            from .editor import mm_to_px
            self._set_zoom(min((self.canvas.viewport().width() - 96) / mm_to_px(width),
                               (self.canvas.viewport().height() - 96) / mm_to_px(height)))

    def _view_windows_menu(self, menu):
        self._populate_windows(menu)
        for view in getattr(self, "_document_views", []):
            if view.isVisible():
                menu.addAction(view.windowTitle(), lambda _checked=False, target=view: self._activate_window(target))

    def _macros(self):
        MacroDialog(self).exec()

    def _properties(self):
        if not hasattr(self, "_properties_dock"):
            dock = QDockWidget("Document Properties", self)
            dock.setWidget(QTextBrowser())
            self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, dock)
            self._properties_dock = dock
        try:
            values = document_properties(self)
            text = "\n\n".join(f"{key}\n{value}" for key, value in values.items())
        except OSError as error:
            text = str(error)
        self._properties_dock.widget().setPlainText(text)
        self._properties_dock.show()

    def _build_help_tab(self):
        groups = []
        group, row = self._group("Help & Learning")
        row.addWidget(self._tool("about", "Search Help", self._search_help, text_mode=True))
        row.addWidget(self._tool("nav", "Tutorials", lambda: TutorialDialog(self).exec(), text_mode=True))
        groups.append(group)
        group, row = self._group("Support")
        row.addWidget(self._tool("comment", "Feedback", lambda: FeedbackDialog(self).exec(), text_mode=True))
        row.addWidget(self._tool("about", "Release Notes", self._release_notes, text_mode=True))
        groups.append(group)
        group, row = self._group("Folio")
        row.addWidget(self._tool("about", "About Folio", self._about, text_mode=True))
        groups.append(group)
        self.ribbon.addTab(self._wrap(groups), "Help")

    def _search_help(self):
        HelpDialog(self).exec()

    def _release_notes(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Bundled Release Notes")
        dialog.resize(720, 500)
        layout = QVBoxLayout(dialog)
        body = QTextBrowser()
        body.setPlainText(RELEASE_NOTES)
        layout.addWidget(body)
        dialog.exec()

    def _close_document_views(self):
        for view in list(getattr(self, "_document_views", [])):
            view.close()
        if hasattr(self, "_split_dock"):
            self._split_dock.close()
            self._split_dock.deleteLater()
            del self._split_dock
