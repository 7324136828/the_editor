from __future__ import annotations

import base64
import copy
import math
from pathlib import Path

from PySide6.QtCore import QRectF, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QImage, QPainter, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QColorDialog, QComboBox, QDialog,
    QDialogButtonBox, QDoubleSpinBox, QFileDialog, QFormLayout, QGridLayout,
    QHBoxLayout, QLabel, QLineEdit, QPlainTextEdit, QPushButton, QSpinBox,
    QTabWidget, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
)

from .design import (
    DesignSettings, FONT_PAIRINGS, MAX_WATERMARK_BYTES, PALETTES,
    SPACING_POLICIES, STYLE_SETS, THEMES, paint_page_decoration,
    theme_tokens, watermark_image,
)
from .document_features import align_objects, checked_features, group_objects


class _ColorButton(QPushButton):
    colorChanged = Signal()

    def __init__(self, value, title, parent=None):
        super().__init__(parent)
        self.title = title
        self.setAccessibleName(title)
        self.set_color(value)
        self.clicked.connect(self.choose_color)

    def set_color(self, value):
        color = QColor(value)
        if not color.isValid():
            raise ValueError("Choose a valid color.")
        self.value = color.name()
        self.setText(self.value.upper())
        text = "#ffffff" if color.lightness() < 140 else "#182638"
        self.setStyleSheet(f"background-color: {self.value}; color: {text}; padding: 5px;")
        self.colorChanged.emit()

    def choose_color(self):
        color = QColorDialog.getColor(QColor(self.value), self, self.title)
        if color.isValid():
            self.set_color(color.name())


def _combo(values, current, default=False):
    combo = QComboBox()
    if default:
        combo.addItem("Use document theme", "")
    for value in values:
        combo.addItem(value, value)
    combo.setCurrentIndex(max(0, combo.findData(current)))
    return combo


def _number(value, low, high, suffix="", decimals=1):
    spin = QDoubleSpinBox()
    spin.setRange(low, high)
    spin.setDecimals(decimals)
    spin.setValue(value)
    spin.setSuffix(suffix)
    return spin


class DesignDialog(QDialog):
    def __init__(self, settings: DesignSettings, parent=None):
        super().__init__(parent)
        settings.validate()
        self.setWindowTitle("Document Design")
        self.resize(710, 570)
        self._watermark_image = settings.watermark_image
        layout = QVBoxLayout(self)
        content = QHBoxLayout()
        tabs = QTabWidget()
        content.addWidget(tabs, 1)
        self.preview = QLabel()
        self.preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview.setFixedSize(245, 345)
        content.addWidget(self.preview)
        layout.addLayout(content)
        theme_page = QWidget()
        theme_form = QFormLayout(theme_page)
        self.theme_combo = _combo(THEMES, settings.theme)
        self.style_set_combo = _combo(STYLE_SETS, settings.style_set)
        self.palette_combo = _combo(PALETTES, settings.color_palette, True)
        self.font_pairing_combo = _combo(FONT_PAIRINGS, settings.font_pairing, True)
        self.spacing_combo = _combo(SPACING_POLICIES, settings.paragraph_spacing)
        for label, widget in (("Theme", self.theme_combo), ("Style set", self.style_set_combo),
                              ("Colors", self.palette_combo), ("Fonts", self.font_pairing_combo),
                              ("Paragraph spacing", self.spacing_combo)):
            theme_form.addRow(label, widget)
        note = QLabel("Themes update paragraphs bound to document styles. Direct text formatting is preserved.")
        note.setWordWrap(True)
        theme_form.addRow(note)
        tabs.addTab(theme_page, "Styles")
        page_page = QWidget()
        page_form = QFormLayout(page_page)
        self.page_color_button = _ColorButton(settings.page_color, "Page color")
        self.border_combo = _combo(("none", "box", "double", "dashed"), settings.page_border)
        self.border_color_button = _ColorButton(settings.page_border_color, "Page border color")
        self.border_width_spin = _number(settings.page_border_width_pt, 0.25, 12, " pt", 2)
        self.border_inset_spin = _number(settings.page_border_inset_mm, 0, 50, " mm")
        for label, widget in (("Page color", self.page_color_button), ("Border", self.border_combo),
                              ("Border color", self.border_color_button), ("Border width", self.border_width_spin),
                              ("Border inset", self.border_inset_spin)):
            page_form.addRow(label, widget)
        tabs.addTab(page_page, "Page")
        watermark_page = QWidget()
        watermark_form = QFormLayout(watermark_page)
        self.watermark_kind_combo = _combo(("none", "text", "image"), settings.watermark_kind)
        self.watermark_text_edit = QLineEdit(settings.watermark_text)
        self.watermark_text_edit.setMaxLength(500)
        self.watermark_image_button = QPushButton("Choose image…")
        self.watermark_image_label = QLabel("Embedded image" if self._watermark_image else "No image selected")
        self.watermark_image_label.setWordWrap(True)
        self.watermark_color_button = _ColorButton(settings.watermark_color, "Watermark text color")
        self.watermark_opacity_spin = _number(settings.watermark_opacity * 100, 0, 100, "%", 0)
        self.watermark_angle_spin = _number(settings.watermark_angle, -180, 180, "°", 0)
        self.watermark_layer_combo = _combo(("behind", "front"), settings.watermark_layer)
        for label, widget in (("Watermark", self.watermark_kind_combo), ("Text", self.watermark_text_edit),
                              ("Image", self.watermark_image_button), ("", self.watermark_image_label),
                              ("Text color", self.watermark_color_button), ("Opacity", self.watermark_opacity_spin),
                              ("Angle", self.watermark_angle_spin), ("Layer", self.watermark_layer_combo)):
            watermark_form.addRow(label, widget)
        tabs.addTab(watermark_page, "Watermark")
        self.error_label = QLabel()
        self.error_label.setWordWrap(True)
        self.error_label.setStyleSheet("color: #a12727;")
        layout.addWidget(self.error_label)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self.watermark_image_button.clicked.connect(self.choose_watermark_image)
        for combo in self.findChildren(QComboBox):
            combo.currentIndexChanged.connect(self.refresh_preview)
        for spin in self.findChildren(QDoubleSpinBox):
            spin.valueChanged.connect(self.refresh_preview)
        for button in self.findChildren(_ColorButton):
            button.colorChanged.connect(self.refresh_preview)
        self.watermark_text_edit.textChanged.connect(self.refresh_preview)
        self.refresh_preview()

    def settings(self):
        value = DesignSettings(
            theme=self.theme_combo.currentData(), style_set=self.style_set_combo.currentData(),
            color_palette=self.palette_combo.currentData(), font_pairing=self.font_pairing_combo.currentData(),
            paragraph_spacing=self.spacing_combo.currentData(), page_color=self.page_color_button.value,
            page_border=self.border_combo.currentData(), page_border_color=self.border_color_button.value,
            page_border_width_pt=self.border_width_spin.value(), page_border_inset_mm=self.border_inset_spin.value(),
            watermark_kind=self.watermark_kind_combo.currentData(), watermark_text=self.watermark_text_edit.text(),
            watermark_image=self._watermark_image, watermark_color=self.watermark_color_button.value,
            watermark_opacity=self.watermark_opacity_spin.value() / 100,
            watermark_angle=self.watermark_angle_spin.value(), watermark_layer=self.watermark_layer_combo.currentData(),
        )
        value.validate()
        return value

    def set_watermark_image(self, path):
        path = Path(path)
        if path.stat().st_size > MAX_WATERMARK_BYTES:
            raise ValueError("Watermark images are limited to 8 MiB.")
        with path.open("rb") as stream:
            payload = stream.read(MAX_WATERMARK_BYTES + 1)
        suffix = path.suffix.lower().lstrip(".")
        if suffix not in ("png", "jpg", "jpeg", "webp"):
            raise ValueError("Select a PNG, JPEG, or WebP watermark image.")
        uri = f"data:image/{suffix};base64," + base64.b64encode(payload).decode("ascii")
        watermark_image(uri)
        self._watermark_image = uri
        self.watermark_image_label.setText(path.name)
        self.watermark_kind_combo.setCurrentIndex(self.watermark_kind_combo.findData("image"))
        self.refresh_preview()

    def choose_watermark_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "Watermark Image", "", "Images (*.png *.jpg *.jpeg *.webp)")
        if path:
            try:
                self.set_watermark_image(path)
            except (OSError, ValueError) as error:
                self.error_label.setText(str(error))

    def refresh_preview(self):
        kind = self.watermark_kind_combo.currentData()
        self.watermark_text_edit.setEnabled(kind == "text")
        self.watermark_color_button.setEnabled(kind == "text")
        self.watermark_image_button.setEnabled(kind == "image")
        try:
            settings = self.settings()
        except ValueError as error:
            self.error_label.setText(str(error))
            self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(False)
            return
        self.error_label.clear()
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(True)
        image = QImage(238, 326, QImage.Format.Format_ARGB32_Premultiplied)
        image.fill(QColor("#ffffff"))
        painter = QPainter(image)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.scale(image.width() / 816, image.height() / 1056)
        rect = QRectF(0, 0, 816, 1056)
        paint_page_decoration(painter, rect, settings, "behind")
        tokens = theme_tokens(settings)
        from PySide6.QtGui import QFont
        font = QFont(tokens["heading_font"], 30)
        painter.setFont(font)
        painter.setPen(QColor(tokens["accents"][0]))
        painter.drawText(QRectF(90, 130, 636, 120), Qt.TextFlag.TextWordWrap, "Document design")
        painter.setFont(QFont(tokens["body_font"], 15))
        painter.setPen(QColor(tokens["text"]))
        painter.drawText(QRectF(90, 270, 636, 370), Qt.TextFlag.TextWordWrap,
                         "Your document’s headings, body text, page color, border, and watermark share a consistent appearance.")
        paint_page_decoration(painter, rect, settings, "front")
        painter.end()
        self.preview.setPixmap(QPixmap.fromImage(image))

    def accept(self):
        try:
            self.settings()
        except ValueError as error:
            self.error_label.setText(str(error))
            return
        super().accept()


class FloatingInspector(QWidget):
    def __init__(self, objects, on_change, parent=None):
        super().__init__(parent)
        self.objects = objects
        self.on_change = on_change
        self._refreshing = False
        layout = QVBoxLayout(self)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Object", "Type", "Page"])
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tree.setMinimumHeight(145)
        layout.addWidget(self.tree)
        geometry = QGridLayout()
        self.x_spin = _number(0, -100000, 100000, " px")
        self.y_spin = _number(0, -100000, 100000, " px")
        self.width_spin = _number(200, 1, 100000, " px")
        self.height_spin = _number(100, 1, 100000, " px")
        for row, (label1, widget1, label2, widget2) in enumerate((
            ("X", self.x_spin, "Y", self.y_spin), ("Width", self.width_spin, "Height", self.height_spin),
        )):
            geometry.addWidget(QLabel(label1), row, 0)
            geometry.addWidget(widget1, row, 1)
            geometry.addWidget(QLabel(label2), row, 2)
            geometry.addWidget(widget2, row, 3)
        self.page_spin = QSpinBox()
        self.page_spin.setRange(1, 10001)
        self.wrap_combo = _combo(("square", "tight", "behind", "front"), "front")
        geometry.addWidget(QLabel("Page"), 2, 0)
        geometry.addWidget(self.page_spin, 2, 1)
        geometry.addWidget(QLabel("Wrap"), 2, 2)
        geometry.addWidget(self.wrap_combo, 2, 3)
        layout.addLayout(geometry)
        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlaceholderText("Floating text box content")
        self.text_edit.setMaximumHeight(90)
        layout.addWidget(self.text_edit)
        self.apply_button = QPushButton("Apply Object Properties")
        self.apply_button.clicked.connect(self.apply_properties)
        layout.addWidget(self.apply_button)
        actions = QGridLayout()
        self.action_buttons = {}
        for index, (name, label) in enumerate((
            ("forward", "Bring Forward"), ("backward", "Send Backward"),
            ("left", "Align Left"), ("center", "Align Center"),
            ("top", "Align Top"), ("distribute", "Distribute"),
            ("group", "Group"), ("ungroup", "Ungroup"), ("delete", "Delete"),
        )):
            button = QPushButton(label)
            button.clicked.connect(lambda checked=False, operation=name: self.perform(operation))
            self.action_buttons[name] = button
            actions.addWidget(button, index // 2, index % 2)
        layout.addLayout(actions)
        self.error_label = QLabel()
        self.error_label.setStyleSheet("color: #a12727;")
        self.error_label.setWordWrap(True)
        layout.addWidget(self.error_label)
        self.tree.itemSelectionChanged.connect(self.selection_changed)
        self.tree.itemChanged.connect(self.item_changed)
        self.refresh()

    def set_objects(self, objects):
        self.objects = objects
        self.refresh()

    def selected_ids(self):
        return {item.data(0, Qt.ItemDataRole.UserRole) for item in self.tree.selectedItems()}

    def select_ids(self, selected):
        self._refreshing = True
        for key, item in self._items.items():
            item.setSelected(key in selected)
        self._refreshing = False
        self.selection_changed()

    def refresh(self, selected=None):
        selected = self.selected_ids() if selected is None else selected
        self._refreshing = True
        self.tree.clear()
        self._items = {}
        for obj in sorted(self.objects, key=lambda value: value.get("z", 0), reverse=True):
            item = QTreeWidgetItem([obj.get("name") or obj.get("text", "")[:24] or obj["kind"].title(), obj["kind"].title(), str(obj.get("page", 0) + 1)])
            item.setData(0, Qt.ItemDataRole.UserRole, obj["id"])
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEditable)
            item.setCheckState(0, Qt.CheckState.Checked if obj.get("visible", True) else Qt.CheckState.Unchecked)
            self._items[obj["id"]] = item
        for obj in sorted(self.objects, key=lambda value: value.get("z", 0), reverse=True):
            item = self._items[obj["id"]]
            parent = self._items.get(obj.get("group"))
            if parent:
                parent.addChild(item)
            else:
                self.tree.addTopLevelItem(item)
        self.tree.expandAll()
        for key, item in self._items.items():
            item.setSelected(key in selected)
        self.tree.resizeColumnToContents(0)
        self._refreshing = False
        self.selection_changed()

    def selection_changed(self):
        if self._refreshing:
            return
        selected = self.selected_ids()
        items = [obj for obj in self.objects if obj["id"] in selected]
        enabled = len(items) == 1
        for widget in (self.x_spin, self.y_spin, self.width_spin, self.height_spin, self.page_spin,
                       self.wrap_combo, self.apply_button, self.text_edit):
            widget.setEnabled(enabled)
        if enabled:
            obj = items[0]
            self.x_spin.setValue(obj["x"])
            self.y_spin.setValue(obj["y"])
            self.width_spin.setValue(obj["width"])
            self.height_spin.setValue(obj["height"])
            self.page_spin.setValue(obj.get("page", 0) + 1)
            self.wrap_combo.setCurrentIndex(self.wrap_combo.findData(obj.get("wrap", "front")))
            self.text_edit.setPlainText(obj.get("text", ""))
            self.text_edit.setEnabled(obj["kind"] == "text")
        for operation, button in self.action_buttons.items():
            button.setEnabled(bool(items) if operation in ("forward", "backward", "delete")
                              else any(obj["kind"] == "group" for obj in items) if operation == "ungroup"
                              else len(items) >= 2)

    def _mutate(self, action):
        original = copy.deepcopy(self.objects)
        selected = self.selected_ids()
        try:
            selection = action()
            checked_features({"objects": self.objects})
        except (ValueError, TypeError, KeyError) as error:
            self.objects[:] = original
            self.error_label.setText(str(error))
            self.refresh(selected)
            return False
        self.error_label.clear()
        self.refresh(selected if selection is None else selection)
        self.on_change()
        return True

    def item_changed(self, item, column):
        if self._refreshing:
            return
        key = item.data(0, Qt.ItemDataRole.UserRole)
        obj = next((value for value in self.objects if value["id"] == key), None)
        if obj is None:
            return
        if column != 0:
            self._refreshing = True
            item.setText(column, obj["kind"].title() if column == 1 else str(obj.get("page", 0) + 1))
            self._refreshing = False
            return
        obj["visible"] = item.checkState(0) == Qt.CheckState.Checked
        obj["name"] = item.text(0).strip()[:200] or obj["kind"].title()
        self.error_label.clear()
        QTimer.singleShot(0, self.on_change)

    def _descendants(self, selected):
        result = set(selected)
        while True:
            children = {obj["id"] for obj in self.objects if obj.get("group") in result}
            if children.issubset(result):
                return result
            result.update(children)

    def apply_properties(self):
        selected = self.selected_ids()
        if len(selected) != 1:
            return
        obj = next(value for value in self.objects if value["id"] in selected)
        values = {"x": self.x_spin.value(), "y": self.y_spin.value(), "width": self.width_spin.value(),
                  "height": self.height_spin.value(), "page": self.page_spin.value() - 1,
                  "wrap": self.wrap_combo.currentData()}
        text = self.text_edit.toPlainText()

        def update():
            if obj["kind"] == "group":
                descendants = self._descendants(selected) - selected
                for child in self.objects:
                    if child["id"] in descendants:
                        child["x"] = values["x"] + (child["x"] - obj["x"]) * values["width"] / obj["width"]
                        child["y"] = values["y"] + (child["y"] - obj["y"]) * values["height"] / obj["height"]
                        child["width"] *= values["width"] / obj["width"]
                        child["height"] *= values["height"] / obj["height"]
                        child["page"] = values["page"]
                        child["wrap"] = values["wrap"]
            obj.update(values)
            if obj["kind"] == "text":
                obj["text"] = text

        self._mutate(update)

    def perform(self, operation):
        selected = self.selected_ids()
        if not selected:
            return

        def update():
            by_id = {obj["id"]: obj for obj in self.objects}
            roots = {key for key in selected if not any(ancestor in selected for ancestor in self._ancestors(by_id[key], by_id))}
            if operation in ("left", "center", "top", "distribute"):
                align_objects(self.objects, roots, operation)
            elif operation == "group":
                return {group_objects(self.objects, roots)}
            elif operation == "ungroup":
                groups = {key for key in roots if by_id[key]["kind"] == "group"}
                selection = set()
                for obj in self.objects:
                    if obj.get("group") in groups:
                        parent = by_id[obj["group"]].get("group")
                        if parent:
                            obj["group"] = parent
                        else:
                            obj.pop("group", None)
                        selection.add(obj["id"])
                self.objects[:] = [obj for obj in self.objects if obj["id"] not in groups]
                return selection
            elif operation == "delete":
                remove = self._descendants(roots)
                self.objects[:] = [obj for obj in self.objects if obj["id"] not in remove]
                return set()
            elif operation in ("forward", "backward"):
                stack = sorted(self.objects, key=lambda obj: obj.get("z", 0))
                moving = self._descendants(roots)
                indexes = range(len(stack) - 2, -1, -1) if operation == "forward" else range(1, len(stack))
                offset = 1 if operation == "forward" else -1
                for index in indexes:
                    if stack[index]["id"] in moving and stack[index + offset]["id"] not in moving:
                        stack[index], stack[index + offset] = stack[index + offset], stack[index]
                for index, obj in enumerate(stack):
                    obj["z"] = index
            else:
                raise ValueError("Unknown object operation.")

        self._mutate(update)

    def _ancestors(self, obj, by_id):
        result = set()
        parent = obj.get("group")
        while parent and parent not in result:
            result.add(parent)
            parent = by_id[parent].get("group")
        return result


class SignatureInputDialog(QDialog):
    def __init__(self, parent=None, name=""):
        super().__init__(parent)
        self.setWindowTitle("Sign Document Field")
        self.resize(460, 240)
        layout = QVBoxLayout(self)
        description = QLabel("Enter your name to sign this field. Folio records your typed name, signing time, and a document text fingerprint locally. This is a standard electronic signature; identity and certificates are not verified.")
        description.setWordWrap(True)
        layout.addWidget(description)
        form = QFormLayout()
        self.name_edit = QLineEdit(name)
        self.name_edit.setMaxLength(200)
        form.addRow("Full name", self.name_edit)
        layout.addLayout(form)
        self.consent_check = QCheckBox("I intend to sign this document with the name above.")
        layout.addWidget(self.consent_check)
        self.error_label = QLabel()
        self.error_label.setWordWrap(True)
        self.error_label.setStyleSheet("color: #a12727;")
        layout.addWidget(self.error_label)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Sign")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.name_edit.textChanged.connect(self.refresh)
        self.consent_check.toggled.connect(self.refresh)
        layout.addWidget(self.buttons)
        self.refresh()

    def signature_name(self):
        return self.name_edit.text().strip()

    def consent(self):
        return self.consent_check.isChecked()

    def refresh(self):
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(bool(self.signature_name()) and self.consent())

    def accept(self):
        if not self.signature_name() or not self.consent():
            self.error_label.setText("Enter your name and confirm your intent to sign.")
            return
        super().accept()


class TableDimensionsDialog(QDialog):
    def __init__(self, columns, rows, widths=None, heights=None, parent=None):
        super().__init__(parent)
        if type(columns) is not int or type(rows) is not int or min(columns, rows) < 1:
            raise ValueError("The table must have at least one row and column.")
        self.columns, self.rows = columns, rows
        self.setWindowTitle("Table Dimensions")
        self.resize(500, 240)
        layout = QVBoxLayout(self)
        hint = QLabel(f"Table: {columns} columns × {rows} rows. Enter comma-separated sizes in points. One value applies to every column or row. Leave blank to keep the current dimensions.")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        form = QFormLayout()
        self.widths_edit = QLineEdit(", ".join(f"{value:g}" for value in widths) if widths else "")
        self.heights_edit = QLineEdit(", ".join(f"{value:g}" for value in heights) if heights else "")
        self.widths_edit.setPlaceholderText("e.g. 120, 180, 120")
        self.heights_edit.setPlaceholderText("e.g. 24")
        form.addRow("Column widths (pt)", self.widths_edit)
        form.addRow("Minimum row heights (pt)", self.heights_edit)
        layout.addLayout(form)
        self.error_label = QLabel()
        self.error_label.setWordWrap(True)
        self.error_label.setStyleSheet("color: #a12727;")
        layout.addWidget(self.error_label)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

    def _sizes(self, text, count, label):
        if not text.strip():
            return None
        try:
            values = [float(item.strip()) for item in text.split(",")]
        except ValueError as error:
            raise ValueError(f"{label} must be comma-separated numbers.") from error
        if len(values) not in (1, count):
            raise ValueError(f"Enter one {label.lower()} value or exactly {count} values.")
        if any(not math.isfinite(value) or not 1 <= value <= 10000 for value in values):
            raise ValueError(f"{label} must be between 1 and 10,000 points.")
        return values * count if len(values) == 1 else values

    def dimensions(self):
        return (self._sizes(self.widths_edit.text(), self.columns, "Column widths"),
                self._sizes(self.heights_edit.text(), self.rows, "Row heights"))

    def accept(self):
        try:
            self.dimensions()
        except ValueError as error:
            self.error_label.setText(str(error))
            return
        super().accept()
