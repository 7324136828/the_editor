from __future__ import annotations

from functools import lru_cache

from PySide6.QtCore import QByteArray, Qt
from PySide6.QtGui import QColor, QIcon, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer

_SVG_OPEN = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16" '
    'fill="none" stroke="{c}" stroke-width="1.3" '
    'stroke-linecap="round" stroke-linejoin="round">'
)

_PATHS = {
    "new": '<path d="M9 1.5H4a1 1 0 0 0-1 1v11a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V5z"/><path d="M9 1.5V5h3.5"/><path d="M8 8v4M6 10h4"/>',
    "open": '<path d="M2 13.5V4a1 1 0 0 1 1-1h3l1.5 2H13a1 1 0 0 1 1 1v1"/><path d="M2 13.5l1.6-5.4a1 1 0 0 1 1-.6h9a1 1 0 0 1 .95 1.3l-1.2 4a1 1 0 0 1-.95.7z"/>',
    "save": '<path d="M3 2h8l2 2v9a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V3a1 1 0 0 1 1-1z"/><path d="M5 2v4h6V2"/><path d="M5 14v-5h6v5"/>',
    "pdf": '<path d="M9 1.5H4a1 1 0 0 0-1 1v11a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V5z"/><path d="M9 1.5V5h3.5"/><path d="M5 11.5V9h1.2a1 1 0 0 1 0 2H5zM9 11.5v-3h.8a1.2 1.2 0 0 1 0 3zM13 8.5h-1.5v3"/>',
    "print": '<path d="M5 6V2h6v4"/><path d="M5 11H3a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1h-2"/><path d="M5 10h6v4H5z"/>',
    "preview": '<path d="M1.5 8s2.5-4.5 6.5-4.5S14.5 8 14.5 8 12 12.5 8 12.5 1.5 8 1.5 8z"/><circle cx="8" cy="8" r="2"/>',
    "undo": '<path d="M6 4L2.5 7.5 6 11"/><path d="M3 7.5h6a4 4 0 0 1 0 8H8"/>',
    "redo": '<path d="M10 4l3.5 3.5L10 11"/><path d="M13 7.5H7a4 4 0 0 0 0 8h1"/>',
    "cut": '<circle cx="4.5" cy="4.5" r="2"/><circle cx="4.5" cy="11.5" r="2"/><path d="M6.2 5.8L13.5 13M6.2 10.2L13.5 3"/>',
    "copy": '<rect x="5" y="5" width="9" height="9" rx="1"/><path d="M11 5V3a1 1 0 0 0-1-1H3a1 1 0 0 0-1 1v7a1 1 0 0 0 1 1h2"/>',
    "paste": '<rect x="3" y="4" width="10" height="11" rx="1"/><path d="M6 4V2.5h4V4"/><path d="M6 4a1 1 0 0 0-1 1v1h6V5a1 1 0 0 0-1-1"/>',
    "bold": '<path d="M4.5 2.5h4a2.5 2.5 0 0 1 0 5h-4zM4.5 7.5h5a3 3 0 0 1 0 6h-5z" stroke-width="1.6"/>',
    "italic": '<path d="M6.5 2.5h5M4.5 13.5h5M9.5 2.5l-3 11"/>',
    "underline": '<path d="M4.5 2.5v6a3.5 3.5 0 0 0 7 0v-6"/><path d="M3.5 13.5h9"/>',
    "strike": '<path d="M3 8h10"/><path d="M10.5 4.5c-.5-1.2-1.6-1.8-2.8-1.8-1.7 0-3 .8-3 2.2 0 .8.5 1.4 1.3 1.8M5.5 11.5c.5 1.2 1.5 1.8 2.7 1.8 1.7 0 3-.9 3-2.2 0-.7-.4-1.3-1-1.6"/>',
    "color": '<path d="M8 2.5v9M5.5 5.5h5M5.5 11h5" stroke-width="0"/><path d="M8 2l-3.5 7h7z"/><path d="M3.5 12.5h9v2h-9z" fill="{c}"/>',
    "highlight": '<path d="M9.5 2l4 4-6.5 6.5H3V8.5z"/><path d="M2.5 14.5h11"/>',
    "align-left": '<path d="M2.5 3.5h11M2.5 6.5h7M2.5 9.5h11M2.5 12.5h7"/>',
    "align-center": '<path d="M2.5 3.5h11M4.5 6.5h7M2.5 9.5h11M4.5 12.5h7"/>',
    "align-right": '<path d="M2.5 3.5h11M6.5 6.5h7M2.5 9.5h11M6.5 12.5h7"/>',
    "align-justify": '<path d="M2.5 3.5h11M2.5 6.5h11M2.5 9.5h11M2.5 12.5h11"/>',
    "bullets": '<circle cx="3" cy="4" r="0.8" fill="{c}"/><circle cx="3" cy="8" r="0.8" fill="{c}"/><circle cx="3" cy="12" r="0.8" fill="{c}"/><path d="M6 4h7.5M6 8h7.5M6 12h7.5"/>',
    "numbering": '<path d="M6 4h7.5M6 8h7.5M6 12h7.5"/><path d="M2.5 3v2.5M2.5 7.5h1.5L2.5 10h1.5M2.5 12.5a.8.8 0 1 1 .8.8H2.5v1"/>',
    "indent": '<path d="M6.5 3.5h7M9.5 6.5h4M9.5 9.5h4M6.5 12.5h7"/><path d="M3 6l3 2-3 2z" fill="{c}"/>',
    "outdent": '<path d="M6.5 3.5h7M9.5 6.5h4M9.5 9.5h4M6.5 12.5h7"/><path d="M6 6L3 8l3 2z" fill="{c}"/>',
    "image": '<rect x="2" y="3" width="12" height="10" rx="1"/><circle cx="5.5" cy="6.5" r="1.2"/><path d="M2 11.5l3.5-3 2.5 2.2 3-3.2 3 3"/>',
    "shape": '<rect x="2" y="2.5" width="6" height="6" rx="1"/><circle cx="11" cy="10.5" r="3.5"/>',
    "diagram": '<rect x="1.5" y="2" width="5" height="3" rx="0.5"/><rect x="9.5" y="2" width="5" height="3" rx="0.5"/><rect x="5.5" y="11" width="5" height="3" rx="0.5"/><path d="M6.5 3.5h3M4 5v3.5h4M12 5v3.5H10.5"/>',
    "table": '<rect x="2" y="3" width="12" height="10" rx="0.5"/><path d="M2 6.3h12M2 9.6h12M7.3 3v10"/>',
    "link": '<path d="M6.5 9.5a3 3 0 0 0 4.5.4l2-2a3 3 0 0 0-4.2-4.2l-1 1"/><path d="M9.5 6.5a3 3 0 0 0-4.5-.4l-2 2a3 3 0 0 0 4.2 4.2l1-1"/>',
    "break": '<path d="M2.5 4.5h11M2.5 7.5h11"/><path d="M2.5 12h4l2-2 2 2h3"/>',
    "rule": '<path d="M2.5 8h11M2.5 5h11M2.5 11h11"/>',
    "date": '<rect x="2" y="3" width="12" height="11" rx="1"/><path d="M2 6.5h12M5.5 1.5v3M10.5 1.5v3"/><path d="M5.5 9.5h2M8.5 9.5h2M5.5 11.8h2"/>',
    "field": '<path d="M4 3.5c-1 0-1.5.8-1.5 1.8v1.4c0 .8-.4 1.3-1 1.3.6 0 1 .5 1 1.3v1.4c0 1 .5 1.8 1.5 1.8M12 3.5c1 0 1.5.8 1.5 1.8v1.4c0 .8.4 1.3 1 1.3-.6 0-1 .5-1 1.3v1.4c0 1-.5 1.8-1.5 1.8"/><path d="M6 8h4"/>',
    "pagesetup": '<rect x="3" y="1.5" width="10" height="13" rx="1"/><path d="M5.5 4.5h5M5.5 7h5M5.5 9.5h3"/>',
    "landscape": '<rect x="1.5" y="4" width="13" height="8" rx="1"/><path d="M4 6.5h8M4 9.5h5"/>',
    "portrait": '<rect x="4" y="1.5" width="8" height="13" rx="1"/><path d="M6.5 4h3M6.5 6.5h3M6.5 9h2"/>',
    "margins": '<rect x="2.5" y="2" width="11" height="12" rx="0.5"/><rect x="5" y="4.5" width="6" height="7" stroke-dasharray="1.5 1"/>',
    "toc": '<path d="M3 3.5h10M3 7h7M3 10.5h8.5M3 14h5"/><circle cx="12.5" cy="10.5" r="0.6" fill="{c}"/><circle cx="13.5" cy="10.5" r="0.6" fill="{c}"/>',
    "heading": '<path d="M3 3v10M11 3v10M3 8h8"/><path d="M13.5 10.5h1.5l-1.5 1.5h1.5"/>',
    "comment": '<path d="M2.5 3h11a1 1 0 0 1 1 1v6a1 1 0 0 1-1 1H7l-3 3v-3h-1.5a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"/>',
    "track": '<path d="M11.5 2.5l2 2L7 11l-2.5.5L5 9z"/><path d="M10 4l2 2"/><path d="M3 13.5h10"/>',
    "spell": '<path d="M3 11.5L7 3l4 8.5M4.5 8.5h5"/><path d="M10 12.5l2 2 3.5-4"/>',
    "find": '<circle cx="7" cy="7" r="4.5"/><path d="M10.5 10.5L14.5 14.5"/>',
    "nav": '<path d="M3 4h10M3 8h6M3 12h8"/><circle cx="12.5" cy="8" r="0.8" fill="{c}"/>',
    "history": '<path d="M8 3a5 5 0 1 1-4.9 6"/><path d="M3 4v3h3"/><path d="M8 6v3l2.5 1.5"/>',
    "author": '<circle cx="8" cy="5" r="2.5"/><path d="M3 14c.5-3 2.5-4.5 5-4.5s4.5 1.5 5 4.5"/>',
    "merge": '<path d="M3 2.5h10M3 7h10M3 11.5h10"/><path d="M8 7v4.5M6 10l2 2 2-2"/>',
    "about": '<circle cx="8" cy="8" r="6.5"/><path d="M8 7.5v4"/><circle cx="8" cy="5" r="0.7" fill="{c}"/>',
    "zoom-in": '<circle cx="7" cy="7" r="4.5"/><path d="M10.5 10.5L14.5 14.5M7 5v4M5 7h4"/>',
    "zoom-out": '<circle cx="7" cy="7" r="4.5"/><path d="M10.5 10.5L14.5 14.5M5 7h4"/>',
    "zoom-fit": '<path d="M2 6V3a1 1 0 0 1 1-1h3M10 2h3a1 1 0 0 1 1 1v3M14 10v3a1 1 0 0 1-1 1h-3M6 14H3a1 1 0 0 1-1-1v-3"/>',
    "focus": '<circle cx="8" cy="8" r="3"/><path d="M8 1.5v3M8 11.5v3M1.5 8h3M11.5 8h3"/>',
    "columns": '<rect x="2" y="3" width="12" height="10" rx="0.5"/><path d="M8 3v10M5 5.5v5M11 5.5v5"/>',
    "clear": '<path d="M4 11.5L10.5 5M13 3.5L6.5 10"/><path d="M3 13.5h10"/><path d="M10.5 5l2.5 2.5L8 12.5H5.5v-2.5z"/>',
    "add-row": '<rect x="2" y="3" width="12" height="10" rx="0.5"/><path d="M2 8h12M8 10v3M6.5 11.5h3"/>',
    "add-col": '<rect x="2" y="3" width="12" height="10" rx="0.5"/><path d="M8 3v10M11 6h3M12.5 4.5v3"/>',
    "del-row": '<rect x="2" y="3" width="12" height="10" rx="0.5"/><path d="M2 8h12M6.5 11.5h3"/>',
    "del-col": '<rect x="2" y="3" width="12" height="10" rx="0.5"/><path d="M8 3v10M11 6h3"/>',
    "merge-cells": '<rect x="2" y="3" width="12" height="10" rx="0.5"/><path d="M2 8h12M8 8V3"/><path d="M6 6L8 4l2 2M6 10l2 2 2-2"/>',
    "review": '<path d="M3 2.5h8l2 2v9a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1v-10a1 1 0 0 1 1-1z"/><path d="M5 8l2 2 4-4"/>',
    "foliomark": '<path d="M4 13V3h8v2.5H6.5V8h5v2.5H6.5V13z" fill="{c}" stroke="none"/>',
    "toc-refresh": '<path d="M13 8a5 5 0 1 1-1.5-3.5"/><path d="M13 2.5V5h-2.5"/>',
}


@lru_cache(maxsize=256)
def _render(name: str, color: str, size: int) -> QPixmap:
    path = _PATHS.get(name)
    if path is None:
        path = _PATHS["about"]
    svg = (_SVG_OPEN + path + "</svg>").replace("{c}", color)
    renderer = QSvgRenderer(QByteArray(svg.encode("utf-8")))
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    return pixmap


def icon(name: str, color: str = "#33415c", size: int = 16) -> QIcon:
    return QIcon(_render(name, color, size))


def mark(size: int = 20) -> QPixmap:
    return _render("foliomark", "#ffffff", size)
