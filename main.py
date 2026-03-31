"""
Honey Batchr - Batch Printing Application
PyQt6 application matching Adobe Acrobat Batch Print layout
"""

import sys
import os
import json
import math
from pathlib import Path
from datetime import datetime
from typing import List

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QTableWidget, QTableWidgetItem, QFileDialog, QLabel,
    QSpinBox, QDoubleSpinBox, QComboBox, QCheckBox, QMessageBox, QMenu, QFrame,
    QStatusBar, QMenuBar, QButtonGroup, QAbstractItemView, QSizePolicy,
    QRadioButton, QGroupBox, QStackedWidget, QHeaderView, QStyleFactory,
    QDialog, QSlider, QLineEdit
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import (
    QIcon, QAction, QDragEnterEvent, QDropEvent, QActionGroup,
    QPalette, QColor, QPixmap, QImage, QPainter
)

CONFIG_FILE = Path.home() / ".honeybatchr" / "config.json"

def _parse_page_range(rng: str, total: int) -> list[int]:
    """Convert '1,5-9,12' to 0-based page indices, clamped to [0, total)."""
    result: list[int] = []
    for part in rng.split(","):
        part = part.strip()
        if "-" in part:
            try:
                a, b = part.split("-", 1)
                result.extend(range(max(0, int(a) - 1), min(total, int(b))))
            except ValueError:
                pass
        else:
            try:
                n = int(part)
                if 1 <= n <= total:
                    result.append(n - 1)
            except ValueError:
                pass
    return result if result else list(range(total))


def _slot_to_grid(slot: int, cols: int, rows: int, order: str) -> tuple[int, int]:
    """Map a sheet slot index to (col, row) for the given page order."""
    if order.startswith("Vertical"):
        col_i, row_i = slot // rows, slot % rows
    else:
        col_i, row_i = slot % cols, slot // cols
    if "Reversed" in order:
        col_i, row_i = cols - 1 - col_i, rows - 1 - row_i
    return col_i, row_i


# File types that PyMuPDF can open (compose + print via QPainter)
_FITZ_EXTS = frozenset({
    ".pdf", ".xps", ".epub", ".cbz", ".fb2", ".svg",
    ".jpg", ".jpeg", ".png", ".bmp", ".gif",
    ".tif", ".tiff", ".pnm", ".pgm", ".ppm", ".pbm", ".pam",
})


# ── Palette factories ──────────────────────────────────────────────────────────

def _light_palette() -> QPalette:
    p = QPalette()
    pairs = [
        (QPalette.ColorRole.Window,          QColor(240, 240, 240)),
        (QPalette.ColorRole.WindowText,      QColor(0, 0, 0)),
        (QPalette.ColorRole.Base,            QColor(255, 255, 255)),
        (QPalette.ColorRole.AlternateBase,   QColor(247, 247, 247)),
        (QPalette.ColorRole.Text,            QColor(0, 0, 0)),
        (QPalette.ColorRole.Button,          QColor(240, 240, 240)),
        (QPalette.ColorRole.ButtonText,      QColor(0, 0, 0)),
        (QPalette.ColorRole.Highlight,       QColor(76, 163, 224)),
        (QPalette.ColorRole.HighlightedText, QColor(255, 255, 255)),
        (QPalette.ColorRole.Link,            QColor(0, 0, 204)),
        (QPalette.ColorRole.ToolTipBase,     QColor(255, 255, 220)),
        (QPalette.ColorRole.ToolTipText,     QColor(0, 0, 0)),
    ]
    for role, color in pairs:
        p.setColor(QPalette.ColorGroup.All, role, color)
    disabled = QColor(160, 160, 160)
    for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText):
        p.setColor(QPalette.ColorGroup.Disabled, role, disabled)
    return p


def _dark_palette() -> QPalette:
    p = QPalette()
    pairs = [
        (QPalette.ColorRole.Window,          QColor(45, 45, 45)),
        (QPalette.ColorRole.WindowText,      QColor(220, 220, 220)),
        (QPalette.ColorRole.Base,            QColor(30, 30, 30)),
        (QPalette.ColorRole.AlternateBase,   QColor(55, 55, 55)),
        (QPalette.ColorRole.Text,            QColor(220, 220, 220)),
        (QPalette.ColorRole.Button,          QColor(55, 55, 55)),
        (QPalette.ColorRole.ButtonText,      QColor(220, 220, 220)),
        (QPalette.ColorRole.Highlight,       QColor(42, 130, 218)),
        (QPalette.ColorRole.HighlightedText, QColor(255, 255, 255)),
        (QPalette.ColorRole.Link,            QColor(100, 180, 255)),
        (QPalette.ColorRole.ToolTipBase,     QColor(60, 60, 60)),
        (QPalette.ColorRole.ToolTipText,     QColor(220, 220, 220)),
    ]
    for role, color in pairs:
        p.setColor(QPalette.ColorGroup.All, role, color)
    disabled = QColor(110, 110, 110)
    for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText):
        p.setColor(QPalette.ColorGroup.Disabled, role, disabled)
    return p


# ── Drag-and-drop table ────────────────────────────────────────────────────────

class DroppableTable(QTableWidget):
    def __init__(self, on_drop, parent=None):
        super().__init__(parent)
        self._on_drop = on_drop
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e: QDragEnterEvent | None):  # type: ignore[override]
        if e is None:
            return
        mime = e.mimeData()
        if mime is not None and mime.hasUrls():
            e.acceptProposedAction()
        else:
            super().dragEnterEvent(e)

    def dragMoveEvent(self, e):  # type: ignore[override]
        if e is None:
            return
        mime = e.mimeData()
        if mime is not None and mime.hasUrls():
            e.acceptProposedAction()
        else:
            super().dragMoveEvent(e)

    def dropEvent(self, e: QDropEvent | None):  # type: ignore[override]
        if e is None:
            return
        mime = e.mimeData()
        if mime is not None and mime.hasUrls():
            self._on_drop([u.toLocalFile() for u in mime.urls()])
            e.acceptProposedAction()
        else:
            super().dropEvent(e)


# ── Stylesheet ─────────────────────────────────────────────────────────────────

STYLESHEET = """
QGroupBox {
    border: 1px solid #c0c0c0;
    border-radius: 3px;
    margin-top: 10px;
    padding-top: 6px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 8px;
    padding: 0 4px;
}
QPushButton#modeBtn {
    border: 1px solid #a0a0a0;
    border-radius: 2px;
    padding: 6px 4px;
    background-color: #e8e8e8;
    min-height: 40px;
}
QPushButton#modeBtn:checked {
    background-color: #cce0ff;
    border-color: #5588cc;
}
QPushButton#modeBtn:hover:!checked {
    background-color: #d8d8d8;
}
QTableWidget {
    border: 1px solid #b0b0b0;
    gridline-color: #e0e0e0;
    selection-background-color: #cce0ff;
    alternate-background-color: #f7f7f7;
}
QHeaderView::section {
    background-color: #f0f0f0;
    border: none;
    border-right: 1px solid #d0d0d0;
    border-bottom: 1px solid #d0d0d0;
    padding: 3px 4px;
    font-weight: normal;
}
QFrame#rightPanel {
    background-color: #f8f8f8;
    border: 1px solid #d0d0d0;
}
"""


# ── Page Configuration Dialog ─────────────────────────────────────────────────

class PageConfigDialog(QDialog):
    """Per-file page configuration: preview, print range, copies, duplex."""

    def __init__(self, entry: dict, nup_settings: dict, parent=None):
        super().__init__(parent)
        self.entry = entry
        self._nup = max(1, nup_settings.get("pages_per_sheet", 1))
        self._page_order = nup_settings.get("page_order", "Horizontal")

        self._page_images: list[QImage] = []
        self._page_count = 0
        self._sheet_index = 0
        self._doc_w_inch = 8.5
        self._doc_h_inch = 11.0

        self.setWindowTitle("Page Configuration Options")
        self.resize(660, 540)
        self.setMinimumSize(600, 480)

        self._load_document()
        self._init_ui()
        self._render_current_sheet()

    # ── Document loading ───────────────────────────────────────────────────────

    def _load_document(self):
        path = self.entry.get("path", "")
        ext = os.path.splitext(path)[1].lower()
        try:
            import fitz  # type: ignore[import]  # PyMuPDF
            doc = fitz.open(path)
            self._page_count = len(doc)
            if self._page_count > 0:
                r = doc[0].rect
                self._doc_w_inch = r.width / 72.0
                self._doc_h_inch = r.height / 72.0
            for i in range(self._page_count):
                pix = doc[i].get_pixmap(matrix=fitz.Matrix(72 / 72, 72 / 72))
                img = QImage(
                    bytes(pix.samples), pix.width, pix.height,
                    pix.stride, QImage.Format.Format_RGB888
                )
                self._page_images.append(img.copy())
            doc.close()
            return
        except Exception:
            pass

        # Fallback: image files via Qt directly
        if ext in (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"):
            img = QImage(path)
            if not img.isNull():
                self._page_images = [img]
                self._page_count = 1
                self._doc_w_inch = img.width() / 96.0
                self._doc_h_inch = img.height() / 96.0
                return

        # Last resort: parse page count from stored range string
        rng = self.entry.get("range", "All")
        if isinstance(rng, str) and " - " in rng:
            try:
                self._page_count = int(rng.split(" - ")[1])
            except ValueError:
                self._page_count = 1
        else:
            self._page_count = 1

    def _total_sheets(self) -> int:
        return max(1, math.ceil(max(1, self._page_count) / self._nup))

    # ── UI construction ────────────────────────────────────────────────────────

    def _init_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)
        root.addLayout(self._build_preview_panel(), stretch=0)
        root.addLayout(self._build_options_panel(), stretch=1)

    def _build_preview_panel(self) -> QVBoxLayout:
        col = QVBoxLayout()
        col.setSpacing(4)

        title = QLabel("Preview")
        f = title.font()
        f.setBold(True)
        title.setFont(f)
        col.addWidget(title)

        # Info grid
        info = QGridLayout()
        info.setHorizontalSpacing(8)
        info.setVerticalSpacing(2)
        self._zoom_lbl = QLabel("—")
        self._zoom_lbl.setFixedWidth(90)  # prevent column reflow when text changes
        w_str = f"{self._doc_w_inch:.1f} x {self._doc_h_inch:.1f} inch"
        for row_idx, (lbl, val_widget) in enumerate([
            ("Zoom:",     self._zoom_lbl),
            ("Document:", QLabel(w_str)),
            ("Paper:",    QLabel("8.5 x 11.0 inch")),
        ]):
            info.addWidget(QLabel(lbl), row_idx, 0)
            info.addWidget(val_widget, row_idx, 1)
        col.addLayout(info)

        # Canvas
        self._canvas = QLabel()
        self._canvas.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._canvas.setMinimumSize(280, 300)
        self._canvas.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored
        )  # stop pixmap sizeHint from reflowing the layout
        self._canvas.setStyleSheet(
            "background: white; border: 1px solid #b0b0b0;"
        )
        col.addWidget(self._canvas, stretch=1)

        # Navigation
        nav = QHBoxLayout()
        self._prev_btn = QPushButton("<")
        self._prev_btn.setFixedWidth(28)
        self._prev_btn.clicked.connect(self._prev_sheet)
        nav.addWidget(self._prev_btn)

        self._slider = QSlider(Qt.Orientation.Horizontal)
        self._slider.setMinimum(0)
        self._slider.setMaximum(max(0, self._total_sheets() - 1))
        self._slider.setValue(0)
        self._slider.valueChanged.connect(self._on_slider)
        nav.addWidget(self._slider, stretch=1)

        self._next_btn = QPushButton(">")
        self._next_btn.setFixedWidth(28)
        self._next_btn.clicked.connect(self._next_sheet)
        nav.addWidget(self._next_btn)
        col.addLayout(nav)

        self._page_lbl = QLabel(f"Page 1 of {self._total_sheets()}")
        self._page_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        col.addWidget(self._page_lbl)

        return col

    def _build_options_panel(self) -> QVBoxLayout:
        col = QVBoxLayout()
        col.setSpacing(8)

        # ── Print Range ────────────────────────────────────────────────────
        rg = QGroupBox("Print Range")
        rg_lay = QVBoxLayout(rg)
        rg_lay.setSpacing(4)

        rbg = QButtonGroup(self)
        rbg.setExclusive(True)

        self._r_view = QRadioButton("Current view")
        self._r_view.setEnabled(False)
        rbg.addButton(self._r_view)
        rg_lay.addWidget(self._r_view)

        self._r_page = QRadioButton("Current page")
        self._r_page.setEnabled(False)
        rbg.addButton(self._r_page)
        rg_lay.addWidget(self._r_page)

        self._r_all = QRadioButton("All pages")
        self._r_all.setChecked(True)
        rbg.addButton(self._r_all)
        rg_lay.addWidget(self._r_all)

        pages_row = QHBoxLayout()
        self._r_pages = QRadioButton("Pages:")
        rbg.addButton(self._r_pages)
        pages_row.addWidget(self._r_pages)

        stored_range = self.entry.get("print_range", "")
        default_txt = stored_range if stored_range and stored_range != "All" else "1"
        if self._page_count > 1:
            default_txt = f"1-{self._page_count}"
        self._pages_edit = QLineEdit(default_txt)
        self._pages_edit.setFixedWidth(80)
        self._pages_edit.setEnabled(False)
        pages_row.addWidget(self._pages_edit)
        pages_row.addWidget(QLabel(f"/ {self._page_count}"))
        pages_row.addStretch()
        rg_lay.addLayout(pages_row)

        self._r_pages.toggled.connect(self._pages_edit.setEnabled)

        hint_row = QHBoxLayout()
        hint_row.addSpacing(20)
        hint = QLabel("Sample: 1,5-9,12")
        hint.setStyleSheet("color: gray; font-size: 11px;")
        hint_row.addWidget(hint)
        info_icon = QLabel("ℹ")
        info_icon.setToolTip(
            "Enter page numbers and/or ranges separated by commas.\n"
            "Example: 1,3,5-12"
        )
        info_icon.setStyleSheet("color: #0066cc;")
        hint_row.addWidget(info_icon)
        hint_row.addStretch()
        rg_lay.addLayout(hint_row)

        subset_row = QHBoxLayout()
        subset_row.addWidget(QLabel("Subset:"))
        self._subset_combo = QComboBox()
        self._subset_combo.addItems([
            "All pages in range", "Odd pages only", "Even pages only"
        ])
        stored_subset = self.entry.get("page_subset", "All pages in range")
        idx = self._subset_combo.findText(stored_subset)
        if idx >= 0:
            self._subset_combo.setCurrentIndex(idx)
        subset_row.addWidget(self._subset_combo)
        subset_row.addStretch()
        rg_lay.addLayout(subset_row)

        col.addWidget(rg)

        # ── Print Specifications ────────────────────────────────────────────
        sg = QGroupBox("Print Specifications")
        sg_lay = QVBoxLayout(sg)
        sg_lay.setSpacing(4)

        copies_row = QHBoxLayout()
        copies_row.addWidget(QLabel("Copies:"))
        self._copies_spin = QSpinBox()
        self._copies_spin.setRange(1, 999)
        self._copies_spin.setValue(self.entry.get("copies_override", 1))
        self._copies_spin.setFixedWidth(60)
        copies_row.addWidget(self._copies_spin)
        copies_row.addStretch()
        sg_lay.addLayout(copies_row)

        method_row = QHBoxLayout()
        method_row.addWidget(QLabel("Method:"))
        self._duplex_chk = QCheckBox("Print on both sides of paper")
        self._duplex_chk.setChecked(self.entry.get("duplex_override", False))
        self._duplex_chk.toggled.connect(self._toggle_flip)
        method_row.addWidget(self._duplex_chk)
        method_row.addStretch()
        sg_lay.addLayout(method_row)

        flip_w = QWidget()
        fi = QVBoxLayout(flip_w)
        fi.setContentsMargins(60, 0, 0, 0)
        fi.setSpacing(2)
        self._flip_long = QRadioButton("Flip on long edge")
        self._flip_long.setChecked(True)
        self._flip_short = QRadioButton("Flip on short edge")
        fi.addWidget(self._flip_long)
        fi.addWidget(self._flip_short)
        sg_lay.addWidget(flip_w)
        self._toggle_flip(self._duplex_chk.isChecked())

        self._reverse_chk = QCheckBox("Reverse pages")
        self._reverse_chk.setChecked(self.entry.get("reverse_pages", False))
        sg_lay.addWidget(self._reverse_chk)

        self._attach_chk = QCheckBox("Included all supported attachments")
        self._attach_chk.setEnabled(False)
        sg_lay.addWidget(self._attach_chk)

        col.addWidget(sg)
        col.addStretch()

        # OK / Cancel
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        ok_btn = QPushButton("OK")
        ok_btn.setDefault(True)
        ok_btn.setFixedWidth(80)
        ok_btn.clicked.connect(self._on_ok)
        btn_row.addWidget(ok_btn)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedWidth(80)
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        col.addLayout(btn_row)

        return col

    # ── Preview rendering ──────────────────────────────────────────────────────

    def _toggle_flip(self, checked: bool):
        self._flip_long.setEnabled(checked)
        self._flip_short.setEnabled(checked)

    def _render_current_sheet(self):
        cw = self._canvas.width() - 4
        ch = self._canvas.height() - 4
        if cw < 20 or ch < 20:
            return

        sheet = QPixmap(cw, ch)
        sheet.fill(QColor(255, 255, 255))

        cols = math.ceil(math.sqrt(self._nup))
        rows = math.ceil(self._nup / cols)
        cell_w = cw // cols
        cell_h = ch // rows
        first = self._sheet_index * self._nup
        reversed_order = "Reversed" in self._page_order
        vertical_order = self._page_order.startswith("Vertical")

        painter = QPainter(sheet)
        painter.setPen(QColor(190, 190, 190))

        for slot in range(self._nup):
            s = (self._nup - 1 - slot) if reversed_order else slot
            if vertical_order:
                col_idx, row_idx = s // rows, s % rows
            else:
                col_idx, row_idx = s % cols, s // cols

            x = col_idx * cell_w + 2
            y = row_idx * cell_h + 2
            cw2 = cell_w - 4
            ch2 = cell_h - 4
            page_idx = first + slot

            if 0 <= page_idx < len(self._page_images):
                pix = QPixmap.fromImage(self._page_images[page_idx]).scaled(
                    cw2, ch2,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
                painter.drawPixmap(
                    x + (cw2 - pix.width()) // 2,
                    y + (ch2 - pix.height()) // 2,
                    pix,
                )
            elif page_idx < self._page_count:
                painter.fillRect(x, y, cw2, ch2, QColor(225, 225, 225))
                painter.setPen(QColor(120, 120, 120))
                painter.drawText(
                    x, y, cw2, ch2,
                    Qt.AlignmentFlag.AlignCenter,
                    f"Page\n{page_idx + 1}",
                )
                painter.setPen(QColor(190, 190, 190))

            painter.drawRect(x, y, cw2, ch2)

        painter.end()
        self._canvas.setPixmap(sheet)

        # Zoom calculation
        pw, ph = 8.5, 11.0
        dw, dh = self._doc_w_inch, self._doc_h_inch
        if self._nup == 1:
            zoom = min(pw / dw, ph / dh) * 100
        else:
            zoom = min((pw / cols) / dw, (ph / rows) / dh) * 100
        self._zoom_lbl.setText(f"{zoom:.2f}%")

        total = self._total_sheets()
        self._page_lbl.setText(f"Page {self._sheet_index + 1} of {total}")
        self._prev_btn.setEnabled(self._sheet_index > 0)
        self._next_btn.setEnabled(self._sheet_index < total - 1)

    def resizeEvent(self, a0):  # type: ignore[override]
        super().resizeEvent(a0)
        self._render_current_sheet()

    def showEvent(self, a0):  # type: ignore[override]
        super().showEvent(a0)
        self._render_current_sheet()

    # ── Navigation ─────────────────────────────────────────────────────────────

    def _prev_sheet(self):
        if self._sheet_index > 0:
            self._sheet_index -= 1
            self._slider.setValue(self._sheet_index)
            self._render_current_sheet()

    def _next_sheet(self):
        if self._sheet_index < self._total_sheets() - 1:
            self._sheet_index += 1
            self._slider.setValue(self._sheet_index)
            self._render_current_sheet()

    def _on_slider(self, value: int):
        self._sheet_index = value
        self._render_current_sheet()

    # ── Accept ─────────────────────────────────────────────────────────────────

    def _on_ok(self):
        self.entry["copies_override"] = self._copies_spin.value()
        self.entry["duplex_override"] = self._duplex_chk.isChecked()
        self.entry["reverse_pages"] = self._reverse_chk.isChecked()
        self.entry["print_range"] = (
            self._pages_edit.text() if self._r_pages.isChecked() else "All"
        )
        self.entry["page_subset"] = self._subset_combo.currentText()
        self.accept()


# ── Main window ────────────────────────────────────────────────────────────────

class BatchPrintApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Honey Batchr")
        self.setMinimumSize(820, 580)
        self.resize(1020, 700)

        icon_path = os.path.join(os.path.dirname(__file__), "resources", "badger.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.file_entries: List[dict] = []
        self._theme_actions: List[QAction] = []
        self.config = self.load_config()
        self.init_ui()

    # ── UI construction ────────────────────────────────────────────────────────

    def init_ui(self):
        self._build_menu()
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 6, 8, 8)
        root.setSpacing(4)

        root.addLayout(self._build_name_row())
        root.addLayout(self._build_copies_row())

        mid = QHBoxLayout()
        mid.setSpacing(6)
        left = QVBoxLayout()
        left.setSpacing(4)
        left.addWidget(self._build_files_group())
        left.addWidget(self._build_handling_group())
        mid.addLayout(left, stretch=1)
        mid.addWidget(self._build_right_panel())
        root.addLayout(mid, stretch=1)

        root.addLayout(self._build_bottom_row())

        self.update_button_states()

    def _build_name_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.addWidget(QLabel("Name:"))
        self.printer_combo = QComboBox()
        self.printer_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.populate_printers()
        row.addWidget(self.printer_combo, stretch=1)
        props_btn = QPushButton("Properties")
        props_btn.clicked.connect(self.printer_properties)
        row.addWidget(props_btn)
        adv_btn = QPushButton("Advanced")
        adv_btn.clicked.connect(self.printer_advanced)
        row.addWidget(adv_btn)
        return row

    def _build_copies_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.addWidget(QLabel("Copies:"))
        self.copies_spin = QSpinBox()
        self.copies_spin.setRange(1, 999)
        self.copies_spin.setValue(self.config.get("copies", 1))
        self.copies_spin.setFixedWidth(58)
        self.copies_spin.valueChanged.connect(self._on_copies_changed)
        row.addWidget(self.copies_spin)

        self.collate_check = QCheckBox("Collate")
        self.collate_check.setChecked(self.config.get("collate", True))
        self.collate_check.setEnabled(self.copies_spin.value() > 1)
        row.addWidget(self.collate_check)
        row.addSpacing(12)

        self.grayscale_check = QCheckBox("Print as grayscale")
        self.grayscale_check.setChecked(self.config.get("grayscale", False))
        row.addWidget(self.grayscale_check)
        row.addSpacing(12)

        self.print_as_image_check = QCheckBox("Print as image")
        self.print_as_image_check.setChecked(self.config.get("print_as_image", False))
        row.addWidget(self.print_as_image_check)
        row.addSpacing(12)

        self.bleed_marks_check = QCheckBox("Bleed Marks")
        self.bleed_marks_check.setChecked(self.config.get("bleed_marks", False))
        row.addWidget(self.bleed_marks_check)
        row.addStretch()
        return row

    def _build_files_group(self) -> QGroupBox:
        group = QGroupBox("Files for printing")
        layout = QVBoxLayout(group)
        layout.setSpacing(4)

        # Toolbar
        toolbar = QHBoxLayout()
        add_btn = QPushButton("Add Files...")
        add_btn.clicked.connect(self.add_files)
        toolbar.addWidget(add_btn)
        toolbar.addStretch()
        self.move_up_btn = QPushButton("Move up")
        self.move_up_btn.clicked.connect(self.move_up)
        toolbar.addWidget(self.move_up_btn)
        self.move_down_btn = QPushButton("Move down")
        self.move_down_btn.clicked.connect(self.move_down)
        toolbar.addWidget(self.move_down_btn)
        self.remove_btn = QPushButton("Remove")
        self.remove_btn.clicked.connect(self.remove_selected)
        toolbar.addWidget(self.remove_btn)
        layout.addLayout(toolbar)

        # Table
        self.file_table = DroppableTable(self.add_files_to_list)
        self.file_table.setColumnCount(8)
        self.file_table.setHorizontalHeaderLabels(
            ["Number", "Name", "Modified", "Range", "Copies", "Size", "Location", "State"]
        )
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.file_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.file_table.setAlternatingRowColors(True)
        vh = self.file_table.verticalHeader()
        if vh is not None:
            vh.setVisible(False)
        hdr = self.file_table.horizontalHeader()
        assert hdr is not None
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        hdr.resizeSection(0, 60)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        hdr.resizeSection(2, 120)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        hdr.resizeSection(3, 70)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        hdr.resizeSection(4, 55)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        hdr.resizeSection(5, 70)
        hdr.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(7, QHeaderView.ResizeMode.Fixed)
        hdr.resizeSection(7, 60)
        self.file_table.setMinimumHeight(150)
        self.file_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.file_table.customContextMenuRequested.connect(self.show_table_context_menu)
        self.file_table.itemSelectionChanged.connect(self.update_button_states)
        layout.addWidget(self.file_table)

        # Total
        total_row = QHBoxLayout()
        total_row.addStretch()
        self.total_label = QLabel("Total 0 File(s)")
        total_row.addWidget(self.total_label)
        layout.addLayout(total_row)

        # Bottom options
        opts = QHBoxLayout()
        self.set_range_check = QCheckBox("Set current page as print range for all opened files")
        opts.addWidget(self.set_range_check)
        self.loop_check = QCheckBox("Loop printing in list order")
        self.loop_check.setEnabled(False)
        opts.addWidget(self.loop_check)
        opts.addStretch()
        page_cfg = QLabel('<a href="#">Page Configuration Options...</a>')
        page_cfg.setOpenExternalLinks(False)
        page_cfg.linkActivated.connect(self.page_configuration_options)
        opts.addWidget(page_cfg)
        layout.addLayout(opts)

        return group

    def _build_handling_group(self) -> QGroupBox:
        group = QGroupBox("Print Handling")
        layout = QVBoxLayout(group)
        layout.setSpacing(6)

        # Mode buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(2)
        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)
        mode_labels = ["Scale", "Tile Large\nPages", "Multiple Pages\nPer Sheet", "Booklet"]
        self.mode_btns = []
        for i, label in enumerate(mode_labels):
            btn = QPushButton(label)
            btn.setObjectName("modeBtn")
            btn.setCheckable(True)
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            self.mode_group.addButton(btn, i)
            self.mode_btns.append(btn)
            btn_row.addWidget(btn)
        self.mode_btns[2].setChecked(True)  # Multiple Pages Per Sheet default
        self.mode_group.idClicked.connect(self._on_mode_changed)
        layout.addLayout(btn_row)

        # Stacked pages for each mode
        self.mode_stack = QStackedWidget()

        # 0 — Scale
        scale_w = QWidget()
        sl = QHBoxLayout(scale_w)
        sl.setContentsMargins(4, 4, 4, 4)
        sl.addWidget(QLabel("Scale:"))
        self.scale_spin = QSpinBox()
        self.scale_spin.setRange(1, 1000)
        self.scale_spin.setValue(100)
        self.scale_spin.setSuffix(" %")
        self.scale_spin.setFixedWidth(80)
        sl.addWidget(self.scale_spin)
        sl.addStretch()
        self.mode_stack.addWidget(scale_w)

        # 1 — Tile Large Pages
        tile_w = QWidget()
        tl = QHBoxLayout(tile_w)
        tl.setContentsMargins(4, 4, 4, 4)
        tl.addWidget(QLabel("Tile large pages only (no additional options)"))
        tl.addStretch()
        self.mode_stack.addWidget(tile_w)

        # 2 — Multiple Pages Per Sheet
        nup_w = QWidget()
        gl = QGridLayout(nup_w)
        gl.setContentsMargins(4, 4, 4, 4)
        gl.setHorizontalSpacing(8)
        gl.setVerticalSpacing(6)

        gl.addWidget(QLabel("Pages per sheet:"), 0, 0)
        pps_row = QHBoxLayout()
        self.pages_per_sheet_combo = QComboBox()
        self.pages_per_sheet_combo.addItems(["1", "2", "4", "6", "8", "9", "16"])
        self.pages_per_sheet_combo.setCurrentText(str(self.config.get("pages_per_sheet", 1)))
        self.pages_per_sheet_combo.setFixedWidth(64)
        self.pages_per_sheet_combo.currentTextChanged.connect(self._update_nup_grid)
        pps_row.addWidget(self.pages_per_sheet_combo)
        self.nup_cols_spin = QSpinBox()
        self.nup_cols_spin.setRange(1, 16)
        self.nup_cols_spin.setValue(2)
        self.nup_cols_spin.setFixedWidth(48)
        pps_row.addWidget(self.nup_cols_spin)
        pps_row.addWidget(QLabel("x"))
        self.nup_rows_spin = QSpinBox()
        self.nup_rows_spin.setRange(1, 16)
        self.nup_rows_spin.setValue(2)
        self.nup_rows_spin.setFixedWidth(48)
        pps_row.addWidget(self.nup_rows_spin)
        pps_row.addStretch()
        gl.addLayout(pps_row, 0, 1)

        gl.addWidget(QLabel("Page Order:"), 1, 0)
        self.page_order_combo = QComboBox()
        self.page_order_combo.addItems([
            "Horizontal", "Horizontal Reversed", "Vertical", "Vertical Reversed"
        ])
        self.page_order_combo.setCurrentText(self.config.get("page_order", "Horizontal"))
        self.page_order_combo.setFixedWidth(190)
        gl.addWidget(self.page_order_combo, 1, 1)

        margin_row = QHBoxLayout()
        self.margins_check = QCheckBox("Margins:")
        self.margins_check.setChecked(self.config.get("margins_enabled", True))
        margin_row.addWidget(self.margins_check)
        self.margins_spin = QDoubleSpinBox()
        self.margins_spin.setRange(0.0, 10.0)
        self.margins_spin.setDecimals(3)
        self.margins_spin.setValue(self.config.get("margins", 0.200))
        self.margins_spin.setFixedWidth(78)
        margin_row.addWidget(self.margins_spin)
        margin_row.addWidget(QLabel("inch"))
        info = QLabel("  ⓘ")
        info.setToolTip("Margin added around each page on the sheet")
        info.setStyleSheet("color: #0066cc; font-size: 13px;")
        margin_row.addWidget(info)
        margin_row.addStretch()
        gl.addLayout(margin_row, 2, 0, 1, 2)

        self.print_page_border_check = QCheckBox("Print Page Border")
        self.print_page_border_check.setChecked(self.config.get("print_page_border", True))
        gl.addWidget(self.print_page_border_check, 3, 0, 1, 2)

        self.mode_stack.addWidget(nup_w)

        # 3 — Booklet
        booklet_w = QWidget()
        bl = QHBoxLayout(booklet_w)
        bl.setContentsMargins(4, 4, 4, 4)
        bl.addWidget(QLabel("Booklet subset:"))
        self.booklet_combo = QComboBox()
        self.booklet_combo.addItems(["Both sides", "Front side only", "Back side only"])
        bl.addWidget(self.booklet_combo)
        bl.addStretch()
        self.mode_stack.addWidget(booklet_w)

        self.mode_stack.setCurrentIndex(2)
        layout.addWidget(self.mode_stack)

        return group

    def _build_right_panel(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("rightPanel")
        panel.setFixedWidth(210)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        # Duplex
        self.duplex_check = QCheckBox("Print on both sides of paper")
        self.duplex_check.setChecked(self.config.get("duplex", False))
        self.duplex_check.toggled.connect(self._toggle_duplex)
        layout.addWidget(self.duplex_check)

        flip_indent = QWidget()
        fi_layout = QVBoxLayout(flip_indent)
        fi_layout.setContentsMargins(16, 0, 0, 0)
        fi_layout.setSpacing(2)
        self.flip_long_radio = QRadioButton("Flip on long edge")
        self.flip_long_radio.setChecked(True)
        self.flip_short_radio = QRadioButton("Flip on short edge")
        fi_layout.addWidget(self.flip_long_radio)
        fi_layout.addWidget(self.flip_short_radio)
        layout.addWidget(flip_indent)
        self._toggle_duplex(self.duplex_check.isChecked())

        self.auto_rotate_check = QCheckBox("Auto-Rotate")
        self.auto_rotate_check.setChecked(self.config.get("auto_rotate", True))
        layout.addWidget(self.auto_rotate_check)

        self.auto_center_check = QCheckBox("Auto-Center")
        self.auto_center_check.setChecked(self.config.get("auto_center", True))
        layout.addWidget(self.auto_center_check)

        layout.addWidget(self._make_hsep())

        layout.addWidget(QLabel("Orientation"))
        self.orientation_combo = QComboBox()
        self.orientation_combo.addItems(["Portrait", "Landscape", "Auto"])
        self.orientation_combo.setCurrentText(self.config.get("orientation", "Portrait"))
        layout.addWidget(self.orientation_combo)

        layout.addWidget(self._make_hsep())

        layout.addWidget(QLabel("Print What"))
        self.print_what_combo = QComboBox()
        self.print_what_combo.addItems([
            "Document",
            "Document and markups",
            "Document and stamps",
            "Form fields only",
            "Markups only",
        ])
        self.print_what_combo.setCurrentText(
            self.config.get("print_what", "Document and markups")
        )
        layout.addWidget(self.print_what_combo)

        layout.addWidget(self._make_hsep())

        layout.addWidget(QLabel("Output"))
        self.simulate_overprint_check = QCheckBox("Simulate Overprinting")
        self.simulate_overprint_check.setChecked(self.config.get("simulate_overprint", False))
        layout.addWidget(self.simulate_overprint_check)

        layout.addStretch()
        return panel

    def _build_bottom_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        page_setting_btn = QPushButton("Page Setting")
        page_setting_btn.clicked.connect(self.page_setting)
        row.addWidget(page_setting_btn)
        row.addStretch()
        ok_btn = QPushButton("OK")
        ok_btn.setDefault(True)
        ok_btn.setFixedWidth(80)
        ok_btn.clicked.connect(self.print_files)
        row.addWidget(ok_btn)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedWidth(80)
        cancel_btn.clicked.connect(self.close)
        row.addWidget(cancel_btn)
        return row

    def _make_hsep(self) -> QFrame:
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        return line

    def _build_menu(self):
        menubar = QMenuBar(self)
        self.setMenuBar(menubar)

        file_menu = menubar.addMenu("File")
        if file_menu:
            add_act = QAction("Add Files...", self)
            add_act.setShortcut("Ctrl+O")
            add_act.triggered.connect(self.add_files)
            file_menu.addAction(add_act)
            file_menu.addSeparator()
            save_act = QAction("Save Settings", self)
            save_act.setShortcut("Ctrl+S")
            save_act.triggered.connect(self.save_config)
            file_menu.addAction(save_act)
            file_menu.addSeparator()
            exit_act = QAction("Exit", self)
            exit_act.triggered.connect(self.close)
            file_menu.addAction(exit_act)

        settings_menu = menubar.addMenu("Settings")
        if settings_menu:
            theme_menu = settings_menu.addMenu("Theme")
            if theme_menu:
                theme_group = QActionGroup(self)
                theme_group.setExclusive(True)
                current = self.config.get("theme", "Fusion Light")
                # Fusion Light / Dark first, then any other native Qt styles
                native = [s for s in QStyleFactory.keys() if s != "Fusion"]
                for name in ["Fusion Light", "Fusion Dark"] + native:
                    act = QAction(name, self)
                    act.setCheckable(True)
                    act.setChecked(name == current)
                    act.triggered.connect(
                        lambda _, n=name: self.apply_theme(n)
                    )
                    theme_group.addAction(act)
                    theme_menu.addAction(act)
                    self._theme_actions.append(act)

        help_menu = menubar.addMenu("Help")
        if help_menu:
            about_act = QAction("About Honey Batchr", self)
            about_act.triggered.connect(self._show_about)
            help_menu.addAction(about_act)

    # ── UI helpers ─────────────────────────────────────────────────────────────

    def _toggle_duplex(self, checked: bool):
        self.flip_long_radio.setEnabled(checked)
        self.flip_short_radio.setEnabled(checked)

    def _on_copies_changed(self, value: int):
        self.collate_check.setEnabled(value > 1)

    def _on_mode_changed(self, idx: int):
        self.mode_stack.setCurrentIndex(idx)

    def _update_nup_grid(self, text: str):
        try:
            n = int(text)
            cols = math.ceil(math.sqrt(n))
            rows = math.ceil(n / cols)
            self.nup_cols_spin.setValue(cols)
            self.nup_rows_spin.setValue(rows)
        except ValueError:
            pass

    def update_button_states(self):
        has_sel = bool(self.file_table.selectedItems())
        self.move_up_btn.setEnabled(has_sel)
        self.move_down_btn.setEnabled(has_sel)
        self.remove_btn.setEnabled(has_sel)

    def _show_about(self):
        QMessageBox.information(self, "About", "Honey Batchr\nBatch printing made simple.")

    # ── Theme ──────────────────────────────────────────────────────────────────

    def apply_theme(self, name: str):
        app = QApplication.instance()
        if not isinstance(app, QApplication):
            return
        if name == "Fusion Light":
            app.setStyle("Fusion")
            app.setPalette(_light_palette())
            app.setStyleSheet(STYLESHEET)
        elif name == "Fusion Dark":
            app.setStyle("Fusion")
            app.setPalette(_dark_palette())
            app.setStyleSheet(STYLESHEET)
        else:
            app.setStyle(name)
            style = app.style()
            if style is not None:
                app.setPalette(style.standardPalette())
            app.setStyleSheet("")
        self.config["theme"] = name
        for act in self._theme_actions:
            act.setChecked(act.text() == name)
        # Persist only the theme key without touching other settings
        try:
            CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
            saved: dict = {}
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE) as f:
                    saved = json.load(f)
            saved["theme"] = name
            with open(CONFIG_FILE, "w") as f:
                json.dump(saved, f, indent=2)
        except Exception:
            pass

    def page_configuration_options(self):
        rows = {item.row() for item in self.file_table.selectedItems()}
        if rows:
            entry = self.file_entries[min(rows)]
        elif self.file_entries:
            entry = self.file_entries[0]
        else:
            QMessageBox.information(self, "Page Configuration", "Add files to the queue first.")
            return
        nup_settings = {
            "pages_per_sheet": int(self.pages_per_sheet_combo.currentText()),
            "page_order": self.page_order_combo.currentText(),
            "orientation": self.orientation_combo.currentText(),
        }
        PageConfigDialog(entry, nup_settings, self).exec()

    def page_setting(self):
        QMessageBox.information(self, "Page Setting", "Page setting dialog is not yet implemented.")

    def printer_properties(self):
        try:
            import subprocess
            printer_name = self.printer_combo.currentText()
            subprocess.Popen(
                ["rundll32", "printui.dll,PrintUIEntry", "/e", "/n", printer_name]
            )
        except Exception as e:
            QMessageBox.information(self, "Properties", f"Could not open printer properties:\n{e}")

    def printer_advanced(self):
        QMessageBox.information(self, "Advanced", "Advanced printer settings.")

    # ── File management ────────────────────────────────────────────────────────

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Files to Print", "",
            "All Printable (*.pdf *.doc *.docx *.txt *.xls *.xlsx *.ppt *.pptx "
            "*.jpg *.jpeg *.png *.bmp *.tif *.tiff);;"
            "PDF Files (*.pdf);;"
            "Word Documents (*.doc *.docx);;"
            "Excel Spreadsheets (*.xls *.xlsx);;"
            "PowerPoint (*.ppt *.pptx);;"
            "Images (*.jpg *.jpeg *.png *.bmp *.tif *.tiff);;"
            "Text Files (*.txt);;"
            "All Files (*.*)"
        )
        if files:
            self.add_files_to_list(files)

    @staticmethod
    def _pdf_page_count(path: str) -> int | None:
        """Return PDF page count using pypdf if available, else None."""
        try:
            from pypdf import PdfReader
            return len(PdfReader(path).pages)
        except Exception:
            return None

    def add_files_to_list(self, files: List[str]):
        existing = {e["path"] for e in self.file_entries}
        added = 0
        for path in files:
            if not os.path.isfile(path) or path in existing:
                continue
            stat = os.stat(path)
            kb = stat.st_size / 1024
            size_str = f"{kb:.1f} KB" if kb < 1024 else f"{kb / 1024:.1f} MB"
            pages = None
            if path.lower().endswith(".pdf"):
                pages = self._pdf_page_count(path)
            page_range = f"1 - {pages}" if pages else "All"
            self.file_entries.append({
                "path": path,
                "name": os.path.basename(path),
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M"),
                "size": size_str,
                "location": os.path.dirname(path),
                "range": page_range,
            })
            existing.add(path)
            added += 1
        if added:
            self._refresh_table()

    def _refresh_table(self):
        self.file_table.setRowCount(0)
        copies = str(self.copies_spin.value())
        for i, entry in enumerate(self.file_entries):
            row = self.file_table.rowCount()
            self.file_table.insertRow(row)
            for col, val in enumerate([
                str(i + 1),
                entry["name"],
                entry["modified"],
                entry.get("range", "All"),
                copies,
                entry["size"],
                entry["location"],
                "Waiting for Print",
            ]):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
                self.file_table.setItem(row, col, item)
        n = len(self.file_entries)
        self.total_label.setText(f"Total {n} File(s)")
        self.update_button_states()

    def remove_selected(self):
        rows = sorted({item.row() for item in self.file_table.selectedItems()}, reverse=True)
        for row in rows:
            self.file_entries.pop(row)
        self._refresh_table()

    def move_up(self):
        rows = sorted({item.row() for item in self.file_table.selectedItems()})
        if not rows or rows[0] == 0:
            return
        for r in rows:
            self.file_entries[r], self.file_entries[r - 1] = \
                self.file_entries[r - 1], self.file_entries[r]
        self._refresh_table()
        for r in rows:
            self.file_table.selectRow(r - 1)

    def move_down(self):
        rows = sorted({item.row() for item in self.file_table.selectedItems()}, reverse=True)
        if not rows or rows[-1] == len(self.file_entries) - 1:
            return
        for r in rows:
            self.file_entries[r], self.file_entries[r + 1] = \
                self.file_entries[r + 1], self.file_entries[r]
        self._refresh_table()
        for r in rows:
            self.file_table.selectRow(r + 1)

    def show_table_context_menu(self, position):
        menu = QMenu()
        a1 = menu.addAction("Remove")
        a2 = menu.addAction("Open File")
        a3 = menu.addAction("Open Containing Folder")
        if a1:
            a1.triggered.connect(self.remove_selected)
        if a2:
            a2.triggered.connect(self.open_selected_file)
        if a3:
            a3.triggered.connect(self.open_containing_folder)
        menu.exec(self.file_table.mapToGlobal(position))

    def open_selected_file(self):
        for row in {item.row() for item in self.file_table.selectedItems()}:
            path = self.file_entries[row]["path"]
            if os.path.exists(path):
                os.startfile(path)

    def open_containing_folder(self):
        for row in {item.row() for item in self.file_table.selectedItems()}:
            path = self.file_entries[row]["path"]
            if os.path.exists(path):
                os.startfile(os.path.dirname(path))

    # ── Printing ───────────────────────────────────────────────────────────────

    def populate_printers(self):
        try:
            import win32print
            for p in win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS, None, 2
            ):
                self.printer_combo.addItem(p["pPrinterName"])
            default = win32print.GetDefaultPrinter()
            idx = self.printer_combo.findText(default)
            if idx >= 0:
                self.printer_combo.setCurrentIndex(idx)
        except Exception:
            self.printer_combo.addItem("Default Printer")

    def print_files(self):
        if not self.file_entries:
            QMessageBox.warning(self, "Warning", "No files in the print queue.")
            return

        self.save_config(silent=True)
        printer_name = self.printer_combo.currentText()

        fitz_entries = [
            e for e in self.file_entries
            if os.path.splitext(e["path"])[1].lower() in _FITZ_EXTS
        ]
        other_entries = [
            e for e in self.file_entries
            if os.path.splitext(e["path"])[1].lower() not in _FITZ_EXTS
        ]

        errors: list[str] = []

        if fitz_entries:
            try:
                tmp = self._compose_nup_pdf(fitz_entries)
                self._print_pdf_qt(tmp, printer_name)
            except Exception as e:
                errors.append(f"Rendered print failed: {e}")

        # Office / unsupported formats — best-effort via ShellExecute
        if other_entries:
            try:
                import win32api
                for entry in other_entries:
                    if os.path.exists(entry["path"]):
                        win32api.ShellExecute(
                            0, "printto", entry["path"], f'"{printer_name}"', ".", 0
                        )
            except Exception as e:
                errors.append(f"ShellExecute print failed: {e}")

        if errors:
            QMessageBox.warning(self, "Print Errors", "\n".join(errors))
        else:
            self.status_bar.showMessage(
                f"Sent {len(self.file_entries)} file(s) to {printer_name}"
            )

    # ── Print helpers ──────────────────────────────────────────────────────────

    def _compose_nup_pdf(self, entries: list) -> str:
        """Render all entries into a single N-up PDF and return its temp path."""
        import fitz  # type: ignore[import]
        import tempfile

        nup = max(1, int(self.pages_per_sheet_combo.currentText()))
        cols = math.ceil(math.sqrt(nup))
        rows = math.ceil(nup / cols)
        order = self.page_order_combo.currentText()
        margin = self.margins_spin.value() * 72 if self.margins_check.isChecked() else 0.0
        draw_border = self.print_page_border_check.isChecked() and nup > 1

        # US Letter in points
        pw, ph = 8.5 * 72, 11.0 * 72
        cell_w = (pw - margin * (cols + 1)) / cols
        cell_h = (ph - margin * (rows + 1)) / rows

        out = fitz.open()

        for entry in entries:
            try:
                src = fitz.open(entry["path"])
            except Exception:
                continue

            n = len(src)
            indices = list(range(n))

            rng = entry.get("print_range", "All")
            if rng and rng != "All":
                indices = _parse_page_range(rng, n)

            subset = entry.get("page_subset", "All pages in range")
            if subset == "Odd pages only":
                indices = [i for i in indices if i % 2 == 0]
            elif subset == "Even pages only":
                indices = [i for i in indices if i % 2 == 1]

            if entry.get("reverse_pages", False):
                indices.reverse()

            indices = indices * max(1, entry.get("copies_override", 1))

            for sheet_start in range(0, max(len(indices), 1), nup):
                slot_indices = indices[sheet_start:sheet_start + nup]
                page = out.new_page(width=pw, height=ph)
                for slot, idx in enumerate(slot_indices):
                    col_i, row_i = _slot_to_grid(slot, cols, rows, order)
                    x0 = margin + col_i * (cell_w + margin)
                    y0 = margin + row_i * (cell_h + margin)
                    rect = fitz.Rect(x0, y0, x0 + cell_w, y0 + cell_h)
                    page.show_pdf_page(rect, src, idx)
                    if draw_border:
                        page.draw_rect(rect, color=(0.7, 0.7, 0.7), width=0.5)

            src.close()

        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp_path = tmp.name
        tmp.close()
        out.save(tmp_path)
        out.close()
        return tmp_path

    def _print_pdf_qt(self, pdf_path: str, printer_name: str):
        """Print a composed PDF via QPrinter/QPainter (handles duplex, color, copies)."""
        import fitz  # type: ignore[import]
        from PyQt6.QtPrintSupport import QPrinter, QPrinterInfo
        from PyQt6.QtGui import QImage
        from PyQt6.QtCore import QRect

        # Locate QPrinterInfo by name
        printer_info = next(
            (p for p in QPrinterInfo.availablePrinters()
             if p.printerName() == printer_name),
            QPrinterInfo(),
        )

        printer = QPrinter(printer_info, QPrinter.PrinterMode.HighResolution)
        printer.setFullPage(True)
        printer.setCopyCount(self.copies_spin.value())
        printer.setColorMode(
            QPrinter.ColorMode.GrayScale
            if self.grayscale_check.isChecked()
            else QPrinter.ColorMode.Color
        )
        if self.duplex_check.isChecked():
            printer.setDuplex(
                QPrinter.DuplexMode.DuplexLongSide
                if self.flip_long_radio.isChecked()
                else QPrinter.DuplexMode.DuplexShortSide
            )
        else:
            printer.setDuplex(QPrinter.DuplexMode.DuplexNone)

        painter = QPainter()
        if not painter.begin(printer):
            try:
                os.unlink(pdf_path)
            except OSError:
                pass
            raise RuntimeError(f"Could not open printer '{printer_name}'")

        try:
            doc = fitz.open(pdf_path)
            device = painter.device()
            dw = device.width() if device else 1654
            dh = device.height() if device else 2339
            render_dpi = min(printer.resolution(), 300)

            for i in range(len(doc)):
                page = doc[i]
                if i > 0:
                    printer.newPage()
                mat = fitz.Matrix(render_dpi / 72, render_dpi / 72)
                cs = fitz.csGRAY if self.grayscale_check.isChecked() else fitz.csRGB
                pix = page.get_pixmap(matrix=mat, colorspace=cs)
                fmt = (
                    QImage.Format.Format_Grayscale8
                    if self.grayscale_check.isChecked()
                    else QImage.Format.Format_RGB888
                )
                img = QImage(
                    bytes(pix.samples), pix.width, pix.height, pix.stride, fmt
                )
                painter.drawImage(QRect(0, 0, dw, dh), img)
            doc.close()
        finally:
            painter.end()
            try:
                os.unlink(pdf_path)
            except OSError:
                pass

    # ── Config ─────────────────────────────────────────────────────────────────

    def load_config(self) -> dict:
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE) as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading config: {e}")
        return {
            "copies": 1,
            "collate": True,
            "grayscale": False,
            "print_as_image": False,
            "bleed_marks": False,
            "duplex": False,
            "auto_rotate": True,
            "auto_center": True,
            "orientation": "Portrait",
            "print_what": "Document and markups",
            "simulate_overprint": False,
            "pages_per_sheet": 1,
            "page_order": "Horizontal",
            "margins": 0.200,
            "margins_enabled": True,
            "print_page_border": True,
            "theme": "Fusion Light",
        }

    def save_config(self, silent: bool = False):
        try:
            CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
            config = {
                "copies": self.copies_spin.value(),
                "collate": self.collate_check.isChecked(),
                "grayscale": self.grayscale_check.isChecked(),
                "print_as_image": self.print_as_image_check.isChecked(),
                "bleed_marks": self.bleed_marks_check.isChecked(),
                "duplex": self.duplex_check.isChecked(),
                "auto_rotate": self.auto_rotate_check.isChecked(),
                "auto_center": self.auto_center_check.isChecked(),
                "orientation": self.orientation_combo.currentText(),
                "print_what": self.print_what_combo.currentText(),
                "simulate_overprint": self.simulate_overprint_check.isChecked(),
                "pages_per_sheet": int(self.pages_per_sheet_combo.currentText()),
                "page_order": self.page_order_combo.currentText(),
                "margins": self.margins_spin.value(),
                "margins_enabled": self.margins_check.isChecked(),
                "print_page_border": self.print_page_border_check.isChecked(),
                "theme": self.config.get("theme", "Fusion Light"),
            }
            with open(CONFIG_FILE, "w") as f:
                json.dump(config, f, indent=2)
            if not silent:
                self.status_bar.showMessage("Settings saved")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save settings:\n{e}")


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)

    # Apply saved theme before the window is constructed so it renders correctly
    theme = "Fusion Light"
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE) as f:
                theme = json.load(f).get("theme", "Fusion Light")
    except Exception:
        pass

    if theme == "Fusion Dark":
        app.setStyle("Fusion")
        app.setPalette(_dark_palette())
        app.setStyleSheet(STYLESHEET)
    elif theme == "Fusion Light" or theme not in QStyleFactory.keys():
        app.setStyle("Fusion")
        app.setPalette(_light_palette())
        app.setStyleSheet(STYLESHEET)
    else:
        app.setStyle(theme)

    window = BatchPrintApp()
    if len(sys.argv) > 1:
        window.add_files_to_list(sys.argv[1:])
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
