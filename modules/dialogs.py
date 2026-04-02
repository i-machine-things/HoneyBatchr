"""Application dialogs."""

import os
import math

from PyQt6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QSpinBox, QComboBox, QCheckBox,
    QButtonGroup, QRadioButton, QGroupBox, QSlider, QLineEdit, QSizePolicy,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QImage, QPainter, QPixmap, QColor


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
            try:
                self._page_count = len(doc)
                if self._page_count > 0:
                    r = doc[0].rect
                    self._doc_w_inch = r.width / 72.0
                    self._doc_h_inch = r.height / 72.0
                for i in range(self._page_count):
                    pix = doc[i].get_pixmap(matrix=fitz.Matrix(72 / 72, 72 / 72), alpha=False)
                    img = QImage(
                        bytes(pix.samples), pix.width, pix.height,
                        pix.stride, QImage.Format.Format_RGB888
                    )
                    self._page_images.append(img.copy())
            finally:
                doc.close()
            return
        except Exception as e:
            print(f"fitz load failed for {path!r}: {e}")

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
        self._zoom_lbl.setFixedWidth(90)
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
        )
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
        if stored_range and stored_range != "All":
            default_txt = stored_range
        elif self._page_count > 1:
            default_txt = f"1-{self._page_count}"
        else:
            default_txt = "1"
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
        flip_short_saved = self.entry.get("flip_short_edge", False)
        self._flip_long.setChecked(not flip_short_saved)
        self._flip_short = QRadioButton("Flip on short edge")
        self._flip_short.setChecked(flip_short_saved)
        fi.addWidget(self._flip_long)
        fi.addWidget(self._flip_short)
        sg_lay.addWidget(flip_w)
        self._toggle_flip(self._duplex_chk.isChecked())

        self._reverse_chk = QCheckBox("Reverse pages")
        self._reverse_chk.setChecked(self.entry.get("reverse_pages", False))
        sg_lay.addWidget(self._reverse_chk)

        self._attach_chk = QCheckBox("Include all supported attachments")
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
            self._slider.blockSignals(True)
            self._slider.setValue(self._sheet_index)
            self._slider.blockSignals(False)
            self._render_current_sheet()

    def _next_sheet(self):
        if self._sheet_index < self._total_sheets() - 1:
            self._sheet_index += 1
            self._slider.blockSignals(True)
            self._slider.setValue(self._sheet_index)
            self._slider.blockSignals(False)
            self._render_current_sheet()

    def _on_slider(self, value: int):
        self._sheet_index = value
        self._render_current_sheet()

    # ── Accept ─────────────────────────────────────────────────────────────────

    def _on_ok(self):
        if self._r_pages.isChecked():
            page_range = self._pages_edit.text().strip()
            if not page_range:
                from PyQt6.QtWidgets import QMessageBox
                QMessageBox.warning(self, "Invalid Range", "Please enter a page range.")
                return
        self.entry["copies_override"] = self._copies_spin.value()
        self.entry["duplex_override"] = self._duplex_chk.isChecked()
        self.entry["flip_short_edge"] = self._flip_short.isChecked()
        self.entry["reverse_pages"] = self._reverse_chk.isChecked()
        self.entry["print_range"] = (
            self._pages_edit.text().strip() if self._r_pages.isChecked() else "All"
        )
        self.entry["page_subset"] = self._subset_combo.currentText()
        self.accept()
