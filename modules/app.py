"""Main application window."""

import os
import math
import sys
from datetime import datetime
from typing import List

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QTableWidgetItem, QFileDialog, QLabel,
    QSpinBox, QDoubleSpinBox, QComboBox, QCheckBox, QMessageBox, QMenu, QFrame,
    QStatusBar, QMenuBar, QButtonGroup, QAbstractItemView, QSizePolicy,
    QRadioButton, QGroupBox, QTabWidget, QHeaderView, QStyleFactory,
    QApplication,
)
from PyQt6.QtCore import Qt, QItemSelectionModel
from PyQt6.QtGui import QIcon, QAction, QActionGroup

from modules.config import CONFIG_FILE, load_config, write_config, update_config_value
from modules.utils import FITZ_EXTS, pdf_page_count
from modules.themes import light_palette, dark_palette, STYLESHEET
from modules.widgets import DroppableTable
from modules.dialogs import PageConfigDialog
from modules.printing import populate_printers, compose_nup_pdf, print_pdf_qt


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
        self.config = load_config()
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
        populate_printers(self.printer_combo)
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
        if hdr is None:
            raise RuntimeError("File table has no horizontal header")
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

        self.mode_tabs = QTabWidget()

        # 0 — Scale
        scale_w = QWidget()
        sl = QVBoxLayout(scale_w)
        sl.setContentsMargins(8, 8, 8, 8)
        sl.setSpacing(6)
        self.scale_mode_group = QButtonGroup(self)
        self.scale_mode_group.setExclusive(True)
        self.scale_none_radio = QRadioButton("None")
        self.scale_fit_radio = QRadioButton("Fit to printer margins")
        self.scale_reduce_radio = QRadioButton("Reduce to printer margins")
        self.scale_custom_radio = QRadioButton("Custom scale")
        for btn in (self.scale_none_radio, self.scale_fit_radio,
                    self.scale_reduce_radio, self.scale_custom_radio):
            self.scale_mode_group.addButton(btn)
            sl.addWidget(btn)
        custom_row = QHBoxLayout()
        custom_row.setContentsMargins(20, 0, 0, 0)
        self.scale_spin = QSpinBox()
        self.scale_spin.setRange(1, 1000)
        self.scale_spin.setValue(100)
        self.scale_spin.setSuffix(" %")
        self.scale_spin.setFixedWidth(80)
        self.scale_spin.setEnabled(False)
        custom_row.addWidget(self.scale_spin)
        custom_row.addStretch()
        sl.addLayout(custom_row)
        self.scale_custom_radio.toggled.connect(self.scale_spin.setEnabled)
        self.scale_paper_source_check = QCheckBox("Choose paper source by PDF page size")
        sl.addWidget(self.scale_paper_source_check)
        sl.addStretch()
        self.scale_fit_radio.setChecked(True)
        self.mode_tabs.addTab(scale_w, "Scale")

        # 1 — Tile Large Pages
        tile_w = QWidget()
        tl = QGridLayout(tile_w)
        tl.setContentsMargins(8, 8, 8, 8)
        tl.setHorizontalSpacing(8)
        tl.setVerticalSpacing(6)

        tl.addWidget(QLabel("Page Zoom:"), 0, 0)
        zoom_row = QHBoxLayout()
        self.tile_zoom_spin = QSpinBox()
        self.tile_zoom_spin.setRange(1, 1000)
        self.tile_zoom_spin.setValue(100)
        self.tile_zoom_spin.setFixedWidth(80)
        zoom_row.addWidget(self.tile_zoom_spin)
        zoom_row.addWidget(QLabel("%"))
        zoom_row.addStretch()
        tl.addLayout(zoom_row, 0, 1)

        tl.addWidget(QLabel("Overlap:"), 1, 0)
        overlap_row = QHBoxLayout()
        self.tile_overlap_spin = QDoubleSpinBox()
        self.tile_overlap_spin.setRange(0.0, 10.0)
        self.tile_overlap_spin.setDecimals(3)
        self.tile_overlap_spin.setValue(0.005)
        self.tile_overlap_spin.setFixedWidth(80)
        overlap_row.addWidget(self.tile_overlap_spin)
        overlap_row.addWidget(QLabel("inch"))
        overlap_row.addStretch()
        tl.addLayout(overlap_row, 1, 1)

        self.tile_cut_marks_check = QCheckBox("Cut Marks")
        tl.addWidget(self.tile_cut_marks_check, 2, 0, 1, 2)

        self.tile_labels_check = QCheckBox("Labels")
        tl.addWidget(self.tile_labels_check, 3, 0, 1, 2)

        tl.setRowStretch(4, 1)
        self.mode_tabs.addTab(tile_w, "Tile Large Pages")

        # 2 — Multiple Pages Per Sheet
        nup_w = QWidget()
        gl = QGridLayout(nup_w)
        gl.setContentsMargins(8, 8, 8, 8)
        gl.setHorizontalSpacing(8)
        gl.setVerticalSpacing(6)

        gl.addWidget(QLabel("Pages per sheet:"), 0, 0)
        pps_row = QHBoxLayout()
        self.pages_per_sheet_combo = QComboBox()
        self.pages_per_sheet_combo.addItems(["2", "4", "6", "9"])
        self.pages_per_sheet_combo.setCurrentText(str(self.config.get("pages_per_sheet", 2)))
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
        info.setToolTip(
            "Refers to the blank space surrounding the page.\n"
            "Setting the margins to adjust the distance between\n"
            "two columns or rows."
        )
        info.setStyleSheet("color: #0066cc; font-size: 13px;")
        margin_row.addWidget(info)
        margin_row.addStretch()
        gl.addLayout(margin_row, 2, 0, 1, 2)

        self.print_page_border_check = QCheckBox("Print Page Border")
        self.print_page_border_check.setChecked(self.config.get("print_page_border", True))
        gl.addWidget(self.print_page_border_check, 3, 0, 1, 2)

        self.mode_tabs.addTab(nup_w, "Multiple Pages Per Sheet")

        # 3 — Booklet
        booklet_w = QWidget()
        bl = QHBoxLayout(booklet_w)
        bl.setContentsMargins(8, 8, 8, 8)
        bl.addWidget(QLabel("Booklet subset:"))
        self.booklet_combo = QComboBox()
        self.booklet_combo.addItems(["Both sides", "Front side only", "Back side only"])
        bl.addWidget(self.booklet_combo)
        bl.addStretch()
        self.mode_tabs.addTab(booklet_w, "Booklet")

        self.mode_tabs.setCurrentIndex(self.config.get("handling_tab", 0))
        layout.addWidget(self.mode_tabs)

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
        self.mode_tabs.setCurrentIndex(idx)

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
            app.setPalette(light_palette())
            app.setStyleSheet(STYLESHEET)
        elif name == "Fusion Dark":
            app.setStyle("Fusion")
            app.setPalette(dark_palette())
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
        update_config_value("theme", name)

    # ── Dialogs ────────────────────────────────────────────────────────────────

    def page_configuration_options(self):
        rows = {item.row() for item in self.file_table.selectedItems()}
        if rows:
            entry = self.file_entries[min(rows)]
        elif self.file_entries:
            entry = self.file_entries[0]
        else:
            QMessageBox.information(self, "Page Configuration", "Add files to the queue first.")
            return
        nup_active = self.mode_tabs.currentIndex() == 2
        nup_settings = {
            "pages_per_sheet": int(self.pages_per_sheet_combo.currentText()) if nup_active else 1,
            "page_order": self.page_order_combo.currentText(),
            "orientation": self.orientation_combo.currentText(),
        }
        PageConfigDialog(entry, nup_settings, self).exec()

    def page_setting(self):
        QMessageBox.information(self, "Page Setting", "Page setting dialog is not yet implemented.")

    def printer_properties(self):
        if sys.platform != "win32":
            QMessageBox.information(self, "Properties", "Printer properties are only available on Windows.")
            return
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
                pages = pdf_page_count(path)
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

    def _reselect_rows(self, rows: list[int]):
        sel = self.file_table.selectionModel()
        model = self.file_table.model()
        if sel is None or model is None:
            return
        for r in rows:
            sel.select(
                model.index(r, 0),
                QItemSelectionModel.SelectionFlag.Select | QItemSelectionModel.SelectionFlag.Rows,
            )

    def move_up(self):
        rows = sorted({item.row() for item in self.file_table.selectedItems()})
        if not rows or rows[0] == 0:
            return
        for r in rows:
            self.file_entries[r], self.file_entries[r - 1] = \
                self.file_entries[r - 1], self.file_entries[r]
        self._refresh_table()
        self._reselect_rows([r - 1 for r in rows])

    def move_down(self):
        rows = sorted({item.row() for item in self.file_table.selectedItems()})
        if not rows or rows[-1] == len(self.file_entries) - 1:
            return
        for r in reversed(rows):
            self.file_entries[r], self.file_entries[r + 1] = \
                self.file_entries[r + 1], self.file_entries[r]
        self._refresh_table()
        self._reselect_rows([r + 1 for r in rows])

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

    @staticmethod
    def _open_path(path: str) -> None:
        """Open a file or folder with the default application, cross-platform."""
        import subprocess
        platform = sys.platform
        try:
            if platform == "win32":
                os.startfile(path)  # type: ignore[attr-defined]
            elif platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            QMessageBox.warning(None, "Open Failed", f"Could not open:\n{path}\n\n{e}")

    def open_selected_file(self):
        for row in {item.row() for item in self.file_table.selectedItems()}:
            path = self.file_entries[row]["path"]
            if os.path.exists(path):
                self._open_path(path)

    def open_containing_folder(self):
        for row in {item.row() for item in self.file_table.selectedItems()}:
            folder = os.path.dirname(self.file_entries[row]["path"])
            if os.path.exists(folder):
                self._open_path(folder)

    # ── Printing ───────────────────────────────────────────────────────────────

    def print_files(self):
        if not self.file_entries:
            QMessageBox.warning(self, "Warning", "No files in the print queue.")
            return

        self.save_config(silent=True)
        printer_name = self.printer_combo.currentText()

        fitz_entries = [
            e for e in self.file_entries
            if os.path.splitext(e["path"])[1].lower() in FITZ_EXTS
        ]
        other_entries = [
            e for e in self.file_entries
            if os.path.splitext(e["path"])[1].lower() not in FITZ_EXTS
        ]

        errors: list[str] = []

        if fitz_entries:
            try:
                nup = max(1, int(self.pages_per_sheet_combo.currentText())) \
                    if self.mode_tabs.currentIndex() == 2 else 1
                order = self.page_order_combo.currentText()
                margin_pts = self.margins_spin.value() * 72 if self.margins_check.isChecked() else 0.0
                draw_border = self.print_page_border_check.isChecked() and nup > 1

                tmp = compose_nup_pdf(fitz_entries, nup, order, margin_pts, draw_border)
                print_pdf_qt(
                    tmp,
                    printer_name,
                    copies=self.copies_spin.value(),
                    grayscale=self.grayscale_check.isChecked(),
                    duplex=self.duplex_check.isChecked(),
                    flip_long=self.flip_long_radio.isChecked(),
                )
            except Exception as e:
                errors.append(f"Rendered print failed: {e}")

        if other_entries:
            platform = sys.platform
            if platform != "win32":
                errors.append(
                    f"Cannot print {len(other_entries)} non-PDF file(s): "
                    "ShellExecute is only available on Windows."
                )
            else:
                try:
                    import win32api
                    safe_name = printer_name.replace('"', '')
                    for entry in other_entries:
                        if os.path.exists(entry["path"]):
                            win32api.ShellExecute(
                                0, "printto", entry["path"],
                                f'"{safe_name}"', ".", 0
                            )
                except Exception as e:
                    errors.append(f"ShellExecute print failed: {e}")

        if errors:
            QMessageBox.warning(self, "Print Errors", "\n".join(errors))
        else:
            self.status_bar.showMessage(
                f"Sent {len(self.file_entries)} file(s) to {printer_name}"
            )

    # ── Config ─────────────────────────────────────────────────────────────────

    def save_config(self, silent: bool = False):
        try:
            data = {
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
                "handling_tab": self.mode_tabs.currentIndex(),
                "pages_per_sheet": int(self.pages_per_sheet_combo.currentText()),
                "page_order": self.page_order_combo.currentText(),
                "margins": self.margins_spin.value(),
                "margins_enabled": self.margins_check.isChecked(),
                "print_page_border": self.print_page_border_check.isChecked(),
                "theme": self.config.get("theme", "Fusion Light"),
            }
            write_config(data)
            if not silent:
                self.status_bar.showMessage("Settings saved")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save settings:\n{e}")
