"""Print execution: compose N-up PDFs, drive QPrinter, enumerate printers."""

import os
import math
import tempfile

from PyQt6.QtWidgets import QComboBox

from modules.utils import parse_page_range, slot_to_grid


def populate_printers(combo: QComboBox) -> None:
    """Fill *combo* with available printers, selecting the system default."""
    try:
        import win32print
        for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS, None, 2
        ):
            combo.addItem(p["pPrinterName"])
        default = win32print.GetDefaultPrinter()
        idx = combo.findText(default)
        if idx >= 0:
            combo.setCurrentIndex(idx)
    except Exception:
        combo.addItem("Default Printer")


def compose_nup_pdf(
    entries: list,
    nup: int,
    order: str,
    margin_pts: float,
    draw_border: bool,
) -> str:
    """Render all entries into a single N-up PDF and return its temp path."""
    import fitz  # type: ignore[import]

    cols = math.ceil(math.sqrt(nup))
    rows = math.ceil(nup / cols)

    # US Letter in points
    pw, ph = 8.5 * 72, 11.0 * 72
    cell_w = (pw - margin_pts * (cols + 1)) / cols
    cell_h = (ph - margin_pts * (rows + 1)) / rows

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
            indices = parse_page_range(rng, n)

        subset = entry.get("page_subset", "All pages in range")
        if subset == "Odd pages only":
            indices = [i for i in indices if i % 2 == 0]
        elif subset == "Even pages only":
            indices = [i for i in indices if i % 2 == 1]

        if entry.get("reverse_pages", False):
            indices.reverse()

        indices = indices * max(1, entry.get("copies_override", 1))

        if not indices:
            src.close()
            continue

        for sheet_start in range(0, len(indices), nup):
            slot_indices = indices[sheet_start:sheet_start + nup]
            page = out.new_page(width=pw, height=ph)
            for slot, idx in enumerate(slot_indices):
                col_i, row_i = slot_to_grid(slot, cols, rows, order)
                x0 = margin_pts + col_i * (cell_w + margin_pts)
                y0 = margin_pts + row_i * (cell_h + margin_pts)
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


def print_pdf_qt(
    pdf_path: str,
    printer_name: str,
    copies: int,
    grayscale: bool,
    duplex: bool,
    flip_long: bool,
) -> None:
    """Print a composed PDF via QPrinter/QPainter."""
    import fitz  # type: ignore[import]
    from PyQt6.QtPrintSupport import QPrinter, QPrinterInfo
    from PyQt6.QtGui import QImage, QPainter
    from PyQt6.QtCore import QRect

    matches = [p for p in QPrinterInfo.availablePrinters()
               if p.printerName() == printer_name]
    if not matches:
        raise RuntimeError(f"Printer '{printer_name}' not found")
    printer_info = matches[0]

    printer = QPrinter(printer_info, QPrinter.PrinterMode.HighResolution)
    printer.setFullPage(True)
    printer.setCopyCount(copies)
    printer.setColorMode(
        QPrinter.ColorMode.GrayScale if grayscale else QPrinter.ColorMode.Color
    )
    if duplex:
        printer.setDuplex(
            QPrinter.DuplexMode.DuplexLongSide
            if flip_long
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
            cs = fitz.csGRAY if grayscale else fitz.csRGB
            pix = page.get_pixmap(matrix=mat, colorspace=cs)
            fmt = (
                QImage.Format.Format_Grayscale8
                if grayscale
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
