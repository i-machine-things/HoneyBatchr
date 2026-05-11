"""Print execution: compose N-up PDFs, drive QPrinter, enumerate printers."""

import os
import sys
import math
import tempfile


class PrintCanceledError(Exception):
    """Raised when the user cancels a print job."""

from PyQt6.QtWidgets import QComboBox

from modules.utils import parse_page_range, slot_to_grid

# Paper sizes in inches (portrait: width x height)
PAPER_SIZES: dict[str, tuple[float, float]] = {
    "Letter":  (8.5,   11.0),
    "Legal":   (8.5,   14.0),
    "Tabloid": (11.0,  17.0),
    "A3":      (11.69, 16.54),
    "A4":      (8.27,  11.69),
    "A5":      (5.83,   8.27),
}


def populate_printers(combo: QComboBox) -> None:
    """Fill *combo* with available printers, selecting the system default."""
    if sys.platform == "win32":
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
            return
        except Exception:
            pass

    # Cross-platform: use Qt printer enumeration (works on Linux via CUPS, macOS via CUPS)
    from PyQt6.QtPrintSupport import QPrinterInfo
    default = QPrinterInfo.defaultPrinter()
    for p in QPrinterInfo.availablePrinters():
        combo.addItem(p.printerName())
    if not default.isNull():
        idx = combo.findText(default.printerName())
        if idx >= 0:
            combo.setCurrentIndex(idx)
    if combo.count() == 0:
        combo.addItem("Default Printer")



def compose_nup_pdf(
    entries: list,
    nup: int,
    order: str,
    margin_pts: float,
    draw_border: bool,
    orientation: str = "Portrait",
    auto_rotate: bool = True,
    auto_center: bool = True,
    paper_size: str = "Letter",
    page_margin_left: float = 0.0,
    page_margin_right: float = 0.0,
    page_margin_top: float = 0.0,
    page_margin_bottom: float = 0.0,
) -> str:
    """Render all entries into a single N-up PDF and return its temp path."""
    import fitz  # type: ignore[import]

    cols = max(1, math.ceil(math.sqrt(nup)))
    rows = max(1, math.ceil(nup / cols))

    # Pre-compute filtered indices for every entry (needed for Auto detection and rendering)
    def _filtered_indices(entry: dict, page_count: int) -> list:
        idx = list(range(page_count))
        rng = entry.get("print_range", "All")
        if rng and rng != "All":
            idx = parse_page_range(rng, page_count)
        subset = entry.get("page_subset", "All pages in range")
        if subset == "Odd pages only":
            idx = [i for i in idx if i % 2 == 0]
        elif subset == "Even pages only":
            idx = [i for i in idx if i % 2 == 1]
        if entry.get("reverse_pages", False):
            idx.reverse()
        return idx * max(1, entry.get("copies_override", 1))

    # Determine sheet orientation
    if orientation == "Auto":
        # Compute filtered indices and count orientations in a single pass per file
        landscape = 0
        portrait = 0
        for entry in entries:
            try:
                s = fitz.open(entry["path"])
                for i in _filtered_indices(entry, len(s)):
                    r = s[i].rect
                    if r.width > r.height:
                        landscape += 1
                    else:
                        portrait += 1
                s.close()
            except (OSError, RuntimeError):
                pass
        use_landscape = landscape > portrait
    else:
        use_landscape = (orientation == "Landscape")

    base_w, base_h = PAPER_SIZES.get(paper_size, PAPER_SIZES["Letter"])
    if use_landscape:
        pw, ph = base_h * 72, base_w * 72
    else:
        pw, ph = base_w * 72, base_h * 72

    # Printable area after page margins
    for _margin_name, _margin_val in (
        ("page_margin_left", page_margin_left),
        ("page_margin_right", page_margin_right),
        ("page_margin_top", page_margin_top),
        ("page_margin_bottom", page_margin_bottom),
    ):
        if _margin_val < 0:
            raise ValueError(f"{_margin_name} cannot be negative (got {_margin_val}).")

    ml = page_margin_left * 72
    mr = page_margin_right * 72
    mt = page_margin_top * 72
    mb = page_margin_bottom * 72
    printable_w = pw - ml - mr
    printable_h = ph - mt - mb

    if printable_w <= 0 or printable_h <= 0:
        raise ValueError(
            f"Page margins ({page_margin_left + page_margin_right:.3f} in wide, "
            f"{page_margin_top + page_margin_bottom:.3f} in tall) exceed the "
            f"{paper_size} paper dimensions."
        )

    net_w = printable_w - margin_pts * (cols + 1)
    net_h = printable_h - margin_pts * (rows + 1)
    if net_w <= 0 or net_h <= 0:
        raise ValueError(
            f"Cell gap ({margin_pts / 72:.3f} in) is too large for "
            f"{cols}\u00d7{rows} cells on {paper_size} paper."
        )

    cell_w = net_w / cols
    cell_h = net_h / rows

    out = fitz.open()

    for entry in entries:
        try:
            src = fitz.open(entry["path"])
        except (OSError, RuntimeError):
            continue

        indices = _filtered_indices(entry, len(src))

        if not indices:
            src.close()
            continue

        for sheet_start in range(0, len(indices), nup):
            slot_indices = indices[sheet_start:sheet_start + nup]
            page = out.new_page(width=pw, height=ph)
            for slot, idx in enumerate(slot_indices):
                col_i, row_i = slot_to_grid(slot, cols, rows, order)
                x0 = ml + margin_pts + col_i * (cell_w + margin_pts)
                y0 = mt + margin_pts + row_i * (cell_h + margin_pts)
                cell_rect = fitz.Rect(x0, y0, x0 + cell_w, y0 + cell_h)

                # Auto-rotate: if page and cell have mismatched orientations, rotate 90°
                rotate = 0
                if auto_rotate:
                    src_rect = src[idx].rect
                    page_is_landscape = src_rect.width > src_rect.height
                    cell_is_landscape = cell_w > cell_h
                    if page_is_landscape != cell_is_landscape:
                        rotate = 90

                # keep_proportion=auto_center: True centres with whitespace padding;
                # False stretches to fill the full cell rect
                page.show_pdf_page(
                    cell_rect, src, idx,
                    rotate=rotate,
                    keep_proportion=auto_center,
                )

                if draw_border:
                    page.draw_rect(
                        fitz.Rect(x0, y0, x0 + cell_w, y0 + cell_h),
                        color=(0.7, 0.7, 0.7),
                        width=0.5,
                    )

        src.close()

    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp_path = tmp.name
    tmp.close()
    try:
        out.save(tmp_path)
    except Exception:
        out.close()
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise
    out.close()
    return tmp_path


def split_for_manual_duplex(pdf_path: str, reverse_back: bool) -> tuple[str, str | None]:
    """Split a composed PDF into front and back temp files for two-pass manual duplex.

    Front pass: sheets at positions 0, 2, 4, ...
    Back pass:  sheets at positions 1, 3, 5, ... (reversed when *reverse_back* is True,
                which is correct for face-down output trays so the sheets realign).

    When *reverse_back* is True and the page count is odd, the last physical sheet
    (front-side only) feeds through the printer first on the back pass.  A blank page
    is prepended to the back-side PDF so that sheet passes through without receiving
    any image, keeping every sheet's back page in the correct position.

    Returns (front_path, back_path).  back_path is None when the source has only one
    sheet.  Caller owns both files and must unlink them after printing.
    """
    import fitz

    doc = fitz.open(pdf_path)
    n = len(doc)
    doc.close()

    front_indices = list(range(0, n, 2))
    back_indices  = list(range(1, n, 2))

    # With face-down reload (reverse_back=True), the last-printed sheet feeds first.
    # When the page count is odd, that sheet has no back side and needs a blank page
    # as the first entry in the back-side PDF to absorb it.
    needs_leading_blank = reverse_back and len(front_indices) > len(back_indices)

    if reverse_back:
        back_indices = list(reversed(back_indices))

    def _subset(indices: list, leading_blank: bool = False) -> str:
        sub = fitz.open(pdf_path)
        sub.select(indices)
        if leading_blank:
            ref = sub[0].rect
            sub.insert_page(0, width=ref.width, height=ref.height)
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        path = tmp.name
        tmp.close()
        try:
            sub.save(path)
        except Exception:
            sub.close()
            try:
                os.unlink(path)
            except OSError:
                pass
            raise
        sub.close()
        return path

    front_path = _subset(front_indices)
    back_path = None
    if back_indices:
        try:
            back_path = _subset(back_indices, leading_blank=needs_leading_blank)
        except Exception:
            try:
                os.unlink(front_path)
            except OSError:
                pass
            raise
    return front_path, back_path


def print_pdf_qt(
    pdf_path: str,
    printer_name: str,
    copies: int,
    grayscale: bool,
    duplex: bool,
    flip_long: bool,
    paper_size: str = "Letter",
    cancel_check=None,
    page_progress=None,
) -> None:
    """Print a composed PDF via QPrinter/QPainter."""
    import fitz  # type: ignore[import]
    from PyQt6.QtPrintSupport import QPrinter, QPrinterInfo
    from PyQt6.QtGui import QImage, QPainter, QPageLayout, QPageSize
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

    # Apply selected paper size to the print driver
    ps_id = getattr(QPageSize.PageSizeId, paper_size, QPageSize.PageSizeId.Letter)
    printer.setPageSize(QPageSize(ps_id))

    # Match QPrinter orientation to the composed PDF's page dimensions
    try:
        probe = fitz.open(pdf_path)
        if len(probe) > 0:
            r = probe[0].rect
            if r.width > r.height:
                printer.setPageOrientation(QPageLayout.Orientation.Landscape)
            else:
                printer.setPageOrientation(QPageLayout.Orientation.Portrait)
        probe.close()
    except (OSError, RuntimeError):
        pass

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

        total_pages = len(doc)
        for i in range(total_pages):
            if cancel_check and cancel_check():
                printer.abort()
                doc.close()
                raise PrintCanceledError()
            if page_progress:
                page_progress(i + 1, total_pages)
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
            # Preserve aspect ratio: centre image on the page
            scale = min(dw / img.width(), dh / img.height())
            iw = int(img.width() * scale)
            ih = int(img.height() * scale)
            x = (dw - iw) // 2
            y = (dh - ih) // 2
            painter.drawImage(QRect(x, y, iw, ih), img)
        doc.close()
    finally:
        painter.end()
        try:
            os.unlink(pdf_path)
        except OSError:
            pass
