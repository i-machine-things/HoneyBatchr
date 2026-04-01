"""Shared utility functions and constants."""

# File types that PyMuPDF can open (compose + print via QPainter)
FITZ_EXTS = frozenset({
    ".pdf", ".xps", ".epub", ".cbz", ".fb2", ".svg",
    ".jpg", ".jpeg", ".png", ".bmp", ".gif",
    ".tif", ".tiff", ".pnm", ".pgm", ".ppm", ".pbm", ".pam",
})


def parse_page_range(rng: str, total: int) -> list[int]:
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


def slot_to_grid(slot: int, cols: int, rows: int, order: str) -> tuple[int, int]:
    """Map a sheet slot index to (col, row) for the given page order."""
    if order.startswith("Vertical"):
        col_i, row_i = slot // rows, slot % rows
    else:
        col_i, row_i = slot % cols, slot // cols
    if "Reversed" in order:
        col_i, row_i = cols - 1 - col_i, rows - 1 - row_i
    return col_i, row_i


def pdf_page_count(path: str) -> int | None:
    """Return PDF page count using pypdf if available, else None."""
    try:
        from pypdf import PdfReader
        return len(PdfReader(path).pages)
    except Exception:
        return None
