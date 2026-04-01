"""Custom Qt widgets."""

from PyQt6.QtWidgets import QTableWidget
from PyQt6.QtGui import QDragEnterEvent, QDropEvent


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
            self._on_drop([p for u in mime.urls() if (p := u.toLocalFile())])
            e.acceptProposedAction()
        else:
            super().dropEvent(e)
