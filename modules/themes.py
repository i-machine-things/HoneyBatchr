"""Qt palette factories and global stylesheet."""

from PyQt6.QtGui import QPalette, QColor

STYLESHEET = """
QToolTip {
    background-color: #ffffdc;
    color: #000000;
    border: 1px solid #a0a0a0;
    padding: 4px;
}
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


def light_palette() -> QPalette:
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


def dark_palette() -> QPalette:
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
