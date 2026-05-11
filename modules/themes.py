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
    selection-color: #000000;
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

DARK_STYLESHEET = """
QToolTip {
    background-color: #3c3c3c;
    color: #dcdcdc;
    border: 1px solid #666666;
    padding: 4px;
}
QGroupBox {
    border: 1px solid #555555;
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
    border: 1px solid #666666;
    border-radius: 2px;
    padding: 6px 4px;
    background-color: #404040;
    min-height: 40px;
    color: #dcdcdc;
}
QPushButton#modeBtn:checked {
    background-color: #1e4d80;
    border-color: #4488cc;
}
QPushButton#modeBtn:hover:!checked {
    background-color: #4a4a4a;
}
QTableWidget {
    border: 1px solid #555555;
    gridline-color: #555555;
    selection-background-color: #2a82da;
    selection-color: #ffffff;
    alternate-background-color: #373737;
}
QHeaderView::section {
    background-color: #3d3d3d;
    color: #dcdcdc;
    border: none;
    border-right: 1px solid #555555;
    border-bottom: 1px solid #555555;
    padding: 3px 4px;
    font-weight: normal;
}
QFrame#rightPanel {
    background-color: #333333;
    border: 1px solid #555555;
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
