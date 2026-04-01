"""
Honey Batchr - Batch Printing Application
Entry point.
"""

import sys
import json

from PyQt6.QtWidgets import QApplication, QStyleFactory

from modules.config import CONFIG_FILE
from modules.themes import light_palette, dark_palette, STYLESHEET
from modules.app import BatchPrintApp


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
        app.setPalette(dark_palette())
        app.setStyleSheet(STYLESHEET)
    elif theme == "Fusion Light" or theme not in QStyleFactory.keys():
        app.setStyle("Fusion")
        app.setPalette(light_palette())
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
