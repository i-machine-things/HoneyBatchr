"""HoneyBatchr — JobDocs plugin module.

Registers HoneyBatchr as the active print provider so that "Print Selected"
in the Job, Quote, and Search tabs sends files to the HoneyBatchr window
instead of the built-in print dialog.

Install via JobDocs → Install Plugin → i-machine-things/HoneyBatchr.
"""

import sys
from pathlib import Path

# Make this plugin's own 'modules' package importable without shadowing
# JobDocs' modules.  We register under a unique top-level name so normal
# 'from modules.job import …' calls in JobDocs are unaffected.
_plugin_dir = Path(__file__).parent
_modules_dir = _plugin_dir / 'modules'

# Register the plugin's modules package under the 'honeybatchr' namespace.
import importlib.util as _ilu
import types as _types

if 'honeybatchr' not in sys.modules:
    _pkg = _types.ModuleType('honeybatchr')
    _pkg.__path__ = [str(_modules_dir)]
    _pkg.__package__ = 'honeybatchr'
    sys.modules['honeybatchr'] = _pkg

def _load(name: str, path: Path):
    """Load a module from an explicit path under the honeybatchr namespace."""
    full = f'honeybatchr.{name}'
    if full in sys.modules:
        return sys.modules[full]
    spec = _ilu.spec_from_file_location(full, path)
    mod = _ilu.module_from_spec(spec)
    mod.__package__ = 'honeybatchr'
    sys.modules[full] = mod
    spec.loader.exec_module(mod)
    return mod

# Pre-load dependencies in correct order so intra-package imports resolve.
_load('config',   _modules_dir / 'config.py')
_load('utils',    _modules_dir / 'utils.py')
_load('themes',   _modules_dir / 'themes.py')
_load('widgets',  _modules_dir / 'widgets.py')
_load('dialogs',  _modules_dir / 'dialogs.py')
_load('printing', _modules_dir / 'printing.py')
_app_mod = _load('app', _modules_dir / 'app.py')
BatchPrintApp = _app_mod.BatchPrintApp

from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel, QPushButton
from PyQt6.QtCore import Qt
from core.base_module import BaseModule


class HoneyBatchrModule(BaseModule):
    """Wraps HoneyBatchr as a JobDocs print-provider plugin."""

    def get_name(self) -> str:
        return "Print Queue"

    def get_order(self) -> int:
        return 90

    def initialize(self, app_context) -> None:
        super().initialize(app_context)
        app_context.register_print_provider(self)

    def get_widget(self) -> QWidget:
        if self._widget is None:
            self._widget = self._build_widget()
        return self._widget

    def _build_widget(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        title = QLabel("<h2>HoneyBatchr</h2>")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        info = QLabel(
            "HoneyBatchr is active as your print provider.\n\n"
            "When you select files and press <b>Print Selected</b> in the Job, "
            "Quote, or Search tabs, they will be sent here automatically."
        )
        info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        info.setWordWrap(True)
        layout.addWidget(info)

        btn = QPushButton("Open Print Queue")
        btn.setFixedWidth(200)
        btn.clicked.connect(self._show_window)
        layout.addWidget(btn, alignment=Qt.AlignmentFlag.AlignCenter)

        return w

    # ------------------------------------------------------------------
    # Print provider interface (called by JobDocs)
    # ------------------------------------------------------------------

    def add_files_to_list(self, paths: list) -> None:
        """Receive files from JobDocs and open HoneyBatchr with them loaded."""
        self._show_window()
        self._get_window().add_files_to_list(paths)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _get_window(self) -> BatchPrintApp:
        if not hasattr(self, '_hb_window') or self._hb_window is None:
            self._hb_window = BatchPrintApp()
        return self._hb_window

    def _show_window(self) -> None:
        win = self._get_window()
        win.show()
        win.raise_()
        win.activateWindow()

    def cleanup(self) -> None:
        if hasattr(self, '_hb_window') and self._hb_window is not None:
            self._hb_window.close()
            self._hb_window = None


def create_module() -> HoneyBatchrModule:
    return HoneyBatchrModule()
