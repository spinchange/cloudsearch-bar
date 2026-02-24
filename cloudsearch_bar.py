"""
Cloud Search Bar
----------------
A floating search bar that queries Google Cloud Search.

Usage:
  python cloudsearch_bar.py

Invoke: configured hotkey (default: Ctrl+Space)
Dismiss: Escape
Search: Enter

Configuration: cloudsearch_bar.ini (same directory as this script)

Requires:
  pip install PyQt6 keyboard
"""

import sys
import webbrowser
import urllib.parse
import configparser
from pathlib import Path

import keyboard

from PyQt6.QtWidgets import (
    QApplication, QLineEdit, QVBoxLayout, QWidget,
    QSystemTrayIcon, QMenu,
)
from PyQt6.QtCore import Qt, pyqtSignal, QObject
from PyQt6.QtGui import QIcon, QPixmap, QColor


# ── Paths ─────────────────────────────────────────────────────────────────────

def _app_dir() -> Path:
    """Directory of the exe or script — works both frozen and in development."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

# ── Config ────────────────────────────────────────────────────────────────────

CONFIG_PATH = _app_dir() / "cloudsearch_bar.ini"

_DEFAULTS = {
    "Search": {
        "hotkey": "ctrl+space",
        "account_index": "0",
        "browser": "",
    },
    "Window": {
        "width": "620",
        "height": "60",
    },
}

def _load_config():
    cfg = configparser.ConfigParser()
    for section, values in _DEFAULTS.items():
        cfg[section] = values
    if CONFIG_PATH.exists():
        cfg.read(CONFIG_PATH)
    return cfg

cfg = _load_config()

HOTKEY        = cfg.get("Search", "hotkey").strip()
ACCOUNT_INDEX = cfg.getint("Search", "account_index")
BROWSER       = cfg.get("Search", "browser").strip()
WINDOW_WIDTH  = cfg.getint("Window", "width")
WINDOW_HEIGHT = cfg.getint("Window", "height")

_account_path = f"/u/{ACCOUNT_INDEX}/" if ACCOUNT_INDEX > 0 else "/"
SEARCH_URL = f"https://cloudsearch.google.com{_account_path}cloudsearch/search?q={{}}"

# ──────────────────────────────────────────────────────────────────────────────


class _Signal(QObject):
    """Bridge: emits from the keyboard hook thread → received on the Qt main thread."""
    triggered = pyqtSignal()


class SearchBar(QWidget):
    def __init__(self):
        super().__init__()
        self._signal = _Signal()
        self._signal.triggered.connect(self._show_and_focus)
        self._init_ui()
        self._init_tray()
        self._init_hotkey()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _init_ui(self):
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(WINDOW_WIDTH, WINDOW_HEIGHT)

        self.line_edit = QLineEdit(self)
        self.line_edit.setPlaceholderText("Search Google Workspace…")
        self.line_edit.setStyleSheet("""
            QLineEdit {
                background-color: #ffffff;
                border: 2px solid #4285F4;
                border-radius: 25px;
                padding: 10px 24px;
                font-size: 20px;
                color: #3c4043;
            }
        """)
        self.line_edit.returnPressed.connect(self._execute_search)

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.line_edit)
        self.setLayout(layout)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.hide()

    def _center(self):
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - WINDOW_WIDTH) // 2
        y = screen.height() // 3
        self.move(x, y)

    # ── Tray ──────────────────────────────────────────────────────────────────

    def _init_tray(self):
        icon_path = _app_dir() / "icon.png"
        if icon_path.exists():
            icon = QIcon(str(icon_path))
        else:
            px = QPixmap(16, 16)
            px.fill(QColor("#4285F4"))
            icon = QIcon(px)

        self.tray = QSystemTrayIcon(icon, self)
        self.tray.setToolTip(f"Cloud Search  ({HOTKEY})")

        menu = QMenu()
        show_action = menu.addAction(f"Show  ({HOTKEY})")
        show_action.triggered.connect(self._show_and_focus)
        menu.addSeparator()
        settings_action = menu.addAction(f"Open settings…")
        settings_action.triggered.connect(self._open_settings)
        menu.addSeparator()
        quit_action = menu.addAction("Quit")
        quit_action.triggered.connect(QApplication.quit)

        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self._on_tray_click)
        self.tray.show()

    def _on_tray_click(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self._show_and_focus()

    # ── Hotkey ────────────────────────────────────────────────────────────────

    def _init_hotkey(self):
        keyboard.add_hotkey(HOTKEY, self._signal.triggered.emit)

    # ── Actions ───────────────────────────────────────────────────────────────

    def _show_and_focus(self):
        self._center()
        self.show()
        self.raise_()
        self.activateWindow()
        self.line_edit.setFocus()
        self.line_edit.selectAll()

    def _open_settings(self):
        """Open the config file in the default text editor."""
        import os
        os.startfile(str(CONFIG_PATH))

    def _execute_search(self):
        query = self.line_edit.text().strip()
        if query:
            url = SEARCH_URL.format(urllib.parse.quote_plus(query))
            if BROWSER:
                try:
                    webbrowser.get(BROWSER).open(url)
                except webbrowser.Error:
                    webbrowser.open(url)  # fall back to system default
            else:
                webbrowser.open(url)
            self.line_edit.clear()
            self.hide()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    bar = SearchBar()
    sys.exit(app.exec())
