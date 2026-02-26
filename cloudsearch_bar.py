"""
Cloud Search Bar
----------------
A floating search bar that searches local files and Google Cloud Search.

Usage:
  python cloudsearch_bar.py

Invoke:   configured hotkey (default: Ctrl+Space)
Dismiss:  Escape
Navigate: Down/Up arrows
Open:     Enter or click
Folder:   Shift+Enter on a file result (opens containing folder)

Configuration: cloudsearch_bar.ini (same directory as this script)

Requires:
  pip install PyQt6 keyboard pywin32 rapidfuzz
"""

import sys
import os
import re
import json
import webbrowser
import urllib.parse
import urllib.request
import configparser
import threading
from pathlib import Path
from datetime import datetime

import keyboard
from rapidfuzz import fuzz as rfuzz

from PyQt6.QtWidgets import (
    QApplication, QLineEdit, QVBoxLayout, QHBoxLayout, QWidget,
    QSystemTrayIcon, QMenu, QListWidget, QListWidgetItem,
    QFrame, QGraphicsDropShadowEffect, QLabel, QFileIconProvider,
    QDialog, QFormLayout, QGroupBox, QSpinBox, QCheckBox, QPushButton,
)
from PyQt6.QtCore import (
    Qt, pyqtSignal, QObject, QTimer, QSize,
    QPropertyAnimation, QEasingCurve, QFileInfo,
)
from PyQt6.QtGui import QIcon, QPixmap, QColor, QFont


# ── Paths ─────────────────────────────────────────────────────────────────────

def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

APP_DIR = _app_dir()


# ── Config ────────────────────────────────────────────────────────────────────

CONFIG_PATH  = APP_DIR / "cloudsearch_bar.ini"
HISTORY_PATH = APP_DIR / "history.json"

_DEFAULTS = {
    "Search":      {"hotkey": "ctrl+space", "account_index": "0", "browser": ""},
    "Window":      {"width": "620", "height": "60"},
    "LocalSearch": {"enabled": "true", "max_results": "7", "paths": ""},
    "Startup":     {"autostart": "true"},
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
LOCAL_ENABLED = cfg.getboolean("LocalSearch", "enabled")
LOCAL_MAX     = cfg.getint("LocalSearch", "max_results")
_raw_paths    = cfg.get("LocalSearch", "paths").strip()
LOCAL_PATHS   = [p.strip() for p in _raw_paths.split(",") if p.strip()] if _raw_paths else []
AUTOSTART     = cfg.getboolean("Startup", "autostart")

_account_path = f"/u/{ACCOUNT_INDEX}/" if ACCOUNT_INDEX > 0 else "/"
SEARCH_URL    = f"https://cloudsearch.google.com{_account_path}cloudsearch/search?q={{}}"

# Layout constants
SHADOW_MARGIN   = 18
SHADOW_OFFSET_Y = 6
ITEM_HEIGHT     = 54
HEADER_HEIGHT   = 28
HISTORY_MAX          = 20
HISTORY_SHOW         = 5
QUERY_HISTORY_PATH   = APP_DIR / "query_history.json"
QUERY_HISTORY_MAX    = 100
IMAGE_EXTS      = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".ico", ".tiff", ".tif"}


# ── Windows Search ────────────────────────────────────────────────────────────

def search_local_files(query: str, paths: list, max_results: int) -> list:
    """
    Query the Windows Search index for files whose name contains `query`.
    Fetches 4× max_results so fuzzy ranking can pick the best matches.
    Returns a list of (filename, full_path) tuples.
    """
    try:
        import win32com.client
        safe_query = query.replace("'", "''")
        conn = win32com.client.Dispatch("ADODB.Connection")
        conn.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
        rs = win32com.client.Dispatch("ADODB.Recordset")

        scope_clause = ""
        if paths:
            parts = " OR ".join(f"scope='file:{p}'" for p in paths)
            scope_clause = f"AND ({parts})"

        fetch_limit = max_results * 4
        sql = (
            f"SELECT TOP {fetch_limit} System.ItemPathDisplay, System.FileName "
            f"FROM SystemIndex "
            f"WHERE System.FileName LIKE '%{safe_query}%' "
            f"{scope_clause} "
            f"ORDER BY System.DateModified DESC"
        )
        rs.Open(sql, conn)
        results = []
        while not rs.EOF:
            path = rs.Fields["System.ItemPathDisplay"].Value
            name = rs.Fields["System.FileName"].Value
            if path and name:
                results.append((name, path))
            rs.MoveNext()
        rs.Close()
        conn.Close()
        return results
    except Exception:
        return []


# ── Fuzzy ranking ─────────────────────────────────────────────────────────────

def fuzzy_rank(query: str, results: list, top_n: int) -> list:
    """Re-rank results by fuzzy match score on filename and return top_n."""
    q = query.lower()
    scored = [(name, path, rfuzz.WRatio(q, name.lower())) for name, path in results]
    scored.sort(key=lambda x: x[2], reverse=True)
    return [(name, path) for name, path, _ in scored[:top_n]]


# ── History ───────────────────────────────────────────────────────────────────

def load_history() -> list:
    try:
        with open(HISTORY_PATH) as f:
            return json.load(f)
    except Exception:
        return []

def add_to_history(path: str):
    entries = load_history()
    entries = [e for e in entries if e.get("path") != path]
    entries.insert(0, {
        "path": path,
        "name": Path(path).name,
        "opened": datetime.now().isoformat(),
    })
    try:
        with open(HISTORY_PATH, "w") as f:
            json.dump(entries[:HISTORY_MAX], f, indent=2)
    except Exception:
        pass


def load_query_history() -> list:
    try:
        with open(QUERY_HISTORY_PATH) as f:
            return json.load(f)
    except Exception:
        return []

def save_query(query: str):
    if not query:
        return
    entries = load_query_history()
    entries = [q for q in entries if q.lower() != query.lower()]
    entries.insert(0, query)
    try:
        with open(QUERY_HISTORY_PATH, "w") as f:
            json.dump(entries[:QUERY_HISTORY_MAX], f, indent=2)
    except Exception:
        pass


# ── Autostart ─────────────────────────────────────────────────────────────────

def sync_autostart(enable: bool) -> bool:
    """Add or remove the Windows startup registry entry. Exe only.

    Returns True if the registry state matches intent, False if something
    went wrong (e.g. permissions blocked the write).
    """
    if not getattr(sys, "frozen", False):
        return True
    try:
        import winreg
        exe_path = str(Path(sys.executable))
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE,
        )
        if enable:
            winreg.SetValueEx(key, "CloudSearchBar", 0, winreg.REG_SZ, exe_path)
            # Verify the write actually took effect
            value, _ = winreg.QueryValueEx(key, "CloudSearchBar")
            winreg.CloseKey(key)
            return value == exe_path
        else:
            try:
                winreg.DeleteValue(key, "CloudSearchBar")
            except FileNotFoundError:
                pass
            winreg.CloseKey(key)
            return True
    except Exception:
        return False


# ── Distance calculator ───────────────────────────────────────────────────────

_DISTANCE_RE = re.compile(r'^distance\s+(.+?)\s+to\s+(.+)$', re.IGNORECASE)


_US_ZIP_RE = re.compile(r'^\d{5}(-\d{4})?$')

def _nominatim_geocode(place: str):
    """Return (lat, lon) floats or None. Uses OSM Nominatim, no API key."""
    try:
        country = "&countrycodes=us" if _US_ZIP_RE.match(place.strip()) else ""
        url = (
            "https://nominatim.openstreetmap.org/search?q="
            + urllib.parse.quote(place)
            + "&format=json&limit=1"
            + country
        )
        req = urllib.request.Request(url, headers={"User-Agent": "CloudSearchBar/1.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode())
        if data:
            return (float(data[0]["lat"]), float(data[0]["lon"]))
    except Exception:
        pass
    return None


def _osrm_route(lat1, lon1, lat2, lon2):
    """Return {distance_mi, distance_km, time_str} or None. Uses public OSRM."""
    try:
        url = (
            f"https://router.project-osrm.org/route/v1/driving/"
            f"{lon1},{lat1};{lon2},{lat2}?overview=false"
        )
        req = urllib.request.Request(url, headers={"User-Agent": "CloudSearchBar/1.0"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode())
        route = data["routes"][0]
        metres  = route["distance"]
        seconds = int(route["duration"])
        miles   = metres / 1609.344
        km      = metres / 1000.0
        hours, rem = divmod(seconds, 3600)
        mins = rem // 60
        time_str = f"{hours}h {mins}m" if hours else f"{mins}m"
        return {"distance_mi": miles, "distance_km": km, "time_str": time_str}
    except Exception:
        pass
    return None


def _maps_url(origin: str, dest: str) -> str:
    return (
        "https://www.google.com/maps/dir/"
        + urllib.parse.quote(origin)
        + "/"
        + urllib.parse.quote(dest)
        + "/"
    )


def _flights_url(origin: str, dest: str) -> str:
    q = urllib.parse.quote(f"Flights from {origin} to {dest}")
    return f"https://www.google.com/flights?q={q}"


# ── Preview popup ─────────────────────────────────────────────────────────────

class PreviewPopup(QWidget):
    def __init__(self):
        super().__init__(
            None,
            Qt.WindowType.ToolTip
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint,
        )
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setStyleSheet("""
            QWidget { background: #ffffff; border: 1px solid #dadce0; border-radius: 8px; }
            QLabel  { border: none; background: transparent; }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        self._img_label = QLabel()
        self._img_label.setFixedSize(200, 160)
        self._img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._img_label.hide()

        self._meta_label = QLabel()
        self._meta_label.setWordWrap(True)
        self._meta_label.setMaximumWidth(220)
        self._meta_label.setStyleSheet("font-size: 11px; color: #5f6368;")

        layout.addWidget(self._img_label)
        layout.addWidget(self._meta_label)

    def show_for(self, path: str, global_pos):
        p = Path(path)
        try:
            stat = p.stat()
            size = stat.st_size
            modified = datetime.fromtimestamp(stat.st_mtime).strftime("%b %d, %Y  %H:%M")
            size_str = (
                f"{size / 1024:.1f} KB"
                if size < 1024 * 1024
                else f"{size / (1024 * 1024):.1f} MB"
            )
            ext = p.suffix.upper().lstrip(".") or "File"
            meta = f"<b>{p.name}</b><br>{ext} · {size_str}<br>{modified}"
        except Exception:
            meta = p.name

        if p.suffix.lower() in IMAGE_EXTS:
            px = QPixmap(path)
            if not px.isNull():
                self._img_label.setPixmap(
                    px.scaled(200, 160, Qt.AspectRatioMode.KeepAspectRatio,
                               Qt.TransformationMode.SmoothTransformation)
                )
                self._img_label.show()
            else:
                self._img_label.hide()
        else:
            self._img_label.hide()

        self._meta_label.setText(meta)
        self.adjustSize()
        self.move(global_pos.x() + 12, global_pos.y() - self.height() // 2)
        self.show()
        self.raise_()


# ── Signals ───────────────────────────────────────────────────────────────────

class _Signals(QObject):
    show_window    = pyqtSignal()
    search_results = pyqtSignal(list)
    distance_result = pyqtSignal(object)   # emits a dict


# ── Settings Dialog ───────────────────────────────────────────────────────────

class SettingsDialog(QDialog):
    def __init__(self, search_bar: "SearchBar"):
        super().__init__(None)
        self._bar = search_bar
        self.setWindowTitle("CloudSearchBar — Settings")
        icon_path = APP_DIR / "icon.png"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self.setMinimumWidth(440)
        self._build_ui()
        self._load_values()
        self.setStyleSheet("""
            QDialog {
                background: #f8f9fa;
            }
            QGroupBox {
                font-weight: bold;
                color: #3c4043;
                border: 1px solid #dadce0;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 6px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
                color: #1a73e8;
            }
            QLineEdit {
                background: #ffffff;
                border: 1px solid #dadce0;
                border-radius: 4px;
                padding: 4px 8px;
                color: #3c4043;
                font-size: 13px;
            }
            QLineEdit:focus {
                border: 1px solid #4285F4;
            }
            QSpinBox {
                background: #ffffff;
                border: 1px solid #dadce0;
                border-radius: 4px;
                padding: 4px 8px;
                color: #3c4043;
                font-size: 13px;
            }
            QCheckBox {
                color: #3c4043;
                font-size: 13px;
            }
            QLabel {
                color: #3c4043;
                font-size: 13px;
            }
            QPushButton {
                background: #4285F4;
                color: #ffffff;
                border: none;
                border-radius: 4px;
                padding: 6px 20px;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #1a73e8;
            }
            QPushButton#cancel_btn {
                background: #ffffff;
                color: #3c4043;
                border: 1px solid #dadce0;
            }
            QPushButton#cancel_btn:hover {
                background: #f1f3f4;
            }
        """)

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setSpacing(12)
        outer.setContentsMargins(16, 16, 16, 16)

        # ── Search group
        search_box = QGroupBox("Search")
        search_form = QFormLayout(search_box)
        search_form.setSpacing(8)
        search_form.setContentsMargins(12, 16, 12, 12)

        self._hotkey_edit = QLineEdit()
        self._hotkey_edit.setPlaceholderText("e.g. ctrl+space")
        search_form.addRow("Hotkey:", self._hotkey_edit)

        self._browser_edit = QLineEdit()
        self._browser_edit.setPlaceholderText("blank = system default")
        search_form.addRow("Browser:", self._browser_edit)

        self._account_spin = QSpinBox()
        self._account_spin.setRange(0, 9)
        search_form.addRow("Google account index:", self._account_spin)

        outer.addWidget(search_box)

        # ── Local Search group
        local_box = QGroupBox("Local Search")
        local_form = QFormLayout(local_box)
        local_form.setSpacing(8)
        local_form.setContentsMargins(12, 16, 12, 12)

        self._local_enabled_check = QCheckBox()
        local_form.addRow("Enabled:", self._local_enabled_check)

        self._local_max_spin = QSpinBox()
        self._local_max_spin.setRange(1, 20)
        local_form.addRow("Max file results:", self._local_max_spin)

        self._local_paths_edit = QLineEdit()
        self._local_paths_edit.setPlaceholderText("blank = everywhere  (comma-separated paths)")
        local_form.addRow("Search paths:", self._local_paths_edit)

        outer.addWidget(local_box)

        # ── Startup group
        startup_box = QGroupBox("Startup")
        startup_form = QFormLayout(startup_box)
        startup_form.setSpacing(8)
        startup_form.setContentsMargins(12, 16, 12, 12)

        self._autostart_check = QCheckBox()
        startup_form.addRow("Launch at startup:", self._autostart_check)

        outer.addWidget(startup_box)

        # ── Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("cancel_btn")
        cancel_btn.clicked.connect(self.reject)

        save_btn = QPushButton("Save")
        save_btn.setDefault(True)
        save_btn.clicked.connect(self._save)

        btn_row.addWidget(cancel_btn)
        btn_row.addWidget(save_btn)
        outer.addLayout(btn_row)

    def _load_values(self):
        self._hotkey_edit.setText(HOTKEY)
        self._browser_edit.setText(BROWSER)
        self._account_spin.setValue(ACCOUNT_INDEX)
        self._local_enabled_check.setChecked(LOCAL_ENABLED)
        self._local_max_spin.setValue(LOCAL_MAX)
        self._local_paths_edit.setText(", ".join(LOCAL_PATHS))
        self._autostart_check.setChecked(AUTOSTART)

    def _save(self):
        global HOTKEY, BROWSER, ACCOUNT_INDEX, SEARCH_URL
        global LOCAL_ENABLED, LOCAL_MAX, LOCAL_PATHS, AUTOSTART

        old_hotkey    = HOTKEY
        old_autostart = AUTOSTART

        # Read widget values
        new_hotkey        = self._hotkey_edit.text().strip() or HOTKEY
        new_browser       = self._browser_edit.text().strip()
        new_account_index = self._account_spin.value()
        new_local_enabled = self._local_enabled_check.isChecked()
        new_local_max     = self._local_max_spin.value()
        raw_paths         = self._local_paths_edit.text().strip()
        new_local_paths   = [p.strip() for p in raw_paths.split(",") if p.strip()]
        new_autostart     = self._autostart_check.isChecked()

        # Write config file cleanly
        new_cfg = configparser.ConfigParser()
        new_cfg["Search"] = {
            "hotkey":        new_hotkey,
            "account_index": str(new_account_index),
            "browser":       new_browser,
        }
        new_cfg["Window"] = {
            "width":  str(WINDOW_WIDTH),
            "height": str(WINDOW_HEIGHT),
        }
        new_cfg["LocalSearch"] = {
            "enabled":     "true" if new_local_enabled else "false",
            "max_results": str(new_local_max),
            "paths":       ", ".join(new_local_paths),
        }
        new_cfg["Startup"] = {
            "autostart": "true" if new_autostart else "false",
        }
        with open(CONFIG_PATH, "w") as f:
            new_cfg.write(f)

        # Update globals
        HOTKEY        = new_hotkey
        BROWSER       = new_browser
        ACCOUNT_INDEX = new_account_index
        _account_path = f"/u/{ACCOUNT_INDEX}/" if ACCOUNT_INDEX > 0 else "/"
        SEARCH_URL    = f"https://cloudsearch.google.com{_account_path}cloudsearch/search?q={{}}"
        LOCAL_ENABLED = new_local_enabled
        LOCAL_MAX     = new_local_max
        LOCAL_PATHS   = new_local_paths
        AUTOSTART     = new_autostart

        # Rebind hotkey if changed
        if new_hotkey != old_hotkey:
            try:
                keyboard.remove_hotkey(old_hotkey)
            except Exception:
                pass
            keyboard.add_hotkey(new_hotkey, self._bar._signals.show_window.emit)

        # Sync autostart registry if changed
        if new_autostart != old_autostart:
            sync_autostart(new_autostart)

        self.accept()


# ── Search Bar ────────────────────────────────────────────────────────────────

class SearchBar(QWidget):
    def __init__(self):
        super().__init__()
        self._signals = _Signals()
        self._signals.show_window.connect(self._show_and_focus)
        self._signals.search_results.connect(self._display_results)
        self._signals.distance_result.connect(self._display_distance_result)
        self._search_gen = 0   # incremented each query; stale threads compare and discard
        self._anim = None
        self.line_edit = None      # guard for eventFilter before _init_ui completes
        self.results_list = None   # guard for eventFilter before _init_ui completes
        self._preview = PreviewPopup()
        self._icon_provider = QFileIconProvider()
        self._init_ui()
        self._init_tray()
        self._init_hotkey()
        if not sync_autostart(AUTOSTART) and AUTOSTART:
            QTimer.singleShot(2000, self._warn_autostart_failed)

    # ── UI ────────────────────────────────────────────────────────────────────

    def _init_ui(self):
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedWidth(WINDOW_WIDTH + SHADOW_MARGIN * 2)

        # ── Container (the visible card with drop shadow)
        self.container = QFrame(self)
        self.container.move(SHADOW_MARGIN, SHADOW_MARGIN)
        self.container.setFixedWidth(WINDOW_WIDTH)

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(24)
        shadow.setOffset(0, SHADOW_OFFSET_Y)
        shadow.setColor(QColor(0, 0, 0, 70))
        self.container.setGraphicsEffect(shadow)

        # ── Search input
        self.line_edit = QLineEdit()
        self.line_edit.setPlaceholderText("Search…")
        self.line_edit.setFixedHeight(WINDOW_HEIGHT)
        self.line_edit.returnPressed.connect(self._on_enter)
        self.line_edit.installEventFilter(self)
        self._style_input(results_open=False)

        # ── Results list
        self.results_list = QListWidget()
        self.results_list.setMouseTracking(True)
        self.results_list.setIconSize(QSize(24, 24))
        self.results_list.hide()
        self.results_list.itemActivated.connect(self._open_result)
        self.results_list.itemEntered.connect(self._on_item_hovered)
        self.results_list.installEventFilter(self)
        self._style_results()

        # ── Container layout
        c_layout = QVBoxLayout(self.container)
        c_layout.setContentsMargins(0, 0, 0, 0)
        c_layout.setSpacing(0)
        c_layout.addWidget(self.line_edit)
        c_layout.addWidget(self.results_list)

        self._set_content_height(WINDOW_HEIGHT)

        # ── Debounce timer — drives local search and distance queries
        self._search_timer = QTimer()
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(250)
        self._search_timer.timeout.connect(self._fire_search)
        self.line_edit.textChanged.connect(self._on_text_changed)

    def _set_content_height(self, h: int):
        self.container.setFixedHeight(h)
        self.setFixedHeight(h + SHADOW_MARGIN * 2 + SHADOW_OFFSET_Y)

    def _style_input(self, results_open: bool):
        radius = (
            "border-top-left-radius: 25px; border-top-right-radius: 25px; "
            "border-bottom-left-radius: 0; border-bottom-right-radius: 0;"
            if results_open else
            "border-radius: 25px;"
        )
        self.line_edit.setStyleSheet(f"""
            QLineEdit {{
                background-color: #ffffff;
                border: 2px solid #4285F4;
                {radius}
                padding: 10px 24px;
                font-size: 20px;
                color: #3c4043;
            }}
        """)

    def _style_results(self):
        self.results_list.setStyleSheet("""
            QListWidget {
                background-color: #ffffff;
                border: 2px solid #4285F4;
                border-top: none;
                border-bottom-left-radius: 12px;
                border-bottom-right-radius: 12px;
                outline: none;
            }
            QListWidget::item {
                padding: 4px 12px;
                border-bottom: 1px solid #e8eaed;
                color: #3c4043;
            }
            QListWidget::item:selected {
                background-color: #e8f0fe;
                color: #1a73e8;
            }
            QListWidget::item:last-child {
                border-bottom: none;
            }
            QListWidget::item:disabled {
                background-color: #f8f9fa;
                color: #80868b;
                font-size: 11px;
                padding: 3px 12px;
            }
        """)

    def keyPressEvent(self, event):
        k = event.key()
        if k == Qt.Key.Key_Escape:
            self._hide()
        elif k == Qt.Key.Key_Down and self.results_list.isVisible():
            self.results_list.setFocus()
            # Skip non-selectable header (row 0)
            self.results_list.setCurrentRow(1 if self.results_list.count() > 1 else 0)

    def eventFilter(self, source, event):
        t = event.type()

        # ── Line edit: intercept Shift+Enter
        if self.line_edit is not None and source is self.line_edit and t == event.Type.KeyPress:
            if (event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter)
                    and event.modifiers() & Qt.KeyboardModifier.ShiftModifier):
                self._open_folder_of_current()
                return True

        # ── Results list: keyboard navigation + preview hide on leave
        if self.results_list is not None and source is self.results_list:
            if t == event.Type.KeyPress:
                k = event.key()
                mods = event.modifiers()
                if k == Qt.Key.Key_Escape:
                    self._hide()
                    return True
                if k in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                    if mods & Qt.KeyboardModifier.ShiftModifier:
                        self._open_folder_of_current()
                    else:
                        item = self.results_list.currentItem()
                        if item:
                            self._open_result(item)
                    return True
                if k == Qt.Key.Key_Up and self.results_list.currentRow() <= 1:
                    self.line_edit.setFocus()
                    self.results_list.clearSelection()
                    return True
            if t == event.Type.Leave:
                self._preview.hide()

        return super().eventFilter(source, event)

    def _open_folder_of_current(self):
        if not self.results_list.isVisible():
            return
        item = self.results_list.currentItem()
        if item:
            path = item.data(Qt.ItemDataRole.UserRole)
            if path:
                os.startfile(str(Path(path).parent))
                self._hide()

    def _center(self):
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width()) // 2
        y = screen.height() // 3
        self.move(x, y)

    # ── Tray ──────────────────────────────────────────────────────────────────

    def _init_tray(self):
        icon_path = APP_DIR / "icon.png"
        icon = QIcon(str(icon_path)) if icon_path.exists() else QIcon()

        self.tray = QSystemTrayIcon(icon, self)
        self.tray.setToolTip(f"Cloud Search  ({HOTKEY})")

        menu = QMenu()
        menu.addAction(f"Show  ({HOTKEY})").triggered.connect(self._show_and_focus)
        menu.addSeparator()
        menu.addAction("Open settings…").triggered.connect(self._open_settings)
        menu.addSeparator()
        menu.addAction("Quit").triggered.connect(QApplication.quit)

        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self._on_tray_click)
        self.tray.show()

    def _on_tray_click(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self._show_and_focus()

    # ── Hotkey ────────────────────────────────────────────────────────────────

    def _init_hotkey(self):
        keyboard.add_hotkey(HOTKEY, self._signals.show_window.emit)

    def _warn_autostart_failed(self):
        self.tray.showMessage(
            "CloudSearchBar — Autostart Warning",
            "Could not write the startup registry entry. "
            "CloudSearchBar may not launch automatically on login. "
            "Try running as administrator.",
            QSystemTrayIcon.MessageIcon.Warning,
            6000,
        )

    # ── Local search ──────────────────────────────────────────────────────────

    def _on_text_changed(self, text):
        self._search_timer.stop()
        if text.strip():
            self._show_query_suggestions(text.strip())
            self._search_timer.start()
        else:
            self._clear_results()
            self._show_recently_opened()

    def _fire_search(self):
        query = self.line_edit.text().strip()
        if not query:
            return
        self._search_gen += 1
        m = _DISTANCE_RE.match(query)
        if m:
            self._run_distance_search(m.group(1).strip(), m.group(2).strip(), self._search_gen)
            return
        if LOCAL_ENABLED:
            self._run_local_search()

    def _run_local_search(self):
        query = self.line_edit.text().strip()
        if not query:
            return
        def _worker():
            results = search_local_files(query, LOCAL_PATHS, LOCAL_MAX)
            self._signals.search_results.emit(results)
        threading.Thread(target=_worker, daemon=True).start()

    def _run_distance_search(self, origin: str, dest: str, gen: int):
        def _worker():
            info = {"origin": origin, "dest": dest, "_gen": gen}
            o = _nominatim_geocode(origin)
            d = _nominatim_geocode(dest)
            if o and d:
                route = _osrm_route(o[0], o[1], d[0], d[1])
                if route:
                    info.update(route)
            self._signals.distance_result.emit(info)
        threading.Thread(target=_worker, daemon=True).start()

    def _display_distance_result(self, info: dict):
        if info.get("_gen") != self._search_gen:
            return   # stale result from a superseded query
        origin = info["origin"]
        dest   = info["dest"]
        maps_link    = _maps_url(origin, dest)
        flights_link = _flights_url(origin, dest)

        self.results_list.clear()

        # Non-selectable header
        header_text = f"{origin.title()} → {dest.title()}"
        h_item = QListWidgetItem(header_text)
        h_item.setFlags(Qt.ItemFlag.NoItemFlags)
        h_font = QFont()
        h_font.setPointSize(9)
        h_item.setFont(h_font)
        h_item.setSizeHint(QSize(WINDOW_WIDTH, HEADER_HEIGHT))
        self.results_list.addItem(h_item)

        f_font = QFont()
        f_font.setPointSize(10)

        # Row 1 — distance/time or error
        if "distance_mi" in info:
            mi  = info["distance_mi"]
            km  = info["distance_km"]
            t   = info["time_str"]
            row1_text = f"  \U0001f5fa  {mi:.0f} mi · {km:.0f} km · {t} driving"
            row1_data = maps_link
        else:
            row1_text = "  \u26a0  Could not calculate — Open in Google Maps"
            row1_data = maps_link

        row1 = QListWidgetItem(row1_text)
        row1.setData(Qt.ItemDataRole.UserRole, row1_data)
        row1.setFont(f_font)
        row1.setSizeHint(QSize(WINDOW_WIDTH, ITEM_HEIGHT))
        self.results_list.addItem(row1)

        # Row 2 — Flights
        row2 = QListWidgetItem("  \u2708  Search Google Flights")
        row2.setData(Qt.ItemDataRole.UserRole, flights_link)
        row2.setFont(f_font)
        row2.setForeground(QColor("#1a73e8"))
        row2.setSizeHint(QSize(WINDOW_WIDTH, ITEM_HEIGHT))
        self.results_list.addItem(row2)

        list_height = HEADER_HEIGHT + 2 * ITEM_HEIGHT
        self.results_list.setFixedHeight(list_height)
        self.results_list.show()
        self._set_content_height(WINDOW_HEIGHT + list_height)
        self._style_input(results_open=True)

    def _display_results(self, results: list):
        query = self.line_edit.text().strip()
        if not query:
            return
        results = fuzzy_rank(query, results, LOCAL_MAX)
        self._populate_list(
            items=results,
            header=f"{len(results)} local result{'s' if len(results) != 1 else ''}"
                   if results else "No local files found",
            cloud_label=f'Search Cloud Search for "{query}" →',
        )

    def _show_recently_opened(self):
        history = load_history()
        if not history:
            return
        items = [(e["name"], e["path"]) for e in history[:HISTORY_SHOW] if Path(e["path"]).exists()]
        if not items:
            return
        self._populate_list(
            items=items,
            header="Recently opened",
            cloud_label="Search Cloud Search…",
        )

    def _show_query_suggestions(self, text: str):
        history = load_query_history()
        matches = [q for q in history
                   if q.lower().startswith(text.lower())
                   and q.lower() != text.lower()][:5]
        if not matches:
            return
        self.results_list.clear()

        h_item = QListWidgetItem("Recent searches")
        h_item.setFlags(Qt.ItemFlag.NoItemFlags)
        h_font = QFont()
        h_font.setPointSize(9)
        h_item.setFont(h_font)
        h_item.setSizeHint(QSize(WINDOW_WIDTH, HEADER_HEIGHT))
        self.results_list.addItem(h_item)

        f_font = QFont()
        f_font.setPointSize(10)
        for q in matches:
            item = QListWidgetItem(f"  \U0001f550  {q}")
            item.setData(Qt.ItemDataRole.UserRole, {"query": q})
            item.setFont(f_font)
            item.setSizeHint(QSize(WINDOW_WIDTH, ITEM_HEIGHT))
            self.results_list.addItem(item)

        list_height = HEADER_HEIGHT + len(matches) * ITEM_HEIGHT
        self.results_list.setFixedHeight(list_height)
        self.results_list.show()
        self._set_content_height(WINDOW_HEIGHT + list_height)
        self._style_input(results_open=True)

    def _populate_list(self, items: list, header: str, cloud_label: str):
        self.results_list.clear()

        # ── Non-selectable scope header
        h_item = QListWidgetItem(header)
        h_item.setFlags(Qt.ItemFlag.NoItemFlags)
        h_font = QFont()
        h_font.setPointSize(9)
        h_item.setFont(h_font)
        h_item.setSizeHint(QSize(WINDOW_WIDTH, HEADER_HEIGHT))
        self.results_list.addItem(h_item)

        # ── File results with system icon
        f_font = QFont()
        f_font.setPointSize(10)
        for name, path in items:
            item = QListWidgetItem()
            item.setText(f"  {name}\n  {path}")
            item.setData(Qt.ItemDataRole.UserRole, path)
            item.setFont(f_font)
            item.setIcon(self._icon_provider.icon(QFileInfo(path)))
            item.setSizeHint(QSize(WINDOW_WIDTH, ITEM_HEIGHT))
            self.results_list.addItem(item)

        # ── Pinned Cloud Search item
        cloud_item = QListWidgetItem(f"  {cloud_label}")
        cloud_item.setData(Qt.ItemDataRole.UserRole, None)
        cloud_item.setForeground(QColor("#1a73e8"))
        c_font = QFont()
        c_font.setPointSize(10)
        c_font.setItalic(True)
        cloud_item.setFont(c_font)
        cloud_item.setSizeHint(QSize(WINDOW_WIDTH, ITEM_HEIGHT))
        self.results_list.addItem(cloud_item)

        # ── Resize
        total_items = self.results_list.count()
        list_height = HEADER_HEIGHT + (total_items - 1) * ITEM_HEIGHT
        max_height  = HEADER_HEIGHT + (LOCAL_MAX + 1) * ITEM_HEIGHT
        list_height = min(list_height, max_height)

        self.results_list.setFixedHeight(list_height)
        self.results_list.show()
        self._set_content_height(WINDOW_HEIGHT + list_height)
        self._style_input(results_open=True)

    def _clear_results(self):
        self._preview.hide()
        self.results_list.clear()
        self.results_list.hide()
        self._set_content_height(WINDOW_HEIGHT)
        self._style_input(results_open=False)

    # ── Item hover preview ────────────────────────────────────────────────────

    def _on_item_hovered(self, item: QListWidgetItem):
        path = item.data(Qt.ItemDataRole.UserRole)
        if not path or (isinstance(path, str) and path.startswith("https://")):
            self._preview.hide()
            return
        rect       = self.results_list.visualItemRect(item)
        global_pos = self.results_list.mapToGlobal(rect.topRight())
        self._preview.show_for(path, global_pos)

    # ── Animation ─────────────────────────────────────────────────────────────

    def _animate_opacity(self, start: float, end: float, on_finish=None):
        if self._anim and self._anim.state() == QPropertyAnimation.State.Running:
            self._anim.stop()
        anim = QPropertyAnimation(self, b"windowOpacity")
        anim.setDuration(150)
        anim.setStartValue(start)
        anim.setEndValue(end)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        if on_finish:
            anim.finished.connect(on_finish)
        anim.start()
        self._anim = anim  # prevent GC

    # ── Actions ───────────────────────────────────────────────────────────────

    def _force_foreground(self):
        """
        Use the AttachThreadInput trick to steal focus from any foreground window.
        Qt's activateWindow() alone is blocked by Windows focus-stealing prevention
        when invoked from a background process (e.g. while a terminal has focus).
        """
        import ctypes
        user32   = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        hwnd             = int(self.winId())
        foreground_hwnd  = user32.GetForegroundWindow()
        foreground_tid   = user32.GetWindowThreadProcessId(foreground_hwnd, None)
        our_tid          = kernel32.GetCurrentThreadId()
        if foreground_tid and foreground_tid != our_tid:
            user32.AttachThreadInput(foreground_tid, our_tid, True)
            user32.SetForegroundWindow(hwnd)
            user32.BringWindowToTop(hwnd)
            user32.AttachThreadInput(foreground_tid, our_tid, False)
        else:
            user32.SetForegroundWindow(hwnd)

    def _show_and_focus(self):
        self._clear_results()
        self._center()
        self.setWindowOpacity(0.0)
        self.show()
        self._force_foreground()
        self.raise_()
        self.activateWindow()
        self.line_edit.setFocus()
        self.line_edit.clear()   # triggers _on_text_changed → shows history
        self._animate_opacity(0.0, 1.0)

    def _hide(self):
        self._preview.hide()
        def _finish():
            self.hide()
            self.setWindowOpacity(1.0)
            self._clear_results()
        self._animate_opacity(self.windowOpacity(), 0.0, on_finish=_finish)

    def _open_settings(self):
        dlg = SettingsDialog(self)
        dlg.exec()

    def _on_enter(self):
        if self.results_list.isVisible():
            item = self.results_list.currentItem()
            if item:
                self._open_result(item)
                return
        self._cloud_search()

    def _open_result(self, item: QListWidgetItem):
        if not (item.flags() & Qt.ItemFlag.ItemIsSelectable):
            return  # header — ignore
        data = item.data(Qt.ItemDataRole.UserRole)
        if isinstance(data, dict) and "query" in data:
            self._clear_results()
            self.line_edit.setText(data["query"])
            return
        if isinstance(data, str) and data.startswith("https://"):
            save_query(self.line_edit.text().strip())
            webbrowser.open(data)
            self._hide()
            return
        if data is None:
            self._cloud_search()
        else:
            save_query(self.line_edit.text().strip())
            os.startfile(data)
            add_to_history(data)
            self._hide()

    def _cloud_search(self):
        query = self.line_edit.text().strip()
        if query:
            save_query(query)
            url = SEARCH_URL.format(urllib.parse.quote_plus(query))
            if BROWSER:
                try:
                    webbrowser.get(BROWSER).open(url)
                except webbrowser.Error:
                    webbrowser.open(url)
            else:
                webbrowser.open(url)
        self.line_edit.clear()
        self._hide()


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    bar = SearchBar()
    sys.exit(app.exec())
