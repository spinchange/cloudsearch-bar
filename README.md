# CloudSearchBar

A lightweight, keyboard-driven Windows search bar that lives in the system tray. Invoke it from anywhere with a global hotkey, search local files, calculate driving distances, and open results instantly.

---

## ⚠️ Important — Google Cloud Search

The Cloud Search integration in this app targets **Google Workspace (enterprise) accounts** only. If you sign in with a regular personal Google account, Cloud Search will not work.

**Personal Google account users:** Google has a dedicated search app for you — the [Google App for Windows Labs experiment](https://blog.google/products-and-platforms/products/search/google-app-windows-labs/). Press `Alt+Space` to search your local files, Google Drive, installed apps, and the web. Enable it at [labs.google.com](https://labs.google.com).

---

## Features

- **Global hotkey** — invoke the bar from anywhere (`Ctrl+Space` by default)
- **Local file search** — queries the Windows Search index with fuzzy re-ranking; works across all indexed locations or scoped to specific paths
- **Distance calculator** — type `distance X to Y` (city names, addresses, or zip codes) to get driving distance, estimated drive time, and one-click links to Google Maps and Google Flights. No API key required (uses OSM Nominatim + public OSRM)
- **Search history autocomplete** — past queries surface instantly as you type
- **Recently opened files** — shown when the bar is empty
- **Google Cloud Search** — fallback to Cloud Search for Workspace users
- **System tray** — runs silently in the background; settings and quit accessible from the tray icon
- **Launch at startup** — optional Windows startup registry entry
- **File preview** — hover a result to see size, date modified, and image thumbnail

---

## Requirements

- Windows 10/11
- For local file search: Windows Search index must be enabled (on by default)
- For Cloud Search: Google Workspace account

---

## Installation

Download `CloudSearchBar_Setup.exe` from [Releases](../../releases) and run it. The installer requires no admin rights by default and places the app in your user profile.

After install, the app launches automatically and registers a startup entry so it runs on every login.

---

## Usage

| Action | How |
|---|---|
| Open the bar | `Ctrl+Space` (configurable) |
| Dismiss | `Escape` |
| Navigate results | `↑` / `↓` |
| Open a result | `Enter` or click |
| Open containing folder | `Shift+Enter` on a file result |
| Calculate distance | Type `distance chicago to new york` |
| Cloud Search fallback | Press `Enter` with no result selected |

---

## Configuration

Open **Settings** from the tray icon menu, or edit `cloudsearch_bar.ini` directly:

```ini
[Search]
hotkey = ctrl+space
account_index = 0      ; 0 = first Google account, 1 = second, etc.
browser =              ; blank = system default

[Window]
width = 620
height = 60

[LocalSearch]
enabled = true
max_results = 7
paths =                ; blank = search everywhere; comma-separated to restrict

[Startup]
autostart = true
```

---

## Building from source

```bash
pip install PyQt6 keyboard pywin32 rapidfuzz pyinstaller
python -m PyInstaller CloudSearchBar.spec
```

To rebuild the installer, compile `installer/CloudSearchBar.iss` with [Inno Setup 6](https://jrsoftware.org/isinfo.php).

---

## License

MIT
