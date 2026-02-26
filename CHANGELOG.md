# Changelog

All notable changes to CloudSearchBar are listed here.

---

## [Unreleased] — feature/local-file-search

### Added
- **Distance calculator** — type `distance X to Y` (cities, addresses, zip codes) to get driving distance (mi + km), estimated drive time, and one-click links to Google Maps directions and Google Flights. Works offline gracefully: falls back to a direct Maps link if geocoding or routing fails. No API key required (OSM Nominatim + public OSRM).
- **Search history autocomplete** — as you type, matching past queries appear instantly in the dropdown under a "Recent searches" header. Queries are saved whenever you open a file, do a Cloud Search, or act on a distance result. Stored in `query_history.json`.
- **Settings UI** — tray menu → "Open settings…" opens a dialog to change hotkey, browser, Google account index, local search paths, and autostart. Replaces manual ini editing.
- **Inno Setup installer** — `CloudSearchBar_Setup.exe` for easy distribution; per-user install by default, no admin rights required.

### Changed
- Debounce timer now runs unconditionally (previously gated on `LOCAL_ENABLED`), so the distance calculator works even when local file search is disabled.

---

## [1.0.0] — Initial release

### Added
- Frameless floating search bar, centered on screen, invoked by global hotkey (`Ctrl+Space`)
- Local file search via Windows Search index (ADODB / `win32com`) with fuzzy re-ranking (`rapidfuzz`)
- Google Cloud Search fallback — **Workspace accounts only**. Personal Google account users should use the [Google App for Windows Labs experiment](https://blog.google/products-and-platforms/products/search/google-app-windows-labs/) instead (`Alt+Space`, available at labs.google.com)
- Recently opened files shown when search bar is empty
- File preview popup on hover (size, date, image thumbnail)
- System tray icon with Show / Settings / Quit menu
- Startup registry entry (`HKCU\...\Run`) with self-healing check on launch
- Opacity fade-in / fade-out animation
- `AttachThreadInput` focus-steal fix for background invocation
- `Shift+Enter` opens containing folder of a file result
- Configurable via `cloudsearch_bar.ini`
