# Changelog

All notable changes to CloudSearchBar are listed here.

---

## [1.1.1] — 2026-07-06

### Fixed
- **Crash on "Recent searches" rows** — hovering or pressing Shift+Enter on a query suggestion (or a distance result) aborted the app; these handlers now only act on real file paths, and folder-open failures show a tray warning instead of crashing.
- **Local file search stopped working after the first query** — COM is now initialized on every search worker thread (`CoInitialize`), so Windows Search results keep coming instead of silently going empty.
- **System-wide installs broke startup and settings** — config, history, and log now fall back to `%LOCALAPPDATA%\CloudSearchBar` when the install folder isn't writable (e.g. Program Files); settings-save failures warn instead of crashing.
- **Malformed ini values killed startup** — invalid config values now fall back to defaults with a logged warning.
- **Failed file opens polluted history** — files that fail to open are removed from "Recently opened" instead of re-added, and showing the bar no longer checks file existence on the GUI thread (which could hang on dead network paths).
- **Stale tray labels after hotkey change** — the tray tooltip and "Show" menu entry now update immediately when the hotkey is changed in Settings.

### Changed
- **Single-instance guard** — a second launch now exits immediately instead of running a duplicate.
- **Distance queries debounce at 1s** (file searches remain 250ms) to respect the OSM Nominatim rate policy.
- **Windows Search queries escape LIKE wildcards** (`%`, `_`, `[`) so they match literally instead of acting as wildcards.

---

## [1.1.0] — 2026-02-26

### Added
- **Distance calculator** — type `distance X to Y` (cities, addresses, zip codes) to get driving distance (mi + km), estimated drive time, and one-click links to Google Maps directions and Google Flights. Works offline gracefully: falls back to a direct Maps link if geocoding or routing fails. No API key required (OSM Nominatim + public OSRM).
- **Search history autocomplete** — as you type, matching past queries appear instantly in the dropdown under a "Recent searches" header. Queries are saved whenever you open a file, do a Cloud Search, or act on a distance result. Stored in `query_history.json`.
- **Settings UI** — tray menu → "Open settings…" opens a dialog to change hotkey, browser, Google account index, local search paths, and autostart. Replaces manual ini editing.
- **Inno Setup installer** — `CloudSearchBar_Setup.exe` for easy distribution; per-user install by default, no admin rights required.
- **Haversine fallback** — when OSRM times out on very long routes, distance results fall back to straight-line aerial distance (`~X mi (aerial)`) instead of failing silently.

### Fixed
- **Race condition: stale local search overwriting distance results** — generation counter applied to local search threads so results from a superseded query can never replace a valid distance result already on screen.
- **Race condition: stale results on history selection** — results list is cleared immediately when a history suggestion is selected, preventing stale items from showing while distance APIs are in flight.
- **Zip code geocoding resolving to wrong country** — 5-digit US zip codes are now detected and Nominatim lookups are restricted to `countrycodes=us`, preventing matches against identically-numbered postal codes in Estonia, Finland, etc. City/state and international queries remain unrestricted.
- **OSRM timeout too short** — timeout increased from 8s to 15s to handle slow first connections on long-distance routes.

### Changed
- Debounce timer now runs unconditionally (previously gated on `LOCAL_ENABLED`), so the distance calculator works even when local file search is disabled.
- `search_results` signal now carries a `(results, gen)` tuple to support generation-based stale result discarding.

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
