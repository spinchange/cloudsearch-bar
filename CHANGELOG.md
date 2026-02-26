# Changelog

All notable changes to CloudSearchBar are listed here.

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
