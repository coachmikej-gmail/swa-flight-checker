# Changelog

All notable changes to `swa_deals_gui.pyw` (formerly `swa_deals_gui.py`,
and before that `bwi_deals_gui.py`; and its companion script,
`build_bwi_deals.py`), newest first.

## 1.12.0 — Price history and chart
- Prices are now saved locally to a small SQLite database
  (`price_history.db`, next to the script/exe) after every fetch. A route's
  price is only written when it differs from the last one stored for that
  route, so the file grows with price *changes*, not with every fetch. Each
  home airport keeps its own history, and a route's first fetch is saved as
  its starting point.
- Double-click any row to open a price-history chart for that route: a step
  line (a price holds until it changes) with labeled points, hover tooltips
  (price, date/time, departure date), and a summary of current / lowest /
  highest price. A dashed segment extends the latest price to today to show
  it's still current. Double-clicking a column header or empty space does
  nothing.
- The chart is drawn with Tkinter's built-in canvas, so there's no new
  dependency (matplotlib would have added ~30 MB to the .exe).
- The status bar reports how many prices were saved after each fetch, and a
  failure to write history never fails the fetch itself.
- Added a tip next to the auto-refresh controls pointing out the
  double-click, and `price_history.db` to `.gitignore`.

## 1.11.0 — Price-change column, network resilience, packaging
- Added a **Change vs. Last Fetch** column showing the dollar change per
  destination since the previous fetch (`+$10`, `-$10`, `$0`; blank when
  there's nothing to compare against). It's sortable like the other columns.
- Network resilience, aimed at the intermittent `getaddrinfo failed` (DNS)
  errors seen after standby or login, when Wi-Fi/VPN/DNS aren't back yet:
  - Each request now retries up to 4 times with a short pause before giving
    up. Only connection-level failures are retried; HTTP errors fail
    immediately.
  - Auto-refresh no longer pops up an error dialog on a connection failure.
    It shows a status-bar message and quietly retries every 30 seconds,
    leaving the last good data on screen. Manual fetches still show a dialog,
    now with a clearer message.
  - The home-airport list retries every 30 seconds if it can't load at
    startup, so the dropdown fills in without restarting the app.
  - Other errors (e.g. Southwest changing its site) still show immediately
    and are not retried.
- The status line (e.g. "Loaded N destinations…") now sits at the bottom of
  the window instead of above the table.
- Window opens wider (1100px) so all columns are visible.
- Exports are saved next to the `.exe` when running as a PyInstaller build
  (previously they would have gone to a temp folder that's deleted on exit).
- Script file renamed `swa_deals_gui.py` → `swa_deals_gui.pyw`. On Windows,
  double-clicking a `.pyw` file runs it with `pythonw.exe`, which opens
  the GUI without a console window flashing up behind it. Command-line use
  still works: `python swa_deals_gui.pyw`.
- Updated `SWA Flight Checker.spec` to point at the new filename.

## 1.10.0 — Renamed to SWA Flight Checker
- App name changed from "BWI Flight Deals" to "SWA Flight Checker" (shown
  in the window title bar and the Help > About box), reflecting that it's
  no longer BWI-only since 1.9.0's home-airport picker.
- Script file renamed `bwi_deals_gui.py` → `swa_deals_gui.py` to match. Run
  it the same way, just with the new filename: `python swa_deals_gui.py`.

## 1.9.0 — Choose your home airport
- Added a "Home airport" dropdown, populated from Southwest's own list of
  origin markets (118+ airports) with a friendly "CODE — City, ST" label.
  Defaults to BWI and works immediately even before the full list finishes
  loading in the background.
- All the BWI-specific plumbing is now parameterized by the selected
  origin: distance column header, PDF/XLSX titles, and output filenames
  (`{CODE}_Flight_Deals.pdf` / `.xlsx`, so different home airports don't
  overwrite each other's exports).
- Price-change highlighting only compares fetches for the *same* home
  airport — switching origins doesn't produce misleading red/green rows
  by diffing unrelated routes.
- Renamed a couple of internal functions to match (`fetch_station_coords`
  → `fetch_station_directory`, `fetch_bwi_fares` → `fetch_fares`); no
  behavior change for BWI itself.

## 1.8.0 — About box
- Added a Help > About menu with the app name, version, a brief
  description, and the author's contact info.
- Window title now shows the app name and current version.
- `APP_NAME`/`APP_VERSION`/`APP_DESCRIPTION`/`AUTHOR_NAME`/`AUTHOR_EMAIL`
  are defined once near the top of the file — bump `APP_VERSION` there
  alongside new CHANGELOG entries.

## 1.7.0 — Auto-refresh
- Added a "minutes" input box and a Start/Stop Auto-Refresh button. When
  started, it fetches immediately and then re-fetches on that interval
  until stopped.
- Leaving the box blank and clicking Start is a no-op (auto-refresh stays
  off) rather than an error — only an actual invalid value (non-numeric,
  or under a 1-minute floor meant to avoid hammering southwest.com) pops
  an error dialog.
- If a fetch is still in flight when the next interval fires, that tick is
  skipped instead of stacking up overlapping requests.
- The interval box is disabled while auto-refresh is running, and the
  scheduled job is cancelled cleanly on Stop or on window close.

## 1.6.0 — Price-change highlighting
- The app now remembers the previous fetch's prices by airport code. On the
  next fetch, each row is compared against it and highlighted:
  - red if the price went up
  - green if the price went down
  - no highlight if it's unchanged or wasn't in the previous fetch
- Status bar shows an "N up, N down, N unchanged" summary after the second
  fetch onward.

## 1.5.0 — City/State columns
- Split the single "Destination" column into separate **City** and
  **State** columns (state holds the country name for international
  destinations, since Southwest's data doesn't distinguish the two).
- Both columns are independently sortable, same as before.

## 1.4.0 — Miles per $ column
- Added a **Miles per $** column (distance ÷ price) to the GUI, PDF, and
  XLSX exports, so cheap-but-short and pricey-but-long routes are easy to
  tell apart from pure fare price.
- XLSX version is a live formula, not a hardcoded value.

## 1.3.0 — Click-to-sort columns
- Clicking any column header in the results table sorts by that column;
  clicking again reverses direction (▲/▼ shown in the header).
- Each row's price rank is now computed once at fetch time and stays
  attached to the row — sorting by another column no longer scrambles it.
- Fixed a latent bug where "Departing" dates were compared as bare
  "M/D" strings with no year, which sorted incorrectly once dates crossed
  a year boundary (e.g. December 2026 vs. January 2027). Dates are now
  kept as real `date` objects internally and only formatted for display.

## 1.2.0 — XLSX export
- Added "Export to XLSX" / "Open XLSX", producing the same styled workbook
  layout as `build_bwi_deals.py` (banded rows, live `RANK()` formula,
  summary block with formulas for count/min/max/average).

## 1.1.0 — PDF export GUI
- Initial Tkinter GUI: "Fetch Latest Deals" pulls live fares from
  southwest.com on a background thread, shows them in a table, and
  "Export to PDF" writes a formatted landscape table via reportlab.

## 1.0.0 — Live-fetch data pipeline (`build_bwi_deals.py`)
- Established the core data source: southwest.com's own
  `stations-data` bundle (airport lat/long) and the `lfma.js` file backing
  the flight-deals page (per-route price/date hashes), fetched with plain
  HTTP requests — no browser/JS execution needed.
- Computes great-circle distance from BWI to each destination and writes
  results to a styled XLSX.
