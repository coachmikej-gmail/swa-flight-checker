# SWA Flight Checker

A desktop GUI that pulls live one-way Southwest Airlines fares departing a
home airport of your choice, straight from southwest.com — no manual
copy/paste, no headless browser. Results show distance, miles-per-dollar
value, and price changes since your last check, and can be sorted, exported
to PDF/XLSX, and auto-refreshed on a timer.

> ⚠️ **Personal / educational use only.** See [Disclaimer](#disclaimer) below.

## Features

- Pick any Southwest home airport from a dropdown (118+ origins)
- One-click fetch of current one-way fares for every destination from that
  airport
- Distance (great-circle miles) and miles-per-dollar value for each route
- Click any column header to sort; click again to reverse
- Row highlighting (red/green) showing price changes since your last fetch
- Local price history: each route's price is saved to a small SQLite file
  next to the app (`price_history.db`) — only when it changes — and
  double-clicking a row opens a price-over-time chart
- Price watches: set a target price for any destination and get a Telegram
  message when the fare drops to or below it
- Optional auto-refresh on a timer you set, in minutes
- Export results to a formatted PDF table or an Excel workbook (with live
  formulas, not just hardcoded values)
- Portable — a single `.pyw` file, or build it into a standalone `.exe`

## Price watches & Telegram alerts

1. **Set up the bot (once):** *Alerts → Telegram Setup...* has step-by-step
   instructions. In short: message `@BotFather` in Telegram, send `/newbot`,
   paste the token it gives you, send your new bot any message, then click
   **Detect Chat ID**, **Send Test Message**, and **Save**.
   **To alert a channel or group** (e.g. one you share with a partner) instead
   of your private chat: add the bot to it — for a channel, as an
   administrator with permission to post messages — post a message there, then
   click **Detect Chat ID** and pick the channel from the list.
2. **Set a watch:** fetch fares, then right-click a row (or select it and use
   *Alerts → Set Price Watch for Selected Row...*) and enter your target price.
3. **See your watches:** *Alerts → View Price Watches...* lists each watch with
   its current price and status, and lets you edit or remove them.

A watch alerts when the price is **at or below** your target. It won't repeat
on every refresh: you're alerted again only if the price drops *further*, or
after it has gone back above your target and dipped again. A failed send is
retried on the next fetch.

The bot token, chat ID, and your watches are stored in `settings.json` next to
the app. The token is in plain text, so keep that file private — it's
git-ignored, so it won't be committed, but don't share the folder with it inside.

## Requirements

- Python 3.9+
- `openpyxl` and `reportlab` (`pip install openpyxl reportlab`)
- Tkinter (bundled with most Python installs on Windows/macOS; on Linux you
  may need `sudo apt install python3-tk` or your distro's equivalent)

## Running it

```bash
python swa_deals_gui.pyw
```

On Windows, double-clicking `swa_deals_gui.pyw` runs it with `pythonw.exe`
directly — no console window.

## Building a standalone .exe (Windows)

```bash
pip install pyinstaller
python -m PyInstaller "SWA Flight Checker.spec"
```

The finished exe lands in `dist/SWA Flight Checker.exe`. Exports (PDF/XLSX)
are saved next to wherever you run the exe from.

## How it works

Southwest's flight-deals page renders its fare cards from data embedded in
two of its own JS files (`stations-data` for airport coordinates,
`lfma.js` for live prices). This app fetches those same files directly with
plain HTTP requests and parses out the fare/coordinate data — no scheduled
flight search, no headless browser, no reverse-engineered private API calls.
See `CHANGELOG.md` for the full version history.

## Disclaimer

This is a personal hobby project, built for my own use in checking flight
deals more conveniently. A few things worth knowing before you use or share
it:

- **Not affiliated with, endorsed by, or supported by Southwest Airlines.**
  "Southwest," "Southwest Airlines," and any related marks belong to their
  respective owner.
- **Fares are informational only, not a substitute for booking on
  southwest.com.** Prices, availability, and dates change constantly and
  may be wrong or stale by the time you see them here — always confirm
  and book directly through Southwest.
- **Reads data Southwest publishes for its own public web page**, rather
  than a documented, supported API. That means it can break at any time if
  Southwest changes their site, and heavy or automated use could run afoul
  of southwest.com's Terms & Conditions. Use it for occasional personal
  checking, not for high-frequency polling, redistribution of the data, or
  any commercial purpose.
- **No warranty.** Provided as-is, with no guarantee of accuracy or
  availability. Use at your own risk.

If you're a Southwest representative and have concerns about this project,
please reach out (contact below) and I'll address them promptly.

## License

Released under the [MIT License](LICENSE) — the code itself is free to use,
modify, and redistribute. That's separate from the usage disclaimer above,
which is about how you use this tool against southwest.com, not about your
rights to the code.

## Contact

Mike Jacobson — coachmikej@gmail.com
