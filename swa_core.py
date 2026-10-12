"""
Shared, GUI-free logic for SWA Flight Checker: fetching Southwest fares, the
local price-history database, settings.json (Telegram config + price watches),
Telegram sending, and the price-watch alert rules.

Used by both swa_deals_gui.pyw (the app) and swa_watcher.py (the headless
background checker), so the two can't drift apart.
"""
import json
import math
import os
import re
import sqlite3
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import namedtuple
from contextlib import closing
from datetime import datetime

DEFAULT_ORIGIN = "BWI"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"}
STATIONS_DATA_URL = "https://www.southwest.com/swa-ui/bootstrap/dotcom-landing-page/1/data.js"
LFMA_DATA_URL = "https://www.southwest.com/swa-resources/scripts/landing-pages/lfma.js"

# When packaged with PyInstaller --onefile, __file__ points into a temp folder
# that's deleted on exit, so save next to the .exe itself instead.
if getattr(sys, "frozen", False):
    OUTPUT_DIR = os.path.dirname(sys.executable)
else:
    OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))


def output_path(origin_code, extension):
    """Per-origin output filename, e.g. BWI_Flight_Deals.pdf / ORD_Flight_Deals.xlsx
    — keeps exports for different home airports from overwriting each other."""
    return os.path.join(OUTPUT_DIR, f"{origin_code}_Flight_Deals.{extension}")


# Local price history (SQLite). Lives next to the script/exe like the exports.
HISTORY_DB_PATH = os.path.join(OUTPUT_DIR, "price_history.db")
HISTORY_TIME_FORMAT = "%Y-%m-%d %H:%M:%S"


def _open_history_db():
    conn = sqlite3.connect(HISTORY_DB_PATH)
    conn.execute(
        """CREATE TABLE IF NOT EXISTS price_history (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               origin TEXT NOT NULL,
               dest TEXT NOT NULL,
               recorded_at TEXT NOT NULL,
               price INTEGER NOT NULL,
               depart_date TEXT
           )"""
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_route ON price_history (origin, dest, id)")
    return conn


def record_prices(origin_code, rows, when=None):
    """Append each route's price to the history, but only if it differs from the
    last price stored for that route (or the route has no history yet), so the
    table grows with price *changes*, not with every fetch. Returns how many
    rows were written. `rows` are fetch_deals() rows."""
    stamp = (when or datetime.now()).strftime(HISTORY_TIME_FORMAT)
    with closing(_open_history_db()) as conn:
        last = dict(conn.execute(
            """SELECT p.dest, p.price FROM price_history p
               JOIN (SELECT dest, MAX(id) AS last_id FROM price_history
                     WHERE origin = ? GROUP BY dest) m ON p.id = m.last_id""",
            (origin_code,),
        ))
        changed = [
            (origin_code, code, stamp, price, date_obj.isoformat())
            for _rank, _city, _state, code, _dist, date_obj, price, _mpd in rows
            if last.get(code) != price
        ]
        with conn:  # one transaction
            conn.executemany(
                "INSERT INTO price_history (origin, dest, recorded_at, price, depart_date) VALUES (?, ?, ?, ?, ?)",
                changed,
            )
    return len(changed)


def load_price_history(origin_code, dest_code):
    """Oldest-first list of (recorded_at: datetime, price, depart_date: str|None)."""
    with closing(_open_history_db()) as conn:
        found = conn.execute(
            "SELECT recorded_at, price, depart_date FROM price_history "
            "WHERE origin = ? AND dest = ? ORDER BY id",
            (origin_code, dest_code),
        ).fetchall()
    return [(datetime.strptime(at, HISTORY_TIME_FORMAT), price, dep) for at, price, dep in found]


# --- Settings (settings.json): Telegram bot config + price watches -----------
# Holds the bot token, which is a secret, so it's git-ignored. Kept next to the
# script/exe like the exports and history database.
SETTINGS_PATH = os.path.join(OUTPUT_DIR, "settings.json")


def default_settings():
    return {"telegram": {"bot_token": "", "chat_id": ""}, "watches": []}


def load_settings():
    """Returns (settings, warning). A missing file is normal (fresh install). A
    file that isn't valid JSON is moved aside to settings.json.bad rather than
    silently overwritten, and `warning` says so."""
    settings = default_settings()
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        return settings, None
    except (json.JSONDecodeError, UnicodeDecodeError, OSError) as exc:
        try:
            os.replace(SETTINGS_PATH, SETTINGS_PATH + ".bad")
            moved = " (the old file was kept as settings.json.bad)"
        except OSError:
            moved = ""
        return settings, f"Couldn't read settings.json ({exc}); starting with empty settings{moved}."

    if not isinstance(data, dict):
        return settings, "settings.json isn't in the expected format; starting with empty settings."
    telegram = data.get("telegram")
    if isinstance(telegram, dict):
        settings["telegram"]["bot_token"] = str(telegram.get("bot_token") or "")
        settings["telegram"]["chat_id"] = str(telegram.get("chat_id") or "")
    for watch in data.get("watches") or []:
        # Skip malformed entries instead of failing the whole file.
        if (isinstance(watch, dict) and watch.get("origin") and watch.get("dest")
                and isinstance(watch.get("target"), (int, float)) and watch["target"] > 0):
            settings["watches"].append({
                "origin": str(watch["origin"]),
                "dest": str(watch["dest"]),
                "city": str(watch.get("city") or ""),
                "state": str(watch.get("state") or ""),
                "target": watch["target"],
                "created": watch.get("created"),
                "last_alert_price": watch.get("last_alert_price"),
                "last_alert_at": watch.get("last_alert_at"),
            })
    return settings, None


def save_settings(settings):
    """Write via a temp file + rename so a crash mid-write can't leave a
    half-written (unparseable) settings.json."""
    tmp_path = SETTINGS_PATH + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)
    # On Windows the rename fails with "Access is denied" if another process has the
    # target open at that instant (the app and the background watcher both use this
    # file, and antivirus scanners touch fresh files). That clears within moments,
    # so retry briefly before giving up.
    attempts = 8
    for attempt in range(1, attempts + 1):
        try:
            os.replace(tmp_path, SETTINGS_PATH)
            return
        except PermissionError:
            if attempt == attempts:
                raise
            time.sleep(0.05 * attempt)


def fmt_money(amount):
    return f"${amount:,.0f}" if float(amount).is_integer() else f"${amount:,.2f}"


# --- Telegram -------------------------------------------------------------------
TELEGRAM_API_URL = "https://api.telegram.org/bot{token}/{method}"


class TelegramError(Exception):
    """A Telegram call failed; the message is safe to show (never contains the token)."""


def telegram_call(token, method, params=None, timeout=10):
    """POST to the Telegram Bot API and return the `result` field. Telegram
    explains failures in the JSON body ({"ok": false, "description": ...}), e.g.
    "Unauthorized" for a bad token or "chat not found" for a wrong chat ID."""
    url = TELEGRAM_API_URL.format(token=token, method=method)
    request = urllib.request.Request(url, data=urllib.parse.urlencode(params or {}).encode("utf-8"))
    try:
        with urllib.request.urlopen(request, timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        try:
            description = json.loads(exc.read().decode("utf-8")).get("description")
        except (ValueError, OSError):
            description = None
        raise TelegramError(description or f"Telegram returned HTTP {exc.code}") from None
    except NETWORK_ERRORS as exc:
        # `from None` + our own message: keeps the request URL (which contains the token) out of any traceback.
        reason = getattr(exc, "reason", exc)
        raise TelegramError(f"Couldn't reach Telegram ({reason})") from None
    except ValueError:
        raise TelegramError("Telegram sent a reply that couldn't be read") from None
    if not body.get("ok"):
        raise TelegramError(body.get("description") or "Telegram rejected the request")
    return body.get("result")


def send_telegram(token, chat_id, text):
    telegram_call(token, "sendMessage", {"chat_id": chat_id, "text": text})


TELEGRAM_CHAT_KINDS = {"private": "private chat", "group": "group", "supergroup": "group", "channel": "channel"}


def detect_telegram_chats(token):
    """Every chat the bot has seen recently, newest first, as (chat_id, name, kind) —
    private chats, groups, and channels. That includes the "bot was added" event
    Telegram sends when you add the bot to a channel/group, so a freshly added
    channel shows up even before anyone posts in it."""
    updates = telegram_call(token, "getUpdates", {"limit": 100}) or []
    found = {}
    for update in reversed(updates):
        source = (update.get("message") or update.get("edited_message") or update.get("channel_post")
                  or update.get("edited_channel_post") or update.get("my_chat_member"))
        chat = (source or {}).get("chat") or {}
        if chat.get("id") is None or str(chat["id"]) in found:
            continue
        name = chat.get("title") or " ".join(
            part for part in (chat.get("first_name"), chat.get("last_name")) if part
        ) or chat.get("username") or "chat"
        found[str(chat["id"])] = (str(chat["id"]), name, TELEGRAM_CHAT_KINDS.get(chat.get("type"), "chat"))
    if not found:
        raise TelegramError(
            "No chats found. Send your bot a message (e.g. \"hi\"). For a channel or group, add the "
            "bot to it (channels: as an administrator who can post messages) and post a message there. "
            "Then click Detect again."
        )
    return list(found.values())


def booking_url(origin, dest, depart_date):
    """Deep link to Southwest's booking page for this route/date — the same URL
    format the flight-deals page's own \"Book now\" links use."""
    query = urllib.parse.urlencode({
        "promoCode": "", "adultPassengersCount": 1, "lapInfantPassengersCount": 0,
        "departureDate": depart_date.isoformat(), "destinationAirportCode": dest,
        "fareType": "USD", "originationAirportCode": origin, "passengerType": "ADULT",
        "returnDate": "", "tripType": "oneway", "departureTimeOfDay": "ALL_DAY",
        "reset": "true", "returnTimeOfDay": "ALL_DAY",
    })
    return f"https://www.southwest.com/air/booking/select-depart.html?{query}"


def format_alert(watch, row):
    """Plain-text Telegram message (no Markdown, so city names with ()_ can't break formatting)."""
    _rank, city, state, code, _dist, depart_date, price, _mpd = row
    return (
        f"✈️ SWA price alert\n"
        f"{watch['origin']} → {code}  ({city}, {state})\n"
        f"Now {fmt_money(price)} one-way — at or below your {fmt_money(watch['target'])} watch\n"
        f"Departing {fmt_date(depart_date)}\n"
        f"{booking_url(watch['origin'], code, depart_date)}"
    )


# --- Price-watch alert rules (shared by the app and the background watcher) ----
AlertPlan = namedtuple("AlertPlan", "jobs unsent_hits rearmed")


def watch_key(watch):
    return (watch["origin"], watch["dest"])


def telegram_is_configured(settings):
    cfg = settings["telegram"]
    return bool(cfg["bot_token"] and cfg["chat_id"])


def plan_watch_alerts(settings, origin_code, rows, pending=()):
    """Decide which watches should alert for a fresh fetch of `origin_code`.

    Rules, so a watch doesn't spam you on every refresh:
    - price at or below the target and no alert sent yet -> alert
    - already alerted -> alert again only if the price drops *below* the last alerted price
    - price back above the target -> re-arm (the next dip alerts again)
    A watch whose send failed (or with Telegram not set up) isn't marked as alerted,
    so it's retried on the next fetch. `pending` holds keys with a send already in flight.

    Re-arming is applied to the in-memory `settings` here; the caller persists it
    (the keys are returned in `.rearmed`). Returns AlertPlan(jobs, unsent_hits,
    rearmed): jobs are (key, message_text, price) to send; unsent_hits are
    "ATL $79" strings for watches that hit but couldn't be sent (no Telegram)."""
    by_code = {row[3]: row for row in rows}
    configured = telegram_is_configured(settings)
    jobs, unsent_hits, rearmed = [], [], []
    for watch in settings["watches"]:
        if watch["origin"] != origin_code or watch["dest"] not in by_code:
            continue
        row, key = by_code[watch["dest"]], watch_key(watch)
        price, last = row[6], watch.get("last_alert_price")
        if price > watch["target"]:
            if last is not None:
                watch["last_alert_price"] = watch["last_alert_at"] = None
                rearmed.append(key)
            continue
        if key in pending or (last is not None and price >= last):
            continue
        if configured:
            jobs.append((key, format_alert(watch, row), price))
        else:
            unsent_hits.append(f"{watch['dest']} {fmt_money(price)}")
    return AlertPlan(jobs, unsent_hits, rearmed)


def mark_alert_sent(watch, price, when=None):
    watch["last_alert_price"] = price
    watch["last_alert_at"] = (when or datetime.now()).strftime(HISTORY_TIME_FORMAT)


def sync_alert_state(settings):
    """Pull each watch's alert state (last_alert_*) from settings.json into the
    in-memory `settings`. The background watcher updates that file while the app
    may be open; without this, the app could repeat an alert the watcher just sent."""
    on_disk, _warning = load_settings()
    saved = {watch_key(w): (w["last_alert_price"], w["last_alert_at"]) for w in on_disk["watches"]}
    for watch in settings["watches"]:
        if watch_key(watch) in saved:
            watch["last_alert_price"], watch["last_alert_at"] = saved[watch_key(watch)]


def apply_alert_changes(changes):
    """Write alert-state changes {key: (last_alert_price, last_alert_at)} into
    settings.json by re-reading it first, so edits made elsewhere in the meantime
    (e.g. a watch added in the app) aren't overwritten by a stale copy."""
    settings, _warning = load_settings()
    for watch in settings["watches"]:
        if watch_key(watch) in changes:
            watch["last_alert_price"], watch["last_alert_at"] = changes[watch_key(watch)]
    save_settings(settings)


# Floor on the auto-refresh interval so a typo (or "0") can't hammer southwest.com.
MIN_AUTO_REFRESH_MINUTES = 1


FETCH_ATTEMPTS = 4
FETCH_RETRY_DELAY_SECONDS = 2

NETWORK_ERRORS = (urllib.error.URLError, TimeoutError, ConnectionError)
# How long the GUI waits before quietly retrying after a connection failure.
# Right after standby or login, Wi-Fi/VPN/DNS can take 20-60s to come back, which
# is longer than fetch()'s own short retries cover.
NETWORK_RETRY_SECONDS = 30


def is_network_error(exc):
    """True for connection-level failures (DNS, no route, timeout) as opposed to
    the server answering with an HTTP error or us failing to parse its data."""
    return isinstance(exc, NETWORK_ERRORS) and not isinstance(exc, urllib.error.HTTPError)


def fetch(url):
    """GET a URL, retrying connection-level failures (DNS lookup errors like
    getaddrinfo 11001, dropped Wi-Fi/VPN, timeouts, resets). Those are
    usually momentary, so a short pause and retry clears them. HTTP errors
    (404, 403, ...) mean the server answered, so they're raised right away."""
    req = urllib.request.Request(url, headers=HEADERS)
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return resp.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError:
            raise
        except NETWORK_ERRORS:
            if attempt == FETCH_ATTEMPTS:
                raise
            time.sleep(FETCH_RETRY_DELAY_SECONDS * attempt)  # 2s, 4s, 6s


def fetch_station_directory():
    """Southwest ships airport reference data (lat/long, display name) as
    one module inside a large browserify bundle (many unrelated modules in
    one file). We don't need a JS engine to get at it — the module body is
    `module.exports = {...}` with plain JSON, so we locate it by module
    name and slice out the JSON between the `module.exports = ` marker and
    the bundle's closing `},{}]` for that module, then json.loads it
    directly. Returns code -> {"lat", "lon", "name"}."""
    js = fetch(STATIONS_DATA_URL)
    marker = '"swa-bootstrap-dotcom-landing-page/air/stations-data"'
    idx = js.index(marker)
    start = js.index("module.exports = ", idx) + len("module.exports = ")
    end = js.index("\n},{}]", start)
    stations = json.loads(js[start:end].rstrip().rstrip(";"))
    return {
        s["id"]: {
            "lat": float(s["latitude"]),
            "lon": float(s["longitude"]),
            "name": s.get("displayName") or s["id"],
        }
        for s in stations
        # A handful of entries (e.g. code-share-only stations) have no
        # coordinates published; skip them rather than crash on float(None).
        if s.get("latitude") is not None and s.get("longitude") is not None
    }


def fetch_fares(origin_code):
    """lfma.js ("Low Fare Market Ad") is the data file behind Southwest's
    own flight-deals page — it's regenerated periodically with the current
    live fares baked in as plain JS assignments like
    priceHash['BWI-MCO']='183'. We regex those straight out instead of
    executing the file, which also means we get every route Southwest has
    fare data for, not just the subset its own widget chooses to render.

    Returns (fares, origin_label): fares is a list of
    (dest_name, dest_code, departing_date, price) for routes departing
    origin_code, and origin_label is origin_code's own "City, ST" name
    (pulled from the same stationNameHash lfma.js uses for destinations)."""
    js = fetch(LFMA_DATA_URL)
    prices = dict(re.findall(r"priceHash\['([A-Z]+-[A-Z]+)'\]='(\d+)'", js))
    dates = dict(re.findall(r"departDateHash\['([A-Z]+-[A-Z]+)'\]='([\d/]+)'", js))
    names = dict(re.findall(r"stationNameHash\['([A-Z]+)'\]=\"(.*?)\";", js))

    fares = []
    for route, price in prices.items():
        origin, dest = route.split("-")
        if origin != origin_code or route not in dates or dest not in names:
            continue
        # Keep the year (not just "M/D") — needed so dates that roll over
        # into the next calendar year still sort chronologically later.
        date_obj = datetime.strptime(dates[route], "%m/%d/%Y").date()
        fares.append((names[dest], dest, date_obj, int(price)))
    return fares, names.get(origin_code, origin_code)


def fetch_origin_airports():
    """Airport codes Southwest treats as an origin market (i.e. you can pick
    them as a departure city), paired with a human-readable "City, ST"
    name, sorted alphabetically by that name. The code list comes from
    lfma.js's originArray; the names come from fetch_station_directory()."""
    js = fetch(LFMA_DATA_URL)
    match = re.search(r"originArray=\[(.*?)\];", js)
    if not match:
        raise RuntimeError("Could not find originArray in Southwest's data feed — its page structure may have changed.")
    codes = re.findall(r"'([A-Z]+)'", match.group(1))

    directory = fetch_station_directory()
    airports = [(code, directory[code]["name"]) for code in codes if code in directory]
    airports.sort(key=lambda pair: pair[1])
    return airports


def haversine_miles(lat1, lon1, lat2, lon2):
    """Great-circle (as-the-crow-flies) distance in miles, not flown route
    miles — actual flights zigzag around weather/traffic and are longer."""
    R = 3958.8  # mean Earth radius, miles
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlambda / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def fmt_date(d):
    return f"{d.month}/{d.day:02d}"


def split_destination(name):
    """Southwest's station names are "City, ST" (domestic) or "City, Country"
    (international) — split on the last comma into (city, state_or_country)."""
    city, _sep, region = name.rpartition(",")
    return (city.strip(), region.strip()) if _sep else (name.strip(), "")


def fetch_deals(origin_code):
    """Returns (origin_label, rows) for flights departing origin_code.
    rows is a list of (rank, city, state, airport, distance_mi,
    departing_date, price, miles_per_dollar), sorted by price ascending.
    rank is each row's fixed price rank (1 = cheapest) and stays attached
    to the row even if it's later displayed in a different sort order."""
    directory = fetch_station_directory()
    if origin_code not in directory:
        raise RuntimeError(f"{origin_code!r} isn't a station Southwest publishes coordinates for.")

    fares, origin_label = fetch_fares(origin_code)
    if not fares:
        raise RuntimeError(
            f"No fares found for {origin_code} — it may not be a Southwest "
            f"origin market, or southwest.com's page structure may have changed."
        )

    origin = directory[origin_code]
    rows = []
    for dest_name, code, date_obj, price in fares:
        if code not in directory:
            continue
        dest = directory[code]
        dist = round(haversine_miles(origin["lat"], origin["lon"], dest["lat"], dest["lon"]))
        city, state = split_destination(dest_name)
        rows.append((city, state, code, dist, date_obj, price, dist / price))

    rows.sort(key=lambda r: r[5])
    ranked = [
        (rank, city, state, code, dist, date_obj, price, mpd)
        for rank, (city, state, code, dist, date_obj, price, mpd) in enumerate(rows, start=1)
    ]
    return origin_label, ranked
