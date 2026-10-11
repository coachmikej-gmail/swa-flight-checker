"""
GUI for fetching live Southwest Airlines one-way flight deals from a chosen
home airport (BWI by default) and exporting them as a PDF table or an Excel
workbook.

Data is fetched live from southwest.com (see fetch_station_directory and
fetch_fares) the same way build_bwi_deals.py does it for BWI specifically.

See CHANGELOG.md for the version history of this file.
"""
import json
import math
import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.request
from datetime import datetime
from tkinter import messagebox, ttk

# Shown in the Help > About box — bump APP_VERSION alongside CHANGELOG.md entries.
APP_NAME = "SWA Flight Checker"
APP_VERSION = "1.11.0"
APP_DESCRIPTION = (
    "Fetches live one-way Southwest Airlines flight deals departing your "
    "chosen home airport directly from southwest.com, with distance and "
    "miles-per-dollar stats, sortable results, PDF/XLSX export, and "
    "optional auto-refresh."
)
AUTHOR_NAME = "Mike Jacobson"
AUTHOR_EMAIL = "coachmikej@gmail.com"

# Treeview column -> index into a display row tuple (rank, city, state, code, dist, date, price, mpd).
# Keep this in sync with the tuple shape returned by fetch_deals().
COLUMN_INDEX = {"rank": 0, "city": 1, "state": 2, "airport": 3, "distance": 4, "departing": 5, "price": 6, "mpd": 7}
# Columns sorted as case-insensitive text; everything else sorts numerically/chronologically.
TEXT_COLUMNS = {"city", "state", "airport"}

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

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


def export_pdf(rows, path, origin_code, origin_label):
    doc = SimpleDocTemplate(
        path,
        pagesize=landscape(letter),
        leftMargin=0.5 * inch, rightMargin=0.5 * inch,
        topMargin=0.5 * inch, bottomMargin=0.5 * inch,
    )
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(
        f"Southwest Airlines &mdash; One-Way Flight Deals from {origin_label} ({origin_code})",
        styles["Title"],
    ))
    captured = datetime.now().strftime("%m/%d/%Y %H:%M")
    story.append(Paragraph(
        f"Live data fetched from southwest.com &mdash; captured {captured}. "
        f"Distances are great-circle (straight-line) miles between airports, not flown route miles.",
        styles["Normal"],
    ))
    story.append(Spacer(1, 0.2 * inch))

    table_data = [["Rank", "City", "State", "Airport", f"Distance from {origin_code} (mi)", "Departing", "One-Way Price ($)", "Miles per $"]]
    for rank, city, state, code, dist, date_obj, price, mpd in rows:
        table_data.append([str(rank), city, state, code, f"{dist:,}", fmt_date(date_obj), f"${price:,}", f"{mpd:.2f}"])

    table = Table(
        table_data,
        colWidths=[0.5 * inch, 1.9 * inch, 0.6 * inch, 0.7 * inch, 1.3 * inch, 0.9 * inch, 1.0 * inch, 0.8 * inch],
        repeatRows=1,
    )
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#004C97")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (2, 0), (3, -1), "CENTER"),
        ("ALIGN", (4, 0), (-1, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(table)

    story.append(Spacer(1, 0.25 * inch))
    prices = [r[6] for r in rows]
    dists = [r[4] for r in rows]
    mpds = [r[7] for r in rows]
    summary = (
        f"Destinations: {len(rows)} &nbsp;|&nbsp; "
        f"Cheapest: ${min(prices):,} &nbsp;|&nbsp; Most expensive: ${max(prices):,} &nbsp;|&nbsp; "
        f"Average fare: ${sum(prices) / len(prices):,.2f} &nbsp;|&nbsp; "
        f"Average distance: {sum(dists) / len(dists):,.0f} mi &nbsp;|&nbsp; "
        f"Average miles per $: {sum(mpds) / len(mpds):,.2f}"
    )
    story.append(Paragraph(summary, styles["Normal"]))

    doc.build(story)


def export_xlsx(rows, path, origin_code, origin_label):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{origin_code} Flight Deals"

    font_name = "Arial"
    header_font = Font(name=font_name, size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="004C97", end_color="004C97", fill_type="solid")
    normal_font = Font(name=font_name, size=11)
    title_font = Font(name=font_name, size=14, bold=True)
    note_font = Font(name=font_name, size=9, italic=True, color="666666")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    num_cols = 8

    ws["A1"] = f"Southwest Airlines — One-Way Flight Deals from {origin_label} ({origin_code})"
    ws["A1"].font = title_font
    ws.merge_cells("A1:H1")

    ws["A2"] = (
        f"Live data fetched from southwest.com — captured {datetime.now().strftime('%m/%d/%Y %H:%M')}. "
        f"Distances are great-circle (straight-line) miles between airports, not flown route miles."
    )
    ws["A2"].font = note_font
    ws.merge_cells("A2:H2")

    headers = ["City", "State", "Airport", f"Distance from {origin_code} (mi)", "Departing", "One-Way Price ($)", "Miles per $", "Rank (Cheapest = 1)"]
    header_row = 4
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border

    start_row = header_row + 1
    for i, (_rank, city, state, code, dist, date_obj, price, _mpd) in enumerate(rows):
        r = start_row + i
        ws.cell(row=r, column=1, value=city).font = normal_font
        ws.cell(row=r, column=2, value=state).font = normal_font
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3, value=code).font = normal_font
        ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")
        dist_cell = ws.cell(row=r, column=4, value=dist)
        dist_cell.font = normal_font
        dist_cell.number_format = "#,##0"
        dist_cell.alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=5, value=fmt_date(date_obj)).font = normal_font
        ws.cell(row=r, column=5).alignment = Alignment(horizontal="center")
        price_cell = ws.cell(row=r, column=6, value=price)
        price_cell.font = normal_font
        price_cell.number_format = "$#,##0"
        price_cell.alignment = Alignment(horizontal="center")
        # Live formulas, not the Python-computed values, so the sheet stays
        # correct if someone edits distance/price cells by hand in Excel.
        mpd_cell = ws.cell(row=r, column=7, value=f"=D{r}/F{r}")  # Distance / Price
        mpd_cell.font = normal_font
        mpd_cell.number_format = "0.00"
        mpd_cell.alignment = Alignment(horizontal="center")
        # RANK() reads the whole price column, so it's correct regardless of
        # which order the rows were written in (e.g. exported while sorted
        # by distance) — it doesn't depend on this row's position.
        rank_cell = ws.cell(row=r, column=8, value=f"=RANK(F{r},$F${start_row}:$F${start_row + len(rows) - 1},1)")
        rank_cell.font = normal_font
        rank_cell.alignment = Alignment(horizontal="center")
        for col in range(1, num_cols + 1):
            ws.cell(row=r, column=col).border = border
        if i % 2 == 1:
            for col in range(1, num_cols + 1):
                ws.cell(row=r, column=col).fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    end_row = start_row + len(rows) - 1

    summary_row = end_row + 2
    ws.cell(row=summary_row, column=1, value="Summary").font = Font(name=font_name, size=11, bold=True)
    labels = ["Number of destinations", "Cheapest fare ($)", "Most expensive fare ($)", "Average fare ($)",
              "Shortest distance (mi)", "Longest distance (mi)", "Average distance (mi)",
              "Average miles per $"]
    formulas = [
        f"=COUNTA(A{start_row}:A{end_row})",
        f"=MIN(F{start_row}:F{end_row})",
        f"=MAX(F{start_row}:F{end_row})",
        f"=AVERAGE(F{start_row}:F{end_row})",
        f"=MIN(D{start_row}:D{end_row})",
        f"=MAX(D{start_row}:D{end_row})",
        f"=AVERAGE(D{start_row}:D{end_row})",
        f"=AVERAGE(G{start_row}:G{end_row})",
    ]
    for i, (label, formula) in enumerate(zip(labels, formulas)):
        r = summary_row + 1 + i
        ws.cell(row=r, column=1, value=label).font = normal_font
        fc = ws.cell(row=r, column=2, value=formula)
        fc.font = Font(name=font_name, size=11, bold=True)
        if "fare" in label:
            fc.number_format = "$#,##0.00" if "Average" in label else "$#,##0"
        elif "distance" in label:
            fc.number_format = "#,##0"
        elif "miles per $" in label:
            fc.number_format = "0.00"

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 12
    ws.column_dimensions["H"].width = 18
    ws.row_dimensions[header_row].height = 30
    ws.freeze_panes = f"A{start_row}"

    wb.save(path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        # Wide enough for all 9 result columns (their widths sum to ~960px)
        # plus the scrollbar and window padding.
        self.geometry("1100x560")

        menubar = tk.Menu(self)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.config(menu=menubar)

        self.rows = []
        self.display_rows = []
        self.sort_column = None
        self.sort_reverse = False
        self.price_changes = {}  # airport code -> "up" / "down", vs. the previous fetch
        self.price_deltas = {}   # airport code -> dollar change vs. the previous fetch
        self.result_queue = queue.Queue()
        self.fetching = False
        self.auto_refresh_active = False
        self.auto_refresh_job = None  # after() id for the next scheduled tick, so it can be cancelled
        self.auto_refresh_ms = None

        self.origin_code = DEFAULT_ORIGIN
        self.origin_label = ""
        self.previous_origin_code = None  # last-fetched origin, so price-change diffs don't cross origins
        self.origin_lookup = {}  # dropdown display string -> airport code
        self.origins_queue = queue.Queue()
        self.fetch_is_auto = False  # True when the running fetch was started by auto-refresh
        self.network_retry_job = None  # after() id of a pending quiet retry (fares)
        self.origins_retry_job = None  # after() id of a pending retry (airport list)
        self.last_pdf_path = None
        self.last_xlsx_path = None

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        origin_frame = ttk.Frame(self, padding=(10, 0))
        origin_frame.pack(fill="x")
        ttk.Label(origin_frame, text="Home airport:").pack(side="left")
        # Bootstrapped with just BWI so the dropdown is usable immediately;
        # start_load_origins() fills in the full list in the background.
        bootstrap_label = f"{DEFAULT_ORIGIN} — Baltimore/Washington, MD"
        self.origin_lookup[bootstrap_label] = DEFAULT_ORIGIN
        self.origin_var = tk.StringVar(value=bootstrap_label)
        self.origin_combo = ttk.Combobox(
            origin_frame, textvariable=self.origin_var, state="readonly",
            width=40, values=[bootstrap_label],
        )
        self.origin_combo.pack(side="left", padx=(4, 0))
        self.origin_combo.bind("<<ComboboxSelected>>", self.on_origin_selected)

        self.fetch_btn = ttk.Button(top, text="Fetch Latest Deals", command=self.start_fetch)
        self.fetch_btn.pack(side="left")

        self.export_pdf_btn = ttk.Button(top, text="Export to PDF", command=self.do_export_pdf, state="disabled")
        self.export_pdf_btn.pack(side="left", padx=(8, 0))

        self.open_pdf_btn = ttk.Button(top, text="Open PDF", command=self.open_pdf, state="disabled")
        self.open_pdf_btn.pack(side="left", padx=(8, 0))

        self.export_xlsx_btn = ttk.Button(top, text="Export to XLSX", command=self.do_export_xlsx, state="disabled")
        self.export_xlsx_btn.pack(side="left", padx=(8, 0))

        self.open_xlsx_btn = ttk.Button(top, text="Open XLSX", command=self.open_xlsx, state="disabled")
        self.open_xlsx_btn.pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(top, mode="indeterminate", length=150)
        self.progress.pack(side="right")

        refresh = ttk.Frame(self, padding=(10, 0))
        refresh.pack(fill="x")
        ttk.Label(refresh, text="Auto-refresh every:").pack(side="left")
        # Blank by default — auto-refresh stays off until a value is entered
        # and "Start Auto-Refresh" is clicked.
        self.interval_var = tk.StringVar(value="")
        self.interval_entry = ttk.Entry(refresh, textvariable=self.interval_var, width=6)
        self.interval_entry.pack(side="left", padx=(4, 4))
        ttk.Label(refresh, text="minutes").pack(side="left")
        self.auto_refresh_btn = ttk.Button(refresh, text="Start Auto-Refresh", command=self.toggle_auto_refresh)
        self.auto_refresh_btn.pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="Click \"Fetch Latest Deals\" to load current fares from southwest.com.")
        # Packed with side="bottom" *before* the tree so it reserves the
        # bottom strip of the window; the expanding tree then fills the rest.
        ttk.Label(self, textvariable=self.status_var, padding=(10, 4)).pack(side="bottom", fill="x")

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        columns = ("rank", "city", "state", "airport", "distance", "departing", "price", "change", "mpd")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        self.column_labels = {
            "rank": ("Rank", 50),
            "city": ("City", 190),
            "state": ("State", 60),
            "airport": ("Airport", 70),
            "distance": (f"Distance from {DEFAULT_ORIGIN} (mi)", 150),
            "departing": ("Departing", 90),
            "price": ("One-Way Price ($)", 120),
            "change": ("Change vs. Last Fetch", 140),
            "mpd": ("Miles per $", 90),
        }
        for col, (label, width) in self.column_labels.items():
            self.tree.heading(col, text=label, command=lambda c=col: self.sort_by(c))
            anchor = "w" if col == "city" else "center"
            self.tree.column(col, width=width, anchor=anchor)

        self.tree.tag_configure("price_up", background="#f8d7da")    # light red
        self.tree.tag_configure("price_down", background="#d4edda")  # light green

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        vsb.pack(side="right", fill="y", pady=10)

        self.start_load_origins()

    def start_load_origins(self):
        threading.Thread(target=self._load_origins_worker, daemon=True).start()
        self.after(100, self._poll_origins_queue)

    def _load_origins_worker(self):
        try:
            airports = fetch_origin_airports()
            self.origins_queue.put(("ok", airports))
        except Exception as exc:
            kind = "network_error" if is_network_error(exc) else "error"
            self.origins_queue.put((kind, str(exc)))

    def _poll_origins_queue(self):
        try:
            status, payload = self.origins_queue.get_nowait()
        except queue.Empty:
            self.after(200, self._poll_origins_queue)
            return

        if status == "network_error":
            # Typically the network isn't back yet after standby/login — keep
            # trying quietly so the dropdown fills in without a restart.
            self.status_var.set(
                f"Couldn't load the full airport list ({payload}); showing BWI only. "
                f"Retrying in {NETWORK_RETRY_SECONDS}s..."
            )
            self.origins_retry_job = self.after(NETWORK_RETRY_SECONDS * 1000, self.start_load_origins)
            return
        if status == "error":
            # Non-fatal — the BWI-only bootstrap entry still works fine.
            self.status_var.set(f"Couldn't load the full airport list ({payload}); showing BWI only.")
            return

        lookup = {}
        for code, name in payload:
            # Southwest's display names are "City, ST - CODE"; drop the
            # " - CODE" suffix since the code is already shown up front.
            city_state = name.rsplit(" - ", 1)[0] if " - " in name else name
            lookup[f"{code} — {city_state}"] = code

        self.origin_lookup = lookup
        labels = sorted(lookup)
        self.origin_combo.config(values=labels)

        # Keep whatever's currently selected (still just BWI at this point,
        # unless the user was fast) pointed at the right dropdown entry.
        current = next((label for label, code in lookup.items() if code == self.origin_code), None)
        if current:
            self.origin_var.set(current)

        self.status_var.set(f"Loaded {len(labels)} home airports. Click \"Fetch Latest Deals\" to load current fares from southwest.com.")

    def on_origin_selected(self, _event=None):
        label = self.origin_var.get()
        self.origin_code = self.origin_lookup.get(label, label.split("—", 1)[0].strip())

    def start_fetch(self, auto=False):
        self._cancel_network_retry()  # a fetch is starting now, so any pending retry is redundant
        self.fetching = True
        self.fetch_is_auto = auto
        self.fetch_btn.config(state="disabled")
        self.export_pdf_btn.config(state="disabled")
        self.open_pdf_btn.config(state="disabled")
        self.export_xlsx_btn.config(state="disabled")
        self.open_xlsx_btn.config(state="disabled")
        self.status_var.set("Fetching live fares from southwest.com...")
        self.progress.start(12)
        threading.Thread(target=self._fetch_worker, daemon=True).start()
        self.after(100, self._poll_queue)

    def _fetch_worker(self):
        # Runs on a background thread — network I/O here so the Tk mainloop
        # (and the window) stays responsive. Tkinter widgets aren't
        # thread-safe, so this thread must never touch self.tree/self.rows
        # etc. directly; it only hands the result to _poll_queue via the
        # queue, which runs back on the main thread.
        origin_code = self.origin_code  # snapshot: whatever's selected right now
        try:
            origin_label, rows = fetch_deals(origin_code)
            self.result_queue.put(("ok", origin_code, origin_label, rows))
        except Exception as exc:
            kind = "network_error" if is_network_error(exc) else "error"
            self.result_queue.put((kind, str(exc)))

    def _cancel_network_retry(self):
        if self.network_retry_job is not None:
            self.after_cancel(self.network_retry_job)
            self.network_retry_job = None

    def _network_retry(self):
        self.network_retry_job = None
        if self.auto_refresh_active and not self.fetching:
            self.start_fetch(auto=True)

    def _poll_queue(self):
        # Polls (rather than blocking) so the mainloop keeps servicing the
        # UI while we wait; reschedules itself every 100ms until the
        # worker thread above has put something in the queue.
        try:
            result = self.result_queue.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_queue)
            return

        self.progress.stop()
        self.fetch_btn.config(state="normal")
        self.fetching = False

        if result[0] == "network_error":
            detail = result[1]
            if self.fetch_is_auto and self.auto_refresh_active:
                # Unattended refresh (often firing the instant the PC wakes from
                # standby, before Wi-Fi/DNS is back): no popup, just say so and
                # try again shortly. The table keeps showing the last good data.
                self.status_var.set(
                    f"Couldn't reach southwest.com ({detail}) — network may still be "
                    f"reconnecting. Retrying in {NETWORK_RETRY_SECONDS}s..."
                )
                self.network_retry_job = self.after(NETWORK_RETRY_SECONDS * 1000, self._network_retry)
                return
            self.status_var.set(f"Couldn't reach southwest.com ({detail})")
            messagebox.showerror(
                "Fetch failed",
                "Couldn't reach southwest.com. Check your internet connection — it can "
                "take a little while to come back after standby or login — then try again."
                f"\n\n{detail}",
            )
            return
        if result[0] == "error":
            message = result[1]
            self.status_var.set(f"Error: {message}")
            messagebox.showerror("Fetch failed", message)
            return

        _, origin_code, origin_label, payload = result

        # Snapshot the OLD self.rows before we overwrite it below — this is
        # the "previous fetch" we compare the new prices against. On the
        # very first fetch (or after switching to a different home airport)
        # there's nothing comparable, so previous_prices stays empty and
        # nothing gets highlighted.
        previous_prices = {}
        if self.previous_origin_code == origin_code:
            previous_prices = {row[3]: row[6] for row in self.rows}  # airport code -> price

        self.origin_code = origin_code
        self.origin_label = origin_label
        self.previous_origin_code = origin_code
        self.column_labels["distance"] = (f"Distance from {origin_code} (mi)", self.column_labels["distance"][1])
        self.rows = payload
        self.display_rows = list(payload)
        self.sort_column = None
        self.sort_reverse = False

        # Diff by airport code (not by row position — sort order/rank can
        # differ between fetches) and only flag actual increases/decreases.
        # Codes with no prior price (new to this fetch) are left untouched
        # in price_changes, which populate_tree treats as "no highlight".
        self.price_changes = {}
        self.price_deltas = {}  # airport code -> new price minus previous price (only codes seen last fetch)
        up = down = same = 0
        for row in self.rows:
            code, price = row[3], row[6]
            if code not in previous_prices:
                continue
            self.price_deltas[code] = price - previous_prices[code]
            if price > previous_prices[code]:
                self.price_changes[code] = "up"
                up += 1
            elif price < previous_prices[code]:
                self.price_changes[code] = "down"
                down += 1
            else:
                same += 1

        for col, (label, _width) in self.column_labels.items():
            self.tree.heading(col, text=label)
        self.populate_tree()
        self.export_pdf_btn.config(state="normal")
        self.export_xlsx_btn.config(state="normal")

        timestamp = datetime.now().strftime("%m/%d/%Y %H:%M")
        status = f"Loaded {len(self.rows)} destinations from {origin_label} ({origin_code}) as of {timestamp}."
        if previous_prices:
            status += f" Since last fetch: {up} up, {down} down, {same} unchanged."
        self.status_var.set(status)

    def populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        for rank, city, state, code, dist, date_obj, price, mpd in self.display_rows:
            change = self.price_changes.get(code)
            tags = (f"price_{change}",) if change else ()
            delta = self.price_deltas.get(code)
            if delta is None:
                delta_text = ""  # no previous price to compare against
            elif delta == 0:
                delta_text = "$0"
            else:
                delta_text = f"{'+' if delta > 0 else '-'}${abs(delta):,}"
            self.tree.insert(
                "", "end",
                values=(rank, city, state, code, f"{dist:,}", fmt_date(date_obj), f"${price:,}", delta_text, f"{mpd:.2f}"),
                tags=tags,
            )

    def sort_by(self, column):
        """Re-sorts self.display_rows in place (clicking the same header
        again reverses direction). This only changes display/export order
        — each row's Rank value was fixed at fetch time in fetch_deals()
        and is never recalculated here, so it keeps showing each
        destination's true price rank even when the table is sorted by,
        say, Distance instead of Price."""
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False

        idx = COLUMN_INDEX.get(column)
        if column == "change":
            # Deltas live outside the row tuples; rows with no previous price sort as 0.
            key = lambda row: self.price_deltas.get(row[3], 0)
        elif column in TEXT_COLUMNS:
            key = lambda row: str(row[idx]).lower()
        else:
            key = lambda row: row[idx]
        self.display_rows.sort(key=key, reverse=self.sort_reverse)

        arrow = " ▼" if self.sort_reverse else " ▲"
        for col, (label, _width) in self.column_labels.items():
            self.tree.heading(col, text=label + (arrow if col == column else ""))

        self.populate_tree()

    def do_export_pdf(self):
        path = output_path(self.origin_code, "pdf")
        try:
            export_pdf(self.display_rows, path, self.origin_code, self.origin_label)
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))
            return
        self.last_pdf_path = path
        self.open_pdf_btn.config(state="normal")
        self.status_var.set(f"PDF saved to {path}")
        messagebox.showinfo("Export complete", f"PDF saved to:\n{path}")

    def do_export_xlsx(self):
        path = output_path(self.origin_code, "xlsx")
        try:
            export_xlsx(self.display_rows, path, self.origin_code, self.origin_label)
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))
            return
        self.last_xlsx_path = path
        self.open_xlsx_btn.config(state="normal")
        self.status_var.set(f"XLSX saved to {path}")
        messagebox.showinfo("Export complete", f"XLSX saved to:\n{path}")

    def open_pdf(self):
        os.startfile(self.last_pdf_path)

    def open_xlsx(self):
        os.startfile(self.last_xlsx_path)

    def show_about(self):
        messagebox.showinfo(
            f"About {APP_NAME}",
            f"{APP_NAME}\n"
            f"Version {APP_VERSION}\n\n"
            f"{APP_DESCRIPTION}\n\n"
            f"Contact: {AUTHOR_NAME} <{AUTHOR_EMAIL}>",
        )

    def toggle_auto_refresh(self):
        if self.auto_refresh_active:
            self.stop_auto_refresh()
        else:
            self.start_auto_refresh()

    def start_auto_refresh(self):
        text = self.interval_var.get().strip()
        if not text:
            # Blank box means "leave auto-refresh off" — not an error, just a no-op.
            self.status_var.set("Enter a number of minutes to enable auto-refresh.")
            return

        try:
            minutes = float(text)
        except ValueError:
            messagebox.showerror("Invalid interval", f'"{text}" isn\'t a number.')
            return
        if minutes < MIN_AUTO_REFRESH_MINUTES:
            messagebox.showerror(
                "Invalid interval",
                f"Enter {MIN_AUTO_REFRESH_MINUTES} minute(s) or more.",
            )
            return

        self.auto_refresh_ms = int(minutes * 60 * 1000)
        self.auto_refresh_active = True
        self.interval_entry.config(state="disabled")
        self.auto_refresh_btn.config(text="Stop Auto-Refresh")
        self.status_var.set(f"Auto-refresh on: fetching every {minutes:g} minute(s).")

        # Fetch right away, then keep firing every interval from here on.
        if not self.fetching:
            self.start_fetch(auto=True)
        self.auto_refresh_job = self.after(self.auto_refresh_ms, self._auto_refresh_tick)

    def stop_auto_refresh(self):
        if self.auto_refresh_job is not None:
            self.after_cancel(self.auto_refresh_job)
            self.auto_refresh_job = None
        self._cancel_network_retry()
        self.auto_refresh_active = False
        self.interval_entry.config(state="normal")
        self.auto_refresh_btn.config(text="Start Auto-Refresh")
        if not self.fetching:
            self.status_var.set("Auto-refresh off.")

    def _auto_refresh_tick(self):
        if not self.auto_refresh_active:
            return
        # If a fetch is still running (slow network, etc.) skip this tick
        # rather than piling up overlapping requests — the next one will
        # pick it up in another interval.
        if not self.fetching:
            self.start_fetch(auto=True)
        self.auto_refresh_job = self.after(self.auto_refresh_ms, self._auto_refresh_tick)

    def _on_close(self):
        if self.auto_refresh_job is not None:
            self.after_cancel(self.auto_refresh_job)
        self._cancel_network_retry()
        if self.origins_retry_job is not None:
            self.after_cancel(self.origins_retry_job)
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
