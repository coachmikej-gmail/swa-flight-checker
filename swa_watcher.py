"""
Headless price-watch checker for SWA Flight Checker.

Meant to be run on a schedule (see install_watcher_task.ps1) so your price
watches are checked even while the app is closed. Each run:

  1. reads settings.json (your Telegram config + price watches),
  2. fetches current fares for each home airport that has a watch,
  3. records any price changes in price_history.db,
  4. sends a Telegram message for each watch that has hit its target, using
     the same alert rules as the app (swa_core.plan_watch_alerts), and
  5. appends what it did to watcher.log.

Examples:
  python swa_watcher.py                            run once, right now
  python swa_watcher.py --dry-run                  show what it would send; change nothing
  python swa_watcher.py --only-between 6-23        skip quietly outside 6:00-23:59 (used by the scheduled task)
"""
import argparse
import json
import logging
import logging.handlers
import os
import sqlite3
import sys
import time
from datetime import datetime

import swa_core as core

LOG_PATH = os.path.join(core.OUTPUT_DIR, "watcher.log")
STATE_PATH = os.path.join(core.OUTPUT_DIR, "watcher_state.json")

# If every fetch fails this many runs in a row, send one Telegram heads-up (so a
# broken watcher doesn't fail silently), plus one message when it recovers.
FAILURES_BEFORE_NOTICE = 3
# A freshly woken PC often has no network for a few seconds, so each home airport's
# fetch is retried (on top of swa_core.fetch's own short retries) before giving up.
DEFAULT_NET_ATTEMPTS = 4
DEFAULT_NET_WAIT_SECONDS = 30

log = logging.getLogger("swa_watcher")


def setup_logging():
    log.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S")
    file_handler = logging.handlers.RotatingFileHandler(LOG_PATH, maxBytes=200_000, backupCount=2, encoding="utf-8")
    file_handler.setFormatter(formatter)
    log.addHandler(file_handler)
    if sys.stderr is not None:  # None under pythonw, which is how the scheduled task runs
        console = logging.StreamHandler()
        console.setFormatter(formatter)
        log.addHandler(console)


def parse_hours(text):
    """'6-23' -> (6, 23): run only during hours 6 through 23 (06:00 until midnight)."""
    try:
        start, end = (int(part) for part in text.split("-"))
    except ValueError:
        raise argparse.ArgumentTypeError("use START-END hours, e.g. 6-23") from None
    if not 0 <= start <= end <= 23:
        raise argparse.ArgumentTypeError("hours must satisfy 0 <= START <= END <= 23")
    return start, end


def load_state():
    state = {"consecutive_failures": 0, "failure_notified": False}
    try:
        with open(STATE_PATH, encoding="utf-8") as f:
            state.update(json.load(f))
    except (OSError, ValueError):
        pass
    return state


def save_state(state):
    try:
        with open(STATE_PATH, "w", encoding="utf-8") as f:
            json.dump(state, f)
    except OSError as exc:
        log.warning("Couldn't save %s: %s", STATE_PATH, exc)


def fetch_with_patience(origin, attempts, wait_seconds):
    for attempt in range(1, attempts + 1):
        try:
            return core.fetch_deals(origin)
        except Exception as exc:
            if core.is_network_error(exc) and attempt < attempts:
                log.warning("Network problem fetching %s (attempt %d/%d): %s - retrying in %ds",
                            origin, attempt, attempts, exc, wait_seconds)
                time.sleep(wait_seconds)
                continue
            raise


def notify(settings, text):
    """Best-effort Telegram message; returns True if it was sent."""
    if not core.telegram_is_configured(settings):
        return False
    cfg = settings["telegram"]
    try:
        core.send_telegram(cfg["bot_token"], cfg["chat_id"], text)
        return True
    except Exception as exc:
        log.error("Couldn't send Telegram message: %s", exc)
        return False


def track_failures(settings, all_failed, detail):
    """Count consecutive fully-failed runs; message once when it becomes a streak, and once on recovery."""
    state = load_state()
    if all_failed:
        state["consecutive_failures"] += 1
        if state["consecutive_failures"] >= FAILURES_BEFORE_NOTICE and not state["failure_notified"]:
            runs = state["consecutive_failures"]
            if notify(settings, f"⚠️ SWA watcher: couldn't fetch fares in {runs} runs in a row ({detail}). "
                                f"Your price watches aren't being checked right now; it will keep trying."):
                state["failure_notified"] = True
    else:
        if state["failure_notified"]:
            notify(settings, "✅ SWA watcher is working again.")
        state = {"consecutive_failures": 0, "failure_notified": False}
    save_state(state)


def run(args):
    """One check. Returns a process exit code (0 = fine, 1 = something failed)."""
    settings, warning = core.load_settings()
    if warning:
        log.warning(warning)
    watches = settings["watches"]
    if not watches:
        log.info("No price watches in %s; nothing to do.", core.SETTINGS_PATH)
        return 0
    if not core.telegram_is_configured(settings):
        log.warning("Telegram isn't set up in settings.json, so alerts can't be sent.")
    cfg = settings["telegram"]

    origins = sorted({watch["origin"] for watch in watches})
    fetched, errors = 0, []
    alerts_sent = alerts_failed = 0
    changes = {}  # (origin, dest) -> (last_alert_price, last_alert_at), written back once at the end

    for origin in origins:
        try:
            _label, rows = fetch_with_patience(origin, args.net_attempts, args.net_wait)
        except Exception as exc:
            log.error("Couldn't fetch fares for %s: %s", origin, exc)
            errors.append(str(exc))
            continue
        fetched += 1

        recorded = 0
        if not args.dry_run:
            try:
                recorded = core.record_prices(origin, rows)
            except (sqlite3.Error, OSError) as exc:
                log.warning("Couldn't record price history for %s: %s", origin, exc)
        log.info("%s: %d routes fetched, %d price change(s) recorded", origin, len(rows), recorded)

        plan = core.plan_watch_alerts(settings, origin, rows)
        for key in plan.rearmed:
            changes[key] = (None, None)
        for hit in plan.unsent_hits:
            log.warning("%s: watch hit (%s) but Telegram isn't set up", origin, hit)
        for key, text, price in plan.jobs:
            if args.dry_run:
                log.info("DRY RUN - would send alert for %s -> %s at %s:\n%s", key[0], key[1], core.fmt_money(price), text)
                continue
            try:
                core.send_telegram(cfg["bot_token"], cfg["chat_id"], text)
            except Exception as exc:  # a failed send is retried next run; just record it
                alerts_failed += 1
                log.error("Alert for %s -> %s FAILED: %s (will retry next run)", key[0], key[1], exc)
                continue
            alerts_sent += 1
            changes[key] = (price, datetime.now().strftime(core.HISTORY_TIME_FORMAT))
            log.info("Alert sent: %s -> %s at %s", key[0], key[1], core.fmt_money(price))

    if changes and not args.dry_run:
        try:
            core.apply_alert_changes(changes)
        except OSError as exc:
            log.error("Couldn't save alert state to settings.json: %s (alerts may repeat next run)", exc)

    all_failed = fetched == 0
    if not args.dry_run:
        track_failures(settings, all_failed, errors[0] if errors else "unknown error")
    log.info("Done: %d/%d home airport(s) checked, %d alert(s) sent, %d failed.",
             fetched, len(origins), alerts_sent, alerts_failed)
    return 1 if all_failed or alerts_failed else 0


def main(argv=None, now=None):
    parser = argparse.ArgumentParser(description="Check SWA price watches and send Telegram alerts.")
    parser.add_argument("--only-between", type=parse_hours, metavar="START-END",
                        help="only run during these hours (inclusive, 24h clock), e.g. 6-23; otherwise exit quietly")
    parser.add_argument("--dry-run", action="store_true",
                        help="fetch and evaluate, but send nothing and change no files")
    parser.add_argument("--net-attempts", type=int, default=DEFAULT_NET_ATTEMPTS, help=argparse.SUPPRESS)
    parser.add_argument("--net-wait", type=int, default=DEFAULT_NET_WAIT_SECONDS, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)

    setup_logging()
    now = now or datetime.now()
    if args.only_between and not args.only_between[0] <= now.hour <= args.only_between[1]:
        log.info("Outside %d:00-%d:59 (it's %s); skipping this run.",
                 args.only_between[0], args.only_between[1], now.strftime("%H:%M"))
        return 0
    try:
        return run(args)
    except Exception:
        log.exception("Watcher crashed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
