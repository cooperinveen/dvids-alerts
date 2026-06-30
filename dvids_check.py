#!/usr/bin/env python3
"""
DVIDS B-Roll watcher.

Checks the DVIDS Search API for newly published B-Roll videos and posts each
new one to a Microsoft Teams channel via a Workflows ("incoming webhook")
Adaptive Card.

State is kept in state.json (last-seen publish date + recently-posted IDs) so
nothing is missed or double-posted even if a scheduled run is skipped.

Reads two secrets from the environment:
  DVIDS_API_KEY    - DVIDS API key (key-...)
  TEAMS_WEBHOOK_URL - Teams Workflows webhook URL

Zero third-party dependencies (stdlib only).
"""

import json
import os
import sys
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

# --- config -----------------------------------------------------------------

SEARCH_URL = "https://api.dvidshub.net/search"
STATE_FILE = os.path.join(os.path.dirname(__file__), "state.json")

# DVIDS search filters — mirrors the dvidshub.net B-Roll video search.
SEARCH_PARAMS = {
    "q": "",
    "type": "video",
    "category": "B-Roll",
    "sort": "publishdate",
    "sortdir": "desc",
    "max_results": "50",
}

# How many recently-posted IDs to remember (dedupe guard).
SEEN_IDS_CAP = 300

# On the very first run (no state file), seed silently instead of flooding the
# channel with the entire current page of results. Override by setting
# SEED_AND_POST=1 (used by the manual "test" run).
SEED_AND_POST = os.environ.get("SEED_AND_POST") == "1"


# --- helpers ----------------------------------------------------------------

def log(msg):
    print(msg, flush=True)


def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return None
    except (json.JSONDecodeError, OSError) as e:
        log(f"WARNING: could not read state file ({e}); treating as first run.")
        return None


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)
        f.write("\n")


def parse_dt(value):
    """Parse a DVIDS ISO8601 timestamp to an aware datetime, or None."""
    if not value:
        return None
    try:
        # DVIDS uses e.g. 2026-06-30T16:28:00Z
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except (ValueError, AttributeError):
        return None


def fetch_results(api_key):
    params = dict(SEARCH_PARAMS)
    params["api_key"] = api_key
    url = SEARCH_URL + "?" + urllib.parse.urlencode(params)

    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = resp.read().decode("utf-8")

    data = json.loads(body)
    if isinstance(data, dict) and data.get("errors"):
        raise RuntimeError(f"DVIDS API returned errors: {data['errors']}")
    return data.get("results", []) if isinstance(data, dict) else []


def build_card(item):
    """Build the Teams Adaptive Card envelope for one video."""
    title = item.get("title") or "(untitled)"
    desc = item.get("short_description") or ""
    url = item.get("url") or ""
    thumb = item.get("thumbnail") or ""

    facts = []
    when = parse_dt(item.get("date_published"))
    if when:
        facts.append({"title": "Published", "value": when.strftime("%d %b %Y %H:%M UTC")})
    for label, key in (("Unit", "unit_name"), ("Branch", "branch")):
        if item.get(key):
            facts.append({"title": label, "value": str(item[key])})
    loc = ", ".join(p for p in (item.get("city"), item.get("state"), item.get("country")) if p)
    if loc:
        facts.append({"title": "Location", "value": loc})
    if item.get("duration"):
        try:
            secs = int(item["duration"])
            facts.append({"title": "Duration", "value": f"{secs // 60}m {secs % 60}s"})
        except (ValueError, TypeError):
            pass

    body = [
        {"type": "TextBlock", "text": "🎥 New B-Roll on DVIDS",
         "weight": "Bolder", "size": "Medium", "color": "Accent"},
        {"type": "TextBlock", "text": title, "weight": "Bolder",
         "size": "Large", "wrap": True, "spacing": "Small"},
    ]
    if thumb:
        body.append({"type": "Image", "url": thumb, "size": "Stretch",
                     "altText": title, "spacing": "Medium"})
    if desc:
        body.append({"type": "TextBlock", "text": desc, "wrap": True, "spacing": "Medium"})
    if facts:
        body.append({"type": "FactSet", "facts": facts, "spacing": "Medium"})

    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": body,
        "actions": [{"type": "Action.OpenUrl", "title": "▶ Watch on DVIDS", "url": url}] if url else [],
    }
    return {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": card,
        }],
    }


def post_to_teams(webhook_url, card):
    payload = json.dumps(card).encode("utf-8")
    req = urllib.request.Request(
        webhook_url, data=payload,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status


# --- main -------------------------------------------------------------------

def main():
    api_key = os.environ.get("DVIDS_API_KEY")
    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL")
    if not api_key or not webhook_url:
        log("ERROR: DVIDS_API_KEY and TEAMS_WEBHOOK_URL must both be set.")
        return 1

    try:
        results = fetch_results(api_key)
    except Exception as e:
        log(f"ERROR fetching from DVIDS: {e}")
        return 1

    log(f"Fetched {len(results)} result(s) from DVIDS.")
    state = load_state()
    first_run = state is None
    if first_run:
        state = {"last_published": None, "seen_ids": []}

    seen_ids = set(state.get("seen_ids") or [])
    last_published = parse_dt(state.get("last_published"))

    # Oldest first, so the channel reads chronologically.
    results = list(reversed(results))

    new_items = []
    for item in results:
        item_id = item.get("id")
        if not item_id or item_id in seen_ids:
            continue
        pub = parse_dt(item.get("date_published"))
        # If we have a watermark, only take strictly newer items.
        if last_published and pub and pub <= last_published:
            continue
        new_items.append(item)

    # Decide whether to actually post.
    posting = not (first_run and not SEED_AND_POST)
    if first_run and not SEED_AND_POST:
        log(f"First run: seeding state with {len(new_items)} item(s), posting none. "
            f"(Set SEED_AND_POST=1 to post on a manual run.)")

    posted = 0
    for item in new_items:
        if posting:
            try:
                status = post_to_teams(webhook_url, build_card(item))
                log(f"Posted: {item.get('id')} — {item.get('title')!r} (HTTP {status})")
                posted += 1
            except Exception as e:
                log(f"ERROR posting {item.get('id')}: {e} — will retry next run.")
                # Don't record as seen, so the next run tries again.
                continue
        seen_ids.add(item.get("id"))
        pub = parse_dt(item.get("date_published"))
        if pub and (last_published is None or pub > last_published):
            last_published = pub

    # Persist trimmed state.
    state["last_published"] = last_published.strftime("%Y-%m-%dT%H:%M:%SZ") if last_published else None
    # Keep the most recent IDs only.
    state["seen_ids"] = (list(seen_ids))[-SEEN_IDS_CAP:]
    save_state(state)

    log(f"Done. New: {len(new_items)}, posted: {posted}, "
        f"watermark: {state['last_published']}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
