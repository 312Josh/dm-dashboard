#!/usr/bin/env python3
"""
Chris Byrnes TAM — bulk tech stack sweep.

Cache-first resumable pipeline. Each parent has a JSON cache file under
data/tam-sweep/cache/{slug}.json. Stages update the cache in place; the
assembler joins cache back to the 820-row CSV at the end.

Usage:
    python tam-stack-sweep.py extract-parents
    python tam-stack-sweep.py curl <slug> <url>
    python tam-stack-sweep.py set-website <slug> <url> [alt_url...]
    python tam-stack-sweep.py set-marketplaces <slug> '<json>'
    python tam-stack-sweep.py set-hr <slug> '<vendor>'
    python tam-stack-sweep.py list-missing-websites [--limit N]
    python tam-stack-sweep.py list-missing-curl [--limit N]
    python tam-stack-sweep.py list-missing-marketplaces [--limit N]
    python tam-stack-sweep.py status
    python tam-stack-sweep.py assemble
    python tam-stack-sweep.py rollup
"""
from __future__ import annotations

import csv
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

ROOT = Path.home() / "dm-dashboard"
CACHE = ROOT / "data" / "tam-sweep" / "cache"
LOGS = ROOT / "data" / "tam-sweep" / "logs"
INPUT_CSV = Path.home() / "Downloads" / "Chris-TAM - Sheet1.csv"
OUTPUT_CSV = Path.home() / "Desktop" / "chris-tam-stack-inventory.csv"
ROLLUP_MD = Path.home() / "Desktop" / "chris-tam-stack-rollup.md"

UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"
)

# Fingerprint categories — each regex returns first match as vendor name.
FINGERPRINTS = {
    "website_cms": [
        (r"popmenu", "Popmenu"),
        (r"bentobox|getbento", "BentoBox"),
        (r"owner\.com|getowner|mrkt\.ninja|ordersave\.com|powered by owner", "Owner.com"),
        (r"spothopper|shopperscms", "SpotHopper"),
        (r"squarespace", "Squarespace"),
        (r"wix\.com", "Wix"),
        (r"wp-content|wordpress", "WordPress"),
        (r"shopify", "Shopify"),
        (r"duda\.co", "Duda"),
        (r"godaddy\.com|websitebuilder\.godaddy", "GoDaddy"),
        (r"skymarketing|skyordering", "Sky Marketing"),
    ],
    "oo": [
        (r"order\.toasttab\.com", "Toast OO (native)"),
        (r"order\.thanx\.com", "Thanx OO"),
        (r"toasttab", "Toast OO (whitelabel candidate)"),
        (r"popmenu.*order|order.*popmenu|popmenu.*\.com/s/", "Popmenu OO"),
        (r"chownow", "ChowNow"),
        (r"olocdn|\.olo\.com|ordering\.olo", "Olo"),
        (r"appfront", "Appfront"),
        (r"menufy", "Menufy"),
        (r"slicelife|slice\.com", "Slice"),
        (r"beyondmenu", "BeyondMenu"),
        (r"owner\.com.*order|order.*owner\.com|ordersave", "Owner.com OO"),
        (r"skyordering", "Sky Ordering"),
    ],
    "reservations": [
        (r"opentable", "OpenTable"),
        (r"resy\.com", "Resy"),
        (r"sevenrooms", "SevenRooms"),
        (r"exploretock|tockhq", "Tock"),
        (r"yelpreservations|yelp\.com/reservations", "Yelp Reservations"),
        (r"tables\.toasttab\.com", "Toast Tables"),
        (r"waitlist\.me|waitlistme", "Waitlist Me"),
        (r"nowait\.com", "NoWait (Yelp)"),
        (r"tablein\.com", "Tablein"),
        (r"eatapp\.co", "Eat App"),
        (r"gloriafood", "GloriaFood"),
    ],
    "loyalty": [
        (r"toasttab\.com/[^\"']*/rewardsLookup", "Toast Loyalty"),
        (r"toasttab\.com/[^\"']*/rewardsSignup", "Toast Loyalty"),
        (r"punchh", "Punchh"),
        (r"thanx\.com", "Thanx"),
        (r"paytronix", "Paytronix"),
        (r"incentivio", "Incentivio"),
        (r"spendgo", "Spendgo"),
        (r"fivestars", "Fivestars"),
        (r"como\.com", "Como"),
        (r"loyalzoo", "Loyalzoo"),
        (r"marsello", "Marsello"),
        (r"belly\.com|bellycard", "Belly"),
        (r"tangoesign|tangocard", "Tango"),
    ],
    "gift_cards": [
        (r"toasttab\.com/[^\"']*/giftcards", "Toast Gift Cards"),
        (r"giftup\.com", "GiftUp"),
        (r"yiftee", "Yiftee"),
        (r"squareup\.com/gift", "Square Gift"),
        (r"givex", "Givex"),
        (r"valutec", "Valutec"),
        (r"factor4gift", "Factor4"),
        (r"fiserv.*gift|giftcardsource", "Fiserv Gift"),
    ],
    "marketing": [
        (r"klaviyo", "Klaviyo"),
        (r"attentivemobile|attn\.tv", "Attentive"),
        (r"mailchimp", "Mailchimp"),
        (r"constantcontact", "Constant Contact"),
        (r"birdeye", "Birdeye"),
        (r"podium\.com", "Podium"),
        (r"yext\.com", "Yext"),
        (r"fishbowl\.com", "Fishbowl"),
        (r"reviewtrackers", "ReviewTrackers"),
        (r"reputation\.com", "Reputation.com"),
        (r"marsello", "Marsello Marketing"),
        (r"textedly", "Textedly"),
        (r"simpletexting", "SimpleTexting"),
        (r"sendlane", "Sendlane"),
        (r"emotive\.io", "Emotive"),
        (r"hubspot", "HubSpot"),
    ],
    "pos_other": [
        (r"squareup\.com|square\.site", "Square"),
        (r"clover\.com", "Clover"),
        (r"lightspeedhq|upserve", "Lightspeed/Upserve"),
        (r"aloha(enterprise|cloud)?", "Aloha"),
        (r"micros\.com|oracle.*micros", "Micros/Oracle"),
        (r"revel(systems)?\.com", "Revel"),
        (r"touchbistro", "TouchBistro"),
        (r"heartland.*payment|heartlandhps", "Heartland"),
    ],
    "scheduling": [
        (r"7shifts", "7shifts"),
        (r"hotschedules|fourth\.com", "HotSchedules"),
        (r"joinhomebase|homebase\.com", "Homebase"),
        (r"deputy\.com", "Deputy"),
        (r"wheniwork", "When I Work"),
        (r"sling\.is", "Sling"),
    ],
    "delivery_platforms": [
        (r"chownow\.com.*direct", "ChowNow Direct"),
        (r"flipdish", "Flipdish"),
        (r"lunchbox\.io", "Lunchbox"),
        (r"bbot\.menu|bbot\.com", "Bbot (DoorDash)"),
    ],
    # Catering mention on the restaurant's own site — signals either:
    # (a) they already do catering somewhere (direct/phone), or
    # (b) they route to ezCater/CaterCow/etc. which another regex catches.
    # Presence of this signal = Toast Catering is a live conversation, not greenfield.
    "catering_mention": [
        (r"/catering[/\"'?#]|catering-menu|catering_menu|cateringmenu", "Catering page"),
        (r"we (also )?(offer|provide|do|cater)\s+catering|catering (is )?available|book\s+catering|order\s+catering", "Catering copy"),
        (r"catering@|cater@", "Catering email"),
        (r"/private[\s\-_]?events?|/private[\s\-_]?dining|private[\s\-_]?event[\s\-_]?space|book[\s\-_]?private[\s\-_]?event", "Private events page"),
        (r"private\s+(events|parties|dining)|book\s+(an?\s+)?private|rehearsal\s+dinner|corporate\s+event|event\s+space\s+(rental|available)|host\s+your\s+(event|party|celebration)|banquet\s+(hall|room)", "Private-events copy"),
        (r"events@|privateevents@|banquet@", "Events email"),
    ],
}

# Toast product URL patterns (the URL itself is evidence).
TOAST_PATTERNS = {
    "toast_oo_native": re.compile(r"order\.toasttab\.com/online/([a-z0-9\-]+)", re.I),
    "toast_oo_whitelabel_host": re.compile(r"https?://order\.([a-z0-9\-\.]+)/", re.I),
    "toast_tables": re.compile(r"tables\.toasttab\.com/restaurants/([a-z0-9\-]+)", re.I),
    "toast_gift": re.compile(r"toasttab\.com/([a-z0-9\-]+)/giftcards", re.I),
    "toast_loyalty": re.compile(r"toasttab\.com/([a-z0-9\-]+)/rewards(?:Lookup|Signup)", re.I),
    "toast_local": re.compile(r"toasttab\.com/local/([a-z0-9\-]+)", re.I),
}


def slugify(name: str) -> str:
    s = name.lower()
    for suffix in [" - parent account", " - parent", " corporate", " parent account", " parent"]:
        if s.endswith(suffix):
            s = s[: -len(suffix)]
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s


def cache_path(slug: str) -> Path:
    return CACHE / f"{slug}.json"


def load_cache(slug: str) -> dict[str, Any]:
    p = cache_path(slug)
    if not p.exists():
        return {}
    return json.loads(p.read_text())


def save_cache(slug: str, data: dict[str, Any]) -> None:
    cache_path(slug).write_text(json.dumps(data, indent=2, ensure_ascii=False))


# ---------- extract-parents ----------

def extract_parents(csv_path: Path | None = None, rep_slug: str | None = None) -> None:
    """Seed per-parent cache JSONs from a TAM CSV. Defaults to the legacy
    INPUT_CSV (Chris) for backward compat. Pass csv_path + rep_slug to
    onboard any other rep's CSV."""
    source = csv_path or INPUT_CSV
    seen: dict[str, dict[str, Any]] = {}
    with source.open() as f:
        reader = csv.DictReader(f)
        for row in reader:
            parent = row["Parent Account Name"].strip()
            if not parent:
                continue
            slug = slugify(parent)
            if slug not in seen:
                seen[slug] = {
                    "parent": parent,
                    "slug": slug,
                    "total_headroom": row.get("Parent Total Headroom", ""),
                    "non_ec_headroom": row.get("Parent Non-EC Headroom", ""),
                    "est_ec_headroom": row.get("Parent Est. EC Headroom", ""),
                    "employee_count": row.get("Parent Unique Employee Count", ""),
                    "non_ec_arr": row.get("Parent Non-EC Live Subs ARR", ""),
                    "state": row.get("State", ""),
                    "city": row.get("City", ""),
                    "restaurant_type": row.get("Restaurant Type", ""),
                    "contact_email": row.get("Business Owner Email", ""),
                    "location_count": 0,
                    "location_names": [],
                    # stage outputs — empty until filled
                    "websites": [],
                    "website_discovery_status": "pending",  # pending | found | not_found
                    "fingerprints": {},
                    "toast_products": {},
                    "marketplaces": {},
                    "marketplaces_status": "pending",
                    "hr_vendor": "",
                    "hr_status": "pending",
                    "notes": [],
                }
            seen[slug]["location_count"] += 1
            loc = row.get("Location Name", "").strip()
            if loc and loc not in seen[slug]["location_names"]:
                seen[slug]["location_names"].append(loc)

    for slug, data in seen.items():
        # only write if no existing cache, so we don't clobber in-progress runs
        p = cache_path(slug)
        if p.exists():
            existing = load_cache(slug)
            # refresh TAM metadata but preserve stage outputs
            for k in [
                "parent", "total_headroom", "non_ec_headroom", "est_ec_headroom",
                "employee_count", "non_ec_arr", "state", "city", "restaurant_type",
                "contact_email", "location_count", "location_names",
            ]:
                existing[k] = data[k]
            # Rep attribution: append rep_slug if not already present
            if rep_slug:
                reps = existing.setdefault("reps", [])
                if rep_slug not in reps:
                    reps.append(rep_slug)
            save_cache(slug, existing)
        else:
            if rep_slug:
                data["reps"] = [rep_slug]
            else:
                data["reps"] = []
            save_cache(slug, data)

    rep_label = f" (rep={rep_slug})" if rep_slug else ""
    print(f"Extracted {len(seen)} unique parents from {source.name}{rep_label} → {CACHE}")


# ---------- curl + grep fingerprint ----------

def _fetch(url: str) -> tuple[str, str | None]:
    try:
        result = subprocess.run(
            ["curl", "-s", "-L", "--max-time", "20", "-A", UA, url],
            capture_output=True, timeout=30,
        )
        return result.stdout.decode("utf-8", errors="replace"), None
    except Exception as e:  # noqa: BLE001
        return "", str(e)


def _scan(html: str) -> tuple[dict[str, str], dict[str, str]]:
    findings: dict[str, str] = {}
    for category, patterns in FINGERPRINTS.items():
        for regex, name in patterns:
            if re.search(regex, html, re.I):
                findings[category] = name
                break
    toast_hits: dict[str, str] = {}
    for key, pattern in TOAST_PATTERNS.items():
        m = pattern.search(html)
        if m:
            toast_hits[key] = m.group(0)
    if "toast_oo_whitelabel_host" in toast_hits:
        host_match = re.match(r"https?://order\.([a-z0-9\-\.]+)/", toast_hits["toast_oo_whitelabel_host"], re.I)
        if host_match:
            host = host_match.group(1)
            if "toasttab" not in host.lower() and "toasttab" in html.lower():
                toast_hits["toast_oo_whitelabel_confirmed"] = f"order.{host} + toasttab string in HTML"
    return findings, toast_hits


SUBPAGE_KEYWORDS = re.compile(
    r'href="([^"]*(?:/order|/reserve|/book|/reservations|/gift|/reward|/loyalty|/menu|/waitlist)[^"]*)"',
    re.I,
)


def _extract_subpages(html: str, base_url: str) -> list[str]:
    from urllib.parse import urljoin, urlparse
    base_host = urlparse(base_url).netloc.lower()
    out = []
    seen = set()
    for m in SUBPAGE_KEYWORDS.finditer(html):
        href = m.group(1)
        abs_url = urljoin(base_url, href)
        parsed = urlparse(abs_url)
        # Only follow same-host hrefs (skip third-party embeds)
        if parsed.netloc.lower() != base_host:
            continue
        # Trim fragments, skip mailto/tel
        if parsed.scheme not in ("http", "https"):
            continue
        key = abs_url.split("#")[0].rstrip("/")
        if key in seen:
            continue
        seen.add(key)
        out.append(abs_url)
        if len(out) >= 6:
            break
    return out


def curl_fingerprint(slug: str, url: str) -> None:
    data = load_cache(slug)
    if not data:
        print(f"ERROR: no cache for {slug}", file=sys.stderr)
        sys.exit(1)

    html, err = _fetch(url)
    if err:
        data.setdefault("fingerprints", {})[url] = {"error": err}
        data.setdefault("notes", []).append(f"curl failed on {url}: {err}")
        save_cache(slug, data)
        print(f"  curl failed: {err}")
        return

    findings, toast_hits = _scan(html)
    entry = {
        "url": url,
        "html_bytes": len(html),
        "fingerprints": findings,
        "toast": toast_hits,
    }
    data.setdefault("fingerprints", {})[url] = entry

    # Follow up to 6 same-host subpages (order/reserve/gift/loyalty/menu)
    subpages = _extract_subpages(html, url)
    for sub in subpages:
        sub_html, sub_err = _fetch(sub)
        if sub_err:
            data["fingerprints"][sub] = {"error": sub_err}
            continue
        sub_find, sub_toast = _scan(sub_html)
        data["fingerprints"][sub] = {
            "url": sub, "html_bytes": len(sub_html),
            "fingerprints": sub_find, "toast": sub_toast,
        }
        # Merge into aggregate findings (first-hit-wins; preserve homepage-level first)
        for cat, vendor in sub_find.items():
            findings.setdefault(cat, vendor)
        for k, v in sub_toast.items():
            toast_hits.setdefault(k, v)

    # Aggregate toast products across all URLs fetched
    agg = data.setdefault("toast_products", {})
    for k, v in toast_hits.items():
        agg.setdefault(k, v)

    save_cache(slug, data)

    print(f"  {slug} ← {url}  ({len(html):,} bytes, +{len(subpages)} subpages)")
    if findings:
        print(f"    vendors: {findings}")
    if toast_hits:
        print(f"    toast:   {toast_hits}")
    if not findings and not toast_hits:
        print("    no fingerprints matched")


def is_multi_brand(data: dict) -> bool:
    """A parent is multi-brand if its locations have >1 distinctly-named brands.
    Normalize names: strip '- {city}' / '- {region}' / ' 1'-' N' suffixes that denote
    same-brand multi-location."""
    locs = data.get("location_names", [])
    if len(locs) <= 1:
        return False
    normalized = set()
    for l in locs:
        n = re.sub(r"\s*-\s*[A-Za-z\s\.\']+$", "", l)  # strip trailing "- City"
        n = re.sub(r"\s+\d+$", "", n)  # strip trailing location number
        n = re.sub(r"[^a-z0-9]+", "", n.lower())
        if n:
            normalized.add(n)
    return len(normalized) > 1


def set_website(slug: str, urls: list[str]) -> None:
    data = load_cache(slug)
    if not data:
        print(f"ERROR: no cache for {slug}", file=sys.stderr)
        sys.exit(1)
    clean = [u for u in urls if u]
    # Guard: refuse to mark a multi-brand parent as "not_found" with no URLs —
    # each brand needs its own per-location website discovery.
    if not clean and is_multi_brand(data):
        locs = data.get("location_names", [])
        print(
            f"  REFUSED: {slug} is multi-brand with {len(locs)} distinct locations. "
            f"Use set-brand-website for each brand instead of marking not_found.\n"
            f"  Distinct location names: {locs}",
            file=sys.stderr,
        )
        data["website_discovery_status"] = "needs_per_location_discovery"
        save_cache(slug, data)
        sys.exit(2)
    data["websites"] = clean
    data["website_discovery_status"] = "found" if clean else "not_found"
    save_cache(slug, data)
    print(f"  {slug} websites set: {clean or '[]'}")


def set_brand_website(slug: str, brand_key: str, url: str) -> None:
    """Attach a website to a specific brand/location-name pattern within a parent.

    brand_key is a substring that identifies location names belonging to this
    brand (case-insensitive). E.g., parent "Pineapple Hospitality" can have:
      brand 'Agave' → eatagavesocial.com
      brand 'Craft Street' → craftstreetkitchen.com
      brand 'Shaker' → shakerandpeel.com
      brand 'ZimZari' → zimzari.com
    """
    data = load_cache(slug)
    if not data:
        print(f"ERROR: no cache for {slug}", file=sys.stderr)
        sys.exit(1)
    brands = data.setdefault("brand_websites", {})
    brands[brand_key.lower()] = url
    # Also track URL in websites list so curl stage picks it up
    if url and url not in data.get("websites", []):
        data.setdefault("websites", []).append(url)
    data["website_discovery_status"] = "found"
    save_cache(slug, data)
    print(f"  {slug} brand '{brand_key}' → {url}")


def set_marketplaces(slug: str, mp_json: str) -> None:
    data = load_cache(slug)
    if not data:
        print(f"ERROR: no cache for {slug}", file=sys.stderr)
        sys.exit(1)
    mp = json.loads(mp_json)
    data["marketplaces"] = mp
    data["marketplaces_status"] = "done"
    save_cache(slug, data)
    print(f"  {slug} marketplaces: {mp}")


def set_hr(slug: str, vendor: str) -> None:
    data = load_cache(slug)
    if not data:
        print(f"ERROR: no cache for {slug}", file=sys.stderr)
        sys.exit(1)
    data["hr_vendor"] = vendor
    data["hr_status"] = "done"
    save_cache(slug, data)
    print(f"  {slug} hr: {vendor}")


# ---------- status / list helpers ----------

def iter_cache():
    for p in sorted(CACHE.glob("*.json")):
        yield p.stem, json.loads(p.read_text())


def status() -> None:
    total = 0
    missing_site = 0
    not_found_site = 0
    missing_curl = 0
    missing_mp = 0
    missing_hr = 0
    for slug, d in iter_cache():
        total += 1
        if d["website_discovery_status"] == "pending":
            missing_site += 1
        elif d["website_discovery_status"] == "not_found":
            not_found_site += 1
        if d["websites"] and not d.get("fingerprints"):
            missing_curl += 1
        if d["marketplaces_status"] == "pending":
            missing_mp += 1
        if d["hr_status"] == "pending":
            missing_hr += 1
    print(f"Total parents:          {total}")
    print(f"Websites pending:       {missing_site}")
    print(f"Websites not_found:     {not_found_site}")
    print(f"Curl pending:           {missing_curl}")
    print(f"Marketplaces pending:   {missing_mp}")
    print(f"HR pending:             {missing_hr}")


def list_missing(stage: str, limit: int | None) -> None:
    out = []
    for slug, d in iter_cache():
        if stage == "websites" and d["website_discovery_status"] == "pending":
            out.append((slug, d["parent"], d.get("state", ""), d.get("city", "")))
        elif stage == "curl" and d["websites"] and not d.get("fingerprints"):
            out.append((slug, d["parent"], d["websites"][0]))
        elif stage == "marketplaces" and d["marketplaces_status"] == "pending" and d["website_discovery_status"] == "found":
            out.append((slug, d["parent"]))
        elif stage == "hr" and d["hr_status"] == "pending" and d["website_discovery_status"] == "found":
            out.append((slug, d["parent"]))
    if limit:
        out = out[:limit]
    for row in out:
        print("\t".join(str(x) for x in row))


# ---------- assemble ----------

def build_primary_play(d: dict[str, Any]) -> str:
    fp = {}
    for fpdata in d.get("fingerprints", {}).values():
        if isinstance(fpdata, dict) and "fingerprints" in fpdata:
            for cat, vendor in fpdata["fingerprints"].items():
                fp.setdefault(cat, vendor)
    toast = d.get("toast_products", {})
    mp = d.get("marketplaces", {})

    # Highest-value displacement first
    if fp.get("website_cms") in ("Popmenu", "BentoBox", "Owner.com"):
        return f"Displace {fp['website_cms']} (website + OO consolidation)"
    if fp.get("oo") in ("ChowNow", "Olo", "Appfront"):
        return f"Displace {fp['oo']} OO"
    if mp and "ezcater" not in mp.get("present", []) and "catering" in d.get("restaurant_type", "").lower():
        return "Toast Catering greenfield (no EzCater detected)"
    if fp.get("reservations") in ("OpenTable", "Resy", "SevenRooms", "Tock"):
        return f"Displace {fp['reservations']} (Toast Tables)"
    if not toast.get("toast_loyalty") and fp.get("loyalty") is None:
        return "Loyalty greenfield — pitch Toast Loyalty"
    return "Assess — no obvious displacement signal"


def confidence(d: dict[str, Any]) -> str:
    if d["website_discovery_status"] != "found":
        return "low"
    fp = d.get("fingerprints", {})
    if not fp or all(isinstance(v, dict) and v.get("error") for v in fp.values()):
        return "low"
    return "high"


def aggregate_fingerprints(d: dict[str, Any], url_filter: set[str] | None = None) -> dict[str, str]:
    agg: dict[str, str] = {}
    for url, fpdata in d.get("fingerprints", {}).items():
        if url_filter is not None and url not in url_filter:
            continue
        if isinstance(fpdata, dict) and "fingerprints" in fpdata:
            for cat, vendor in fpdata["fingerprints"].items():
                agg.setdefault(cat, vendor)
    return agg


def aggregate_toast(d: dict[str, Any], url_filter: set[str] | None = None) -> dict[str, str]:
    agg: dict[str, str] = {}
    for url, fpdata in d.get("fingerprints", {}).items():
        if url_filter is not None and url not in url_filter:
            continue
        if isinstance(fpdata, dict) and "toast" in fpdata:
            for k, v in fpdata["toast"].items():
                agg.setdefault(k, v)
    return agg


def pick_brand_urls(d: dict[str, Any], location_name: str) -> set[str]:
    """For a given location name, return the set of URLs (brand sites + subpages)
    that apply. If parent has no brand_websites mapping, all URLs apply."""
    brands = d.get("brand_websites", {})
    if not brands:
        return set(d.get("fingerprints", {}).keys())

    loc_lc = location_name.lower()
    matched_root_url: str | None = None
    # Longest brand_key match wins
    best = (0, None)
    for bkey, burl in brands.items():
        if bkey in loc_lc and len(bkey) > best[0]:
            best = (len(bkey), burl)
    matched_root_url = best[1]
    if not matched_root_url:
        # No match — use all URLs so we don't blank the row
        return set(d.get("fingerprints", {}).keys())

    # Include that brand's root URL and any subpages under its host
    from urllib.parse import urlparse
    target_host = urlparse(matched_root_url).netloc.lower()
    out = set()
    for url in d.get("fingerprints", {}).keys():
        if urlparse(url).netloc.lower() == target_host:
            out.add(url)
    return out


def assemble() -> None:
    # Build parent → stack dict
    parent_data: dict[str, dict[str, Any]] = {}
    for slug, d in iter_cache():
        parent_data[d["parent"]] = d

    new_cols = [
        "Detected Website(s)",
        "Website/CMS Vendor",
        "OO Vendor",
        "Reservations Vendor",
        "Loyalty Vendor",
        "Gift Cards Vendor",
        "Marketing Vendor",
        "HR/Payroll/Scheduling Vendor",
        "Toast OO",
        "Toast Tables",
        "Toast Gift Cards",
        "Toast Loyalty",
        "Marketplaces Present",
        "Primary Displacement Play",
        "Confidence",
        "Notes",
    ]

    with INPUT_CSV.open() as fin, OUTPUT_CSV.open("w", newline="") as fout:
        reader = csv.DictReader(fin)
        fieldnames = reader.fieldnames + new_cols
        writer = csv.DictWriter(fout, fieldnames=fieldnames)
        writer.writeheader()
        for row in reader:
            parent = row["Parent Account Name"].strip()
            loc = row.get("Location Name", "").strip()
            d = parent_data.get(parent, {})
            # Per-location attribution: pick URLs matching this location's brand
            url_filter = pick_brand_urls(d, loc) if d else None
            agg_fp = aggregate_fingerprints(d, url_filter) if d else {}
            toast = aggregate_toast(d, url_filter) if d else {}
            mp = d.get("marketplaces", {}) if d else {}

            # Toast OO — only count confirmed native or confirmed whitelabel
            if "toast_oo_native" in toast:
                toast_oo = f"native → {toast['toast_oo_native']}"
            elif "toast_oo_whitelabel_confirmed" in toast:
                toast_oo = f"whitelabel → {toast['toast_oo_whitelabel_confirmed']}"
            else:
                toast_oo = "—"

            toast_tables = f"✅ {toast['toast_tables']}" if "toast_tables" in toast else "—"
            toast_gift = f"✅ {toast['toast_gift']}" if "toast_gift" in toast else "—"
            if "toast_loyalty" in toast:
                m = TOAST_PATTERNS["toast_loyalty"].search(toast["toast_loyalty"])
                prog = m.group(1) if m else ""
                toast_loyalty = f"✅ program={prog}" if prog else "✅"
            else:
                toast_loyalty = "—"

            mp_present = ",".join(mp.get("present", [])) if mp else ""

            row.update({
                "Detected Website(s)": ", ".join(d.get("websites", [])),
                "Website/CMS Vendor": agg_fp.get("website_cms", "—"),
                "OO Vendor": agg_fp.get("oo", "—"),
                "Reservations Vendor": agg_fp.get("reservations", "—"),
                "Loyalty Vendor": agg_fp.get("loyalty", "—"),
                "Gift Cards Vendor": agg_fp.get("gift_cards", "—"),
                "Marketing Vendor": agg_fp.get("marketing", "—"),
                "HR/Payroll/Scheduling Vendor": d.get("hr_vendor", "—") or "—",
                "Toast OO": toast_oo,
                "Toast Tables": toast_tables,
                "Toast Gift Cards": toast_gift,
                "Toast Loyalty": toast_loyalty,
                "Marketplaces Present": mp_present,
                "Primary Displacement Play": build_primary_play(d) if d else "—",
                "Confidence": confidence(d) if d else "low",
                "Notes": "; ".join(d.get("notes", [])) if d else "",
            })
            writer.writerow(row)

    print(f"Wrote {OUTPUT_CSV}")


# ---------- rollup ----------

def rollup() -> None:
    parents = [d for _, d in iter_cache()]
    total = len(parents)
    covered = [p for p in parents if p["website_discovery_status"] == "found"]

    def count_vendor(cat: str, vendor: str) -> int:
        n = 0
        for p in covered:
            agg = aggregate_fingerprints(p)
            if agg.get(cat) == vendor:
                n += 1
        return n

    def count_toast(key: str) -> int:
        return sum(1 for p in covered if key in p.get("toast_products", {}))

    popmenu = count_vendor("website_cms", "Popmenu")
    bentobox = count_vendor("website_cms", "BentoBox")
    owner = count_vendor("website_cms", "Owner.com")
    chownow = count_vendor("oo", "ChowNow")
    olo = count_vendor("oo", "Olo")
    appfront = count_vendor("oo", "Appfront")

    toast_oo_native = count_toast("toast_oo_native")
    toast_oo_wl = count_toast("toast_oo_whitelabel_confirmed")
    toast_tables = count_toast("toast_tables")
    toast_gift = count_toast("toast_gift")
    toast_loyalty = count_toast("toast_loyalty")

    # EzCater greenfield
    ezcater_greenfield = 0
    for p in covered:
        mp = p.get("marketplaces", {}) or {}
        if "ezcater" not in (mp.get("present") or []) and "catering" in p.get("restaurant_type", "").lower():
            ezcater_greenfield += 1

    # Top 10 opportunities — headroom × confidence × # displaceable products
    ranked = []
    for p in parents:
        try:
            hr = int(re.sub(r"[^\d]", "", p.get("total_headroom", "0")) or 0)
        except ValueError:
            hr = 0
        agg = aggregate_fingerprints(p)
        n_displaceable = sum(
            1 for cat in ["website_cms", "oo", "reservations", "loyalty"]
            if agg.get(cat) and agg[cat] != "—" and "Toast" not in agg.get(cat, "")
        )
        score = hr * (1 + n_displaceable)
        ranked.append((score, hr, n_displaceable, p))
    ranked.sort(key=lambda x: x[0], reverse=True)

    lines = [
        "# Chris Byrnes TAM — Tech Stack Rollup",
        "",
        f"**Generated**: {__import__('datetime').date.today().isoformat()}",
        f"**Parents covered**: {len(covered)} / {total} ({len(covered)*100//max(total,1)}%)",
        "",
        "## Competitor concentration",
        "",
        "| Category | Vendor | # parents | % of covered |",
        "|---|---|---:|---:|",
    ]
    for cat, vendor, n in [
        ("Website/CMS", "Popmenu", popmenu),
        ("Website/CMS", "BentoBox", bentobox),
        ("Website/CMS", "Owner.com", owner),
        ("OO", "ChowNow", chownow),
        ("OO", "Olo", olo),
        ("OO", "Appfront", appfront),
    ]:
        pct = n * 100 // max(len(covered), 1)
        lines.append(f"| {cat} | {vendor} | {n} | {pct}% |")

    lines += [
        "",
        "## Toast footprint",
        "",
        f"- Toast OO (native): {toast_oo_native}",
        f"- Toast OO (whitelabel confirmed): {toast_oo_wl}",
        f"- Toast Tables: {toast_tables}",
        f"- Toast Gift Cards: {toast_gift}",
        f"- Toast Loyalty (active): {toast_loyalty}  ← **DO NOT cold-pitch loyalty to these**",
        "",
        "## Toast Catering greenfield",
        "",
        f"- {ezcater_greenfield} catering-type parents with NO EzCater presence detected",
        "",
        "## Top 10 displacement opportunities (headroom × displaceable vendors)",
        "",
        "| Rank | Parent | Headroom | Displaceable | Primary play |",
        "|---|---|---:|---:|---|",
    ]
    for i, (score, hr, nd, p) in enumerate(ranked[:10], 1):
        play = build_primary_play(p)
        lines.append(f"| {i} | {p['parent']} | ${hr:,} | {nd} | {play} |")

    ROLLUP_MD.write_text("\n".join(lines) + "\n")
    print(f"Wrote {ROLLUP_MD}")


# ---------- main ----------

def main() -> None:
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    cmd = sys.argv[1]
    if cmd == "extract-parents":
        # extract-parents [csv_path [rep_slug]]
        csv_arg = Path(sys.argv[2]).expanduser() if len(sys.argv) > 2 else None
        rep_arg = sys.argv[3] if len(sys.argv) > 3 else None
        extract_parents(csv_arg, rep_arg)
    elif cmd == "curl":
        curl_fingerprint(sys.argv[2], sys.argv[3])
    elif cmd == "curl-all":
        # Curl every URL already set on this slug, serially (no race).
        slug = sys.argv[2]
        data = load_cache(slug)
        if not data:
            print(f"ERROR: no cache for {slug}", file=sys.stderr)
            sys.exit(1)
        for url in data.get("websites", []):
            if url:
                curl_fingerprint(slug, url)
    elif cmd == "set-website":
        set_website(sys.argv[2], sys.argv[3:])
    elif cmd == "set-brand-website":
        set_brand_website(sys.argv[2], sys.argv[3], sys.argv[4])
    elif cmd == "set-marketplaces":
        set_marketplaces(sys.argv[2], sys.argv[3])
    elif cmd == "set-hr":
        set_hr(sys.argv[2], sys.argv[3])
    elif cmd == "list-missing-websites":
        limit = int(sys.argv[3]) if len(sys.argv) > 3 and sys.argv[2] == "--limit" else None
        list_missing("websites", limit)
    elif cmd == "list-missing-curl":
        limit = int(sys.argv[3]) if len(sys.argv) > 3 and sys.argv[2] == "--limit" else None
        list_missing("curl", limit)
    elif cmd == "list-missing-marketplaces":
        limit = int(sys.argv[3]) if len(sys.argv) > 3 and sys.argv[2] == "--limit" else None
        list_missing("marketplaces", limit)
    elif cmd == "list-missing-hr":
        limit = int(sys.argv[3]) if len(sys.argv) > 3 and sys.argv[2] == "--limit" else None
        list_missing("hr", limit)
    elif cmd == "status":
        status()
    elif cmd == "assemble":
        assemble()
    elif cmd == "rollup":
        rollup()
    else:
        print(f"Unknown command: {cmd}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
