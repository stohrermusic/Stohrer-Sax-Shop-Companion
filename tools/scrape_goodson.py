"""
Scrape pad set size data from saxgourmet.com and convert to pad preset format.

Outputs a JSON file with all presets under a "Steve Goodson" library key,
ready to merge into the website's pad_presets.json.
"""

import urllib.request
import json
import re
import time
import sys
import html as html_module


# All instrument page URLs from saxgourmet.com/pad-set-sizes-2/
URLS = [
    "https://www.saxgourmet.com/pad_set_sizes/saxgourmet-super-400-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/saxgourmet-super-400-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/saxgourmet-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/orpheo-tenor-steve-goodson-designed-only/",
    "https://www.saxgourmet.com/pad_set_sizes/orpheo-alto-steve-goodson-designed-only/",
    "https://www.saxgourmet.com/pad_set_sizes/steve-goodson-model-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/steve-goodson-model-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/steve-goodson-model-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/saxgourmet-voodoo-rex-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/saxgourmet-voodoo-rex-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/saxgourmet-category-five-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/buescher-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/buescher-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/buescher-c-melody/",
    "https://www.saxgourmet.com/pad_set_sizes/buescher-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/buffet-s-3-prestige-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/buffet-superdynaction-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/cannonball-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/cannonball-soprano/",
    "https://www.saxgourmet.com/pad_set_sizes/cannonball-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/conn-18m-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/conn-28m-constellation-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/conn-50m-director-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/conn-11m-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/conn-wonder-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/conn-10m-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/conn-16m-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/conn-djh-modified-110-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/dolnet-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/keilwerth-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/king-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/king-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/king-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/martin-the-martin-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/martin-c-melody/",
    "https://www.saxgourmet.com/pad_set_sizes/martin-the-martin-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/pan-american-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-balanced-action-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-cigar-cutter-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/couesnon-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-mark-vi-alto-2/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-mark-vii-alto-2/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-sa-80-ii-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-series-iii-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-balanced-action-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-162-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-1242-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-as100-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-bundy-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-mark-vi-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-mark-vii-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-mark-vi-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-s-80-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-1256-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-mark-vi-soprano/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-super-80-soprano/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-1244-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-bundy-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-mark-vi-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-mark-vii-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-sa-80-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-paris-series-iii-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-signet-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-ts100-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-ts200-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/selmer-usa-signet-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/sml-super-47-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/unison-steve-goodson-model-alto/",
    "https://www.saxgourmet.com/pad_set_sizes/unison-steve-goodson-model-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/yamaha-yts-82z-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/yamaha-yts-62-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/yamaha-yts-52-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/yamaha-yts-23-tenor/",
    "https://www.saxgourmet.com/pad_set_sizes/yamaha-yss-62-soprano/",
    "https://www.saxgourmet.com/pad_set_sizes/yamaha-yss-61-soprano/",
    "https://www.saxgourmet.com/pad_set_sizes/yamaha-ybs-62-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/yamaha-ybs-52-baritone/",
    "https://www.saxgourmet.com/pad_set_sizes/1039/",
    "https://www.saxgourmet.com/pad_set_sizes/1037/",
    "https://www.saxgourmet.com/pad_set_sizes/1030/",
    "https://www.saxgourmet.com/pad_set_sizes/yas-23-alto/",
]


def parse_page(html):
    """Extract title and size/qty data from a saxgourmet.com pad size page.

    The data is NOT in <table> tags. It's in <h2> tags with <br /> separators:
        <h1>Conn 10M Tenor</h1>
        <h2>Size&nbsp;&nbsp;&nbsp;Qty</h2>
        <h2>10.0 &nbsp; &nbsp; &nbsp;&nbsp; 2<br />
        17.5&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 1<br />
        ...</h2>

    Returns (title, rows) where rows is list of [size_str, qty_str].
    """
    # Extract title from <h1>
    h1_match = re.search(r'<h1[^>]*>(.*?)</h1>', html, re.DOTALL)
    title = ""
    if h1_match:
        title = re.sub(r'<[^>]+>', '', h1_match.group(1))
        title = html_module.unescape(title).strip()

    # Extract all <h2> blocks
    h2_blocks = re.findall(r'<h2[^>]*>(.*?)</h2>', html, re.DOTALL)

    rows = []
    for block in h2_blocks:
        # Split on <br /> or <br> to get individual lines
        lines = re.split(r'<br\s*/?\s*>', block)
        for line in lines:
            # Strip HTML tags
            text = re.sub(r'<[^>]+>', '', line)
            # Decode HTML entities (&#8217; -> ', &#8211; -> –, etc.)
            text = html_module.unescape(text)
            # Replace smart quotes used as decimal points (WordPress mangling)
            text = text.replace('\u2019', '.').replace('\u2018', '.')  # ' and '
            text = text.replace("'", ".").replace("\u2032", ".")  # apostrophe, prime
            # Replace &nbsp; and \xa0 with spaces
            text = text.replace('\xa0', ' ').replace('&nbsp;', ' ')
            # Normalize whitespace
            text = re.sub(r'\s+', ' ', text).strip()

            if not text:
                continue

            # Skip header lines
            if text.lower().startswith('size') or text.lower().startswith('qty'):
                continue

            # Try to parse as "size qty" (two numbers separated by space)
            parts = text.split()
            if len(parts) >= 2:
                # Take first and last parts as size and qty
                size_str = parts[0]
                qty_str = parts[-1]
                # Verify both look numeric
                try:
                    float(size_str)
                    int(float(qty_str))
                    rows.append([size_str, qty_str])
                except ValueError:
                    pass

    return title, rows


def clean_title(raw_title):
    """Clean up a page title into a nice preset name."""
    # Decode HTML entities first
    title = html_module.unescape(raw_title)
    # Remove site suffix
    title = title.split(" | ")[0].split(" – Sax Gourmet")[0].split(" - Sax Gourmet")[0]
    title = title.replace("Pad Set Sizes", "").strip()
    title = title.strip(" -\u2013\u2014:|")
    # Normalize dashes: en-dash/em-dash surrounded by spaces -> hyphen
    title = re.sub(r'\s*[\u2013\u2014]\s*', '-', title)
    # Normalize whitespace
    title = re.sub(r'\s+', ' ', title).strip()
    # Remove leading "Pad Set Sizes for" or similar
    title = re.sub(r'^Pad Set Sizes?\s*(for\s*)?', '', title, flags=re.IGNORECASE).strip()
    return title


def detect_instrument_type(title):
    """Try to detect instrument type from title for consistent naming."""
    title_lower = title.lower()
    for typ in ["soprano", "alto", "tenor", "baritone", "c-melody", "c melody"]:
        if typ in title_lower:
            return typ.title().replace("C-melody", "C-Melody").replace("C melody", "C-Melody")
    return None


def parse_table_data(rows):
    """Parse Size/Qty rows into pad preset format. Returns (pads_string, warnings)."""
    warnings = []
    pad_lines = []

    for row in rows:
        if len(row) < 2:
            continue

        size_str = row[0].strip()
        qty_str = row[1].strip()

        # Skip header rows
        if size_str.lower() in ("size", "pad size", "sz", ""):
            continue
        if qty_str.lower() in ("qty", "quantity", "qt", ""):
            continue

        # Parse size
        try:
            size = float(size_str)
        except ValueError:
            warnings.append(f"  Could not parse size: '{size_str}'")
            continue

        # Parse quantity
        try:
            qty = int(float(qty_str))
        except ValueError:
            warnings.append(f"  Could not parse qty: '{qty_str}' for size {size_str}")
            continue

        # Sanity checks
        if size <= 0 or qty <= 0:
            warnings.append(f"  Invalid size/qty: {size} x {qty}")
            continue

        # Auto-fix suspicious sizes (likely missing decimal point)
        if size >= 100:
            fixed = size / 10
            warnings.append(f"  AUTO-FIXED SIZE {size} -> {fixed:.1f} (missing decimal)")
            size = fixed

        # Flag unusual quantities
        if qty > 10:
            warnings.append(f"  UNUSUAL QTY: {size} x {qty}")

        # Format size: use integer if whole number, else one decimal
        if size == int(size):
            size_fmt = f"{int(size)}.0"
        else:
            size_fmt = f"{size:.1f}" if size * 10 == int(size * 10) else f"{size}"

        pad_lines.append(f"{size_fmt} x {qty}")

    pads_string = "\n".join(pad_lines)
    return pads_string, warnings


def fetch_page(url, retries=2):
    """Fetch a URL with retries and rate limiting."""
    headers = {"User-Agent": "StohrerSaxShopCompanion/scraper"}
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=15) as resp:
                return resp.read().decode("utf-8", errors="replace")
        except Exception as e:
            if attempt < retries:
                time.sleep(2)
            else:
                raise


def main():
    results = {}
    all_warnings = []
    errors = []
    name_counts = {}  # Track duplicate names

    print(f"Scraping {len(URLS)} instrument pages from saxgourmet.com...")
    print()

    for i, url in enumerate(URLS):
        slug = url.rstrip("/").split("/")[-1]
        progress = f"[{i+1}/{len(URLS)}]"

        try:
            html = fetch_page(url)
        except Exception as e:
            msg = f"{progress} ERROR fetching {slug}: {e}"
            print(msg)
            errors.append(msg)
            continue

        raw_title, rows = parse_page(html)

        title = clean_title(raw_title) if raw_title else ""
        if not title:
            # Fallback: derive from URL slug
            title = slug.replace("-", " ").title()

        inst_type = detect_instrument_type(title)

        # Parse data
        if not rows:
            msg = f"{progress} NO DATA FOUND: {slug} (title: {title})"
            print(msg)
            errors.append(msg)
            continue

        pads_string, warnings = parse_table_data(rows)

        if not pads_string:
            msg = f"{progress} NO VALID DATA: {slug} (title: {title})"
            print(msg)
            errors.append(msg)
            continue

        # Add instrument type suffix if not already in title
        preset_name = title
        if inst_type and inst_type.lower() not in title.lower():
            preset_name = f"{title} ({inst_type})"

        # Handle duplicate names
        if preset_name in name_counts:
            name_counts[preset_name] += 1
            preset_name = f"{preset_name} (v{name_counts[preset_name]})"
        else:
            name_counts[preset_name] = 1

        # Count total pads
        total_pads = 0
        for line in pads_string.split("\n"):
            parts = line.split(" x ")
            if len(parts) == 2:
                total_pads += int(parts[1])

        results[preset_name] = {
            "pads": pads_string,
            "notes": "via saxgourmet.com (Steve Goodson)"
        }

        status = "OK"
        if warnings:
            status = "WARNINGS"
            all_warnings.append(f"\n{preset_name}:")
            all_warnings.extend(warnings)

        print(f"{progress} {status}: {preset_name} ({total_pads} pads)")

        # Be polite - small delay between requests
        if i < len(URLS) - 1:
            time.sleep(0.5)

    # Output
    output = {"Steve Goodson": results}

    output_file = "tools/goodson_pad_presets.json"
    with open(output_file, "w") as f:
        json.dump(output, f, indent=2)

    print(f"\n{'='*60}")
    print(f"RESULTS: {len(results)} presets scraped, {len(errors)} errors")
    print(f"Output: {output_file}")

    if all_warnings:
        print(f"\nWARNINGS ({len(all_warnings)} items):")
        for w in all_warnings:
            print(w)

    if errors:
        print(f"\nERRORS ({len(errors)}):")
        for e in errors:
            print(f"  {e}")

    # Print all preset names sorted for review
    print(f"\nALL PRESET NAMES ({len(results)}):")
    for name in sorted(results.keys()):
        pads = results[name]["pads"]
        total = sum(int(line.split(" x ")[1]) for line in pads.split("\n") if " x " in line)
        print(f"  {name} ({total} pads)")


if __name__ == "__main__":
    main()
