"""
Convert Selmer USA catalog pad size chart into pad preset format.
Data extracted from Selmer-Bundy Pad sizes.PDF, page 3 (Saxophone Pad Sizes).

This is a Selmer factory catalog, likely early-to-mid 1980s based on the
SA-80 (Series I, not Series II) being the newest Paris model listed.

Footnotes from catalog:
* With metal tone boosters
** 11.0 and 9.0mm, one each
*** 11.0 and 7.5mm, one each
**** 27.0 and 41.5mm, one each
"""

import json
from collections import Counter

# Raw per-key data extracted from PDF.
# Keys 1-22 are regular tone holes, then special keys.
# Footnote keys use tuples for "one each" entries.

MODELS = {
    # ===== SOPRANO =====
    "Selmer (Paris) Mark VI (Soprano)": {
        1: 9.0, 2: 9.0, 3: 9.0, 4: 9.0,
        5: (11.0, 9.0),  # ** footnote
        6: 18.5, 7: 14.5, 8: 16.0, 9: 14.5, 10: 18.5,
        11: 22.0, 12: 18.5, 13: 22.0, 14: 22.0,
        15: 26.0, 16: 26.0, 17: 30.0, 18: 26.0,
        19: 34.0, 20: 34.0, 21: 38.0, 22: 38.0,
        "U Octave": 8.0, "L Octave": 8.0, "High F#": 9.0,
    },
    "Selmer (Paris) SA-80 (Soprano)": {
        1: 10.0, 2: 10.0, 3: 10.0, 4: 10.0,
        5: (11.0, 7.5),  # *** footnote
        6: 16.0, 7: 14.0, 8: 16.0, 9: 14.0, 10: 16.0,
        11: 22.0, 12: 18.0, 13: 20.0, 14: 22.0,
        15: 26.0, 16: 26.0, 17: 30.0, 18: 26.0,
        19: 34.0, 20: 34.0, 21: 38.0, 22: 38.0,
        "U Octave": 8.0, "L Octave": 8.0, "High F#": 10.0,
    },

    # ===== ALTO =====
    "Selmer (Paris) Mark VI (Alto)": {
        1: 18.5, 2: 18.5, 3: 18.5, 4: 18.5, 5: 14.0,
        6: 24.0, 7: 24.0, 8: 24.0, 9: 24.0, 10: 24.0,
        11: 30.0, 12: 28.0, 13: 28.0, 14: 30.0,
        15: 34.0, 16: 36.0, 17: 42.0, 18: 40.0,
        19: 40.0, 20: 38.0, 21: 44.0, 22: 44.0,
        "U Octave": 9.0, "L Octave": 9.0, "High F#": 18.5,
    },
    "Selmer (Paris) Mark VII (Alto)": {
        1: 18.5, 2: 18.5, 3: 18.5, 4: 18.5, 5: 14.0,
        6: 24.0, 7: 24.0, 8: 24.0, 9: 24.0, 10: 24.0,
        11: 30.0, 12: 28.0, 13: 28.0, 14: 30.0,
        15: 34.0, 16: 36.0, 17: 42.0, 18: 40.0,
        19: 38.0, 20: 38.0, 21: 48.0, 22: 48.0,
        "U Octave": 9.0, "L Octave": 9.0, "High F#": 18.5,
    },
    "Selmer (Paris) SA-80 (Alto)": {
        1: 18.5, 2: 18.5, 3: 18.5, 4: 18.5, 5: 14.5,
        6: 27.0, 7: 24.0, 8: 24.0, 9: 24.0, 10: 27.0,
        11: 30.0, 12: 32.0, 13: 30.0, 14: 30.0,
        15: 34.0, 16: 36.0, 17: 42.0, 18: 40.0,
        19: 38.0, 20: 40.0, 21: 48.0, 22: 48.0,
        "U Octave": 9.0, "L Octave": 9.0, "High F#": 18.5,
    },
    "Selmer USA 162/AS100 (Alto)": {
        1: 18.5, 2: 18.5, 3: 18.5, 4: 18.5, 5: 14.5,
        6: 25.1, 7: 25.1, 8: 23.5, 9: 25.1, 10: 25.1,
        11: 30.2, 12: 31.7, 13: 30.2, 14: 30.2,
        15: 34.0, 16: 37.0, 17: 41.5, 18: 38.5,
        19: 38.5, 20: 38.5, 21: 48.0, 22: 48.0,
        "U Octave": 9.3, "L Octave": 9.3, "High F#": 18.5,
    },
    "Selmer USA 142F/1242/AS200 (Alto)": {
        1: 16.5, 2: 18.5, 3: 18.5, 4: 16.5, 5: 14.5,
        6: 25.1, 7: 27.0, 8: 23.5, 9: 25.1, 10: 25.1,
        11: 30.2, 12: 27.0, 13: 27.0, 14: 25.1,
        15: 34.0, 16: 34.0, 17: 41.5, 18: 41.5,
        19: 41.5, 20: 38.5, 21: 45.5, 22: 45.5,
        "U Octave": 9.3, "L Octave": 9.3, "High F#": 16.5,
    },

    # ===== TENOR =====
    "Selmer (Paris) Mark VI (Tenor)": {
        1: 18.5, 2: 20.0, 3: 20.0, 4: 20.0, 5: 18.5,
        6: 34.0, 7: 32.0, 8: 32.0, 9: 23.0, 10: 32.0,
        11: 36.0, 12: 36.0, 13: 38.0, 14: 42.0,
        15: 42.0, 16: 42.0, 17: 48.0, 18: 40.0,
        19: 44.0, 20: 40.0, 21: 52.0, 22: 52.0,
        "U Octave": 9.0, "L Octave": 9.0, "High F#": 20.0,
    },
    "Selmer (Paris) Mark VII (Tenor)": {
        1: 20.0, 2: 20.0, 3: 20.0, 4: 20.0, 5: 18.5,
        6: 34.0, 7: 32.0, 8: 32.0, 9: 23.0, 10: 30.0,
        11: 36.0, 12: 36.0, 13: 38.0, 14: 42.0,
        15: 42.0, 16: 42.0, 17: 48.0, 18: 40.0,
        19: 44.0, 20: 40.0, 21: 52.0, 22: 52.0,
        "U Octave": 9.0, "L Octave": 9.0, "High F#": 20.0,
    },
    "Selmer (Paris) SA-80 (Tenor)": {
        1: 20.0, 2: 20.0, 3: 20.0, 4: 20.0, 5: 18.5,
        6: 34.0, 7: 32.0, 8: 32.0, 9: 23.0, 10: 30.0,
        11: 38.0, 12: 36.0, 13: 38.0, 14: 38.0,
        15: 42.0, 16: 42.0, 17: 48.0, 18: 40.0,
        19: 44.0, 20: 44.0, 21: 52.0, 22: 52.0,
        "U Octave": 9.0, "L Octave": 9.0, "High F#": 20.0,
    },
    "Selmer USA 164/TS100 (Tenor)": {
        1: 20.0, 2: 20.0, 3: 20.0, 4: 20.0, 5: 18.5,
        6: 34.0, 7: 31.7, 8: 31.7, 9: 28.0, 10: 30.2,
        11: 38.5, 12: 37.0, 13: 38.5, 14: 38.5,
        15: 41.5, 16: 41.5, 17: 48.0, 18: 41.5,
        19: 45.4, 20: 45.4, 21: 52.7, 22: 52.7,
        "U Octave": 9.3, "L Octave": 9.3, "High F#": 20.0,
    },
    "Selmer USA 144F/1244/TS200 (Tenor)": {
        1: 18.5, 2: 18.5, 3: 18.5, 4: 20.0, 5: 18.5,
        6: 34.0, 7: 30.2, 8: 31.7, 9: 27.0, 10: 31.7,
        11: 38.5, 12: 34.0, 13: 38.5, 14: 37.0,
        15: 41.5, 16: 41.5, 17: 48.0, 18: 41.5,
        19: 45.4, 20: 41.5, 21: 49.7, 22: 49.7,
        "U Octave": 9.3, "L Octave": 9.3, "High F#": 18.5,
    },

    # ===== BARITONE =====
    "Selmer (Paris) Mark VI (Baritone)": {
        1: 24.0, 2: 24.0, 3: 30.0, 4: 28.0, 5: 30.0,
        6: 36.0, 7: 34.0, 8: 32.0, 9: 38.0, 10: 38.0,
        11: 40.0, 12: 42.0, 13: 44.0, 14: 30.0,
        15: 46.0, 16: 46.0, 17: 52.0, 18: 44.0,
        19: 54.0, 20: 46.0, 21: 60.0, 22: 60.0,
        23: 64.0,  # Low A
        "U Octave": 9.0, "L Octave": 11.0, "High F#": None,
        "Waterkey": 11.0,
    },
    "Selmer (Paris) SA-80 (Baritone)": {
        1: 24.0, 2: 24.0, 3: 28.0, 4: 28.0, 5: 26.0,
        6: 36.0, 7: 34.0, 8: 32.0, 9: 38.0, 10: 38.0,
        11: 40.0, 12: 42.0, 13: 40.0, 14: 30.0,
        15: 46.0, 16: 40.0, 17: 52.0, 18: 44.0,
        19: 48.0, 20: 48.0, 21: 60.0, 22: 60.0,
        23: 64.0,  # Low A
        "U Octave": 9.0, "L Octave": None, "High F#": 20.0,
        "Waterkey": 11.0,
    },
    "Selmer USA 1256 (Baritone)": {
        1: 30.2, 2: 27.0, 3: 31.7, 4: 30.2, 5: 31.7,
        6: 37.0, 7: 25.1, 8: 37.0, 9: 37.0, 10: 41.5,
        11: (27.0, 41.5),  # **** footnote
        12: 48.0, 13: 48.0, 14: 27.0,
        15: 48.0, 16: 48.0, 17: 49.7, 18: 48.0,
        19: 49.7, 20: 49.7, 21: 58.8, 22: 58.5,
        "U Octave": 9.3, "L Octave": None, "High F#": None,
        "Waterkey": 11.0,
    },
    "Selmer USA 156 (Baritone)": {
        1: 30.2, 2: 27.0, 3: 31.7, 4: 30.2, 5: 31.7,
        6: 37.0, 7: 25.1, 8: 37.0, 9: 37.0, 10: 41.0,
        11: (27.0, 41.5),  # **** footnote
        12: 48.0, 13: 48.0, 14: 27.0,
        15: 48.0, 16: 48.0, 17: 49.7, 18: 48.0,
        19: 49.7, 20: 49.7, 21: 58.8, 22: 58.5,
        23: 58.5,  # Low A
        "U Octave": 9.3, "L Octave": 9.3, "High F#": None,
        "Waterkey": 11.0,
    },
}


def model_to_preset(name, key_data):
    """Convert per-key data to size x qty preset format."""
    sizes = []
    for key, value in key_data.items():
        if value is None:
            continue
        if isinstance(value, tuple):
            sizes.extend(value)
        else:
            sizes.append(value)

    # Count occurrences of each size, sorted descending
    counts = Counter(sizes)
    lines = []
    for size in sorted(counts.keys(), reverse=True):
        qty = counts[size]
        if size == int(size):
            size_fmt = f"{int(size)}.0"
        else:
            size_fmt = f"{size}"
        lines.append(f"{size_fmt} x {qty}")

    return "\n".join(lines), len(sizes)


def main():
    presets = {}
    print("Selmer Catalog Pad Sets\n")

    for name, key_data in MODELS.items():
        pads_str, total = model_to_preset(name, key_data)
        presets[name] = {
            "pads": pads_str,
            "notes": "Selmer factory catalog, early 1980s"
        }
        print(f"{name}: {total} pads")

        # Show the preset
        for line in pads_str.split("\n"):
            print(f"  {line}")
        print()

    # Save output
    output = {"Selmer USA Catalog 1980s": presets}
    output_file = "tools/selmer_catalog_presets.json"
    with open(output_file, "w") as f:
        json.dump(output, f, indent=2)

    print(f"Output saved to {output_file}")
    print(f"Total: {len(presets)} presets")


if __name__ == "__main__":
    main()
