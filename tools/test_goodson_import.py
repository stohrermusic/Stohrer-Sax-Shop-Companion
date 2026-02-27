"""
Test that the Steve Goodson pad preset data is valid and importable.
Validates JSON structure, pad sizes, and round-trips through parse_pad_list.
"""

import json
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def parse_pad_list(pad_input):
    """Simplified version of main.py's parse_pad_list for testing."""
    pad_list = []
    for line in pad_input.strip().splitlines():
        line = line.strip().lower()
        if not line:
            continue
        parts = line.split('x', 1)
        if len(parts) != 2:
            continue
        try:
            size = float(parts[0].strip())
            if size <= 0:
                continue
            qty_str = parts[1].strip()
            pad_list.append({'size': size, 'qty': int(float(qty_str))})
        except ValueError:
            continue
    return pad_list


def main():
    passed = 0
    failed = 0

    def check(name, condition, detail=""):
        nonlocal passed, failed
        if condition:
            passed += 1
        else:
            failed += 1
            msg = f"FAIL: {name}"
            if detail:
                msg += f" - {detail}"
            print(msg)

    # Load website JSON
    website_path = "C:/code/stohrermusic/static/data/pad_presets.json"
    with open(website_path) as f:
        data = json.load(f)

    # Check structure
    check("Steve Goodson library exists", "Steve Goodson" in data)
    check("Is a dict", isinstance(data["Steve Goodson"], dict))

    lib = data["Steve Goodson"]
    check("Has 83 presets", len(lib) == 83, f"got {len(lib)}")

    # Check each preset
    for name, preset in lib.items():
        check(f"{name}: has 'pads' key", "pads" in preset)
        check(f"{name}: has 'notes' key", "notes" in preset)
        check(f"{name}: pads is string", isinstance(preset.get("pads"), str))
        check(f"{name}: pads not empty", len(preset.get("pads", "")) > 0)

        # Parse through parse_pad_list
        pads = parse_pad_list(preset["pads"])
        check(f"{name}: parse_pad_list returns data", len(pads) > 0,
              f"got {len(pads)} pads from: {preset['pads'][:50]}")

        # Count total pads
        total = sum(p["qty"] for p in pads)
        check(f"{name}: reasonable pad count (15-35)", 15 <= total <= 35,
              f"got {total}")

        # Check all sizes are reasonable
        for p in pads:
            check(f"{name}: size {p['size']} in range 5-80",
                  5 <= p["size"] <= 80,
                  f"size {p['size']} out of range")
            check(f"{name}: qty {p['qty']} in range 1-10",
                  1 <= p["qty"] <= 10,
                  f"qty {p['qty']} out of range for size {p['size']}")

        # Verify round-trip: pads string -> parse -> same total
        lines = preset["pads"].strip().split("\n")
        expected_total = 0
        for line in lines:
            parts = line.strip().split(" x ")
            if len(parts) == 2:
                expected_total += int(parts[1])
        check(f"{name}: round-trip total matches", total == expected_total,
              f"parsed {total} vs expected {expected_total}")

    # Check no duplicate names
    names = list(lib.keys())
    check("No duplicate preset names", len(names) == len(set(names)))

    # Check other libraries still intact
    check("Music Center library intact", "Music Center" in data)
    check("Elkhart library intact", "Elkhart" in data)
    check("Stohrer library intact", "Stohrer" in data)
    check("Music Center has presets", len(data["Music Center"]) > 0)

    # Spot-check a known entry
    if "Conn 10M Tenor" in lib:
        pads = parse_pad_list(lib["Conn 10M Tenor"]["pads"])
        sizes = [p["size"] for p in pads]
        check("Conn 10M: has size 10.0", 10.0 in sizes)
        check("Conn 10M: has size 51.0", 51.0 in sizes)
        total = sum(p["qty"] for p in pads)
        check("Conn 10M: 23 total pads", total == 23, f"got {total}")

    # Spot-check auto-fixed YTS-82Z
    if "Yamaha YTS-82Z Tenor" in lib:
        pads = parse_pad_list(lib["Yamaha YTS-82Z Tenor"]["pads"])
        sizes = [p["size"] for p in pads]
        check("YTS-82Z: has 18.5 (auto-fixed from 185)", 18.5 in sizes)
        check("YTS-82Z: no size >= 100", all(s < 100 for s in sizes))

    print(f"\n{'='*50}")
    print(f"Results: {passed} passed, {failed} failed")
    if failed:
        print("SOME TESTS FAILED")
        sys.exit(1)
    else:
        print("ALL TESTS PASSED")


if __name__ == "__main__":
    main()
