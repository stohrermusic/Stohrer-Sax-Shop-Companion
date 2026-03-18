"""
Horn Spread Analysis Tool

Analyzes all tone profiles to find the actual statistical spread of
each descriptor across different horns. This tells us the real-world
range of variation and helps calibrate the gauge scaling.

Usage:
    python tools/analyze_horn_spread.py

Best used with a controlled dataset: same player, same mouthpiece,
different horns. The variation then represents horn character, not
player or setup differences.

Outputs:
- Per-descriptor statistics (min, max, mean, stddev, range)
- Per-descriptor ranking of all profiles
- Suggested gauge scaling based on actual data spread
- Per-note analysis showing which notes vary most between horns
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import math
from toner_engine import (
    load_tone_profiles, compute_fingerprint, MIN_PROFILE_NOTES,
)
from config import TONE_PROFILES_FILE


def main():
    print("=" * 65)
    print("HORN SPREAD ANALYSIS")
    print("=" * 65)

    profiles = load_tone_profiles(TONE_PROFILES_FILE)
    if not profiles:
        print("\nNo profiles found.")
        return

    # Build fingerprints for all profiles with data
    fingerprints = []
    for lib_name, lib_profiles in profiles.items():
        if not isinstance(lib_profiles, dict):
            continue
        for prof_name, prof in lib_profiles.items():
            sessions = prof.get('sessions', [])
            if not sessions:
                continue
            fp = compute_fingerprint(sessions)
            if fp['capture_count'] == 0:
                continue
            fp['_name'] = prof_name
            fp['_lib'] = lib_name
            fp['_profile'] = prof
            fingerprints.append(fp)

    if not fingerprints:
        print("\nNo profiles with capture data found.")
        return

    print(f"\nProfiles analyzed: {len(fingerprints)}")
    for fp in fingerprints:
        p = fp['_profile']
        print(f"  [{fp['_lib']}] {fp['_name']}")
        print(f"    {p.get('horn_type', '?')} | {p.get('horn_make', '')} {p.get('horn_model', '')}"
              f" | {fp['note_count']} notes, {fp['capture_count']} captures")

    # ============================================================
    print(f"\n{'=' * 65}")
    print("DESCRIPTOR SPREAD (Horn Average)")
    print(f"{'=' * 65}")

    desc_keys = ['resonance', 'richness', 'brightness', 'darkness', 'fullness']

    # Collect values
    desc_data = {}
    for key in desc_keys:
        values = [(fp['_name'], fp['descriptors'].get(key, 0))
                  for fp in fingerprints]
        values.sort(key=lambda x: x[1])
        desc_data[key] = values

    # Summary table
    print(f"\n{'Descriptor':>12} {'Min':>7} {'Max':>7} {'Mean':>7} {'StdDev':>7} {'Range':>7}")
    print("-" * 55)
    for key in desc_keys:
        vals = [v for _, v in desc_data[key]]
        if not vals:
            continue
        mn = min(vals)
        mx = max(vals)
        mean = sum(vals) / len(vals)
        if len(vals) > 1:
            variance = sum((v - mean) ** 2 for v in vals) / (len(vals) - 1)
            stddev = math.sqrt(variance)
        else:
            stddev = 0.0
        spread = mx - mn
        print(f"{key:>12} {mn:7.1%} {mx:7.1%} {mean:7.1%} {stddev:7.1%} {spread:7.1%}")

    # Rankings per descriptor
    print(f"\n{'=' * 65}")
    print("RANKINGS (lowest to highest)")
    print(f"{'=' * 65}")

    for key in desc_keys:
        print(f"\n  {key.upper()}:")
        for name, val in desc_data[key]:
            bar_len = int(val * 40)
            bar = "#" * bar_len + "." * (40 - bar_len)
            print(f"    {val:5.1%} {bar} {name[:30]}")

    # ============================================================
    print(f"\n{'=' * 65}")
    print("PER-NOTE VARIATION")
    print("Which notes show the most variation between horns?")
    print(f"{'=' * 65}")

    # Collect all notes across all profiles
    all_notes = set()
    for fp in fingerprints:
        all_notes.update(fp.get('per_note', {}).keys())

    # Sort chromatically
    pitch_classes = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
    note_order = []
    for octave in range(1, 8):
        for pc in pitch_classes:
            n = f"{pc}{octave}"
            if n in all_notes:
                note_order.append(n)

    if note_order:
        print(f"\n{'Note':>6}", end="")
        for key in desc_keys:
            print(f" {key[:5]:>7}", end="")
        print("  profiles")
        print("-" * 55)

        note_variations = []
        for note in note_order:
            # Collect descriptor values for this note across profiles
            note_descs = {}
            count = 0
            for fp in fingerprints:
                pn = fp.get('per_note', {}).get(note)
                if pn and pn.get('descriptors'):
                    count += 1
                    for key in desc_keys:
                        if key not in note_descs:
                            note_descs[key] = []
                        note_descs[key].append(pn['descriptors'].get(key, 0))

            if count < 2:
                continue  # Need at least 2 profiles to compare

            # Compute range for each descriptor
            ranges = {}
            for key in desc_keys:
                vals = note_descs.get(key, [])
                ranges[key] = max(vals) - min(vals) if vals else 0

            total_variation = sum(ranges.values())
            note_variations.append((note, ranges, count, total_variation))

            print(f"{note:>6}", end="")
            for key in desc_keys:
                r = ranges.get(key, 0)
                print(f" {r:7.1%}", end="")
            print(f"  {count}")

        # Most variable notes
        if note_variations:
            note_variations.sort(key=lambda x: x[3], reverse=True)
            print(f"\nMost variable notes (total descriptor spread):")
            for note, ranges, count, total in note_variations[:5]:
                top_desc = max(ranges.items(), key=lambda x: x[1])
                print(f"  {note}: total spread {total:.0%} "
                      f"(biggest: {top_desc[0]} at {top_desc[1]:.0%}, "
                      f"across {count} horns)")

    # ============================================================
    print(f"\n{'=' * 65}")
    print("GAUGE SCALING SUGGESTIONS")
    print(f"{'=' * 65}")

    print("\nBased on the actual spread in this dataset, here's where the")
    print("gauge needles would sit if scaled to use the full range:")

    for key in desc_keys:
        vals = [v for _, v in desc_data[key]]
        if len(vals) < 2:
            continue
        mn = min(vals)
        mx = max(vals)
        spread = mx - mn

        if spread < 0.05:
            print(f"\n  {key}: All profiles read similarly ({mn:.0%}-{mx:.0%}).")
            print(f"    Not enough variation to calibrate. Need more diverse horns,")
            print(f"    or this descriptor may not differentiate horns well.")
        else:
            print(f"\n  {key}: Range {mn:.0%} to {mx:.0%} (spread {spread:.0%})")
            # Suggest a mapping that uses 10%-90% of gauge range
            # for the observed data spread
            margin = spread * 0.2
            suggested_lo = max(0, mn - margin)
            suggested_hi = min(1, mx + margin)
            print(f"    Suggested gauge range: {suggested_lo:.0%} to {suggested_hi:.0%}")
            print(f"    This would map the observed spread to ~20%-80% of the gauge,")
            print(f"    leaving room for horns outside this dataset.")

            # Show where each horn would sit on the rescaled gauge
            for name, val in desc_data[key]:
                if suggested_hi > suggested_lo:
                    rescaled = (val - suggested_lo) / (suggested_hi - suggested_lo)
                else:
                    rescaled = 0.5
                rescaled = max(0, min(1, rescaled))
                bar_pos = int(rescaled * 30)
                bar = "." * bar_pos + "#" + "." * (30 - bar_pos)
                print(f"      {bar} {name[:25]} ({val:.0%})")

    # ============================================================
    print(f"\n{'=' * 65}")
    print("GROUPING ANALYSIS")
    print("Do horns cluster by type, make, or era?")
    print(f"{'=' * 65}")

    # Group by horn type
    by_type = {}
    for fp in fingerprints:
        ht = fp['_profile'].get('horn_type', 'Unknown')
        if ht not in by_type:
            by_type[ht] = []
        by_type[ht].append(fp)

    if len(by_type) > 1:
        print(f"\nBy horn type:")
        for ht, fps in sorted(by_type.items()):
            if len(fps) < 1:
                continue
            print(f"\n  {ht} ({len(fps)} profiles):")
            for key in desc_keys:
                vals = [fp['descriptors'].get(key, 0) for fp in fps]
                mean = sum(vals) / len(vals)
                print(f"    {key:>12}: avg {mean:.0%}", end="")
                if len(vals) > 1:
                    spread = max(vals) - min(vals)
                    print(f"  (spread {spread:.0%})", end="")
                print()

    # Group by make
    by_make = {}
    for fp in fingerprints:
        make = fp['_profile'].get('horn_make', 'Unknown')
        if make not in by_make:
            by_make[make] = []
        by_make[make].append(fp)

    if len(by_make) > 1:
        print(f"\nBy manufacturer:")
        for make, fps in sorted(by_make.items()):
            if not make or len(fps) < 1:
                continue
            descs = {key: sum(fp['descriptors'].get(key, 0) for fp in fps) / len(fps)
                     for key in desc_keys}
            desc_str = ", ".join(f"{k[:5]}={v:.0%}" for k, v in descs.items())
            print(f"  {make} ({len(fps)}): {desc_str}")

    print()


if __name__ == "__main__":
    main()
