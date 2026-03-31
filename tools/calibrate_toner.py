"""
Toner Calibration Tool

Scans tone profiles for expert annotations in the notes field,
compares them against computed descriptors, and reports on alignment.

Usage:
    python tools/calibrate_toner.py

Write annotations in profile notes using natural language:
    "rich horn", "very bright", "dark and warm", "pure tone",
    "considered bright by most players", "full sound", etc.

The tool extracts keywords, matches them to descriptors, and shows:
- Whether the engine agrees with the expert assessment
- Where the scaling is too compressed or too spread
- Suggested threshold adjustments

Over time, as more profiles are annotated, the suggestions get
more reliable.
"""

import sys
import os
import re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from toner_engine import (
    load_tone_presets, compute_fingerprint, MIN_PRESET_NOTES,
)
from config import TONE_PRESETS_FILE

# ============================================================
# Keyword patterns → descriptor mapping
# ============================================================

# Each entry: (regex_pattern, descriptor_key, expected_direction)
# direction: "high" means the expert says this descriptor should read high,
#            "low" means it should read low
ANNOTATION_PATTERNS = [
    # Richness
    (r'\brich\b', 'richness', 'high'),
    (r'\bvery rich\b', 'richness', 'very_high'),
    (r'\bpure\b', 'richness', 'low'),
    (r'\bvery pure\b', 'richness', 'very_low'),
    (r'\bthin\b', 'richness', 'low'),

    # Brightness
    (r'\bbright\b', 'brightness', 'high'),
    (r'\bvery bright\b', 'brightness', 'very_high'),
    (r'\bnot bright\b', 'brightness', 'low'),
    (r'\bedgy\b', 'brightness', 'high'),
    (r'\bcutting\b', 'brightness', 'high'),
    (r'\bpiercing\b', 'brightness', 'very_high'),

    # Darkness
    (r'\bdark\b', 'darkness', 'high'),
    (r'\bvery dark\b', 'darkness', 'very_high'),
    (r'\bwarm\b', 'darkness', 'high'),
    (r'\bmellow\b', 'darkness', 'high'),
    (r'\bnot dark\b', 'darkness', 'low'),

    # Fullness
    (r'\bfull\b', 'fullness', 'high'),
    (r'\bfull sound\b', 'fullness', 'very_high'),
    (r'\bbig sound\b', 'fullness', 'high'),
    (r'\bhollow\b', 'fullness', 'low'),
    (r'\bthin\b', 'fullness', 'low'),

    # Resonance
    (r'\bresonant\b', 'resonance', 'high'),
    (r'\bstuffy\b', 'resonance', 'low'),
    (r'\bfree.?blowing\b', 'resonance', 'high'),
    (r'\bresistant\b', 'resonance', 'low'),
]

# Expected value ranges for each direction
DIRECTION_RANGES = {
    'very_high': (0.7, 1.0),
    'high': (0.45, 1.0),
    'low': (0.0, 0.35),
    'very_low': (0.0, 0.15),
}


def extract_annotations(notes_text):
    """Extract descriptor expectations from free-text notes."""
    if not notes_text:
        return []

    text = notes_text.lower()
    found = []
    for pattern, descriptor, direction in ANNOTATION_PATTERNS:
        if re.search(pattern, text):
            found.append((descriptor, direction, pattern))
    return found


def main():
    print("=" * 60)
    print("TONER CALIBRATION REPORT")
    print("=" * 60)

    profiles = load_tone_presets(TONE_PRESETS_FILE)
    if not profiles:
        print("\nNo profiles found.")
        return

    # Collect all annotated profiles with enough data
    annotated = []
    total_profiles = 0
    total_annotated = 0
    total_with_data = 0

    for lib_name, lib_profiles in profiles.items():
        if not isinstance(lib_profiles, dict):
            continue
        for prof_name, prof in lib_profiles.items():
            total_profiles += 1
            notes = prof.get('notes', '')
            annotations = extract_annotations(notes)

            sessions = prof.get('sessions', [])
            fp = compute_fingerprint(sessions) if sessions else None
            has_data = fp and fp['capture_count'] > 0

            if has_data:
                total_with_data += 1

            if annotations and has_data:
                total_annotated += 1
                annotated.append({
                    'lib': lib_name,
                    'name': prof_name,
                    'notes': notes,
                    'annotations': annotations,
                    'fingerprint': fp,
                })

    print(f"\nProfiles: {total_profiles} total, {total_with_data} with capture data, "
          f"{total_annotated} annotated with tone descriptions")

    if not annotated:
        print("\nNo profiles have both capture data AND tone annotations in their notes.")
        print("Add descriptions like 'rich horn', 'bright', 'dark and warm' to profile notes,")
        print("then run this tool again.")
        return

    # ============================================================
    # Analyze each annotated profile
    # ============================================================
    print(f"\n{'='*60}")
    print("PROFILE-BY-PROFILE ANALYSIS")
    print(f"{'='*60}")

    agreements = 0
    disagreements = 0
    per_descriptor = {}  # descriptor -> list of (expected_direction, actual_value)

    for entry in annotated:
        fp = entry['fingerprint']
        desc = fp['descriptors']
        print(f"\n  [{entry['lib']}] {entry['name']}")
        print(f"  Notes: \"{entry['notes']}\"")
        print(f"  Data: {fp['note_count']} notes, {fp['capture_count']} captures")

        for descriptor, direction, pattern in entry['annotations']:
            actual = desc.get(descriptor, 0)
            lo, hi = DIRECTION_RANGES[direction]
            match = lo <= actual <= hi

            if descriptor not in per_descriptor:
                per_descriptor[descriptor] = []
            per_descriptor[descriptor].append((direction, actual))

            status = "OK" if match else "MISMATCH"
            if match:
                agreements += 1
            else:
                disagreements += 1

            arrow = ""
            if not match:
                if actual < lo:
                    arrow = f" (engine reads {actual:.0%}, expected >{lo:.0%})"
                else:
                    arrow = f" (engine reads {actual:.0%}, expected <{hi:.0%})"

            print(f"    {status}: '{descriptor}' should be {direction} "
                  f"=> actual {actual:.0%}{arrow}")

    # ============================================================
    # Summary statistics
    # ============================================================
    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")

    total_checks = agreements + disagreements
    pct = agreements / total_checks * 100 if total_checks > 0 else 0
    print(f"\nAgreement: {agreements}/{total_checks} ({pct:.0f}%)")

    if disagreements > 0:
        print(f"\n  Mismatches suggest the scaling constants need adjustment.")
    else:
        print(f"\n  Engine agrees with all expert annotations!")

    # Per-descriptor breakdown
    print(f"\n{'='*60}")
    print("PER-DESCRIPTOR ANALYSIS")
    print(f"{'='*60}")

    for descriptor, entries in sorted(per_descriptor.items()):
        high_vals = [v for d, v in entries if d in ('high', 'very_high')]
        low_vals = [v for d, v in entries if d in ('low', 'very_low')]

        print(f"\n  {descriptor.upper()}:")
        if high_vals:
            avg_h = sum(high_vals) / len(high_vals)
            print(f"    Expert says HIGH ({len(high_vals)} profiles): "
                  f"engine avg = {avg_h:.0%}  "
                  f"(range {min(high_vals):.0%} - {max(high_vals):.0%})")
        if low_vals:
            avg_l = sum(low_vals) / len(low_vals)
            print(f"    Expert says LOW ({len(low_vals)} profiles): "
                  f"engine avg = {avg_l:.0%}  "
                  f"(range {min(low_vals):.0%} - {max(low_vals):.0%})")

        if high_vals and low_vals:
            separation = (sum(high_vals) / len(high_vals)) - (sum(low_vals) / len(low_vals))
            if separation < 0.15:
                print(f"    WARNING: Low separation ({separation:.0%}) between high and low. "
                      f"The gauge isn't distinguishing well.")
            elif separation > 0.4:
                print(f"    Good separation ({separation:.0%}) between high and low.")
            else:
                print(f"    Moderate separation ({separation:.0%}).")

    # ============================================================
    # Actionable suggestions
    # ============================================================
    print(f"\n{'='*60}")
    print("SUGGESTIONS")
    print(f"{'='*60}")

    any_suggestions = False
    for descriptor, entries in sorted(per_descriptor.items()):
        high_vals = [v for d, v in entries if d in ('high', 'very_high')]
        low_vals = [v for d, v in entries if d in ('low', 'very_low')]

        if high_vals:
            avg_h = sum(high_vals) / len(high_vals)
            if avg_h < 0.4:
                print(f"\n  {descriptor}: Expert-tagged 'high' profiles average only {avg_h:.0%}.")
                print(f"    => The scaling multiplier for {descriptor} may need to increase.")
                any_suggestions = True
            elif avg_h > 0.95:
                print(f"\n  {descriptor}: Expert-tagged 'high' profiles are pegged at {avg_h:.0%}.")
                print(f"    => The scaling may be too aggressive (everything reads high).")
                any_suggestions = True

        if low_vals:
            avg_l = sum(low_vals) / len(low_vals)
            if avg_l > 0.4:
                print(f"\n  {descriptor}: Expert-tagged 'low' profiles average {avg_l:.0%}.")
                print(f"    => The gauge floor is too high for {descriptor}.")
                any_suggestions = True

    if not any_suggestions:
        print("\n  No scaling adjustments suggested at this time.")
        print("  Add more annotated profiles for better calibration data.")

    print()


if __name__ == "__main__":
    main()
