#!/usr/bin/env python3
"""Saturation metric analysis for tone profiles.

Saturation = sum of H2-H12 linear amplitudes relative to the fundamental.
Higher = more harmonic energy. This is a "total harmonic richness" measure
that's independent of spectral shape (bright vs dark).

Uses fingerprint method: average per-note, then average across notes.
"""

import json
import os
import sys
import math
from collections import defaultdict

# Import toner_engine for descriptors_from_harmonics
sys.path.insert(0, 'C:/code/saxshopcompanion')
from toner_engine import compute_fingerprint


# ============================================
# Data loading
# ============================================

PROFILES_PATH = os.path.expandvars(r'%APPDATA%\StohrerSaxShopCompanion\toner_data.json')

SUBJECTIVE = {
    "Selmer SBA tenor 38k": "piercing ring, bright, saturated",
    "Selmer BA tenor 29k": "moderate, round, full",
    "Keilwerth Shadow Tenor": "bright but weak, unsaturated, uninteresting",
    "Conn NW2 Virtuoso Deluxe 205k": "fat Conn tone, saturated, full",
    "Couesnon Monopole II alto": "lyrical, dark, warm",
    "Conn Stretch Soprano": "warmest soprano, pure, zero brightness",
}


def load_profiles():
    with open(PROFILES_PATH, 'r') as f:
        data = json.load(f)
    return data


def is_synthetic(profile_data):
    """Return True if all captures have f0=440.0 (synthetic sample data)."""
    for session in profile_data.get('sessions', []):
        for cap in session.get('captures', []):
            f0 = cap.get('fundamental_freq', cap.get('freq', 0))
            if abs(f0 - 440.0) > 0.01:
                return False
    return True


def capture_saturation(harmonics_db):
    """Compute saturation: sum of H2-H12 linear amplitudes."""
    total = 0.0
    for i in range(1, min(12, len(harmonics_db))):
        db = harmonics_db[i]
        amp = 10.0 ** (db / 20.0)
        total += amp
    return total


def note_to_midi(note_str):
    """Convert note name like 'C#3' to a MIDI-ish number for sorting."""
    if not note_str:
        return 0
    note_map = {
        'C': 0, 'C#': 1, 'Db': 1, 'D': 2, 'D#': 3, 'Eb': 3,
        'E': 4, 'F': 5, 'F#': 6, 'Gb': 6, 'G': 7, 'G#': 8, 'Ab': 8,
        'A': 9, 'A#': 10, 'Bb': 10, 'B': 11,
    }
    # Parse note name and octave
    if len(note_str) >= 2 and note_str[1] in '#b':
        name = note_str[:2]
        octave_str = note_str[2:]
    else:
        name = note_str[0]
        octave_str = note_str[1:]
    try:
        octave = int(octave_str)
    except (ValueError, IndexError):
        return 0
    pc = note_map.get(name, 0)
    return octave * 12 + pc


# ============================================
# Analysis
# ============================================

def analyze_profile(name, profile_data):
    """Analyze a single profile. Returns dict with all metrics."""
    sax_type = profile_data.get('horn_type', 'Tenor')
    sessions = profile_data.get('sessions', [])

    # Collect all captures grouped by note
    per_note_captures = defaultdict(list)
    all_saturations = []

    for session in sessions:
        for cap in session.get('captures', []):
            note = cap.get('note', '')
            hdb = cap.get('harmonics_db', [])
            f0 = cap.get('fundamental_freq', cap.get('freq', 0))
            if not hdb or f0 <= 0:
                continue
            sat = capture_saturation(hdb)
            all_saturations.append(sat)
            per_note_captures[note].append({
                'harmonics_db': hdb,
                'fundamental_freq': f0,
                'saturation': sat,
            })

    if not per_note_captures:
        return None

    # Per-note average saturation (fingerprint method)
    per_note_sat = {}
    for note, caps in per_note_captures.items():
        per_note_sat[note] = sum(c['saturation'] for c in caps) / len(caps)

    # Fingerprint-level saturation: average across notes
    fingerprint_sat = sum(per_note_sat.values()) / len(per_note_sat)

    # Register analysis: split notes into low/mid/high thirds
    sorted_notes = sorted(per_note_sat.keys(), key=note_to_midi)
    n = len(sorted_notes)
    if n >= 3:
        third = n // 3
        low_notes = sorted_notes[:third]
        mid_notes = sorted_notes[third:2*third]
        high_notes = sorted_notes[2*third:]
    else:
        low_notes = sorted_notes[:1]
        mid_notes = sorted_notes[1:2] if n > 1 else []
        high_notes = sorted_notes[2:] if n > 2 else []

    def mean_sat(notes):
        if not notes:
            return float('nan')
        return sum(per_note_sat[n] for n in notes) / len(notes)

    register_sat = {
        'low': mean_sat(low_notes),
        'mid': mean_sat(mid_notes),
        'high': mean_sat(high_notes),
    }

    # Within-profile stdev (measurement noise across individual captures)
    if len(all_saturations) > 1:
        mean_all = sum(all_saturations) / len(all_saturations)
        within_var = sum((s - mean_all)**2 for s in all_saturations) / (len(all_saturations) - 1)
        within_std = math.sqrt(within_var)
    else:
        within_std = 0.0

    # Get existing descriptors via compute_fingerprint
    fp = compute_fingerprint(sessions, sax_type)
    descriptors = fp.get('descriptors', {})

    return {
        'name': name,
        'sax_type': sax_type,
        'fingerprint_sat': fingerprint_sat,
        'register_sat': register_sat,
        'within_std': within_std,
        'n_captures': len(all_saturations),
        'n_notes': len(per_note_sat),
        'per_note_sat': per_note_sat,
        'descriptors': descriptors,
        'all_saturations': all_saturations,
    }


def main():
    data = load_profiles()

    # Collect real profiles only
    results = []
    for lib_name, lib in data.items():
        for pname, pdata in lib.items():
            if is_synthetic(pdata):
                continue
            result = analyze_profile(pname, pdata)
            if result:
                results.append(result)

    if not results:
        print("No real profiles found!")
        return

    # Sort by saturation descending
    results.sort(key=lambda r: r['fingerprint_sat'], reverse=True)

    # ============================================
    # Section 1: Profile saturation + subjective
    # ============================================
    print("=" * 90)
    print("SATURATION ANALYSIS: Total Harmonic Energy (H2-H12 linear sum)")
    print("=" * 90)
    print()
    print(f"{'Profile':<35s} {'Type':<8s} {'Sat':>6s}  {'Notes':>5s}  {'Caps':>4s}  Subjective")
    print("-" * 90)
    for r in results:
        subj = SUBJECTIVE.get(r['name'], '(no description)')
        print(f"{r['name']:<35s} {r['sax_type']:<8s} {r['fingerprint_sat']:>6.2f}  "
              f"{r['n_notes']:>5d}  {r['n_captures']:>4d}  {subj}")
    print()

    # ============================================
    # Section 2: Register breakdown + noise/signal
    # ============================================
    print("=" * 90)
    print("REGISTER BREAKDOWN & DISCRIMINATION")
    print("=" * 90)
    print()
    print(f"{'Profile':<35s} {'Low':>6s} {'Mid':>6s} {'High':>6s}  {'Within-SD':>10s}")
    print("-" * 90)
    for r in results:
        rs = r['register_sat']
        low_s = f"{rs['low']:.2f}" if not math.isnan(rs['low']) else "  n/a"
        mid_s = f"{rs['mid']:.2f}" if not math.isnan(rs['mid']) else "  n/a"
        high_s = f"{rs['high']:.2f}" if not math.isnan(rs['high']) else "  n/a"
        print(f"{r['name']:<35s} {low_s:>6s} {mid_s:>6s} {high_s:>6s}  {r['within_std']:>10.3f}")

    # Between-profile stats (F-ratio)
    print()
    fp_sats = [r['fingerprint_sat'] for r in results]
    mean_between = sum(fp_sats) / len(fp_sats)
    if len(fp_sats) > 1:
        between_var = sum((s - mean_between)**2 for s in fp_sats) / (len(fp_sats) - 1)
        between_std = math.sqrt(between_var)
    else:
        between_std = 0.0

    within_vars = []
    for r in results:
        sats = r['all_saturations']
        if len(sats) > 1:
            m = sum(sats) / len(sats)
            v = sum((s - m)**2 for s in sats) / (len(sats) - 1)
            within_vars.append(v)
    mean_within_var = sum(within_vars) / len(within_vars) if within_vars else 0.001

    f_ratio = between_var / mean_within_var if mean_within_var > 0 else float('inf')

    print(f"Between-profile std:  {between_std:.3f}")
    print(f"Mean within-profile variance: {mean_within_var:.4f}")
    print(f"F-ratio (discrimination power): {f_ratio:.2f}")
    print("  (F >> 1 means saturation discriminates well between horns)")
    print()

    # ============================================
    # Section 3: Combined table with brightness, saturation, complexity
    # ============================================
    print("=" * 90)
    print("COMBINED DESCRIPTOR TABLE")
    print("=" * 90)
    print()
    print(f"{'Profile':<35s} {'Bright%':>8s} {'Complex%':>9s} {'Full%':>6s} {'Sat':>7s}  Subjective")
    print("-" * 90)
    for r in results:
        d = r['descriptors']
        bright = d.get('brightness', 0) * 100
        # richness is the "complexity" gauge
        complex_pct = d.get('richness', 0) * 100
        full = d.get('fullness', 0) * 100
        subj = SUBJECTIVE.get(r['name'], '')
        print(f"{r['name']:<35s} {bright:>7.1f}% {complex_pct:>8.1f}% {full:>5.1f}% {r['fingerprint_sat']:>7.2f}  {subj}")
    print()

    # ============================================
    # Section 4: Hypothesis checks
    # ============================================
    print("=" * 90)
    print("HYPOTHESIS CHECKS")
    print("=" * 90)
    print()

    # Build lookup
    by_name = {r['name']: r for r in results}

    checks = [
        ("Conn NW2 Virtuoso Deluxe 205k",
         "Conn alto: high saturation despite being dark?",
         lambda r: r['fingerprint_sat'],
         lambda r: r['descriptors'].get('brightness', 0)),
        ("Keilwerth Shadow Tenor",
         "Shadow: low saturation despite bright spectral shape?",
         lambda r: r['fingerprint_sat'],
         lambda r: r['descriptors'].get('brightness', 0)),
        ("Selmer SBA tenor 38k",
         "SBA: high saturation (piercing ring)?",
         lambda r: r['fingerprint_sat'],
         lambda r: r['descriptors'].get('brightness', 0)),
    ]

    # Rank order
    sat_ranked = sorted(results, key=lambda r: r['fingerprint_sat'], reverse=True)
    bright_ranked = sorted(results, key=lambda r: r['descriptors'].get('brightness', 0), reverse=True)

    for name, question, sat_fn, bright_fn in checks:
        r = by_name.get(name)
        if not r:
            print(f"  {name}: NOT FOUND")
            continue
        sat_rank = [x['name'] for x in sat_ranked].index(name) + 1
        bright_rank = [x['name'] for x in bright_ranked].index(name) + 1
        print(f"  {question}")
        print(f"    Saturation: {r['fingerprint_sat']:.2f} (rank {sat_rank}/{len(results)})")
        print(f"    Brightness: {r['descriptors'].get('brightness',0)*100:.1f}% (rank {bright_rank}/{len(results)})")
        sat_match = ""
        if "high" in question.lower() and sat_rank <= 2:
            sat_match = "YES - reads high"
        elif "low" in question.lower() and sat_rank >= len(results) - 1:
            sat_match = "YES - reads low"
        elif "high" in question.lower() and sat_rank > len(results) // 2:
            sat_match = "NO - reads low/mid"
        elif "low" in question.lower() and sat_rank <= len(results) // 2:
            sat_match = "NO - reads high/mid"
        else:
            sat_match = "MIXED - reads mid-range"
        print(f"    Verdict: {sat_match}")
        print()

    # ============================================
    # Section 5: Per-note detail for each profile
    # ============================================
    print("=" * 90)
    print("PER-NOTE SATURATION (sorted by pitch)")
    print("=" * 90)
    for r in results:
        print(f"\n  {r['name']} ({r['sax_type']}):")
        sorted_notes = sorted(r['per_note_sat'].keys(), key=note_to_midi)
        for note in sorted_notes:
            sat = r['per_note_sat'][note]
            bar = '#' * int(sat * 2)  # rough visual
            print(f"    {note:<6s} {sat:>6.2f}  {bar}")

    # ============================================
    # Section 6: Key insight summary
    # ============================================
    print()
    print("=" * 90)
    print("KEY INSIGHTS")
    print("=" * 90)
    print()

    # Does saturation correlate with brightness?
    if len(results) >= 3:
        # Spearman rank correlation (simple version)
        n = len(results)
        sat_ranks = {r['name']: i+1 for i, r in enumerate(sat_ranked)}
        bright_ranks = {r['name']: i+1 for i, r in enumerate(bright_ranked)}
        d_sq_sum = sum((sat_ranks[r['name']] - bright_ranks[r['name']])**2 for r in results)
        rho = 1 - (6 * d_sq_sum) / (n * (n**2 - 1))
        print(f"  Saturation vs Brightness rank correlation (Spearman rho): {rho:.3f}")
        if abs(rho) < 0.3:
            print("    -> Weak correlation: saturation measures something different from brightness")
        elif abs(rho) < 0.7:
            print("    -> Moderate correlation: partially overlapping but distinct")
        else:
            print("    -> Strong correlation: these metrics are measuring similar things")

    # Saturation range and spread
    print(f"\n  Saturation range: {min(fp_sats):.2f} to {max(fp_sats):.2f} (spread: {max(fp_sats)-min(fp_sats):.2f})")
    print(f"  Mean: {mean_between:.2f}, SD: {between_std:.3f}, CV: {between_std/mean_between*100:.1f}%")

    # Does saturation capture what brightness misses?
    conn = by_name.get("Conn NW2 Virtuoso Deluxe 205k")
    shadow = by_name.get("Keilwerth Shadow Tenor")
    if conn and shadow:
        print("\n  Conn alto vs Shadow tenor:")
        print(f"    Brightness: Conn {conn['descriptors'].get('brightness',0)*100:.1f}% vs Shadow {shadow['descriptors'].get('brightness',0)*100:.1f}%")
        print(f"    Saturation: Conn {conn['fingerprint_sat']:.2f} vs Shadow {shadow['fingerprint_sat']:.2f}")
        if conn['fingerprint_sat'] > shadow['fingerprint_sat']:
            print("    -> Saturation correctly captures that Conn sounds 'full/saturated' while Shadow sounds 'weak/unsaturated'")
        else:
            print("    -> Saturation does NOT distinguish these two as expected")


if __name__ == '__main__':
    main()
