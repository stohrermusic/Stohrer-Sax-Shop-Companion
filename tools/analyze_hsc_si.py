#!/usr/bin/env python3
"""
Analyze HSC (Harmonic Spectral Centroid) and SI (Spectral Irregularity) as
candidate descriptors, comparing them to the current brightness and complexity
metrics computed by toner_engine.py.

Uses actual toner_engine code for current descriptors to ensure exact match.

Research script — does not modify any source files or profiles.
"""

import sys
import os
import json
import math

# Add project root so we can import toner_engine
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

from toner_engine import (
    descriptors_from_harmonics,
    compute_fingerprint,
    load_tone_presets,
    flatten_presets,
    BREAK_FREQUENCIES,
    DEFAULT_BREAK_FREQ,
    BRIGHTNESS_HARMONIC_WEIGHTS,
    BRIGHTNESS_DB_FLOOR,
    BRIGHTNESS_DB_RANGE,
)


# ============================================
# PROFILE LOADING
# ============================================

def get_profiles_path():
    """Get the toner_data.json path for the current platform."""
    if sys.platform == 'win32':
        base = os.environ.get('APPDATA', '')
        return os.path.join(base, 'StohrerSaxShopCompanion', 'toner_data.json')
    elif sys.platform == 'darwin':
        return os.path.expanduser(
            '~/Library/Application Support/StohrerSaxShopCompanion/toner_data.json')
    else:
        base = os.environ.get('XDG_CONFIG_HOME', os.path.expanduser('~/.config'))
        return os.path.join(base, 'StohrerSaxShopCompanion', 'toner_data.json')


def load_all_profiles():
    """Load and flatten all tone profiles."""
    path = get_profiles_path()
    if not os.path.exists(path):
        print(f"ERROR: Profile file not found at {path}")
        sys.exit(1)
    print(f"Loading profiles from: {path}")
    nested = load_tone_presets(path)
    flat = flatten_presets(nested)
    print(f"Found {len(flat)} profiles across {len(nested)} libraries\n")
    return flat, nested


# ============================================
# NEW DESCRIPTOR COMPUTATION
# ============================================

def compute_hsc(harmonics_db, f0):
    """Compute Harmonic Spectral Centroid.

    Returns (hsc_harmonic_number, hsc_hz):
        hsc_harmonic_number: weighted centroid by harmonic number (unitless)
        hsc_hz: weighted centroid by frequency (Hz)
    """
    if not harmonics_db or f0 <= 0:
        return 0.0, 0.0

    num_h = 0.0
    num_hz = 0.0
    denom = 0.0

    for i, db in enumerate(harmonics_db):
        n = i + 1  # harmonic number (1 = fundamental)
        linear_amp = 10.0 ** (db / 20.0)
        freq = f0 * n
        num_h += n * linear_amp
        num_hz += freq * linear_amp
        denom += linear_amp

    if denom < 1e-15:
        return 0.0, 0.0

    return num_h / denom, num_hz / denom


def compute_si(harmonics_db):
    """Compute Spectral Irregularity.

    Average of |dB[n+1] - dB[n]| for consecutive harmonics.
    Higher = more jagged spectral envelope, lower = smoother.
    """
    if not harmonics_db or len(harmonics_db) < 2:
        return 0.0

    total = 0.0
    count = 0
    for i in range(len(harmonics_db) - 1):
        total += abs(harmonics_db[i + 1] - harmonics_db[i])
        count += 1

    return total / count if count > 0 else 0.0


# ============================================
# DATA EXTRACTION
# ============================================

def extract_captures(profiles):
    """Extract all captures with computed metrics from all profiles.

    Returns list of dicts, one per capture, and a list of profile-level summaries.
    """
    captures = []
    profile_summaries = []

    for prof_name, prof_data in profiles.items():
        if not isinstance(prof_data, dict):
            continue
        sessions = prof_data.get('sessions', [])
        if not sessions:
            continue

        sax_type = prof_data.get('horn_type', 'Tenor')

        # Check if this is a sample/test profile (all f0=440)
        all_freqs = []
        for session in sessions:
            for cap in session.get('captures', []):
                f0 = cap.get('fundamental_freq', cap.get('freq', 0))
                if f0 > 0:
                    all_freqs.append(f0)
        if all_freqs and all(abs(f - 440.0) < 1.0 for f in all_freqs):
            continue  # Skip sample profiles

        prof_captures = []
        for session in sessions:
            for cap in session.get('captures', []):
                harmonics_db = cap.get('harmonics_db', [])
                f0 = cap.get('fundamental_freq', cap.get('freq', 0))
                note = cap.get('note', '?')
                harmonic_cents = cap.get('harmonic_cents')

                if not harmonics_db or f0 <= 0:
                    continue

                # Current descriptors from toner_engine
                descs = descriptors_from_harmonics(
                    harmonics_db, f0, sax_type, harmonic_cents)

                # New metrics
                hsc_h, hsc_hz = compute_hsc(harmonics_db, f0)
                si = compute_si(harmonics_db)

                entry = {
                    'profile': prof_name,
                    'sax_type': sax_type,
                    'note': note,
                    'f0': f0,
                    'brightness': descs['brightness'],
                    'complexity': descs['richness'],
                    'hsc_h': hsc_h,
                    'hsc_hz': hsc_hz,
                    'si': si,
                    'harmonics_db': harmonics_db,
                }
                captures.append(entry)
                prof_captures.append(entry)

        if prof_captures:
            # Profile-level averages
            n = len(prof_captures)
            summary = {
                'profile': prof_name,
                'sax_type': sax_type,
                'n_captures': n,
                'brightness': sum(c['brightness'] for c in prof_captures) / n,
                'complexity': sum(c['complexity'] for c in prof_captures) / n,
                'hsc_h': sum(c['hsc_h'] for c in prof_captures) / n,
                'hsc_hz': sum(c['hsc_hz'] for c in prof_captures) / n,
                'si': sum(c['si'] for c in prof_captures) / n,
                'f0_mean': sum(c['f0'] for c in prof_captures) / n,
            }

            # Also compute fingerprint-level using toner_engine's method
            fp = compute_fingerprint(sessions, sax_type)
            summary['fp_brightness'] = fp['descriptors'].get('brightness', 0)
            summary['fp_complexity'] = fp['descriptors'].get('richness', 0)

            # Fingerprint-level HSC and SI from per-note averaged data
            fp_hsc_h_vals = []
            fp_hsc_hz_vals = []
            fp_si_vals = []
            for note_name, note_avg in fp.get('per_note', {}).items():
                if note_avg and note_avg.get('harmonics_db'):
                    h_db = note_avg['harmonics_db']
                    nf0 = note_avg.get('fundamental_freq', 0)
                    if nf0 > 0:
                        hh, hhz = compute_hsc(h_db, nf0)
                        fp_hsc_h_vals.append(hh)
                        fp_hsc_hz_vals.append(hhz)
                    fp_si_vals.append(compute_si(h_db))

            summary['fp_hsc_h'] = (sum(fp_hsc_h_vals) / len(fp_hsc_h_vals)
                                   if fp_hsc_h_vals else 0)
            summary['fp_hsc_hz'] = (sum(fp_hsc_hz_vals) / len(fp_hsc_hz_vals)
                                    if fp_hsc_hz_vals else 0)
            summary['fp_si'] = (sum(fp_si_vals) / len(fp_si_vals)
                                if fp_si_vals else 0)

            profile_summaries.append(summary)

    return captures, profile_summaries


# ============================================
# STATISTICS HELPERS
# ============================================

def pearson_r(xs, ys):
    """Compute Pearson correlation coefficient."""
    n = len(xs)
    if n < 3:
        return float('nan')
    mx = sum(xs) / n
    my = sum(ys) / n
    sx = math.sqrt(sum((x - mx) ** 2 for x in xs) / (n - 1)) if n > 1 else 0
    sy = math.sqrt(sum((y - my) ** 2 for y in ys) / (n - 1)) if n > 1 else 0
    if sx == 0 or sy == 0:
        return float('nan')
    cov = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / (n - 1)
    return cov / (sx * sy)


def f_ratio(groups):
    """Compute F-ratio (between-group variance / within-group variance).

    groups: list of lists of values, one list per group.
    """
    all_vals = [v for g in groups for v in g]
    if not all_vals:
        return 0.0
    grand_mean = sum(all_vals) / len(all_vals)

    k = len(groups)
    N = len(all_vals)

    if k < 2 or N <= k:
        return 0.0

    # Between-group sum of squares
    ss_between = sum(len(g) * (sum(g) / len(g) - grand_mean) ** 2
                     for g in groups if g)
    # Within-group sum of squares
    ss_within = sum(sum((v - sum(g) / len(g)) ** 2 for v in g)
                    for g in groups if g)

    df_between = k - 1
    df_within = N - k

    if df_within <= 0 or ss_within == 0:
        return float('inf')

    ms_between = ss_between / df_between
    ms_within = ss_within / df_within
    return ms_between / ms_within


def group_stdev(values):
    """Compute standard deviation of a list."""
    if len(values) < 2:
        return 0.0
    m = sum(values) / len(values)
    return math.sqrt(sum((v - m) ** 2 for v in values) / (len(values) - 1))


# ============================================
# NOTE/REGISTER HELPERS
# ============================================

def note_to_midi(note_str):
    """Convert note name like 'C4', 'Bb3', 'F#5' to MIDI number."""
    if not note_str or len(note_str) < 2:
        return None
    NOTE_MAP = {
        'C': 0, 'D': 2, 'E': 4, 'F': 5, 'G': 7, 'A': 9, 'B': 11
    }
    name = note_str[0].upper()
    if name not in NOTE_MAP:
        return None
    idx = 1
    accidental = 0
    if idx < len(note_str) and note_str[idx] == '#':
        accidental = 1
        idx += 1
    elif idx < len(note_str) and note_str[idx] == 'b':
        accidental = -1
        idx += 1
    try:
        octave = int(note_str[idx:])
    except (ValueError, IndexError):
        return None
    return (octave + 1) * 12 + NOTE_MAP[name] + accidental


def get_register(note_str):
    """Classify a note into low/mid/high register."""
    midi = note_to_midi(note_str)
    if midi is None:
        return 'unknown'
    # Concert pitch ranges:
    # Low: below C4 (MIDI 60)
    # Mid: C4 to B4 (MIDI 60-71)
    # High: C5 and above (MIDI 72+)
    if midi < 60:
        return 'low'
    elif midi < 72:
        return 'mid'
    else:
        return 'high'


# ============================================
# REPORT SECTIONS
# ============================================

def print_separator(title):
    print()
    print("=" * 80)
    print(f"  {title}")
    print("=" * 80)


def section_fingerprint_table(summaries):
    """Section 3: Per-profile fingerprint comparison table."""
    print_separator("FINGERPRINT-LEVEL COMPARISON (sorted by brightness)")

    # Sort by fingerprint brightness
    sorted_s = sorted(summaries, key=lambda s: s['fp_brightness'], reverse=True)

    # Header
    print(f"{'Profile':<30} {'Type':<8} {'Bright':>7} {'HSC(h#)':>8} "
          f"{'HSC(Hz)':>9} {'Complex':>8} {'SI':>7} {'Caps':>5}")
    print("-" * 95)

    for s in sorted_s:
        print(f"{s['profile']:<30} {s['sax_type']:<8} "
              f"{s['fp_brightness']:>6.1%} {s['fp_hsc_h']:>8.2f} "
              f"{s['fp_hsc_hz']:>8.0f} Hz {s['fp_complexity']:>7.1%} "
              f"{s['fp_si']:>6.2f} {s['n_captures']:>5}")

    print()
    print("Brightness = current weighted H2-H6 formula (toner_engine)")
    print("HSC(h#) = Harmonic Spectral Centroid by harmonic number (unitless)")
    print("HSC(Hz) = Harmonic Spectral Centroid in Hz")
    print("Complex = current spectral flatness formula (toner_engine)")
    print("SI = Spectral Irregularity (avg |dB step| between consecutive harmonics)")


def section_rank_comparison(summaries):
    """Section 4: Rank-order comparison."""
    print_separator("RANK-ORDER COMPARISON")

    # Brightness vs HSC
    by_bright = sorted(summaries, key=lambda s: s['fp_brightness'], reverse=True)
    by_hsc_h = sorted(summaries, key=lambda s: s['fp_hsc_h'], reverse=True)
    by_hsc_hz = sorted(summaries, key=lambda s: s['fp_hsc_hz'], reverse=True)

    print("\n--- Brightness ranking ---")
    print(f"{'Rank':<5} {'By Brightness':<30} {'By HSC(h#)':<30} {'By HSC(Hz)':<30}")
    print("-" * 95)
    for i in range(len(summaries)):
        b_name = by_bright[i]['profile'] if i < len(by_bright) else ''
        h_name = by_hsc_h[i]['profile'] if i < len(by_hsc_h) else ''
        hz_name = by_hsc_hz[i]['profile'] if i < len(by_hsc_hz) else ''
        print(f"{i+1:<5} {b_name:<30} {h_name:<30} {hz_name:<30}")

    # Complexity vs SI
    by_complex = sorted(summaries, key=lambda s: s['fp_complexity'], reverse=True)
    by_si = sorted(summaries, key=lambda s: s['fp_si'], reverse=True)

    print("\n--- Complexity ranking ---")
    print(f"{'Rank':<5} {'By Complexity':<30} {'By SI (high=jagged)':<30}")
    print("-" * 65)
    for i in range(len(summaries)):
        c_name = by_complex[i]['profile'] if i < len(by_complex) else ''
        s_name = by_si[i]['profile'] if i < len(by_si) else ''
        print(f"{i+1:<5} {c_name:<30} {s_name:<30}")


def section_correlations(captures, summaries):
    """Section 5: Correlation analysis."""
    print_separator("CORRELATION ANALYSIS (Pearson r)")

    # Per-capture correlations
    brights = [c['brightness'] for c in captures]
    complexs = [c['complexity'] for c in captures]
    hsc_hs = [c['hsc_h'] for c in captures]
    hsc_hzs = [c['hsc_hz'] for c in captures]
    sis = [c['si'] for c in captures]
    f0s = [c['f0'] for c in captures]

    print(f"\nPer-capture level (n={len(captures)} captures):")
    print(f"  Brightness  vs HSC(h#)    r = {pearson_r(brights, hsc_hs):+.4f}")
    print(f"  Brightness  vs HSC(Hz)    r = {pearson_r(brights, hsc_hzs):+.4f}")
    print(f"  Complexity  vs SI         r = {pearson_r(complexs, sis):+.4f}")
    print(f"  HSC(h#)     vs f0         r = {pearson_r(hsc_hs, f0s):+.4f}")
    print(f"  HSC(Hz)     vs f0         r = {pearson_r(hsc_hzs, f0s):+.4f}")
    print(f"  SI          vs f0         r = {pearson_r(sis, f0s):+.4f}")
    print(f"  Brightness  vs f0         r = {pearson_r(brights, f0s):+.4f}")
    print(f"  Complexity  vs f0         r = {pearson_r(complexs, f0s):+.4f}")

    # Per-profile (fingerprint) correlations
    if len(summaries) >= 3:
        fp_b = [s['fp_brightness'] for s in summaries]
        fp_c = [s['fp_complexity'] for s in summaries]
        fp_hsc_h = [s['fp_hsc_h'] for s in summaries]
        fp_hsc_hz = [s['fp_hsc_hz'] for s in summaries]
        fp_si = [s['fp_si'] for s in summaries]
        fp_f0 = [s['f0_mean'] for s in summaries]

        print(f"\nFingerprint level (n={len(summaries)} profiles):")
        print(f"  Brightness  vs HSC(h#)    r = {pearson_r(fp_b, fp_hsc_h):+.4f}")
        print(f"  Brightness  vs HSC(Hz)    r = {pearson_r(fp_b, fp_hsc_hz):+.4f}")
        print(f"  Complexity  vs SI         r = {pearson_r(fp_c, fp_si):+.4f}")
        print(f"  HSC(h#)     vs f0         r = {pearson_r(fp_hsc_h, fp_f0):+.4f}")
        print(f"  HSC(Hz)     vs f0         r = {pearson_r(fp_hsc_hz, fp_f0):+.4f}")
        print(f"  SI          vs f0         r = {pearson_r(fp_si, fp_f0):+.4f}")
        print(f"  Brightness  vs f0         r = {pearson_r(fp_b, fp_f0):+.4f}")
        print(f"  Complexity  vs f0         r = {pearson_r(fp_c, fp_f0):+.4f}")
    else:
        print("\n  (Need >= 3 profiles for fingerprint-level correlations)")

    print()
    print("Interpretation:")
    print("  |r| > 0.9 : very strong correlation (metrics measure essentially same thing)")
    print("  |r| > 0.7 : strong (substantial overlap, but some unique information)")
    print("  |r| > 0.4 : moderate (related but capturing different aspects)")
    print("  |r| < 0.4 : weak (largely independent)")
    print("  HSC(Hz) vs f0 high => HSC(Hz) is confounded by pitch register")


def section_register_analysis(captures, summaries):
    """Section 6: Per-note register analysis for HSC."""
    print_separator("REGISTER ANALYSIS — HSC BY PROFILE AND REGISTER")

    # Group captures by profile and register
    prof_register = {}
    for c in captures:
        prof = c['profile']
        reg = get_register(c['note'])
        if prof not in prof_register:
            prof_register[prof] = {}
        if reg not in prof_register[prof]:
            prof_register[prof][reg] = []
        prof_register[prof][reg].append(c)

    # Display
    profiles_sorted = sorted(prof_register.keys())
    registers = ['low', 'mid', 'high']

    print(f"\n{'Profile':<30}", end='')
    for reg in registers:
        print(f" {'HSC(h#)-' + reg:>14} {'Bright-' + reg:>12} {'n':>4}", end='')
    print()
    print("-" * 110)

    for prof in profiles_sorted:
        print(f"{prof:<30}", end='')
        for reg in registers:
            caps = prof_register[prof].get(reg, [])
            if caps:
                avg_hsc = sum(c['hsc_h'] for c in caps) / len(caps)
                avg_br = sum(c['brightness'] for c in caps) / len(caps)
                print(f" {avg_hsc:>14.2f} {avg_br:>11.1%} {len(caps):>4}", end='')
            else:
                print(f" {'---':>14} {'---':>12} {'0':>4}", end='')
        print()

    # Register effect: how much does HSC change across registers within a profile?
    print("\n--- Register effect (high minus low within profile) ---")
    print(f"{'Profile':<30} {'HSC(h#) delta':>14} {'Bright delta':>14}")
    print("-" * 60)
    for prof in profiles_sorted:
        low_caps = prof_register[prof].get('low', [])
        high_caps = prof_register[prof].get('high', [])
        if low_caps and high_caps:
            hsc_low = sum(c['hsc_h'] for c in low_caps) / len(low_caps)
            hsc_high = sum(c['hsc_h'] for c in high_caps) / len(high_caps)
            br_low = sum(c['brightness'] for c in low_caps) / len(low_caps)
            br_high = sum(c['brightness'] for c in high_caps) / len(high_caps)
            print(f"{prof:<30} {hsc_high - hsc_low:>+14.2f} {br_high - br_low:>+13.1%}")
        else:
            print(f"{prof:<30} {'(insufficient data)':>14}")

    print()
    print("If HSC(h#) is stable across registers for each profile, it is more")
    print("pitch-independent than brightness. Large register effects suggest")
    print("the metric is confounded by note register.")


def section_discrimination(captures, summaries):
    """Section 7: Discrimination power (F-ratio)."""
    print_separator("DISCRIMINATION POWER (F-ratio)")

    # Group captures by profile
    prof_groups = {}
    for c in captures:
        prof = c['profile']
        if prof not in prof_groups:
            prof_groups[prof] = []
        prof_groups[prof].append(c)

    # F-ratio for each metric
    metrics = [
        ('brightness', 'Brightness (current)'),
        ('complexity', 'Complexity (current)'),
        ('hsc_h', 'HSC by harmonic #'),
        ('hsc_hz', 'HSC in Hz'),
        ('si', 'Spectral Irregularity'),
    ]

    print(f"\n{'Metric':<25} {'F-ratio':>10} {'# groups':>10} {'# captures':>12}")
    print("-" * 60)

    for key, label in metrics:
        groups = [
            [c[key] for c in caps]
            for caps in prof_groups.values()
            if len(caps) >= 2  # need at least 2 captures per group
        ]
        if len(groups) >= 2:
            fr = f_ratio(groups)
            total_n = sum(len(g) for g in groups)
            print(f"{label:<25} {fr:>10.2f} {len(groups):>10} {total_n:>12}")
        else:
            print(f"{label:<25} {'(insufficient)':>10}")

    print()
    print("F > 1 means between-horn variance exceeds within-horn variance.")
    print("Higher F = better discrimination between different horns.")
    print("Reference: Brightness F=2.13, Richness F=1.83 (from earlier analysis).")


def section_signal_to_noise(captures, summaries):
    """Section 8: Signal-to-noise — within vs between profile stdev."""
    print_separator("SIGNAL-TO-NOISE: WITHIN-PROFILE vs BETWEEN-PROFILE STDEV")

    # Within-profile stdev
    prof_groups = {}
    for c in captures:
        prof = c['profile']
        if prof not in prof_groups:
            prof_groups[prof] = []
        prof_groups[prof].append(c)

    metrics = ['brightness', 'complexity', 'hsc_h', 'hsc_hz', 'si']
    labels = ['Brightness', 'Complexity', 'HSC(h#)', 'HSC(Hz)', 'SI']

    print(f"\n{'Metric':<14}", end='')
    print(f" {'Within-prof':>12} {'Between-prof':>13} {'Ratio(B/W)':>11} {'Interpret':>12}")
    print("-" * 65)

    for metric, label in zip(metrics, labels):
        # Within-profile stdevs
        within_stdevs = []
        for caps in prof_groups.values():
            if len(caps) >= 3:
                vals = [c[metric] for c in caps]
                within_stdevs.append(group_stdev(vals))

        # Between-profile stdev (of profile means)
        prof_means = []
        for caps in prof_groups.values():
            if len(caps) >= 3:
                prof_means.append(sum(c[metric] for c in caps) / len(caps))

        if within_stdevs and len(prof_means) >= 2:
            avg_within = sum(within_stdevs) / len(within_stdevs)
            between = group_stdev(prof_means)
            ratio = between / avg_within if avg_within > 0 else float('inf')
            quality = "good" if ratio > 1.5 else "ok" if ratio > 1.0 else "poor"
            print(f"{label:<14} {avg_within:>12.4f} {between:>13.4f} "
                  f"{ratio:>11.2f} {quality:>12}")
        else:
            print(f"{label:<14} {'(insufficient data)':>50}")

    print()
    print("Ratio > 1.5: metric varies more between horns than within — good discriminator.")
    print("Ratio ~ 1.0: as much noise within a horn as signal between horns.")
    print("Ratio < 1.0: within-horn noise dominates — poor discriminator.")


def section_per_capture_sample(captures):
    """Bonus: Show a sample of individual captures for sanity checking."""
    print_separator("SAMPLE CAPTURES (first 5 per profile)")

    prof_caps = {}
    for c in captures:
        if c['profile'] not in prof_caps:
            prof_caps[c['profile']] = []
        prof_caps[c['profile']].append(c)

    for prof, caps in sorted(prof_caps.items()):
        print(f"\n  {prof} ({len(caps)} captures):")
        print(f"  {'Note':<8} {'f0':>7} {'Bright':>7} {'HSC(h#)':>8} "
              f"{'HSC(Hz)':>9} {'Complex':>8} {'SI':>7}")
        print(f"  {'-' * 56}")
        for c in caps[:5]:
            print(f"  {c['note']:<8} {c['f0']:>7.1f} {c['brightness']:>6.1%} "
                  f"{c['hsc_h']:>8.2f} {c['hsc_hz']:>8.0f} Hz "
                  f"{c['complexity']:>7.1%} {c['si']:>6.2f}")
        if len(caps) > 5:
            print(f"  ... ({len(caps) - 5} more)")


# ============================================
# MAIN
# ============================================

def main():
    print("HSC & SI Descriptor Analysis")
    print("Comparing candidate descriptors to current brightness & complexity\n")

    flat_profiles, nested = load_all_profiles()
    captures, summaries = extract_captures(flat_profiles)

    if not captures:
        print("ERROR: No valid captures found in profiles.")
        sys.exit(1)

    print(f"Extracted {len(captures)} captures from {len(summaries)} real profiles")
    for s in summaries:
        print(f"  - {s['profile']} ({s['sax_type']}): {s['n_captures']} captures")

    # Run all analysis sections
    section_fingerprint_table(summaries)
    section_rank_comparison(summaries)
    section_correlations(captures, summaries)
    section_register_analysis(captures, summaries)
    section_discrimination(captures, summaries)
    section_signal_to_noise(captures, summaries)
    section_per_capture_sample(captures)

    print_separator("SUMMARY")
    print("""
Key questions this analysis answers:

1. Does HSC rank horns the same as current brightness?
   If so, it validates the current formula. If not, which ranking
   matches Matt's subjective experience better?

2. Is HSC(Hz) confounded by pitch register?
   High correlation with f0 means HSC(Hz) partly just tracks what
   note is being played — not useful as a horn-level descriptor.
   HSC by harmonic number should be more pitch-independent.

3. Does SI capture something different from complexity?
   Low correlation means SI provides new information. High correlation
   means it's redundant with the existing metric.

4. Which metric discriminates horns better?
   Higher F-ratio and higher signal-to-noise ratio = better at
   telling horns apart from each other.

5. Does HSC show the soprano anomaly?
   If soprano HSC reads higher than perception suggests (like current
   brightness does), then HSC has the same register bias problem.
   If HSC is stable across registers, it may solve the soprano issue.
""")


if __name__ == '__main__':
    main()
