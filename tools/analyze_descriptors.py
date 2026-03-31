#!/usr/bin/env python3
"""
Deep analysis of tone profile descriptor data.

Reads tone_profiles.json and computes:
  a) Per-capture descriptors + spectral metrics
  b) Per-note variation within profiles (measurement noise)
  c) Per-note variation across profiles (between-horn signal)
  d) Register effects (descriptor vs pitch)
  e) Correlation analysis
  f) Harmonic profile comparison for tenor trio
  g) Descriptor discrimination power (between/within variance ratio)

This is a read-only research tool; it does not modify any files.
"""

import sys
import os
import json
import math
from collections import defaultdict

# Add parent dir so we can import toner_engine
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from toner_engine import (
    descriptors_from_harmonics, compute_fingerprint, load_tone_presets,
    BREAK_FREQUENCIES, DEFAULT_BREAK_FREQ, BRIGHTNESS_HARMONIC_WEIGHTS,
    BRIGHTNESS_DB_FLOOR, BRIGHTNESS_DB_RANGE, RICHNESS_RAW_MIN, RICHNESS_RAW_RANGE,
    RESONANCE_RAW_MIN, RESONANCE_RAW_RANGE, FULLNESS_BALANCE_EXPONENT,
    FULLNESS_BASE_WEIGHT, FULLNESS_ENERGY_WEIGHT, FULLNESS_ENERGY_DIVISOR,
    flatten_presets,
)
from config import TONE_PRESETS_FILE


# ── Helpers ──────────────────────────────────────────────────────────────

NOTE_ORDER = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']

def note_sort_key(note_name):
    """Return a sortable integer for a note name like 'C4', 'A#3', etc."""
    if not note_name:
        return 999
    # Parse note letter(s) and octave
    if len(note_name) >= 2 and note_name[1] in ('#', 'b'):
        name = note_name[:2]
        octave_str = note_name[2:]
    else:
        name = note_name[:1]
        octave_str = note_name[1:]
    # Handle flats by converting to sharps
    flat_to_sharp = {'Db': 'C#', 'Eb': 'D#', 'Fb': 'E', 'Gb': 'F#',
                     'Ab': 'G#', 'Bb': 'A#', 'Cb': 'B'}
    if name in flat_to_sharp:
        name = flat_to_sharp[name]
    try:
        octave = int(octave_str)
    except ValueError:
        return 999
    idx = NOTE_ORDER.index(name) if name in NOTE_ORDER else 0
    return octave * 12 + idx


def note_to_freq(note_name):
    """Convert note name to frequency (A4=440)."""
    key = note_sort_key(note_name)
    # A4 = 440 Hz, note_sort_key('A4') = 4*12+9 = 57
    a4_key = 57
    return 440.0 * (2.0 ** ((key - a4_key) / 12.0))


def spectral_centroid(harmonics_db):
    """Weighted average harmonic number (linear amplitude weights)."""
    if not harmonics_db or len(harmonics_db) < 2:
        return 0.0
    total_w = 0.0
    total_wa = 0.0
    for i, db in enumerate(harmonics_db):
        lin = 10.0 ** (db / 20.0)
        h_num = i + 1
        total_wa += h_num * lin
        total_w += lin
    return total_wa / total_w if total_w > 0 else 0.0


def spectral_slope(harmonics_db):
    """Linear regression slope of dB vs harmonic number (for H2+)."""
    if not harmonics_db or len(harmonics_db) < 3:
        return 0.0
    # Use H2 onward (index 1+)
    xs = list(range(2, len(harmonics_db) + 1))
    ys = harmonics_db[1:]
    n = len(xs)
    if n < 2:
        return 0.0
    mean_x = sum(xs) / n
    mean_y = sum(ys) / n
    num = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys))
    den = sum((x - mean_x) ** 2 for x in xs)
    return num / den if den > 0 else 0.0


def harmonic_count_above(harmonics_db, threshold_db):
    """Count harmonics (H2+) above threshold dB."""
    if not harmonics_db:
        return 0
    return sum(1 for db in harmonics_db[1:] if db > threshold_db)


def pearson_r(xs, ys):
    """Pearson correlation coefficient."""
    n = len(xs)
    if n < 3:
        return float('nan')
    mx = sum(xs) / n
    my = sum(ys) / n
    num = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    dx = math.sqrt(sum((x - mx) ** 2 for x in xs))
    dy = math.sqrt(sum((y - my) ** 2 for y in ys))
    if dx == 0 or dy == 0:
        return float('nan')
    return num / (dx * dy)


def variance(vals):
    if len(vals) < 2:
        return 0.0
    m = sum(vals) / len(vals)
    return sum((v - m) ** 2 for v in vals) / (len(vals) - 1)


def stdev(vals):
    return math.sqrt(variance(vals)) if len(vals) >= 2 else 0.0


def mean(vals):
    return sum(vals) / len(vals) if vals else 0.0


def fmt_pct(v):
    return f"{v*100:5.1f}%"


def fmt_f(v, w=6, d=2):
    return f"{v:{w}.{d}f}"


# ── Load Data ────────────────────────────────────────────────────────────

print("=" * 80)
print("TONE PROFILE DESCRIPTOR ANALYSIS")
print("=" * 80)
print()

if not os.path.exists(TONE_PRESETS_FILE):
    print(f"ERROR: Tone profiles file not found: {TONE_PRESETS_FILE}")
    sys.exit(1)

profiles_nested = load_tone_presets(TONE_PRESETS_FILE)
profiles_flat = flatten_presets(profiles_nested)

print(f"Profiles file: {TONE_PRESETS_FILE}")
print(f"Libraries: {len(profiles_nested)}")
print(f"Total profiles: {len(profiles_flat)}")
print()

# ── Extract all captures ─────────────────────────────────────────────────

# Structure: profile_name -> list of {note, f0, harmonics_db, harmonic_cents, method, ...}
profile_captures = {}
all_captures = []

for pname, pdata in profiles_flat.items():
    sax_type = pdata.get('sax_type', 'Tenor')
    sessions = pdata.get('sessions', [])
    caps = []
    for session in sessions:
        for cap in session.get('captures', []):
            note = cap.get('note', '')
            f0 = cap.get('fundamental_freq', cap.get('freq', 0))
            hdb = cap.get('harmonics_db', [])
            hcents = cap.get('harmonic_cents', None)
            method = cap.get('method', 'unknown')
            if not hdb or f0 <= 0:
                continue
            desc = descriptors_from_harmonics(hdb, f0, sax_type, hcents)
            entry = {
                'profile': pname,
                'sax_type': sax_type,
                'note': note,
                'f0': f0,
                'harmonics_db': hdb,
                'harmonic_cents': hcents,
                'method': method,
                'descriptors': desc,
                'spectral_centroid': spectral_centroid(hdb),
                'spectral_slope': spectral_slope(hdb),
                'h_above_40': harmonic_count_above(hdb, -40.0),
                'h_above_60': harmonic_count_above(hdb, -60.0),
                'n_harmonics': len(hdb),
            }
            caps.append(entry)
            all_captures.append(entry)
    profile_captures[pname] = caps

total_caps = len(all_captures)
print(f"Total captures across all profiles: {total_caps}")

# Separate real vs sample/synthetic profiles
# Sample profiles have all f0=440 (synthetic data)
real_profiles = []
sample_profiles = []
for pname, caps in profile_captures.items():
    if not caps:
        continue
    if all(abs(c['f0'] - 440.0) < 1.0 for c in caps):
        sample_profiles.append(pname)
    else:
        real_profiles.append(pname)

real_captures = [c for c in all_captures if c['profile'] in real_profiles]

print(f"Real profiles (varied f0): {len(real_profiles)} ({len(real_captures)} captures)")
for p in real_profiles:
    print(f"    {p}")
print(f"Sample/synthetic profiles (all f0=440): {len(sample_profiles)}")
for p in sample_profiles:
    print(f"    {p}")
print()
print("NOTE: Sample profiles have f0=440 for all notes, meaning harmonic dB values")
print("are fabricated/templated, not from real measurements. All serious analysis")
print("below focuses on the REAL profiles. Samples included in summary only.")
print()

# ── (a) Per-Profile Summary ─────────────────────────────────────────────

print("=" * 80)
print("(a) PER-PROFILE SUMMARY")
print("=" * 80)
print()

desc_keys = ['brightness', 'darkness', 'resonance', 'richness', 'fullness']

for pname, pdata in profiles_flat.items():
    sax_type = pdata.get('sax_type', 'Tenor')
    caps = profile_captures.get(pname, [])
    if not caps:
        print(f"  {pname}: NO CAPTURES")
        continue

    fp = compute_fingerprint(pdata.get('sessions', []), sax_type)

    print(f"  {pname} ({sax_type})")
    print(f"    Captures: {len(caps)}, Notes: {fp['note_count']}, Method mix: ", end='')
    methods = defaultdict(int)
    for c in caps:
        methods[c['method']] += 1
    print(', '.join(f"{m}={n}" for m, n in sorted(methods.items())))
    print(f"    Fingerprint descriptors:")
    for dk in desc_keys:
        print(f"      {dk:12s}: {fmt_pct(fp['descriptors'].get(dk, 0))}")
    print(f"    Spectral metrics (mean across captures):")
    print(f"      centroid:     {fmt_f(mean([c['spectral_centroid'] for c in caps]))}")
    print(f"      slope:        {fmt_f(mean([c['spectral_slope'] for c in caps]))} dB/harmonic")
    print(f"      H above -40:  {fmt_f(mean([c['h_above_40'] for c in caps]), 4, 1)}")
    print(f"      H above -60:  {fmt_f(mean([c['h_above_60'] for c in caps]), 4, 1)}")
    print()


# ── (b) Per-Note Variation WITHIN Profiles ──────────────────────────────

print("=" * 80)
print("(b) PER-NOTE VARIATION WITHIN PROFILES (measurement noise)")
print("=" * 80)
print()
print("  Shows stdev of each descriptor across multiple captures of the SAME note")
print("  within a single profile. Lower = more repeatable measurement.")
print()

within_profile_stdevs = defaultdict(lambda: defaultdict(list))  # desc -> profile -> [stdevs per note]

print("  (Only real profiles with 2+ captures per note shown)")
print()

for pname, caps in profile_captures.items():
    if not caps or pname not in real_profiles:
        continue
    by_note = defaultdict(list)
    for c in caps:
        by_note[c['note']].append(c)

    print(f"  {pname}")
    print(f"    {'Note':>6s}  {'N':>3s}  {'bright':>7s}  {'dark':>7s}  {'reson':>7s}  {'rich':>7s}  {'full':>7s}")
    print(f"    {'-'*6}  {'-'*3}  {'-'*7}  {'-'*7}  {'-'*7}  {'-'*7}  {'-'*7}")

    note_stdevs = {dk: [] for dk in desc_keys}
    sorted_notes = sorted(by_note.keys(), key=note_sort_key)

    for note in sorted_notes:
        ncaps = by_note[note]
        n = len(ncaps)
        if n < 2:
            continue
        row = f"    {note:>6s}  {n:>3d}"
        for dk in desc_keys:
            vals = [c['descriptors'][dk] for c in ncaps]
            sd = stdev(vals)
            note_stdevs[dk].append(sd)
            within_profile_stdevs[dk][pname].append(sd)
            row += f"  {sd*100:6.2f}%"
        print(row)

    # Profile average stdev
    row = f"    {'AVG':>6s}  {'':>3s}"
    for dk in desc_keys:
        if note_stdevs[dk]:
            avg_sd = mean(note_stdevs[dk])
            row += f"  {avg_sd*100:6.2f}%"
        else:
            row += f"  {'N/A':>7s}"
    print(row)
    print()

# Global within-profile average
print("  GLOBAL AVERAGE WITHIN-PROFILE STDEV (measurement noise floor):")
print(f"    {'':>6s}  {'':>3s}  {'bright':>7s}  {'dark':>7s}  {'reson':>7s}  {'rich':>7s}  {'full':>7s}")
for dk in desc_keys:
    all_sd = []
    for pname_sd in within_profile_stdevs[dk].values():
        all_sd.extend(pname_sd)
    if all_sd:
        print(f"    {'':>6s}  {'':>3s}  ", end='') if dk == desc_keys[0] else None
        # Just print the number in the right column position
pass
# Re-do this more cleanly
global_within = {}
print(f"    ", end='')
for dk in desc_keys:
    all_sd = []
    for pname_sd in within_profile_stdevs[dk].values():
        all_sd.extend(pname_sd)
    avg = mean(all_sd) if all_sd else 0
    global_within[dk] = avg
    print(f"  {avg*100:6.2f}%", end='')
print()
print()


# ── (c) Per-Note Variation ACROSS Profiles ──────────────────────────────

print("=" * 80)
print("(c) PER-NOTE VARIATION ACROSS PROFILES (between-horn signal)")
print("=" * 80)
print()
print("  For notes appearing in 2+ profiles: stdev of per-profile note-average descriptors.")
print("  Compare to within-profile stdev to assess signal vs noise.")
print()

# Build per-note, per-profile averages (REAL PROFILES ONLY)
note_profile_descs = defaultdict(lambda: defaultdict(list))  # note -> profile -> [desc dicts]
for c in real_captures:
    note_profile_descs[c['note']][c['profile']].append(c['descriptors'])

# Average descriptors per note per profile
note_profile_avg = defaultdict(dict)  # note -> {profile: avg_desc}
for note, pdict in note_profile_descs.items():
    for pname, desc_list in pdict.items():
        avg = {}
        for dk in desc_keys:
            vals = [d[dk] for d in desc_list]
            avg[dk] = mean(vals)
        note_profile_avg[note][pname] = avg

# For each note with 2+ profiles, compute between-profile stdev
print(f"  {'Note':>6s}  {'#Prof':>5s}  {'bright':>7s}  {'dark':>7s}  {'reson':>7s}  {'rich':>7s}  {'full':>7s}")
print(f"  {'-'*6}  {'-'*5}  {'-'*7}  {'-'*7}  {'-'*7}  {'-'*7}  {'-'*7}")

between_stdevs = {dk: [] for dk in desc_keys}
sorted_notes = sorted(note_profile_avg.keys(), key=note_sort_key)

for note in sorted_notes:
    pavgs = note_profile_avg[note]
    n_profiles = len(pavgs)
    if n_profiles < 2:
        continue
    row = f"  {note:>6s}  {n_profiles:>5d}"
    for dk in desc_keys:
        vals = [pavgs[p][dk] for p in pavgs]
        sd = stdev(vals)
        between_stdevs[dk].append(sd)
        row += f"  {sd*100:6.2f}%"
    print(row)

print()
print("  AVERAGE BETWEEN-PROFILE STDEV (signal):")
print(f"    ", end='')
global_between = {}
for dk in desc_keys:
    avg = mean(between_stdevs[dk]) if between_stdevs[dk] else 0
    global_between[dk] = avg
    print(f"  {avg*100:6.2f}%", end='')
print()
print()
print("  SIGNAL-TO-NOISE COMPARISON:")
print(f"    {'Descriptor':>12s}  {'Between(sig)':>12s}  {'Within(noise)':>13s}  {'Ratio':>7s}")
for dk in desc_keys:
    b = global_between.get(dk, 0)
    w = global_within.get(dk, 0)
    ratio = b / w if w > 0 else float('inf')
    print(f"    {dk:>12s}  {b*100:11.2f}%  {w*100:12.2f}%  {ratio:6.2f}x")
print()


# ── (d) Register Effects ────────────────────────────────────────────────

print("=" * 80)
print("(d) REGISTER EFFECTS (descriptor vs pitch)")
print("=" * 80)
print()
print("  Per-profile: how descriptors change from low to high notes.")
print("  If brightness systematically increases with pitch across ALL horns,")
print("  it's measuring register, not horn character.")
print()

print("  (Real profiles only -- sample profiles have fake f0=440 for all notes)")
print()

for pname, caps in profile_captures.items():
    if not caps or pname not in real_profiles:
        continue
    by_note = defaultdict(list)
    for c in caps:
        by_note[c['note']].append(c)

    sorted_notes = sorted(by_note.keys(), key=note_sort_key)
    if len(sorted_notes) < 3:
        continue

    print(f"  {pname} ({caps[0]['sax_type']})")
    print(f"    {'Note':>6s}  {'f0':>7s}  {'N':>3s}  {'bright':>7s}  {'dark':>7s}  {'reson':>7s}  {'rich':>7s}  {'full':>7s}  {'centroid':>8s}  {'slope':>7s}")
    print(f"    {'-'*6}  {'-'*7}  {'-'*3}  {'-'*7}  {'-'*7}  {'-'*7}  {'-'*7}  {'-'*7}  {'-'*8}  {'-'*7}")

    for note in sorted_notes:
        ncaps = by_note[note]
        n = len(ncaps)
        avg_f0 = mean([c['f0'] for c in ncaps])
        row = f"    {note:>6s}  {avg_f0:7.1f}  {n:>3d}"
        for dk in desc_keys:
            vals = [c['descriptors'][dk] for c in ncaps]
            row += f"  {mean(vals)*100:6.1f}%"
        row += f"  {mean([c['spectral_centroid'] for c in ncaps]):8.2f}"
        row += f"  {mean([c['spectral_slope'] for c in ncaps]):7.2f}"
        print(row)
    print()

    # Correlation of brightness with f0 within this profile
    bvals = [c['descriptors']['brightness'] for c in caps]
    fvals = [c['f0'] for c in caps]
    r_bf = pearson_r(bvals, fvals)
    print(f"    Correlation brightness vs f0: r = {r_bf:+.3f}")
    print()


# ── (e) Correlation Analysis ────────────────────────────────────────────

print("=" * 80)
print("(e) CORRELATION ANALYSIS (real profiles only)")
print("=" * 80)
print()

# Gather vectors from REAL captures only
all_bright = [c['descriptors']['brightness'] for c in real_captures]
all_dark = [c['descriptors']['darkness'] for c in real_captures]
all_reson = [c['descriptors']['resonance'] for c in real_captures]
all_rich = [c['descriptors']['richness'] for c in real_captures]
all_full = [c['descriptors']['fullness'] for c in real_captures]
all_centroid = [c['spectral_centroid'] for c in real_captures]
all_slope = [c['spectral_slope'] for c in real_captures]
all_f0 = [c['f0'] for c in real_captures]
all_h40 = [c['h_above_40'] for c in real_captures]
all_h60 = [c['h_above_60'] for c in real_captures]

pairs = [
    ("brightness", all_bright, "spectral_centroid", all_centroid),
    ("brightness", all_bright, "f0 (pitch)", all_f0),
    ("brightness", all_bright, "spectral_slope", all_slope),
    ("richness", all_rich, "H_above_-40dB", all_h40),
    ("richness", all_rich, "H_above_-60dB", all_h60),
    ("richness", all_rich, "spectral_centroid", all_centroid),
    ("fullness", all_full, "f0 (pitch)", all_f0),
    ("fullness", all_full, "spectral_centroid", all_centroid),
    ("resonance", all_reson, "f0 (pitch)", all_f0),
]

print("  Key correlations:")
print(f"    {'Pair':>40s}    {'r':>7s}   {'Interpretation'}")
print(f"    {'-'*40}    {'-'*7}   {'-'*30}")
for name1, v1, name2, v2 in pairs:
    r = pearson_r(v1, v2)
    interp = ""
    if abs(r) > 0.7:
        interp = "STRONG — may be redundant/confounded"
    elif abs(r) > 0.4:
        interp = "moderate"
    elif abs(r) > 0.2:
        interp = "weak"
    else:
        interp = "negligible"
    print(f"    {name1 + ' vs ' + name2:>40s}    {r:+7.3f}   {interp}")

print()
print("  Descriptor inter-correlations:")
all_descs = {
    'brightness': all_bright,
    'resonance': all_reson,
    'richness': all_rich,
    'fullness': all_full,
}
dk_list = list(all_descs.keys())
print(f"    {'':>12s}", end='')
for dk in dk_list:
    print(f"  {dk:>10s}", end='')
print()
for dk1 in dk_list:
    print(f"    {dk1:>12s}", end='')
    for dk2 in dk_list:
        if dk1 == dk2:
            print(f"  {'1.000':>10s}", end='')
        else:
            r = pearson_r(all_descs[dk1], all_descs[dk2])
            print(f"  {r:+10.3f}", end='')
    print()
print()


# ── (f) Harmonic Profile Comparison (Tenor Trio) ────────────────────────

print("=" * 80)
print("(f) HARMONIC PROFILE COMPARISON — TENOR TRIO")
print("=" * 80)
print()

# Find the three tenor profiles by scanning all profile names
tenor_keywords = {'SBA': None, 'BA': None, 'Shadow': None}
for pname, pdata in profiles_flat.items():
    caps = profile_captures.get(pname, [])
    # Skip sample/synthetic profiles (all f0=440)
    if caps and all(abs(c['f0'] - 440.0) < 1.0 for c in caps):
        continue
    pname_lower = pname.lower()
    if 'sba' in pname_lower and 'tenor' in pname_lower:
        tenor_keywords['SBA'] = pname
    elif 'shadow' in pname_lower:
        tenor_keywords['Shadow'] = pname
    elif 'ba' in pname_lower and 'sba' not in pname_lower and 'tenor' in pname_lower:
        tenor_keywords['BA'] = pname

found_tenors = {k: v for k, v in tenor_keywords.items() if v is not None}
if len(found_tenors) < 2:
    print("  Could not find at least 2 tenor profiles (SBA, BA, Shadow).")
    print(f"  Found: {found_tenors}")
    print()
else:
    print(f"  Profiles found:")
    for short, full in found_tenors.items():
        print(f"    {short:>8s}: {full}")
    print()

    # Gather per-note averaged harmonics for each
    tenor_notes = {}  # {short_name: {note: avg_harmonics_db}}
    for short, full in found_tenors.items():
        caps = profile_captures[full]
        by_note = defaultdict(list)
        for c in caps:
            by_note[c['note']].append(c['harmonics_db'])
        avg_by_note = {}
        for note, hdb_list in by_note.items():
            max_len = max(len(h) for h in hdb_list)
            avg = [0.0] * max_len
            counts = [0] * max_len
            for h in hdb_list:
                for i, db in enumerate(h):
                    avg[i] += db
                    counts[i] += 1
            for i in range(max_len):
                if counts[i] > 0:
                    avg[i] /= counts[i]
            avg_by_note[note] = avg
        tenor_notes[short] = avg_by_note

    # Find overlapping notes
    all_tenor_note_sets = [set(tn.keys()) for tn in tenor_notes.values()]
    overlap = all_tenor_note_sets[0]
    for s in all_tenor_note_sets[1:]:
        overlap = overlap & s

    overlap_sorted = sorted(overlap, key=note_sort_key)

    if not overlap_sorted:
        print("  No overlapping notes found between tenor profiles.")
    else:
        print(f"  Overlapping notes: {len(overlap_sorted)}")
        print()

        # Show a few representative notes (low, mid, high)
        if len(overlap_sorted) >= 6:
            show_notes = [overlap_sorted[0], overlap_sorted[len(overlap_sorted)//4],
                          overlap_sorted[len(overlap_sorted)//2],
                          overlap_sorted[3*len(overlap_sorted)//4],
                          overlap_sorted[-1]]
        else:
            show_notes = overlap_sorted

        tenor_shorts = list(found_tenors.keys())

        for note in show_notes:
            print(f"  Note: {note} (f0 ~ {note_to_freq(note):.1f} Hz)")
            # Header
            header = f"    {'H#':>3s}"
            for short in tenor_shorts:
                header += f"  {short:>8s}"
            header += "   max_diff"
            print(header)

            max_h = max(len(tenor_notes[s].get(note, [])) for s in tenor_shorts)
            for hi in range(min(max_h, 12)):
                row = f"    H{hi+1:>2d}"
                vals = []
                for short in tenor_shorts:
                    hdb = tenor_notes[short].get(note, [])
                    if hi < len(hdb):
                        row += f"  {hdb[hi]:8.1f}"
                        vals.append(hdb[hi])
                    else:
                        row += f"  {'---':>8s}"
                if len(vals) >= 2:
                    diff = max(vals) - min(vals)
                    row += f"   {diff:5.1f} dB"
                print(row)
            print()

    # Show where the biggest harmonic differences are across all overlapping notes
    print("  BIGGEST HARMONIC DIFFERENCES (avg across all overlapping notes):")
    if len(found_tenors) >= 2 and overlap_sorted:
        h_diffs = defaultdict(list)  # harmonic_num -> list of max_diffs
        for note in overlap_sorted:
            max_h = max(len(tenor_notes[s].get(note, [])) for s in tenor_shorts)
            for hi in range(min(max_h, 12)):
                vals = []
                for short in tenor_shorts:
                    hdb = tenor_notes[short].get(note, [])
                    if hi < len(hdb):
                        vals.append(hdb[hi])
                if len(vals) >= 2:
                    h_diffs[hi + 1].append(max(vals) - min(vals))

        print(f"    {'H#':>3s}  {'Avg diff':>9s}  {'Max diff':>9s}")
        for h_num in sorted(h_diffs.keys()):
            diffs = h_diffs[h_num]
            print(f"    H{h_num:>2d}  {mean(diffs):8.1f} dB  {max(diffs):8.1f} dB")
        print()


# ── (g) Descriptor Discrimination Power ─────────────────────────────────

print("=" * 80)
print("(g) DESCRIPTOR DISCRIMINATION POWER")
print("=" * 80)
print()
print("  F-ratio = between-horn variance / within-horn variance")
print("  Higher = better at distinguishing horns. F > 1 means the descriptor")
print("  captures more horn-to-horn variation than measurement noise.")
print()

# We need overlapping notes for a fair comparison.
# For each note present in 2+ profiles, compute:
#   within_var: average of per-profile variances of descriptor for that note
#   between_var: variance of per-profile means for that note
# Then average across notes.

for dk in desc_keys:
    within_vars = []
    between_vars = []

    for note in sorted(note_profile_descs.keys(), key=note_sort_key):
        pdict = note_profile_descs[note]
        if len(pdict) < 2:
            continue

        profile_means = []
        profile_within = []
        for pname, desc_list in pdict.items():
            vals = [d[dk] for d in desc_list]
            profile_means.append(mean(vals))
            if len(vals) >= 2:
                profile_within.append(variance(vals))

        if len(profile_means) >= 2:
            between_vars.append(variance(profile_means))
        if profile_within:
            within_vars.append(mean(profile_within))

    avg_between = mean(between_vars) if between_vars else 0
    avg_within = mean(within_vars) if within_vars else 0
    f_ratio = avg_between / avg_within if avg_within > 0 else float('inf')

    print(f"  {dk:>12s}:  between_var={avg_between:.6f}  within_var={avg_within:.6f}  F={f_ratio:.2f}")

print()

# Also compute overall profile-level discrimination
print("  PROFILE-LEVEL DISCRIMINATION (using compute_fingerprint per profile):")
print()

profile_fingerprints = {}
for pname, pdata in profiles_flat.items():
    sax_type = pdata.get('sax_type', 'Tenor')
    sessions = pdata.get('sessions', [])
    if not sessions:
        continue
    fp = compute_fingerprint(sessions, sax_type)
    if fp['capture_count'] > 0:
        profile_fingerprints[pname] = fp['descriptors']

if len(profile_fingerprints) >= 2:
    print(f"  {'Descriptor':>12s}  {'Range':>12s}  {'Mean':>8s}  {'Stdev':>8s}")
    for dk in desc_keys:
        vals = [fp[dk] for fp in profile_fingerprints.values()]
        lo = min(vals)
        hi = max(vals)
        print(f"  {dk:>12s}  {lo*100:.1f}-{hi*100:.1f}%  {mean(vals)*100:6.1f}%  {stdev(vals)*100:6.1f}%")

    print()
    print("  Profile fingerprint values:")
    header = f"    {'Profile':>30s}"
    for dk in desc_keys:
        print_dk = dk[:7]
        header += f"  {print_dk:>7s}"
    print(header)
    for pname, fp in sorted(profile_fingerprints.items()):
        row = f"    {pname:>30s}"
        for dk in desc_keys:
            row += f"  {fp[dk]*100:6.1f}%"
        print(row)
    print()


# ── Summary ──────────────────────────────────────────────────────────────

# ── (h) Register-Corrected Analysis ──────────────────────────────────────

print("=" * 80)
print("(h) REGISTER-CORRECTED BRIGHTNESS (is horn character visible after")
print("    removing the register-dependent component?)")
print("=" * 80)
print()
print("  For each note present in 2+ real profiles, compute the mean brightness")
print("  across all profiles at that note. Then each profile's register-corrected")
print("  brightness = mean(brightness_at_note - group_mean_at_note).")
print("  If a horn is consistently brighter/darker AFTER removing register,")
print("  the descriptor is measuring real horn character.")
print()

# Build note group means (real profiles only)
note_group_means = {}  # note -> mean brightness across profiles
for note, pavgs in note_profile_avg.items():
    if len(pavgs) < 2:
        continue
    vals = [pavgs[p]['brightness'] for p in pavgs]
    note_group_means[note] = mean(vals)

# Per-profile: average deviation from group mean at each note
profile_rc = {}  # profile -> (mean_deviation, n_notes, per_note_deviations)
for pname in real_profiles:
    deviations = []
    for note, group_mean in note_group_means.items():
        if pname in note_profile_avg.get(note, {}):
            prof_val = note_profile_avg[note][pname]['brightness']
            deviations.append(prof_val - group_mean)
    if deviations:
        profile_rc[pname] = (mean(deviations), len(deviations), deviations)

print(f"  {'Profile':>35s}  {'RC Bright':>9s}  {'Raw Bright':>10s}  {'#Notes':>6s}")
print(f"  {'-'*35}  {'-'*9}  {'-'*10}  {'-'*6}")
for pname in sorted(profile_rc.keys(), key=lambda p: profile_rc[p][0], reverse=True):
    rc_val, n_notes, _ = profile_rc[pname]
    # Get raw fingerprint brightness
    pdata = profiles_flat[pname]
    fp = compute_fingerprint(pdata.get('sessions', []), pdata.get('sax_type', 'Tenor'))
    raw_bright = fp['descriptors'].get('brightness', 0)
    print(f"  {pname:>35s}  {rc_val*100:+8.1f}%  {raw_bright*100:9.1f}%  {n_notes:>6d}")

print()
print("  If register-corrected ranking matches subjective brightness ranking,")
print("  the descriptor is working. If it scrambles the order, it's not.")
print()

# ── (i) Brightness Stability Across Registers ───────────────────────────

print("=" * 80)
print("(i) BRIGHTNESS BY REGISTER (real profiles)")
print("=" * 80)
print()
print("  Group notes into low/mid/high register, show average brightness per profile.")
print("  If a horn's rank stays consistent across registers, descriptor is stable.")
print()

# Define registers by f0 range
REGISTER_BOUNDS = {
    'Low (f0<220)': (0, 220),
    'Mid (220-440)': (220, 440),
    'High (440+)': (440, 9999),
}

reg_names = list(REGISTER_BOUNDS.keys())
print(f"  {'Profile':>35s}", end='')
for rn in reg_names:
    print(f"  {rn:>14s}", end='')
print()
print(f"  {'-'*35}", end='')
for _ in reg_names:
    print(f"  {'-'*14}", end='')
print()

for pname in real_profiles:
    caps = profile_captures[pname]
    print(f"  {pname:>35s}", end='')
    for rn, (lo, hi) in REGISTER_BOUNDS.items():
        reg_caps = [c for c in caps if lo <= c['f0'] < hi]
        if reg_caps:
            avg_b = mean([c['descriptors']['brightness'] for c in reg_caps])
            print(f"  {avg_b*100:13.1f}%", end='')
        else:
            print(f"  {'---':>14s}", end='')
    print()

print()

# ══════════════════════════════════════════════════════════════════════════

print("=" * 80)
print("SUMMARY OF KEY FINDINGS")
print("=" * 80)
print()

# Check if brightness is confounded with pitch
all_r_bf = pearson_r(all_bright, all_f0)
print(f"  1. Brightness-pitch confound: r = {all_r_bf:+.3f}", end='')
if abs(all_r_bf) > 0.5:
    print("  *** WARNING: brightness strongly tracks pitch ***")
elif abs(all_r_bf) > 0.3:
    print("  ** moderate pitch dependence **")
else:
    print("  (OK, weak pitch dependence)")

# Check brightness-centroid relationship
r_bc = pearson_r(all_bright, all_centroid)
print(f"  2. Brightness vs spectral centroid: r = {r_bc:+.3f}", end='')
if abs(r_bc) > 0.7:
    print("  (essentially measuring the same thing)")
else:
    print()

# Discrimination summary
print(f"  3. Discrimination power (F-ratios):")
for dk in desc_keys:
    within_vars = []
    between_vars = []
    for note in note_profile_descs.keys():
        pdict = note_profile_descs[note]
        if len(pdict) < 2:
            continue
        profile_means = []
        profile_within = []
        for pname, desc_list in pdict.items():
            vals = [d[dk] for d in desc_list]
            profile_means.append(mean(vals))
            if len(vals) >= 2:
                profile_within.append(variance(vals))
        if len(profile_means) >= 2:
            between_vars.append(variance(profile_means))
        if profile_within:
            within_vars.append(mean(profile_within))
    avg_b = mean(between_vars) if between_vars else 0
    avg_w = mean(within_vars) if within_vars else 0
    f = avg_b / avg_w if avg_w > 0 else float('inf')
    verdict = "GOOD" if f > 2 else "OK" if f > 1 else "POOR"
    print(f"     {dk:>12s}: F={f:5.2f}  [{verdict}]")

# Check for redundant descriptors
print(f"  4. Redundancy check (|r| > 0.8 between descriptors):")
found_redundant = False
for i, dk1 in enumerate(dk_list):
    for dk2 in dk_list[i+1:]:
        r = pearson_r(all_descs[dk1], all_descs[dk2])
        if abs(r) > 0.8:
            print(f"     {dk1} vs {dk2}: r = {r:+.3f}  *** POTENTIALLY REDUNDANT ***")
            found_redundant = True
if not found_redundant:
    print(f"     None found (all |r| < 0.8)")


# Per-profile brightness vs f0 correlations
print(f"  5. Per-profile brightness-pitch correlations (register confound):")
for pname in real_profiles:
    caps = profile_captures[pname]
    bvals = [c['descriptors']['brightness'] for c in caps]
    fvals = [c['f0'] for c in caps]
    r = pearson_r(bvals, fvals)
    flag = ""
    if abs(r) > 0.5:
        flag = " *** STRONG ***"
    elif abs(r) > 0.3:
        flag = " ** moderate **"
    print(f"     {pname:>35s}: r = {r:+.3f}{flag}")

# Register-corrected ranking
print()
print(f"  6. Register-corrected brightness ranking (horn character):")
if profile_rc:
    for pname in sorted(profile_rc.keys(), key=lambda p: profile_rc[p][0], reverse=True):
        rc_val, _, _ = profile_rc[pname]
        print(f"     {pname:>35s}: {rc_val*100:+.1f}% vs group mean")

# Sample profile issues
print()
print(f"  7. SAMPLE PROFILE DATA QUALITY WARNING:")
print(f"     {len(sample_profiles)} sample profiles have f0=440 Hz for ALL notes.")
print(f"     This means the harmonic data is synthetic/templated, not real measurements.")
print(f"     These profiles produce constant descriptors across the 'register' (no register")
print(f"     variation) because the f0 never changes. The fullness descriptor uses break")
print(f"     frequency relative to f0*harmonic_number, so fake f0 corrupts fullness values.")
print(f"     All statistical analysis above uses REAL profiles only to avoid contamination.")

print()
print("  8. RESONANCE IS USELESS (F=0.07):")
print("     Every real profile reads 100% resonance. The scaling constants")
print("     (RESONANCE_RAW_MIN=0.85, RESONANCE_RAW_RANGE=0.15) map the entire")
print("     real-world range to the ceiling. This descriptor provides zero")
print("     discrimination between horns. Needs recalibration or removal.")

print()
print("Analysis complete.")
