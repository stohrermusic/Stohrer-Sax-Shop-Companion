#!/usr/bin/env python3
"""Test whether toner descriptors actually differentiate horns.

Reads groups of recordings, each one player and setup with only the horn
changing. If descriptors can't differentiate horns inside such a group,
they measure the player, not the horn; if the gap between two players'
groups is larger than the spread inside either, they measure the setup.

A descriptor is a viable live-gauge candidate if its within-player spread
is large enough to see on a gauge AND not dwarfed by between-player gap.
Live descriptor gauges were removed 2026-04-06 because absolute single-
preset readouts proved too noisy; this script is what we'd run to vet any
candidate before re-adding one.

The recordings are not in the repository and neither is their description:
who played what is in tools/local/descriptor_corpus.json, which git ignores.
No player's or recording's name is written in this file. Without that file
the suite has nothing to run and says so. Its shape:

    {"recordings": "<folder>",
     "groups": [{"title": "...", "player": "...", "sax_type": "Tenor",
                 "files": {"<label>": "<path under the folder>", ...}}, ...]}

Run:  python tools/test_descriptor_validity.py
"""
import json
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from toner_engine import (
    TonerEngine, analyze_audio_file, compute_fingerprint,
)

CORPUS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "local", "descriptor_corpus.json")
RECORDINGS = ""          # set by load_corpus()


def load_corpus():
    """The local description of the corpus, or None when this machine has none."""
    global RECORDINGS
    if not os.path.isfile(CORPUS_FILE):
        return None
    with open(CORPUS_FILE, encoding="utf-8") as f:
        corpus = json.load(f)
    RECORDINGS = corpus["recordings"]
    return corpus if os.path.isdir(RECORDINGS) else None


def process_file(path, sax_type):
    engine = TonerEngine()
    engine._sax_type = sax_type
    captures = analyze_audio_file(path, engine)
    if not captures:
        return None
    sessions = [{"captures": captures, "mic_type": "condenser"}]
    return compute_fingerprint(sessions, sax_type)


def analyze_group(name, file_dict, sax_type):
    """Process a group of recordings and analyze descriptor differentiation."""
    print(f"\n{'=' * 75}")
    print(f"  {name} ({len(file_dict)} horns, sax_type={sax_type})")
    print(f"{'=' * 75}")

    # Check which files exist
    existing = {}
    for label, rel in file_dict.items():
        full = os.path.join(RECORDINGS, rel)
        if os.path.isfile(full):
            existing[label] = full
        else:
            print(f"  SKIP (not found): {label}")

    if len(existing) < 2:
        print("  Not enough files to compare.")
        return None

    # Process all files
    fingerprints = {}
    for label, path in existing.items():
        sys.stdout.write(f"  {label}...")
        sys.stdout.flush()
        fp = process_file(path, sax_type)
        if fp and fp['capture_count'] > 0:
            fingerprints[label] = fp
            print(f" {fp['capture_count']} caps, {fp['note_count']} notes")
        else:
            print(" FAIL")

    if len(fingerprints) < 2:
        print("  Not enough successful analyses to compare.")
        return None

    # ── Print descriptor table ──
    desc_keys = ['richness', 'warmth', 'even_odd', 'rolloff_shape']
    header_map = {'richness': 'Rich', 'warmth': 'Warm',
                  'even_odd': 'E/O', 'rolloff_shape': 'Roll'}

    print(f"\n  {'Horn':<25}", end="")
    for k in desc_keys:
        print(f" {header_map[k]:>6}", end="")
    print(f" {'Even':>6} {'R.Rate':>6}")
    print(f"  {'-'*25}", end="")
    for _ in desc_keys:
        print(f" {'-'*6}", end="")
    print(f" {'-'*6} {'-'*6}")

    for label in existing:
        fp = fingerprints.get(label)
        if not fp:
            continue
        d = fp['descriptors']
        rr = fp.get('rolloff_rate')
        rr_str = f"{rr:.1f}" if rr is not None else "n/a"
        print(f"  {label:<25}", end="")
        for k in desc_keys:
            print(f" {d.get(k,0):5.1%}", end="")
        print(f" {d.get('evenness',0):5.1%} {rr_str:>6}")

    # ── Compute differentiation power for each descriptor ──
    print("\n  -- Differentiation Analysis --")

    for k in desc_keys:
        vals = [fp['descriptors'].get(k, 0) for fp in fingerprints.values()]
        mn, mx = min(vals), max(vals)
        spread = mx - mn
        mean = sum(vals) / len(vals)
        stdev = (sum((v - mean)**2 for v in vals) / len(vals)) ** 0.5

        # How much of the gauge range does this group use?
        gauge_usage = f"{mn:.1%}-{mx:.1%}"

        # Coefficient of variation (stdev/mean) — higher = more differentiating
        cv = stdev / mean if mean > 0 else 0

        verdict = "GOOD" if spread > 0.15 else "WEAK" if spread > 0.05 else "DEAD"
        print(f"  {header_map[k]:<6} range={spread:5.1%}  "
              f"stdev={stdev:.3f}  CV={cv:.2f}  "
              f"gauge={gauge_usage:<15} [{verdict}]")

    # ── Per-note harmonic comparison (the raw delta approach) ──
    # Pick two horns and compare note-by-note harmonics directly
    labels = list(fingerprints.keys())
    fp_a = fingerprints[labels[0]]
    fp_b = fingerprints[labels[-1]]
    pn_a = fp_a.get('per_note', {})
    pn_b = fp_b.get('per_note', {})
    common = sorted(set(pn_a.keys()) & set(pn_b.keys()),
                    key=lambda n: (int(n[-1]) if n[-1].isdigit() else 0, n))

    if common:
        print(f"\n  -- Raw Harmonic Delta: {labels[0]} vs {labels[-1]} --")
        print(f"  (Direct dB differences per harmonic, {len(common)} common notes)")

        # Average absolute dB difference per harmonic position across all notes
        max_h = 12
        h_diffs = [[] for _ in range(max_h)]
        for note in common:
            a_h = pn_a[note].get('harmonics_db', []) if pn_a[note] else []
            b_h = pn_b[note].get('harmonics_db', []) if pn_b[note] else []
            n_h = min(len(a_h), len(b_h), max_h)
            for i in range(n_h):
                h_diffs[i].append(a_h[i] - b_h[i])

        print(f"  {'H#':<4}", end="")
        for i in range(max_h):
            print(f" {'H'+str(i+1):>6}", end="")
        print()

        # Mean signed difference
        print(f"  {'Mean':<4}", end="")
        for i in range(max_h):
            if h_diffs[i]:
                m = sum(h_diffs[i]) / len(h_diffs[i])
                print(f" {m:+5.1f}", end="")
            else:
                print("   n/a", end="")
        print(" dB")

        # Absolute mean (how big are the differences regardless of direction)
        print(f"  {'|D|':<4}", end="")
        for i in range(max_h):
            if h_diffs[i]:
                m = sum(abs(d) for d in h_diffs[i]) / len(h_diffs[i])
                print(f"  {m:5.1f}", end="")
            else:
                print("   n/a", end="")
        print(" dB")

        # Overall: average absolute delta across all harmonics and notes
        all_abs = [abs(d) for diffs in h_diffs for d in diffs]
        if all_abs:
            overall_avg = sum(all_abs) / len(all_abs)
            print(f"\n  Overall avg |delta|: {overall_avg:.1f} dB across "
                  f"{len(common)} notes x {max_h} harmonics")
            if overall_avg > 3.0:
                print("  -> These horns sound measurably different")
            elif overall_avg > 1.5:
                print("  -> Moderate differences (may be audible)")
            else:
                print("  -> Small differences (borderline audible)")

    # Delta descriptors (spectral_tilt, mid_harmonic) were removed from the
    # app after data analysis showed live comparison is unreliable; the
    # per-pair matrix that used to live here is gone with them.

    return fingerprints


def main():
    corpus = load_corpus()
    if corpus is None:
        print(f"SKIP: no recordings corpus on this machine ({CORPUS_FILE})")
        return

    print("DESCRIPTOR VALIDITY TEST")
    print("Question: Do these descriptors measure HORNS or RECORDING SETUPS?")
    print()
    print("Test methodology:")
    print("  - Each group: same player/setup = isolates horn differences")
    print("  - Cross-player comparison tests whether descriptors are dominated")
    print("    by player/setup rather than horn character")

    results = []
    for group in corpus["groups"]:
        fps = analyze_group(group["title"], group["files"], group["sax_type"])
        results.append((group, fps))

    # ── Cross-player test: the first two groups of one horn type by different players ──
    pair = None
    for i, (ga, fa) in enumerate(results):
        for gb, fb in results[i + 1:]:
            if fa and fb and ga["sax_type"] == gb["sax_type"] and ga["player"] != gb["player"]:
                pair = pair or (ga, fa, gb, fb)
    if pair:
        ga, first, gb, second = pair
        print(f"\n{'=' * 75}")
        print("  CROSS-PLAYER: Is between-player variation > within-player variation?")
        print(f"{'=' * 75}")

        desc_keys = ['richness', 'warmth', 'even_odd', 'rolloff_shape']
        header_map = {'richness': 'Rich', 'warmth': 'Warm',
                      'even_odd': 'E/O', 'rolloff_shape': 'Roll'}

        for k in desc_keys:
            e_vals = [fp['descriptors'].get(k, 0) for fp in first.values()]
            t_vals = [fp['descriptors'].get(k, 0) for fp in second.values()]

            e_mean = sum(e_vals) / len(e_vals)
            t_mean = sum(t_vals) / len(t_vals)
            e_spread = max(e_vals) - min(e_vals)
            t_spread = max(t_vals) - min(t_vals)
            player_gap = abs(e_mean - t_mean)

            # If the gap between players > spread within each player,
            # the descriptor is measuring the player/setup, not the horn
            within = max(e_spread, t_spread)
            ratio = player_gap / within if within > 0 else float('inf')

            if ratio > 1.5:
                verdict = "MEASURES PLAYER/SETUP"
            elif ratio > 0.8:
                verdict = "MIXED (player + horn)"
            else:
                verdict = "Measures horn"

            print(f"  {header_map[k]:<6} {ga['player']} mean={e_mean:.1%} spread={e_spread:.1%}  "
                  f"{gb['player']} mean={t_mean:.1%} spread={t_spread:.1%}  "
                  f"gap={player_gap:.1%}  [{verdict}]")

    # ── Summary ──
    print(f"\n{'=' * 75}")
    print("  SUMMARY")
    print(f"{'=' * 75}")
    print("""
  Current descriptors (computed from raw harmonics on the fly):
    Richness       (Pure <-> Complex)  -- spectral flatness
    Warmth         (Thin <-> Warm)     -- H2 strength
    Even/Odd                           -- even vs odd harmonic balance
    Rolloff Shape                      -- nonlinearity of harmonic rolloff

  How to read this:
    * A descriptor is a viable LIVE-gauge candidate if its within-player
      spread is large enough to see on a gauge AND not dwarfed by the
      between-player gap. Read the [verdict] column above.
    * A descriptor is useful in the Analyze tool (where deltas cancel
      mic/setup confounders) even if it fails the live-gauge test.

  No live descriptor gauges currently exist -- they were removed
  2026-04-06 because mic position alone shifted complexity 10-20%
  between same-horn takes. Re-vet here before re-adding any.
""")


if __name__ == '__main__':
    main()
