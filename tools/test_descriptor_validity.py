#!/usr/bin/env python3
"""Test whether toner descriptors actually differentiate horns.

Uses three test groups:
1. Edinger Head2Head (gold standard): Same player, mpc, mic, room, day.
   If descriptors can't differentiate THESE, they measure the player not the horn.
2. Tyler Tenors: Same player/setup, different horns.
3. Tyler Altos: Same player/setup, different horns.

A descriptor is a viable live-gauge candidate if its within-player spread
is large enough to see on a gauge AND not dwarfed by between-player gap.
Live descriptor gauges were removed 2026-04-06 because absolute single-
preset readouts proved too noisy; this script is what we'd run to vet any
candidate before re-adding one.

Run:  python tools/test_descriptor_validity.py
"""
import sys, os, math
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from toner_engine import (
    TonerEngine, analyze_audio_file, descriptors_from_harmonics,
    compute_fingerprint, compute_rolloff_rate,
)

RECORDINGS = r"C:\sax shop companion\recordings"

EDINGER_FILES = {
    "SML Gold Medal 15840":     "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-SML Gold Medal #15840.wav",
    "SBA 35737":                "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #35737.wav",
    "MkVI 68611":               "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #68611.wav",
    "MkVI 68630 orig neck":     "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #68630 with original neck.wav",
    "MkVI 68630 later neck":    "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #68630 with later Mk VI replacement neck.wav",
    "MkVI 68630 122028 neck":   "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #68630 with neck from #122028.wav",
    "MkVI 71352":               "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #71352.wav",
    "MkVI 76047":               "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #76047.wav",
    "MkVI 85961":               "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #85961.wav",
    "MkVI 97226":               "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #97226.wav",
    "MkVI 99314":               "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #99314.wav",
    "SA80 122028":              "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #122028.wav",
    "Ref54 141956":             "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Selmer #141956.wav",
    "Yamaha YTS-62":            "Thomas Edinger/Thomas Edinger/Saxophones Head2Head-Yamaha YTS-62 Purple Logo #017542.wav",
}

TYLER_TENORS = {
    "MkVI 62418":       "Tyler Anderson/Tenor/Tenor/Mark VI 62418.wav",
    "MkVI 72434":       "Tyler Anderson/Tenor/Tenor/Mark VI 72434.wav",
    "MkVI 100115":      "Tyler Anderson/Tenor/Tenor/Mark VI 100115.wav",
    "MkVI 141672":      "Tyler Anderson/Tenor/Tenor/Mark VI 141672.wav",
    "MkVI 147401":      "Tyler Anderson/Tenor/Tenor/Mark VI 147401.wav",
    "MkVI 157137":      "Tyler Anderson/Tenor/Tenor/Mark VI 157137.wav",
    "MkVI 164746":      "Tyler Anderson/Tenor/Tenor/Mark VI 164746.wav",
    "SBA 38438":        "Tyler Anderson/Tenor/Tenor/Selmer 38438.wav",
    "Conn 274254":      "Tyler Anderson/Tenor/Tenor/Conn 274254.wav",
    "Conn 329232":      "Tyler Anderson/Tenor/Tenor/Conn 329232.wav",
    "Couf Superba":     "Tyler Anderson/Tenor/Tenor/Couf Superba 1.wav",
    "King S20 375203":  "Tyler Anderson/Tenor/Tenor/King Super 20 375203.wav",
    "SML 21321":        "Tyler Anderson/Tenor/Tenor/SML 21321.wav",
    "Tenor Mad 1222":   "Tyler Anderson/Tenor/Tenor/Tenor Madness 1222.wav",
    "Yamaha 875EX":     "Tyler Anderson/Tenor/Tenor/Yamaha 875EX 6598.wav",
}

TYLER_ALTOS = {
    "MkVI 93668":       "Tyler Anderson/Alto/Alto/Selmer Mark VI 93668.wav",
    "MkVI 190366":      "Tyler Anderson/Alto/Alto/Selmer Mark VI 190366.wav",
    "MkVI 219548":      "Tyler Anderson/Alto/Alto/Mark VI 219548.wav",
    "Supreme 842487":   "Tyler Anderson/Alto/Alto/Selmer Supreme 842487.wav",
    "YAS-62 11959":     "Tyler Anderson/Alto/Alto/YAS-62 11959.wav",
    "YAS-62 88666":     "Tyler Anderson/Alto/Alto/YAS-62 88666.wav",
    "Buescher 267734":  "Tyler Anderson/Alto/Alto/Buescher 267734.wav",
    "Buescher 307777":  "Tyler Anderson/Alto/Alto/Buescher 307777.wav",
    "Keilwerth NK":     "Tyler Anderson/Alto/Alto/Keilwerth New King 38431.wav",
}


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
    print(f"\n  -- Differentiation Analysis --")

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
                print(f"   n/a", end="")
        print(" dB")

        # Absolute mean (how big are the differences regardless of direction)
        print(f"  {'|D|':<4}", end="")
        for i in range(max_h):
            if h_diffs[i]:
                m = sum(abs(d) for d in h_diffs[i]) / len(h_diffs[i])
                print(f"  {m:5.1f}", end="")
            else:
                print(f"   n/a", end="")
        print(" dB")

        # Overall: average absolute delta across all harmonics and notes
        all_abs = [abs(d) for diffs in h_diffs for d in diffs]
        if all_abs:
            overall_avg = sum(all_abs) / len(all_abs)
            print(f"\n  Overall avg |delta|: {overall_avg:.1f} dB across "
                  f"{len(common)} notes x {max_h} harmonics")
            if overall_avg > 3.0:
                print(f"  -> These horns sound measurably different")
            elif overall_avg > 1.5:
                print(f"  -> Moderate differences (may be audible)")
            else:
                print(f"  -> Small differences (borderline audible)")

    # Delta descriptors (spectral_tilt, mid_harmonic) were removed from the
    # app after data analysis showed live comparison is unreliable; the
    # per-pair matrix that used to live here is gone with them.

    return fingerprints


def main():
    if not os.path.isdir(RECORDINGS):
        print(f"Recordings not found: {RECORDINGS}")
        return

    print("DESCRIPTOR VALIDITY TEST")
    print("Question: Do these descriptors measure HORNS or RECORDING SETUPS?")
    print()
    print("Test methodology:")
    print("  - Edinger: same player/mpc/mic/room = isolates horn differences")
    print("  - Tyler: same player/setup = isolates horn differences")
    print("  - Cross-player comparison tests whether descriptors are dominated")
    print("    by player/setup rather than horn character")

    edinger = analyze_group(
        "EDINGER HEAD-TO-HEAD (14 tenors, same everything except horn)",
        EDINGER_FILES, "Tenor")

    tyler_t = analyze_group(
        "TYLER TENORS (15 horns, same player/setup)",
        TYLER_TENORS, "Tenor")

    tyler_a = analyze_group(
        "TYLER ALTOS (9 horns, same player/setup)",
        TYLER_ALTOS, "Alto")

    # ── Cross-player test ──
    if edinger and tyler_t:
        print(f"\n{'=' * 75}")
        print("  CROSS-PLAYER: Is between-player variation > within-player variation?")
        print(f"{'=' * 75}")

        desc_keys = ['richness', 'warmth', 'even_odd', 'rolloff_shape']
        header_map = {'richness': 'Rich', 'warmth': 'Warm',
                      'even_odd': 'E/O', 'rolloff_shape': 'Roll'}

        for k in desc_keys:
            e_vals = [fp['descriptors'].get(k, 0) for fp in edinger.values()]
            t_vals = [fp['descriptors'].get(k, 0) for fp in tyler_t.values()]

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

            print(f"  {header_map[k]:<6} Edinger mean={e_mean:.1%} spread={e_spread:.1%}  "
                  f"Tyler mean={t_mean:.1%} spread={t_spread:.1%}  "
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
