"""
Deep variance analysis across Tyler's controlled dataset.
Looking for patterns the current descriptors might miss.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from toner_engine import (
    load_tone_presets, compute_fingerprint,
)
from config import TONER_DATA_FILE
import numpy as np

profiles = load_tone_presets(TONER_DATA_FILE)


def fp(lib, name):
    p = profiles[lib][name]
    sax = p.get("horn_type", "Alto")
    return compute_fingerprint(p["sessions"], sax)


def section(title):
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}\n")


# Collect all Tyler profiles with fingerprints
tyler = {}
for name, data in profiles.get("Tyler Anderson", {}).items():
    f = fp("Tyler Anderson", name)
    tyler[name] = {
        "fp": f,
        "type": data.get("horn_type", "?"),
        "make": data.get("horn_make", "?"),
        "model": data.get("horn_model", "?"),
        "serial": data.get("serial", "?"),
    }

# Also Matt's profiles
matt = {}
for name, data in profiles.get("My Profiles", {}).items():
    if data.get("sessions"):
        f = fp("My Profiles", name)
        matt[name] = {
            "fp": f,
            "type": data.get("horn_type", "?"),
            "make": data.get("horn_make", "?"),
        }

all_profiles = {}
for name, d in tyler.items():
    all_profiles[f"[Tyler] {name}"] = d
for name, d in matt.items():
    all_profiles[f"[Matt] {name}"] = d


# ============================================================
# 1. HARMONIC SHAPE ANALYSIS
# Which harmonics vary most between horns?
# ============================================================
section("WHICH HARMONICS VARY MOST BETWEEN HORNS?")
print("  StdDev of each harmonic's dB level across all profiles.\n")

# Tyler tenors only (controlled comparison)
tenor_harmonics = []
tenor_labels = []
for name, d in tyler.items():
    if d["type"] == "Tenor":
        tenor_harmonics.append(d["fp"]["harmonics_db"][:12])
        tenor_labels.append(name)

arr = np.array(tenor_harmonics)
print("  Tyler's tenors (same player/mpc/reed):")
print(f"  {'Harmonic':<10} {'Mean':>8} {'StdDev':>8} {'Min':>8} {'Max':>8} {'Range':>8}")
print(f"  {'-'*52}")
for i in range(min(12, arr.shape[1])):
    col = arr[:, i]
    print(f"  H{i+1:<9} {col.mean():+7.1f}  {col.std():7.2f}  "
          f"{col.min():+7.1f}  {col.max():+7.1f}  {col.max()-col.min():7.1f}")

# Tyler altos
alto_harmonics = []
for name, d in tyler.items():
    if d["type"] == "Alto":
        alto_harmonics.append(d["fp"]["harmonics_db"][:12])

arr_a = np.array(alto_harmonics)
print("\n  Tyler's altos (same player/mpc/reed):")
print(f"  {'Harmonic':<10} {'Mean':>8} {'StdDev':>8} {'Min':>8} {'Max':>8} {'Range':>8}")
print(f"  {'-'*52}")
for i in range(min(12, arr_a.shape[1])):
    col = arr_a[:, i]
    print(f"  H{i+1:<9} {col.mean():+7.1f}  {col.std():7.2f}  "
          f"{col.min():+7.1f}  {col.max():+7.1f}  {col.max()-col.min():7.1f}")


# ============================================================
# 2. NOTE-TO-NOTE CONSISTENCY (within each horn)
# Does brightness vary a lot across the register?
# ============================================================
section("NOTE-TO-NOTE CONSISTENCY — how uniform is each horn?")
print("  StdDev of brightness across notes within each profile.")
print("  High = horn changes character across register")
print("  Low = uniform tone across register\n")

consistency = []
for name, d in sorted(tyler.items()):
    pn = d["fp"].get("per_note", {})
    if len(pn) < 10:
        continue
    bright_vals = [v["descriptors"]["brightness"] * 100 for v in pn.values()]
    rich_vals = [v["descriptors"]["richness"] * 100 for v in pn.values()]
    full_vals = [v["descriptors"]["fullness"] * 100 for v in pn.values()]
    consistency.append((
        name, d["type"],
        np.std(bright_vals), np.std(rich_vals), np.std(full_vals),
        len(pn)
    ))

consistency.sort(key=lambda x: x[2], reverse=True)
print(f"  {'Profile':<40} {'Type':<6} {'B sd':>6} {'Cx sd':>6} {'F sd':>6} {'Notes':>5}")
print(f"  {'-'*70}")
for name, typ, bs, rs, fs, nc in consistency:
    print(f"  {name:<40} {typ:<6} {bs:5.1f}  {rs:5.1f}  {fs:5.1f}  {nc:4d}")


# ============================================================
# 3. SPECTRAL CENTROID — alternative brightness measure
# ============================================================
section("SPECTRAL CENTROID — weighted average harmonic position")
print("  Higher = energy concentrated in upper harmonics")
print("  May correlate better with perceived 'brightness'\n")

centroids = []
for label, d in sorted(all_profiles.items()):
    h = d["fp"]["harmonics_db"][:12]
    # Convert dB to linear amplitude
    linear = [10 ** (db / 20.0) for db in h]
    total = sum(linear)
    if total > 0:
        centroid = sum((i + 1) * a for i, a in enumerate(linear)) / total
    else:
        centroid = 1.0
    centroids.append((label, d["type"], centroid,
                       d["fp"]["descriptors"]["brightness"] * 100))

centroids.sort(key=lambda x: x[2], reverse=True)
print(f"  {'Profile':<50} {'Type':<6} {'Centr':>6} {'Bright':>7}")
print(f"  {'-'*72}")
for label, typ, c, b in centroids:
    print(f"  {label:<50} {typ:<6} {c:5.2f}  {b:5.1f}%")


# ============================================================
# 4. H2/H1 RATIO — "warmth" or "roundness"
# ============================================================
section("H2 STRENGTH — octave harmonic (roundness/warmth)")
print("  H2 is the octave harmonic. Strong H2 = round, warm quality.\n")

h2_data = []
for label, d in sorted(all_profiles.items()):
    h = d["fp"]["harmonics_db"]
    if len(h) >= 2:
        h2_data.append((label, d["type"], h[1],
                         d["fp"]["descriptors"]["brightness"] * 100,
                         d["fp"]["descriptors"]["fullness"] * 100))

h2_data.sort(key=lambda x: x[2], reverse=True)
print(f"  {'Profile':<50} {'Type':<6} {'H2 dB':>7} {'Bright':>7} {'Full':>6}")
print(f"  {'-'*80}")
for label, typ, h2, b, f in h2_data:
    print(f"  {label:<50} {typ:<6} {h2:+6.1f}  {b:5.1f}%  {f:4.1f}%")


# ============================================================
# 5. UPPER HARMONIC ROLLOFF RATE
# ============================================================
section("UPPER HARMONIC ROLLOFF — how fast do harmonics decay?")
print("  Slope of linear fit to H1-H8 dB levels.")
print("  Steep = fundamental dominates (pure). Shallow = rich overtones.\n")

rolloff_data = []
for label, d in sorted(all_profiles.items()):
    h = d["fp"]["harmonics_db"][:8]
    if len(h) >= 8:
        x = np.arange(len(h))
        slope, intercept = np.polyfit(x, h, 1)
        rolloff_data.append((label, d["type"], slope,
                              d["fp"]["descriptors"]["richness"] * 100,
                              d["fp"]["descriptors"]["brightness"] * 100))

rolloff_data.sort(key=lambda x: x[2], reverse=True)
print(f"  {'Profile':<50} {'Type':<6} {'Slope':>7} {'Cmpx':>6} {'Bright':>7}")
print(f"  {'-'*80}")
for label, typ, s, r, b in rolloff_data:
    print(f"  {label:<50} {typ:<6} {s:+6.2f}  {r:5.1f}%  {b:5.1f}%")


# ============================================================
# 6. H3-H5 vs H6-H8 BALANCE — "edge" vs "warmth"?
# ============================================================
section("MID vs UPPER HARMONIC BALANCE")
print("  Ratio of H3-H5 energy to H6-H8 energy.")
print("  High = presence band dominates (edgy, projecting)")
print("  Low = upper harmonics relatively strong (complex, shimmery)\n")

balance_data = []
for label, d in sorted(all_profiles.items()):
    h = d["fp"]["harmonics_db"][:8]
    if len(h) >= 8:
        mid = [10 ** (h[i] / 20.0) for i in range(2, 5)]  # H3-H5
        upper = [10 ** (h[i] / 20.0) for i in range(5, 8)]  # H6-H8
        mid_sum = sum(mid)
        upper_sum = sum(upper)
        ratio = mid_sum / upper_sum if upper_sum > 0 else 999
        balance_data.append((label, d["type"], ratio,
                              d["fp"]["descriptors"]["brightness"] * 100))

balance_data.sort(key=lambda x: x[2], reverse=True)
print(f"  {'Profile':<50} {'Type':<6} {'Mid/Up':>7} {'Bright':>7}")
print(f"  {'-'*74}")
for label, typ, r, b in balance_data:
    print(f"  {label:<50} {typ:<6} {r:6.2f}  {b:5.1f}%")


# ============================================================
# 7. CORRELATION MATRIX
# ============================================================
section("CORRELATION MATRIX — which metrics move together?")

# Gather all metrics for Tyler's profiles
metrics = []
metric_names = ["brightness", "complexity", "fullness", "centroid",
                "H2_dB", "rolloff", "mid_upper_ratio", "consistency"]
for name, d in tyler.items():
    h = d["fp"]["harmonics_db"][:8]
    if len(h) < 8:
        continue
    linear = [10 ** (db / 20.0) for db in h[:12] if db is not None]
    total = sum(linear) if linear else 1
    centroid = sum((i+1) * a for i, a in enumerate(linear)) / total if total > 0 else 1

    x = np.arange(len(h))
    slope, _ = np.polyfit(x, h, 1)

    mid = sum(10 ** (h[i] / 20.0) for i in range(2, 5))
    upper = sum(10 ** (h[i] / 20.0) for i in range(5, 8))
    ratio = mid / upper if upper > 0 else 10

    pn = d["fp"].get("per_note", {})
    if len(pn) >= 5:
        bright_vals = [v["descriptors"]["brightness"] * 100 for v in pn.values()]
        consist = np.std(bright_vals)
    else:
        consist = 0

    metrics.append([
        d["fp"]["descriptors"]["brightness"] * 100,
        d["fp"]["descriptors"]["richness"] * 100,
        d["fp"]["descriptors"]["fullness"] * 100,
        centroid,
        h[1],
        slope,
        ratio,
        consist,
    ])

m = np.array(metrics)
print(f"  {'':>14}", end="")
for n in metric_names:
    print(f" {n[:8]:>8}", end="")
print()
print(f"  {'-'*80}")

for i, ni in enumerate(metric_names):
    print(f"  {ni:>13}", end="")
    for j, nj in enumerate(metric_names):
        r = np.corrcoef(m[:, i], m[:, j])[0, 1]
        print(f" {r:+8.2f}", end="")
    print()
