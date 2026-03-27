"""
Analyze Tyler Anderson's tone library data.

Key comparisons:
1. Same horn, different player (Conn Virtuoso: Matt vs Tyler)
2. Mark VI tenor variation (7 different VIs, same player/mpc/reed)
3. Cross-make comparison (same player/mpc/reed controls for player)
4. Alto vs Tenor brightness/fullness patterns
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from toner_engine import (
    load_tone_profiles, compute_fingerprint,
    descriptors_from_harmonics,
)
from config import TONE_PROFILES_FILE

profiles = load_tone_profiles(TONE_PROFILES_FILE)


def fp(lib, name):
    """Get fingerprint for a profile."""
    p = profiles[lib][name]
    sax = p.get("horn_type", "Alto")
    return compute_fingerprint(p["sessions"], sax)


def desc_line(label, fp_data, width=40):
    """Format one profile's descriptors."""
    d = fp_data["descriptors"]
    notes = fp_data["note_count"]
    caps = fp_data["capture_count"]
    print(f"  {label:<{width}} B={d['brightness']*100:4.1f}%  "
          f"Cx={d['richness']*100:4.1f}%  "
          f"F={d['fullness']*100:4.1f}%  "
          f"D={d['darkness']*100:4.1f}%  "
          f"({caps} caps, {notes} notes)")


def harmonic_line(label, fp_data, width=40):
    """Show averaged harmonic dB levels."""
    h = fp_data["harmonics_db"]
    bars = "  ".join(f"H{i+1}:{v:+5.1f}" for i, v in enumerate(h[:8]))
    print(f"  {label:<{width}} {bars}")


def section(title):
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}\n")


# ============================================================
# 1. SAME HORN, DIFFERENT PLAYER
# ============================================================
section("CONN VIRTUOSO DELUXE — Matt vs Tyler (same horn!)")
print("  Matt: Morgan 3C mouthpiece, live mic capture")
print("  Tyler: GSNY 7 mouthpiece, Rigotti reed, Neumann TLM studio recording\n")

matt_conn = fp("My Profiles", "Conn NW2 Virtuoso Deluxe 205k")
tyler_conn = fp("Tyler Anderson", "Conn NW2 Virtuoso Deluxe #205k")

desc_line("Matt (Morgan 3C)", matt_conn)
desc_line("Tyler (GSNY 7)", tyler_conn)
print()
harmonic_line("Matt harmonics", matt_conn)
harmonic_line("Tyler harmonics", tyler_conn)

# Per-note comparison on shared notes
matt_notes = matt_conn.get("per_note", {})
tyler_notes = tyler_conn.get("per_note", {})
shared = sorted(set(matt_notes.keys()) & set(tyler_notes.keys()))
if shared:
    print(f"\n  Per-note brightness on {len(shared)} shared notes:")
    print(f"  {'Note':<6} {'Matt':>8} {'Tyler':>8} {'Diff':>8}")
    print(f"  {'-'*32}")
    for note in shared:
        mb = matt_notes[note]["descriptors"]["brightness"] * 100
        tb = tyler_notes[note]["descriptors"]["brightness"] * 100
        diff = mb - tb
        print(f"  {note:<6} {mb:7.1f}% {tb:7.1f}% {diff:+7.1f}%")


# ============================================================
# 2. MARK VI TENOR VARIATION (same player, same mpc, same reed)
# ============================================================
section("MARK VI TENOR VARIATION — 7 horns, same player/mpc/reed")
print("  Player: Tyler Anderson | Mpc: GS Slant 7* | Reed: Rigotti\n")

vi_tenors = []
for name, data in sorted(profiles["Tyler Anderson"].items()):
    if "Mark VI" in name and data.get("horn_type") == "Tenor":
        f = fp("Tyler Anderson", name)
        serial = data.get("serial", "?")
        vi_tenors.append((serial, name, f))

vi_tenors.sort(key=lambda x: x[2]["descriptors"]["brightness"], reverse=True)

for serial, name, f in vi_tenors:
    desc_line(f"VI #{serial}", f, width=25)

# Spread stats
bright_vals = [f["descriptors"]["brightness"]*100 for _, _, f in vi_tenors]
rich_vals = [f["descriptors"]["richness"]*100 for _, _, f in vi_tenors]
full_vals = [f["descriptors"]["fullness"]*100 for _, _, f in vi_tenors]

print(f"\n  Brightness spread: {min(bright_vals):.1f}% — {max(bright_vals):.1f}% "
      f"(range {max(bright_vals)-min(bright_vals):.1f}%)")
print(f"  Complexity spread: {min(rich_vals):.1f}% — {max(rich_vals):.1f}% "
      f"(range {max(rich_vals)-min(rich_vals):.1f}%)")
print(f"  Fullness spread:   {min(full_vals):.1f}% — {max(full_vals):.1f}% "
      f"(range {max(full_vals)-min(full_vals):.1f}%)")


# ============================================================
# 3. CROSS-MAKE ALTO COMPARISON (same player/mpc/reed)
# ============================================================
section("ALTO COMPARISON — different makes, same player/setup")
print("  Player: Tyler Anderson | Mpc: GSNY 7 | Reed: Rigotti\n")

altos = []
for name, data in sorted(profiles["Tyler Anderson"].items()):
    if data.get("horn_type") == "Alto":
        f = fp("Tyler Anderson", name)
        altos.append((name, f))

altos.sort(key=lambda x: x[1]["descriptors"]["brightness"], reverse=True)

for name, f in altos:
    desc_line(name, f, width=40)


# ============================================================
# 4. CROSS-MAKE TENOR COMPARISON (same player/mpc/reed)
# ============================================================
section("TENOR COMPARISON — different makes, same player/setup")
print("  Player: Tyler Anderson | Mpc: GS Slant 7* | Reed: Rigotti\n")

tenors = []
for name, data in sorted(profiles["Tyler Anderson"].items()):
    if data.get("horn_type") == "Tenor":
        f = fp("Tyler Anderson", name)
        tenors.append((name, f))

tenors.sort(key=lambda x: x[1]["descriptors"]["brightness"], reverse=True)

for name, f in tenors:
    desc_line(name, f, width=40)


# ============================================================
# 5. MATT vs TYLER BRIGHTNESS OFFSET
# ============================================================
section("SYSTEMATIC BRIGHTNESS OFFSET — Matt vs Tyler")
print("  If Tyler's recordings are consistently darker, it suggests")
print("  player/mpc/reed effect rather than horn character.\n")

matt_profiles = []
for name, data in profiles.get("My Profiles", {}).items():
    if data.get("sessions"):
        f = fp("My Profiles", name)
        matt_profiles.append((name, data.get("horn_type", "?"), f))

tyler_profiles = []
for name, data in profiles.get("Tyler Anderson", {}).items():
    f = fp("Tyler Anderson", name)
    tyler_profiles.append((name, data.get("horn_type", "?"), f))

matt_bright = [f["descriptors"]["brightness"]*100 for _, _, f in matt_profiles]
tyler_bright = [f["descriptors"]["brightness"]*100 for _, _, f in tyler_profiles]

print(f"  Matt's profiles ({len(matt_bright)} horns):")
print(f"    Brightness: {min(matt_bright):.1f}% — {max(matt_bright):.1f}%, "
      f"mean {sum(matt_bright)/len(matt_bright):.1f}%")

print(f"  Tyler's profiles ({len(tyler_bright)} horns):")
print(f"    Brightness: {min(tyler_bright):.1f}% — {max(tyler_bright):.1f}%, "
      f"mean {sum(tyler_bright)/len(tyler_bright):.1f}%")

offset = (sum(matt_bright)/len(matt_bright)) - (sum(tyler_bright)/len(tyler_bright))
print(f"\n  Systematic offset: Matt reads {offset:+.1f}% brighter on average")
print(f"  (Conn Virtuoso alone: {matt_conn['descriptors']['brightness']*100 - tyler_conn['descriptors']['brightness']*100:+.1f}% difference on same horn)")
