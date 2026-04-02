"""
Batch import WAV recordings into toner profiles.

Re-imports Tyler Anderson's library (replacing old H12 data with H20),
imports Thomas Edinger's Head2Head tenor comparison,
and imports Mario Larios-García's alto and soprano recordings.

Usage:
    python tools/import_wavs.py              # import all
    python tools/import_wavs.py tyler        # Tyler only
    python tools/import_wavs.py edinger      # Edinger only
    python tools/import_wavs.py mario        # Mario only
"""

import sys
import os
import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from toner_engine import (
    TonerEngine, analyze_audio_file,
    load_tone_presets, save_tone_presets,
)
from config import TONER_DATA_FILE

RECORDINGS_BASE = os.path.expanduser("~/Downloads/recordings")


# ============================================================
# TYLER ANDERSON
# ============================================================

TYLER_LIBRARY = "Tyler Anderson"
TYLER_PLAYER = "Tyler Anderson"
TYLER_TENOR_MPC = "GS Slant 7*"
TYLER_ALTO_MPC = "GSNY 7"
TYLER_REED = "Rigotti"
TYLER_MIC_TYPE = "condenser"
TYLER_MIC_MODEL = "Neumann TLM"

TYLER_ALTO_DIR = os.path.join(RECORDINGS_BASE, "Tyler Anderson", "Alto", "Alto")
TYLER_TENOR_DIR = os.path.join(RECORDINGS_BASE, "Tyler Anderson", "Tenor", "Tenor")

TYLER_ALTO_FILES = {
    "Buescher 267734":           ("Buescher", "Alto", "267734"),
    "Buescher 307777":           ("Buescher", "Alto", "307777"),
    "Keilwerth New King 38431":  ("Keilwerth", "New King", "38431"),
    "Mark VI 219548":            ("Selmer", "Mark VI", "219548"),
    "Selmer Mark VI 190366":     ("Selmer", "Mark VI", "190366"),
    "Selmer Mark VI 93668":      ("Selmer", "Mark VI", "93668"),
    "Selmer Supreme 842487":     ("Selmer", "Supreme", "842487"),
    "YAS-62 11959":              ("Yamaha", "YAS-62", "11959"),
    "YAS-62 88666":              ("Yamaha", "YAS-62", "88666"),
    "Matt Stohrer Conn":         ("Conn", "NW2 Virtuoso Deluxe", "205k"),
    # "Yamaha YAS 62 11959" is duplicate of "YAS-62 11959"
    # "Selmer Mark VI Tenor 194784" is misfiled tenor — listed below
}

TYLER_TENOR_FILES = {
    "Conn 274254":               ("Conn", "Tenor", "274254"),
    "Conn 329232":               ("Conn", "Tenor", "329232"),
    "Couf Superba 1":            ("Couf", "Superba", "1"),
    "King Super 20 375203":      ("King", "Super 20", "375203"),
    "Mark VI 100115":            ("Selmer", "Mark VI", "100115"),
    "Mark VI 141672":            ("Selmer", "Mark VI", "141672"),
    "Mark VI 147401":            ("Selmer", "Mark VI", "147401"),
    "Mark VI 157137":            ("Selmer", "Mark VI", "157137"),
    "Mark VI 164746":            ("Selmer", "Mark VI", "164746"),
    "Mark VI 62418":             ("Selmer", "Mark VI", "62418"),
    "Mark VI 72434":             ("Selmer", "Mark VI", "72434"),
    "SML 21321":                 ("SML", "Tenor", "21321"),
    "Selmer 38438":              ("Selmer", "Tenor", "38438"),
    "Tenor Madness 1222":        ("Tenor Madness", "Tenor", "1222"),
    "Yamaha 875EX 6598":         ("Yamaha", "875EX", "6598"),
}

# Misfiled tenor in alto folder
TYLER_MISFILED_TENOR = {
    "Selmer Mark VI Tenor 194784": ("Selmer", "Mark VI", "194784"),
}


# ============================================================
# THOMAS EDINGER — Saxophones Head2Head (all tenors)
# ============================================================

EDINGER_LIBRARY = "Thomas Edinger"
EDINGER_PLAYER = "Thomas Edinger"
EDINGER_MPC = "Philtone Intrepid (.106, refaced by Tommy Occhiuto)"
EDINGER_REED = "BSS Silver Label 2.5"
EDINGER_MIC_TYPE = "condenser"
EDINGER_MIC_MODEL = "Neumann U67 > Neve 1272 > UA Apollo"

EDINGER_DIR = os.path.join(RECORDINGS_BASE, "Thomas Edinger", "Thomas Edinger")

# Map filenames to (make, model, serial, notes)
# All tenors. #68630 recorded with three different necks.
EDINGER_FILES = {
    "Saxophones Head2Head-Selmer #35737": (
        "Selmer", "SBA", "35737",
        "1948. Lacquer (maybe relaq) with EU engraving, no high F#"),
    "Saxophones Head2Head-Selmer #68611": (
        "Selmer", "Mk VI", "68611",
        "1957. Lacquer stripped (previously relacquered wrong color), no high F#, no engraving"),
    "Saxophones Head2Head-Selmer #68630 with original neck": (
        "Selmer", "Mk VI", "68630",
        "1957. Cellulose lacquer (dark honey gold), no high F#, no engraving. Original neck."),
    "Saxophones Head2Head-Selmer #68630 with neck from #122028": (
        "Selmer", "Mk VI", "68630 (neck from #122028)",
        "1957. Same horn as #68630 but with neck from #122028."),
    "Saxophones Head2Head-Selmer #68630 with later Mk VI replacement neck": (
        "Selmer", "Mk VI", "68630 (later replacement neck)",
        "1957. Same horn as #68630 but with a later Mk VI replacement neck."),
    "Saxophones Head2Head-Selmer #71352": (
        "Selmer", "Mk VI", "71352",
        "1957. Cellulose lacquer (dark honey gold), no high F#, no engraving"),
    "Saxophones Head2Head-Selmer #76047": (
        "Selmer", "Mk VI", "76047",
        "1958. Sandblasted finish. Heavily restored, palmkey toneholes restored with rings, EU horn"),
    "Saxophones Head2Head-Selmer #85961": (
        "Selmer", "Mk VI", "85961",
        "1960. Lacquer stripped, US engraving, soldered body-to-bow, no high F#, plugged hole in neck"),
    "Saxophones Head2Head-Selmer #97226": (
        "Selmer", "Mk VI", "97226",
        "1962. Lacquer/silver plated keys, EU engraving, no high F#"),
    "Saxophones Head2Head-Selmer #99314": (
        "Selmer", "Mk VI", "99314",
        "1962. Lacquer with EU engraving, high F#, possible relaq"),
    "Saxophones Head2Head-Selmer #122028": (
        "Selmer", "Mk VI", "122028",
        "1965. Lacquer, no engraving, no high F#. 028 under neck octave key."),
    "Saxophones Head2Head-Selmer #141956": (
        "Selmer", "Mk VI", "141956",
        "1967. Lacquer, silver plated keys, EU engraving, no F#"),
    "Saxophones Head2Head-SML Gold Medal #15840": (
        "SML", "Gold Medal", "15840",
        "1957"),
    "Saxophones Head2Head-Yamaha YTS-62 Purple Logo #017542": (
        "Yamaha", "YTS-62 Purple Logo", "017542",
        "High F#, engraving"),
}


# ============================================================
# MARIO LARIOS-GARCÍA
# ============================================================

MARIO_LIBRARY = "Mario Larios-García"
MARIO_PLAYER = "Mario Larios-García"
MARIO_MIC_TYPE = "condenser"
MARIO_MIC_MODEL = "Audio Technica AT2035"

MARIO_DIR = os.path.join("C:\\sax shop companion\\recordings", "Mario Larios-García")

MARIO_FILES = [
    {
        "file": "Alto Sax Demo.wav",
        "horn_type": "Alto",
        "make": "Buescher",
        "model": "True-Tone",
        "serial": "1926",
        "mpc": "Marvell Carpenter/PureVibe Power",
        "reed": "Boston Sax Shop silver box 2.5",
        "notes": "1926 vintage. Ligature: Marvell Carpenter/PureVibe.",
    },
    {
        "file": "Soprano Sax Demo.wav",
        "horn_type": "Soprano",
        "make": "Yamaha",
        "model": "YSS-475",
        "serial": "",
        "mpc": "AM Mouthpieces Aras 7",
        "reed": "Boston Sax Shop black box 2.5",
        "notes": "",
    },
]


# ============================================================
# IMPORT LOGIC
# ============================================================

def make_profile_name(make, model, serial):
    """Generate a profile name like 'Selmer Mark VI #62418'."""
    return f"{make} {model} #{serial}"


def import_file(filepath, horn_type, make, model, serial, mpc, reed,
                mic_type, mic_model, player, library, notes, profiles):
    """Import one WAV file into the profile library."""
    profile_name = make_profile_name(make, model, serial)
    stem = os.path.splitext(os.path.basename(filepath))[0]

    print(f"  Analyzing: {stem} ...", end="", flush=True)

    engine = TonerEngine()
    engine.sax_type = horn_type
    captures = analyze_audio_file(filepath, engine)

    if not captures:
        print(f" NO STABLE NOTES FOUND")
        return 0

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    for cap in captures:
        cap["timestamp"] = now
        cap["source_file"] = os.path.basename(filepath)

    if library not in profiles:
        profiles[library] = {}

    if profile_name not in profiles[library]:
        profiles[library][profile_name] = {
            "horn_type": horn_type,
            "horn_make": make,
            "horn_model": model,
            "serial": serial,
            "player": player,
            "mouthpiece": mpc,
            "reed": reed,
            "notes": notes,
            "created": now,
            "sessions": [],
        }

    session = {
        "date": now,
        "captures": captures,
        "method": "file",
        "source_notes": f"Recorded by {player}",
        "mic_type": mic_type,
        "mic_model": mic_model,
    }
    profiles[library][profile_name]["sessions"].append(session)

    unique_notes = set(c["note"] for c in captures)
    print(f" {len(captures)} captures, {len(unique_notes)} unique notes")
    return len(captures)


def import_tyler(profiles):
    """Import/re-import all Tyler Anderson recordings."""
    # Clear existing Tyler library to replace H12 data with H20
    if TYLER_LIBRARY in profiles:
        old_count = len(profiles[TYLER_LIBRARY])
        del profiles[TYLER_LIBRARY]
        print(f"  Cleared {old_count} old profiles (replacing with H20 data)")

    total_captures = 0
    total_profiles = 0

    # Altos
    print(f"\n--- Altos ({TYLER_ALTO_MPC}, {TYLER_REED}) ---\n")
    for stem, (make, model, serial) in sorted(TYLER_ALTO_FILES.items()):
        filepath = os.path.join(TYLER_ALTO_DIR, stem + ".wav")
        if not os.path.exists(filepath):
            print(f"  MISSING: {filepath}")
            continue
        n = import_file(filepath, "Alto", make, model, serial,
                        TYLER_ALTO_MPC, TYLER_REED,
                        TYLER_MIC_TYPE, TYLER_MIC_MODEL,
                        TYLER_PLAYER, TYLER_LIBRARY, "", profiles)
        if n > 0:
            total_captures += n
            total_profiles += 1

    # Tenors
    print(f"\n--- Tenors ({TYLER_TENOR_MPC}, {TYLER_REED}) ---\n")
    for stem, (make, model, serial) in sorted(TYLER_TENOR_FILES.items()):
        filepath = os.path.join(TYLER_TENOR_DIR, stem + ".wav")
        if not os.path.exists(filepath):
            print(f"  MISSING: {filepath}")
            continue
        n = import_file(filepath, "Tenor", make, model, serial,
                        TYLER_TENOR_MPC, TYLER_REED,
                        TYLER_MIC_TYPE, TYLER_MIC_MODEL,
                        TYLER_PLAYER, TYLER_LIBRARY, "", profiles)
        if n > 0:
            total_captures += n
            total_profiles += 1

    # Misfiled tenor in alto folder
    for stem, (make, model, serial) in TYLER_MISFILED_TENOR.items():
        filepath = os.path.join(TYLER_ALTO_DIR, stem + ".wav")
        if not os.path.exists(filepath):
            print(f"  MISSING: {filepath}")
            continue
        n = import_file(filepath, "Tenor", make, model, serial,
                        TYLER_TENOR_MPC, TYLER_REED,
                        TYLER_MIC_TYPE, TYLER_MIC_MODEL,
                        TYLER_PLAYER, TYLER_LIBRARY, "", profiles)
        if n > 0:
            total_captures += n
            total_profiles += 1

    return total_profiles, total_captures


def import_edinger(profiles):
    """Import Thomas Edinger's Head2Head recordings."""
    # Clear existing to allow clean re-import
    if EDINGER_LIBRARY in profiles:
        old_count = len(profiles[EDINGER_LIBRARY])
        del profiles[EDINGER_LIBRARY]
        print(f"  Cleared {old_count} old profiles (re-importing)")

    total_captures = 0
    total_profiles = 0

    print(f"\n--- Tenors (Head2Head comparison) ---\n")
    for stem, (make, model, serial, notes) in sorted(EDINGER_FILES.items()):
        filepath = os.path.join(EDINGER_DIR, stem + ".wav")
        if not os.path.exists(filepath):
            print(f"  MISSING: {filepath}")
            continue
        n = import_file(filepath, "Tenor", make, model, serial,
                        EDINGER_MPC, EDINGER_REED,
                        EDINGER_MIC_TYPE, EDINGER_MIC_MODEL,
                        EDINGER_PLAYER, EDINGER_LIBRARY, notes, profiles)
        if n > 0:
            total_captures += n
            total_profiles += 1

    return total_profiles, total_captures


def import_mario(profiles):
    """Import Mario Larios-García's recordings."""
    # Clear existing to allow clean re-import
    if MARIO_LIBRARY in profiles:
        old_count = len(profiles[MARIO_LIBRARY])
        del profiles[MARIO_LIBRARY]
        print(f"  Cleared {old_count} old profiles (re-importing)")

    total_captures = 0
    total_profiles = 0

    for entry in MARIO_FILES:
        filepath = os.path.join(MARIO_DIR, entry["file"])
        if not os.path.exists(filepath):
            print(f"  MISSING: {filepath}")
            continue
        n = import_file(filepath, entry["horn_type"], entry["make"],
                        entry["model"], entry["serial"], entry["mpc"],
                        entry["reed"], MARIO_MIC_TYPE, MARIO_MIC_MODEL,
                        MARIO_PLAYER, MARIO_LIBRARY, entry["notes"],
                        profiles)
        if n > 0:
            total_captures += n
            total_profiles += 1

    return total_profiles, total_captures


def main():
    which = sys.argv[1].lower() if len(sys.argv) > 1 else "all"

    profile_path = TONER_DATA_FILE
    profiles = load_tone_presets(profile_path)

    grand_profiles = 0
    grand_captures = 0

    if which in ("all", "tyler"):
        print(f"\n{'='*50}")
        print("TYLER ANDERSON")
        print(f"{'='*50}")
        p, c = import_tyler(profiles)
        grand_profiles += p
        grand_captures += c

    if which in ("all", "edinger"):
        print(f"\n{'='*50}")
        print("THOMAS EDINGER — Saxophones Head2Head")
        print(f"{'='*50}")
        p, c = import_edinger(profiles)
        grand_profiles += p
        grand_captures += c

    if which in ("all", "mario"):
        print(f"\n{'='*50}")
        print("MARIO LARIOS-GARCÍA")
        print(f"{'='*50}")
        p, c = import_mario(profiles)
        grand_profiles += p
        grand_captures += c

    if grand_captures > 0:
        save_tone_presets(profiles, profile_path)
        print(f"\n{'='*50}")
        print(f"DONE: {grand_profiles} profiles, {grand_captures} total captures")
        print(f"Saved to: {profile_path}")
    else:
        print("\nNo captures to save.")


if __name__ == "__main__":
    main()
