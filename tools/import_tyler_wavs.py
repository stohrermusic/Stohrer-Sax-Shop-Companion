"""
Batch import Tyler Anderson's WAV recordings into toner profiles.

Creates profiles under the "Tyler Anderson" library, one per horn.
Parses horn make/model/serial from filenames.

Usage:
    python tools/import_tyler_wavs.py
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

LIBRARY = "Tyler Anderson"
PLAYER = "Tyler Anderson"

# Tyler's gear
TENOR_MPC = "GS Slant 7*"
ALTO_MPC = "GSNY 7"
REED = "Rigotti"

# Map filenames to (make, model, serial)
# Format: "Filename stem" -> (make, model, serial)
ALTO_FILES = {
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
    # Skip duplicates:
    # "Yamaha YAS 62 11959" is same horn as "YAS-62 11959"
    # "Selmer Mark VI Tenor 194784" is misfiled (tenor in alto folder)
}

TENOR_FILES = {
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

# Also import the misfiled tenor from alto folder
MISFILED_TENOR = {
    "Selmer Mark VI Tenor 194784": ("Selmer", "Mark VI", "194784"),
}

ALTO_DIR = os.path.expanduser("~/Downloads/Altos/Alto")
TENOR_DIR = os.path.expanduser("~/Downloads/Tenors/Tenor")


def make_profile_name(make, model, serial):
    """Generate a profile name like 'Selmer Mark VI #62418'."""
    if model in ("Alto", "Tenor"):
        return f"{make} {model} #{serial}"
    return f"{make} {model} #{serial}"


def import_file(filepath, horn_type, make, model, serial, mpc, profiles):
    """Import one WAV file into the profile library."""
    profile_name = make_profile_name(make, model, serial)
    stem = os.path.splitext(os.path.basename(filepath))[0]

    # Check if profile already has this file
    if profile_name in profiles.get(LIBRARY, {}):
        existing = profiles[LIBRARY][profile_name]
        for session in existing.get("sessions", []):
            for cap in session.get("captures", []):
                if cap.get("source_file") == os.path.basename(filepath):
                    print(f"  SKIP (already imported): {stem}")
                    return 0

    print(f"  Analyzing: {stem} ...", end="", flush=True)

    engine = TonerEngine()
    engine.sax_type = horn_type
    captures = analyze_audio_file(filepath, engine)

    if not captures:
        print(" NO STABLE NOTES FOUND")
        return 0

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    for cap in captures:
        cap["timestamp"] = now
        cap["source_file"] = os.path.basename(filepath)
        cap["source_notes"] = f"Recorded by {PLAYER}"

    # Ensure library exists
    if LIBRARY not in profiles:
        profiles[LIBRARY] = {}

    # Create or get profile
    if profile_name not in profiles[LIBRARY]:
        profiles[LIBRARY][profile_name] = {
            "horn_type": horn_type,
            "horn_make": make,
            "horn_model": model,
            "serial": serial,
            "player": PLAYER,
            "mouthpiece": mpc,
            "reed": REED,
            "notes": "",
            "created": now,
            "sessions": [],
        }

    # Add session
    session = {
        "date": now,
        "captures": captures,
        "method": "file",
        "source_notes": f"Recorded by {PLAYER}",
    }
    profiles[LIBRARY][profile_name]["sessions"].append(session)

    notes = set(c["note"] for c in captures)
    print(f" {len(captures)} captures, {len(notes)} unique notes")
    return len(captures)


def main():
    profile_path = TONER_DATA_FILE
    profiles = load_tone_presets(profile_path)

    total_captures = 0
    total_profiles = 0

    # --- Altos ---
    print(f"\n=== ALTOS ({ALTO_MPC}, {REED}) ===\n")
    for stem, (make, model, serial) in sorted(ALTO_FILES.items()):
        filepath = os.path.join(ALTO_DIR, stem + ".wav")
        if not os.path.exists(filepath):
            print(f"  MISSING: {filepath}")
            continue
        n = import_file(filepath, "Alto", make, model, serial, ALTO_MPC, profiles)
        if n > 0:
            total_captures += n
            total_profiles += 1

    # --- Tenors ---
    print(f"\n=== TENORS ({TENOR_MPC}, {REED}) ===\n")
    for stem, (make, model, serial) in sorted(TENOR_FILES.items()):
        filepath = os.path.join(TENOR_DIR, stem + ".wav")
        if not os.path.exists(filepath):
            print(f"  MISSING: {filepath}")
            continue
        n = import_file(filepath, "Tenor", make, model, serial, TENOR_MPC, profiles)
        if n > 0:
            total_captures += n
            total_profiles += 1

    # --- Misfiled tenor in alto folder ---
    for stem, (make, model, serial) in MISFILED_TENOR.items():
        filepath = os.path.join(ALTO_DIR, stem + ".wav")
        if not os.path.exists(filepath):
            print(f"  MISSING: {filepath}")
            continue
        n = import_file(filepath, "Tenor", make, model, serial, TENOR_MPC, profiles)
        if n > 0:
            total_captures += n
            total_profiles += 1

    # Save
    if total_captures > 0:
        save_tone_presets(profiles, profile_path)
        print(f"\n{'='*50}")
        print(f"DONE: {total_profiles} profiles, {total_captures} total captures")
        print(f"Saved to: {profile_path}")
    else:
        print("\nNo captures to save.")


if __name__ == "__main__":
    main()
