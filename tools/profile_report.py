"""
Single-profile report generator.

Reads a tone profile and prints a detailed human-readable report
covering: setup info, per-note harmonic data, descriptor averages,
register analysis, and notable findings.

Usage:
    python tools/profile_report.py [profile_name]

If no name given, lists available profiles.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from toner_engine import (
    load_tone_presets, compute_fingerprint, BREAK_FREQUENCIES,
)
from config import TONER_DATA_FILE


def _note_sort_key(note_name):
    pc_order = {'C': 0, 'C#': 1, 'D': 2, 'D#': 3, 'E': 4, 'F': 5,
                'F#': 6, 'G': 7, 'G#': 8, 'A': 9, 'A#': 10, 'B': 11}
    try:
        if '#' in note_name:
            pc = note_name[:-1]
            octave = int(note_name[-1])
        else:
            pc = note_name[:-1]
            octave = int(note_name[-1])
        return (octave + 1) * 12 + pc_order.get(pc, 0)
    except (ValueError, KeyError):
        return 0


def report(profile_name, profile_data):
    """Print a detailed report for a single profile."""
    p = profile_data
    sessions = p.get('sessions', [])
    fp = compute_fingerprint(sessions)

    horn_type = p.get('horn_type', '?')
    break_freq = BREAK_FREQUENCIES.get(horn_type, 750)

    print("=" * 65)
    print(f"TONE PROFILE REPORT: {profile_name}")
    print("=" * 65)

    # Setup
    print(f"\nHorn:       {p.get('horn_make', '')} {p.get('horn_model', '')}")
    if p.get('serial'):
        print(f"Serial:     {p['serial']}")
    print(f"Type:       {horn_type} (break freq: {break_freq} Hz)")
    if p.get('player'):
        print(f"Player:     {p['player']}")
    if p.get('mouthpiece'):
        print(f"Mouthpiece: {p['mouthpiece']}")
    if p.get('reed'):
        print(f"Reed:       {p['reed']}")
    if p.get('notes'):
        print(f"Notes:      {p['notes']}")

    # Session summary
    all_captures = []
    for s in sessions:
        all_captures.extend(s.get('captures', []))

    methods = {}
    for c in all_captures:
        m = c.get('method', 'structured')
        methods[m] = methods.get(m, 0) + 1

    total_frames = sum(c.get('n_frames', 0) for c in all_captures)
    unique_notes = set(c.get('note', '') for c in all_captures)

    print(f"\nCaptures:   {len(all_captures)} across {len(sessions)} sessions")
    print(f"Notes:      {len(unique_notes)} unique")
    if methods:
        print(f"Methods:    {', '.join(f'{v} {k}' for k, v in methods.items())}")
    if total_frames:
        avg = total_frames / len(all_captures) if all_captures else 0
        print(f"Avg frames: {avg:.0f} per capture ({avg * 0.033:.1f}s)")

    # Overall descriptors
    print(f"\n{'=' * 65}")
    print("OVERALL DESCRIPTORS (equal weight per note)")
    print(f"{'=' * 65}")
    d = fp.get('descriptors', {})
    print(f"\n  Resonance:  {d.get('resonance', 0):6.1%}")
    print(f"  Richness:   {d.get('richness', 0):6.1%}")
    print(f"  Brightness: {d.get('brightness', 0):6.1%}")
    print(f"  Darkness:   {d.get('darkness', 0):6.1%}")
    print(f"  Fullness:   {d.get('fullness', 0):6.1%}")

    # Overall harmonic profile
    hdb = fp.get('harmonics_db', [])
    if hdb:
        print("\n  Harmonic profile (avg dB relative to fundamental):")
        for i, db in enumerate(hdb):
            bar_len = max(0, int((db + 60) / 60 * 30))
            bar = "#" * bar_len + "." * (30 - bar_len)
            print(f"    H{i+1:2d}: {db:+6.1f} dB  {bar}")

    # Per-note breakdown
    print(f"\n{'=' * 65}")
    print("PER-NOTE BREAKDOWN")
    print(f"{'=' * 65}")

    per_note = fp.get('per_note', {})
    sorted_notes = sorted(per_note.keys(), key=_note_sort_key)

    print(f"\n  {'Note':>5} {'Freq':>7} {'Res':>5} {'Rich':>5} {'Bri':>5} {'Drk':>5} {'Full':>5} {'Frames':>6}  Strongest harmonic")
    print(f"  {'-'*60}")

    for note in sorted_notes:
        pn = per_note[note]
        nd = pn.get('descriptors', {})
        hdb = pn.get('harmonics_db', [])

        # Find captures for this note to get avg freq and frames
        note_caps = [c for c in all_captures if c.get('note') == note]
        avg_freq = sum(c.get('fundamental_freq', 0) for c in note_caps) / len(note_caps) if note_caps else 0
        total_f = sum(c.get('n_frames', 0) for c in note_caps)

        # Strongest harmonic (excluding fundamental)
        strongest = ""
        if len(hdb) > 1:
            max_db = max(hdb[1:])
            max_h = hdb.index(max_db) + 1
            if max_db > -10:
                strongest = f"H{max_h} at {max_db:+.1f} dB"

        print(f"  {note:>5} {avg_freq:6.0f}Hz {nd.get('resonance',0):5.0%} {nd.get('richness',0):5.0%} "
              f"{nd.get('brightness',0):5.0%} {nd.get('darkness',0):5.0%} {nd.get('fullness',0):5.0%} "
              f"{total_f:6d}  {strongest}")

    # Register analysis
    print(f"\n{'=' * 65}")
    print("REGISTER ANALYSIS")
    print(f"{'=' * 65}")

    low_notes = [n for n in sorted_notes if _note_sort_key(n) < 48]
    mid_notes = [n for n in sorted_notes if 48 <= _note_sort_key(n) < 72]
    high_notes = [n for n in sorted_notes if _note_sort_key(n) >= 72]

    for label, notes in [("Low (below C4)", low_notes), ("Mid (C4-B5)", mid_notes), ("High (C6+)", high_notes)]:
        if not notes:
            continue
        descs = [per_note[n].get('descriptors', {}) for n in notes]
        print(f"\n  {label}: {len(notes)} notes")
        for key in ['resonance', 'richness', 'brightness', 'darkness', 'fullness']:
            vals = [d.get(key, 0) for d in descs]
            avg = sum(vals) / len(vals)
            print(f"    {key:>12}: {avg:.0%}", end="")
            if len(vals) > 1:
                print(f"  (range {min(vals):.0%}-{max(vals):.0%})", end="")
            print()

    # Notable findings
    print(f"\n{'=' * 65}")
    print("NOTABLE FINDINGS")
    print(f"{'=' * 65}\n")

    findings = []

    # Notes where upper harmonics are stronger than fundamental
    for note in sorted_notes:
        hdb = per_note[note].get('harmonics_db', [])
        for i, db in enumerate(hdb):
            if i > 0 and db > 0:
                findings.append(f"  {note}: H{i+1} is {db:+.1f} dB above fundamental")

    # Most/least resonant
    if sorted_notes:
        res_sorted = sorted(sorted_notes, key=lambda n: per_note[n].get('descriptors', {}).get('resonance', 0))
        least = res_sorted[0]
        most = res_sorted[-1]
        lr = per_note[least]['descriptors'].get('resonance', 0)
        mr = per_note[most]['descriptors'].get('resonance', 0)
        if mr - lr > 0.05:
            findings.append(f"  Most resonant: {most} ({mr:.0%}), least: {least} ({lr:.0%})")

    # Richest/purest
    if sorted_notes:
        rich_sorted = sorted(sorted_notes, key=lambda n: per_note[n].get('descriptors', {}).get('richness', 0))
        purest = rich_sorted[0]
        richest = rich_sorted[-1]
        pr = per_note[purest]['descriptors'].get('richness', 0)
        rr = per_note[richest]['descriptors'].get('richness', 0)
        if rr - pr > 0.1:
            findings.append(f"  Richest: {richest} ({rr:.0%}), purest: {purest} ({pr:.0%})")

    if findings:
        for f in findings:
            print(f)
    else:
        print("  No notable findings.")

    print(f"\n{'=' * 65}")
    print("NOTE")
    print(f"{'=' * 65}\n")
    print("  Room acoustics, microphone frequency response, and mic placement")
    print("  all affect these readings. Anomalies that appear on a single")
    print("  profile may be artifacts of the recording environment rather than")
    print("  the horn. Patterns that repeat across multiple profiles recorded")
    print("  in the same room likely reflect the room or mic, not the horns.")
    print("  Compare profiles recorded in different environments to distinguish")
    print("  horn character from room character.")
    print()


def main():
    profiles = load_tone_presets(TONER_DATA_FILE)
    if not profiles:
        print("No profiles found.")
        return

    # Build flat list
    all_profiles = []
    for lib_name, lib_profiles in profiles.items():
        if not isinstance(lib_profiles, dict):
            continue
        for prof_name, prof_data in lib_profiles.items():
            all_profiles.append((lib_name, prof_name, prof_data))

    if len(sys.argv) > 1:
        # Search for matching profile
        search = " ".join(sys.argv[1:]).lower()
        for lib, name, data in all_profiles:
            if search in name.lower():
                report(name, data)
                return
        print(f"No profile matching '{search}' found.")
    else:
        print("Available profiles:")
        for lib, name, data in all_profiles:
            sessions = data.get('sessions', [])
            caps = sum(len(s.get('captures', [])) for s in sessions)
            print(f"  [{lib}] {name} ({caps} captures)")
        print("\nUsage: python tools/profile_report.py <profile name>")


if __name__ == "__main__":
    main()
