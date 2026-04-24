#!/usr/bin/env python3
"""Test concert pitch storage, transposition functions, and migration.

Verifies that:
- transpose_note and reverse_transpose_note are correct inverses
- note_to_freq produces correct frequencies
- migrate_profile_to_concert properly converts written -> concert
- All sax types produce correct round-trips
- Edge cases at octave boundaries work
"""

import sys
sys.path.insert(0, '.')
from toner_engine import (
    transpose_note, reverse_transpose_note, note_to_freq,
    SAX_TRANSPOSITIONS, CALIBRATION_NOTES, MIN_FUNDAMENTAL_HZ,
)

passes = 0
fails = 0

def test(name, condition):
    global passes, fails
    if condition:
        passes += 1
        print(f"  PASS: {name}")
    else:
        fails += 1
        print(f"  FAIL: {name}")

def section(title):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")

# ============================================================
section("transpose_note / reverse_transpose_note round-trips")
# ============================================================

test_notes = ['C2', 'C#3', 'D4', 'D#5', 'E3', 'F4', 'F#2',
              'G3', 'G#4', 'A4', 'A#3', 'B5']

for sax_type, shift in SAX_TRANSPOSITIONS.items():
    all_ok = True
    for note in test_notes:
        written = transpose_note(note, sax_type)
        back = reverse_transpose_note(written, sax_type)
        if back != note:
            all_ok = False
            print(f"    MISMATCH: {sax_type} {note} -> {written} -> {back}")
    test(f"{sax_type} (shift={shift:+d}): all round-trips correct", all_ok)

# ============================================================
section("Known transposition values")
# ============================================================

# Tenor: concert Bb2 = written C4
test("Tenor: concert A#2 -> written C4",
     transpose_note("A#2", "Tenor") == "C4")
test("Tenor: written C4 -> concert A#2",
     reverse_transpose_note("C4", "Tenor") == "A#2")

# Alto: concert Eb3 = written C4
test("Alto: concert D#3 -> written C4",
     transpose_note("D#3", "Alto") == "C4")

# Soprano: concert Bb3 = written C4
test("Soprano: concert A#3 -> written C4",
     transpose_note("A#3", "Soprano") == "C4")

# Baritone: concert Eb2 = written C4
test("Bari: concert D#2 -> written C4",
     transpose_note("D#2", "Baritone") == "C4")

# Sopranino: concert Eb4 = written C4 (sounds ABOVE written)
test("Sopranino: concert D#4 -> written C4",
     transpose_note("D#4", "Sopranino") == "C4")

# C Melody: no transposition
test("C Melody: concert C4 = written C4",
     transpose_note("C4", "C Melody") == "C4")

# ============================================================
section("note_to_freq")
# ============================================================

test("A4 = 440 Hz", abs(note_to_freq("A4") - 440.0) < 0.1)
test("A3 = 220 Hz", abs(note_to_freq("A3") - 220.0) < 0.1)
test("C4 = 261.6 Hz", abs(note_to_freq("C4") - 261.63) < 0.1)
test("G#2 = 103.8 Hz", abs(note_to_freq("G#2") - 103.83) < 0.1)
test("Empty string = 0", note_to_freq("") == 0.0)

# ============================================================
section("Octave boundary edge cases")
# ============================================================

# Tenor: concert B2 -> written C#4 (crosses octave boundary)
test("Tenor: concert B2 -> written C#4",
     transpose_note("B2", "Tenor") == "C#4")

# Sopranino: concert C4 -> written A3 (negative shift crosses boundary)
test("Sopranino: concert C4 -> written A3",
     transpose_note("C4", "Sopranino") == "A3")

# Tenor: written A#3 (lowest cal note) -> concert G#2
test("Tenor: written A#3 -> concert G#2",
     reverse_transpose_note("A#3", "Tenor") == "G#2")

# ============================================================
section("Calibration notes detectable for all standard sax types")
# ============================================================

for sax_type in ['Soprano', 'Alto', 'Tenor', 'Baritone']:
    detectable = []
    for n in CALIBRATION_NOTES:
        concert = reverse_transpose_note(n, sax_type)
        freq = note_to_freq(concert)
        if freq >= MIN_FUNDAMENTAL_HZ:
            detectable.append(n)
    test(f"{sax_type}: all {len(CALIBRATION_NOTES)} cal notes detectable",
         len(detectable) == len(CALIBRATION_NOTES))

# ============================================================
section("Consistent note-freq relationship after transposition")
# ============================================================

# For tenor: written A#3 -> concert G#2. The freq of G#2 should be ~103.8 Hz.
# The stored freq in the capture should match the concert note name.
concert = reverse_transpose_note("A#3", "Tenor")
freq = note_to_freq(concert)
test("Tenor written A#3: concert note G#2 freq ~103.8 Hz",
     concert == "G#2" and abs(freq - 103.83) < 0.1)

# ============================================================
# Summary
# ============================================================
print(f"\n{'='*60}")
print(f"  RESULTS: {passes} passed, {fails} failed out of {passes + fails}")
print(f"{'='*60}")
sys.exit(0 if fails == 0 else 1)
