"""Test compute_fingerprint's read-side filter for detection artifacts.

The filter drops two classes of bad captures:
  1. Notes outside SAX_NOTE_RANGES[sax_type] (no altissimo margin —
     altissimo is a harmonic partial, not a true closed-tube fundamental).
  2. Captures where any harmonic value exceeds +20 dB above the labeled
     fundamental — physically impossible, a clear sub-octave detection
     error.

These bogus captures had been silently poisoning preset descriptors
(see the Conn 6M Yanagisawa AC140 case where 5 bad captures shifted the
overall H2 average by ~3.7 dB and the warmth descriptor visibly).
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from toner_engine import (
    compute_fingerprint, _capture_is_plausible,
)

passed = 0
failed = 0


def test(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        failed += 1


def make_capture(note, harmonics_db, fund=440.0):
    return {
        "note": note,
        "fundamental_freq": fund,
        "harmonics_db": harmonics_db,
    }


def realistic_alto_harmonics():
    # Realistic alto mid-register, dB below fundamental
    return [0.0, -8.0, -12.0, -16.0, -22.0, -28.0, -34.0, -40.0]


# ================================================
print("\n=== Test 1: _capture_is_plausible — in-range alto note ===")
cap = make_capture("A4", realistic_alto_harmonics())
test("In-range A4 accepted", _capture_is_plausible(cap, "Alto") is True)

cap = make_capture("C4", realistic_alto_harmonics())
test("In-range C4 accepted", _capture_is_plausible(cap, "Alto") is True)

cap = make_capture("C#3", realistic_alto_harmonics())  # MIDI 49 — alto's old low Bb floor
test("Edge-of-range C#3 accepted (low Bb concert)", _capture_is_plausible(cap, "Alto") is True)

cap = make_capture("C3", realistic_alto_harmonics())  # MIDI 48 — new low A floor
test("Low A C3 (concert) accepted on alto", _capture_is_plausible(cap, "Alto") is True)


# ================================================
print("\n=== Test 2: _capture_is_plausible — out-of-range rejection ===")
cap = make_capture("D2", realistic_alto_harmonics())  # MIDI 38, way below alto range
test("Below-range D2 rejected on alto", _capture_is_plausible(cap, "Alto") is False)

cap = make_capture("C2", realistic_alto_harmonics())
test("Below-range C2 rejected on alto", _capture_is_plausible(cap, "Alto") is False)

cap = make_capture("C6", realistic_alto_harmonics())  # MIDI 84, formerly the upper edge
test("C6 rejected on alto (was altissimo edge)", _capture_is_plausible(cap, "Alto") is False)

cap = make_capture("A5", realistic_alto_harmonics())  # MIDI 81 — new top
test("A5 accepted on alto (high F#6 written)", _capture_is_plausible(cap, "Alto") is True)


# ================================================
print("\n=== Test 3: _capture_is_plausible — impossible H2 rejection ===")
# Sub-octave detection error: real fundamental landed in H2 bin
cap = make_capture("D3", [0.0, 41.0, -5.0, -10.0, -15.0])
test("Capture with H2=+41 dB rejected", _capture_is_plausible(cap, "Alto") is False)

cap = make_capture("D3", [0.0, 21.0, -5.0])  # just over the +20 ceiling
test("Capture with H2=+21 dB rejected", _capture_is_plausible(cap, "Alto") is False)

cap = make_capture("D3", [0.0, 19.0, -5.0])  # just under the ceiling
test("Capture with H2=+19 dB accepted", _capture_is_plausible(cap, "Alto") is True)

# Real low-register alto: H2 at +10 dB is normal physics, not an error
cap = make_capture("D3", [0.0, 10.0, -3.0, -8.0])
test("Capture with realistic low-register H2=+10 dB accepted",
     _capture_is_plausible(cap, "Alto") is True)


# ================================================
print("\n=== Test 4: NaN / inf rejection ===")
cap = make_capture("A4", [0.0, float("nan"), -8.0])
test("Capture with NaN harmonic rejected", _capture_is_plausible(cap, "Alto") is False)

cap = make_capture("A4", [0.0, float("inf"), -8.0])
test("Capture with inf harmonic rejected", _capture_is_plausible(cap, "Alto") is False)


# ================================================
print("\n=== Test 5: compute_fingerprint drops bad captures, keeps good ones ===")
sessions = [{
    "captures": [
        # 4 good in-range alto captures
        make_capture("A4", realistic_alto_harmonics()),
        make_capture("B4", realistic_alto_harmonics()),
        make_capture("C5", realistic_alto_harmonics()),
        make_capture("D5", realistic_alto_harmonics()),
        # 1 out-of-range note (concert D2 — way below alto range)
        make_capture("D2", realistic_alto_harmonics()),
        # 1 in-range note but impossible H2
        make_capture("A4", [0.0, 41.0, -5.0, -10.0]),
    ],
    "mic_type": "condenser",
}]

fp = compute_fingerprint(sessions, "Alto")
test("Bad captures dropped, 4 of 6 kept", fp["capture_count"] == 4)
test("Per-note count matches kept captures", fp["note_count"] == 4)
test("Out-of-range D2 not in per_note", "D2" not in fp["per_note"])

# H2 should be exactly -8.0 (the realistic value); had the +41 dB capture
# leaked through, the average for A4 would be much higher
a4 = fp["per_note"]["A4"]
test("A4 H2 averaged correctly (no contamination)",
     abs(a4["harmonics_db"][1] - (-8.0)) < 0.01)


# ================================================
print("\n=== Test 6: Tenor range bounds (G#2 to E5 concert) ===")
# Note: capture labels always use sharp names (G#) not flats (Ab) because
# PITCH_CLASSES in toner_engine.py is sharp-based.
cap = make_capture("G#2", realistic_alto_harmonics())  # tenor's low Bb concert
test("Low Bb (G#2 concert) accepted on tenor", _capture_is_plausible(cap, "Tenor") is True)

cap = make_capture("E5", realistic_alto_harmonics())  # tenor high F# written
test("High F# (E5 concert) accepted on tenor", _capture_is_plausible(cap, "Tenor") is True)

cap = make_capture("F5", realistic_alto_harmonics())  # one above
test("Altissimo F5 rejected on tenor", _capture_is_plausible(cap, "Tenor") is False)

cap = make_capture("G2", realistic_alto_harmonics())  # below low Bb
test("Below-range G2 rejected on tenor", _capture_is_plausible(cap, "Tenor") is False)


# ================================================
print("\n=== Test 7: Baritone range bounds (C2 low A through A4 high F#) ===")
cap = make_capture("C2", realistic_alto_harmonics())  # bari low A concert
test("Low A (C2 concert) accepted on baritone", _capture_is_plausible(cap, "Baritone") is True)

cap = make_capture("A4", realistic_alto_harmonics())  # bari high F# written
test("High F# (A4 concert) accepted on baritone", _capture_is_plausible(cap, "Baritone") is True)

cap = make_capture("B1", realistic_alto_harmonics())  # below low A
test("B1 rejected on baritone", _capture_is_plausible(cap, "Baritone") is False)


# ================================================
print("\n=== Test 8: Soprano range bounds (G#3 low Bb through E6 high F#) ===")
cap = make_capture("G#3", realistic_alto_harmonics())
test("Low Bb (G#3 concert) accepted on soprano", _capture_is_plausible(cap, "Soprano") is True)

cap = make_capture("E6", realistic_alto_harmonics())
test("High F# (E6 concert) accepted on soprano", _capture_is_plausible(cap, "Soprano") is True)

cap = make_capture("F6", realistic_alto_harmonics())
test("Altissimo F6 rejected on soprano", _capture_is_plausible(cap, "Soprano") is False)


# ================================================
print("\n=== Test 9: Unknown sax_type passes the range check ===")
# An unrecognized sax type shouldn't crash; only the H2 plausibility filter applies.
cap = make_capture("A4", realistic_alto_harmonics())
test("Unknown sax type accepts in-range note", _capture_is_plausible(cap, "Slide Whistle") is True)

cap = make_capture("A4", [0.0, 41.0])
test("Unknown sax type still drops impossible H2",
     _capture_is_plausible(cap, "Slide Whistle") is False)


# ================================================
print(f"\n{'='*40}")
print(f"RESULTS: {passed} passed, {failed} failed")
print(f"{'='*40}")

sys.exit(0 if failed == 0 else 1)
