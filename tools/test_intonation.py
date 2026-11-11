"""intonation.py: the analyzer's captures as each note's cents, as played.

The cases
(the median of takes, the spread, written names by saxophone, the plausibility
filter, the reference pitch, the offset and the per-register figures) and add the
chart's data (spreads, the overlay of several runs, the register changes).
Headless, no audio. Run:
    python tools/test_intonation.py
"""
import math
import os
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Never the user's real settings folder: an isolated profile, set before any
# project module is imported (tools/test_config_isolation.py checks this).
os.environ["SAXSHOP_CONFIG_DIR"] = tempfile.mkdtemp(prefix="ssc-intonation-")

import intonation as I  # noqa: E402

results = []


def check(name, fn):
    try:
        fn()
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


def approx(a, b, tol=1e-6):
    assert a is not None and abs(a - b) <= tol, f"{a} != {b}"


def hz(midi, a4=440.0):
    return a4 * 2 ** ((midi - 69) / 12)


def cap(note, f0, frames=30, harm=(0.0, -3.0, -6.0), **extra):
    """A capture as the analyzer stores it (the keys the run reads)."""
    return {"note": note, "fundamental_freq": f0, "harmonics_db": list(harm), "n_frames": frames,
            "method": "free", **extra}


def detuned(midi, cents, a4=440.0):
    return hz(midi, a4) * 2 ** (cents / 1200.0)


print("Intonation from captures")
print("=" * 60)


def written_names_follow_the_saxophone():
    # A tenor's concert G#2 (MIDI 44) is its written Bb3 (58), low register; concert C4 (60) is
    # written D5 (74), the octave key.
    assert I.written_midi(44, "Tenor") == 58 and I.written_midi(60, "Tenor") == 74, "I.written_midi(44, \"Tenor\") == 58 and I.written_midi(60, \"Tenor\") == 74"
    assert I.register_of_written(58) == 1 and I.register_of_written(73) == 1 and I.register_of_written(74) == 2, "I.register_of_written(58) == 1 and I.register_of_written(73) == 1 and I.register_of_written(74) == 2"
    assert I.written_band(58) == "low" and I.written_band(74) == "octave_key" and I.written_band(86) == "palm_keys", "I.written_band(58) == \"low\" and I.written_band(74) == \"octave_key\" and I.written_band(86) == \"palm_keys\""
    # An alto's concert A4 (69) is its written F#5 (78).
    assert I.written_midi(69, "Alto") == 78, "I.written_midi(69, \"Alto\") == 78"
    try:
        I.written_midi(60, "Ocarina")
    except I.IntonationError as e:
        assert "Ocarina" in str(e), "\"Ocarina\" in str(e)"
    else:
        raise AssertionError("an unknown saxophone type must be refused")
check("written names, as a player writes them, and registers follow the saxophone (tenor G#2 -> Bb3 low, "
      "C4 -> D5 octave key; alto A4 -> F#5)",
      written_names_follow_the_saxophone)


def a_preset_is_a_run_in_written_names_against_equal_temperament():
    caps = [cap("G#2", detuned(44, 6)), cap("G#2", detuned(44, 10)),
            cap("C4", detuned(60, -4)),
            cap("C1", 32.7),                                  # outside the tenor's range
            cap("C4", hz(60), harm=(0.0, 25.0))]              # an impossible harmonic
    run = I.run_from_captures([{"captures": caps, "date": "2026-11-10"}], "Tenor", name="Bench / Mark VI")
    assert run.sax_type == "Tenor" and run.name == "Bench / Mark VI" and run.reference_hz == 440.0, "run.sax_type == \"Tenor\" and run.name == \"Bench / Mark VI\" and run.reference_hz == 440.0"
    assert run.notes == ["Bb3", "D5"], run.notes
    played = run.played()
    approx(played["Bb3"][0], 8.0)                             # the median of 6 and 10
    assert played["Bb3"][1] == 1 and played["D5"][1] == 2, "played[\"Bb3\"][1] == 1 and played[\"D5\"][1] == 2"
    approx(played["D5"][0], -4.0)
    a = run.reading("Bb3")
    assert len(a.takes) == 2 and a.wobble_cents is None and approx(a.take_spread_cents, 4.0) is None, "len(a.takes) == 2 and a.wobble_cents is None and approx(a.take_spread_cents, 4.0) is None"
    assert a.takes[0].seconds == 30 * I.FRAME_SECONDS and a.takes[0].session_date == "2026-11-10", "a.takes[0].seconds == 30 * I.FRAME_SECONDS and a.takes[0].session_date == \"2026-11-10\""
    assert run.left_out == {"implausible": 2, "unlabelled": 0, "off_note": 0}, run.left_out
    assert run.sessions_used == 1 and run.captures_used == 3, "run.sessions_used == 1 and run.captures_used == 3"
check("a preset's captures are a run in written names against equal temperament; the median of takes; "
      "the fingerprint's filter applies and is counted", a_preset_is_a_run_in_written_names_against_equal_temperament)


def the_median_resists_one_bad_take():
    caps = [cap("G#2", detuned(44, -3)), cap("G#2", detuned(44, -5)), cap("G#2", detuned(44, 40))]
    r = I.run_from_captures([{"captures": caps}], "Tenor").reading("Bb3")
    approx(r.cents, -3.0)                                     # the median, not the mean (10.7)
    approx(r.low, -5.0)
    approx(r.high, 40.0)
    approx(r.take_spread_cents, 45.0)
    one = I.run_from_captures([{"captures": [cap("G#2", detuned(44, -3))]}], "Tenor").reading("Bb3")
    assert one.take_spread_cents is None and approx(one.cents, -3.0) is None, "one.take_spread_cents is None and approx(one.cents, -3.0) is None"
check("a note's cents are the median of its takes (one bad take cannot drag it); the spread is max - min, "
      "None under two takes", the_median_resists_one_bad_take)


def cents_follow_the_reference_pitch():
    a4 = cap("A4", 442.0)                                     # an alto's written F#5
    run = I.run_from_captures([{"captures": [a4]}], "Alto", reference_hz=442.0)
    approx(run.played()["F#5"][0], 0.0, 1e-9)
    run = I.run_from_captures([{"captures": [a4]}], "Alto", reference_hz=440.0)
    approx(run.played()["F#5"][0], 1200 * math.log2(442 / 440))
    approx(I.midi_to_hz(57, 440.0), 220.0)                   # concert A3 = a tenor's written B4
    approx(I.midi_to_hz(69, 442.0), 442.0)
check("cents follow the reference pitch (A4 at 442 is 0 c at A=442, +7.85 c at A=440)", cents_follow_the_reference_pitch)


def unlabelled_and_broken_captures_are_counted_not_crashed():
    caps = [cap("", 220.0), cap("A3", 0.0), cap("A3", "x"), {"note": "A3"}, cap("A3", 220.0, frames="many")]
    run = I.run_from_captures([{"captures": caps}], "Tenor")
    # "" has no MIDI number, 0.0, "x" and the missing fundamental no pitch: all four unlabelled
    assert run.left_out == {"implausible": 0, "unlabelled": 4, "off_note": 0}, run.left_out
    assert run.notes == ["B4"] and run.reading("B4").takes[0].frames == 0, "run.notes == [\"B4\"] and run.reading(\"B4\").takes[0].frames == 0"
    empty = I.run_from_captures([], "Tenor")
    assert empty.notes == [] and empty.played() == {} and I.note_spreads(empty) == [], "empty.notes == [] and empty.played() == {} and I.note_spreads(empty) == []"
    assert I.offset_and_shape(empty)["offset"] is None and I.low_register_offset(empty) is None, "I.offset_and_shape(empty)[\"offset\"] is None and I.low_register_offset(empty) is None"
check("captures with no note, no fundamental or broken fields are counted and skipped; an empty run is empty",
      unlabelled_and_broken_captures_are_counted_not_crashed)


def a_preset_and_its_chosen_sessions():
    preset = {"name": "Mark VI", "horn_type": "Tenor",
              "sessions": [{"date": "d1", "captures": [cap("G#2", detuned(44, 6))]},
                           {"date": "d2", "captures": [cap("G#2", detuned(44, 10)), cap("C4", detuned(60, -4))]},
                           {"date": "d3", "captures": [cap("C4", detuned(60, 20))]}]}
    whole = I.run_from_preset(preset)
    assert whole.name == "Mark VI" and whole.sessions_used == 3 and whole.captures_used == 4, "whole.name == \"Mark VI\" and whole.sessions_used == 3 and whole.captures_used == 4"
    approx(whole.played()["Bb3"][0], 8.0)
    approx(whole.played()["D5"][0], 8.0)                     # the median of -4 and 20
    some = I.run_from_preset(preset, session_indices=[1], name="d2 only")
    assert some.name == "d2 only" and some.sessions_used == 1 and some.captures_used == 2, "some.name == \"d2 only\" and some.sessions_used == 1 and some.captures_used == 2"
    approx(some.played()["Bb3"][0], 10.0)
    approx(some.played()["D5"][0], -4.0)
    assert I.run_from_preset(preset, session_indices=[7]).captures_used == 0, "I.run_from_preset(preset, session_indices=[7]).captures_used == 0"
check("a run from a preset takes all its sessions or the chosen ones; the name is the preset's or the caller's",
      a_preset_and_its_chosen_sessions)


def the_charts_points_and_overlay():
    r1 = I.run_from_captures([{"captures": [cap("G#2", detuned(44, 6)), cap("G#2", detuned(44, 10)),
                                            cap("A2", detuned(45, 2)), cap("C4", detuned(60, -4))]}],
                             "Tenor", name="horn A")
    r2 = I.run_from_captures([{"captures": [cap("G#2", detuned(44, -2)), cap("D4", detuned(62, 12))]}],
                             "Tenor", name="horn B")
    spreads = I.note_spreads(r1)
    assert [s.note for s in spreads] == ["Bb3", "B3", "D5"], "[s.note for s in spreads] == [\"Bb3\", \"B3\", \"D5\"]"      # low to high
    assert spreads[0] == I.NoteSpread("Bb3", 1, 8.0, 6.0, 10.0, 2) or (
        spreads[0].note == "Bb3" and approx(spreads[0].cents, 8.0) is None and approx(spreads[0].low, 6.0) is None
        and approx(spreads[0].high, 10.0) is None and spreads[0].takes == 2 and spreads[0].register == 1), "spreads[0] == I.NoteSpread(\"Bb3\", 1, 8.0, 6.0, 10.0, 2) or ( spreads[0].note == \"Bb3\" and approx(spreads[0]..."
    ov = I.overlay([r1, r2])
    assert ov["notes"] == ["Bb3", "B3", "D5", "E5"], ov["notes"]          # the union, low to high
    assert ov["registers"] == {"Bb3": 1, "B3": 1, "D5": 2, "E5": 2}, "ov[\"registers\"] == {\"Bb3\": 1, \"B3\": 1, \"D5\": 2, \"E5\": 2}"
    assert [ln["name"] for ln in ov["lines"]] == ["horn A", "horn B"], "[ln[\"name\"] for ln in ov[\"lines\"]] == [\"horn A\", \"horn B\"]"
    assert set(ov["lines"][1]["points"]) == {"Bb3", "E5"}, "set(ov[\"lines\"][1][\"points\"]) == {\"Bb3\", \"E5\"}"               # a line skips the notes it lacks
    assert I.register_changes(ov["notes"], ov["registers"]) == [2], "I.register_changes(ov[\"notes\"], ov[\"registers\"]) == [2]"       # the dashed line sits before D5
    assert I.register_changes(["Bb3", "B3"], {"Bb3": 1, "B3": 1}) == [], "I.register_changes([\"Bb3\", \"B3\"], {\"Bb3\": 1, \"B3\": 1}) == []"
check("the chart's points are each note's median with its takes' range, low to high; the overlay is the union "
      "of the runs' notes with one register change", the_charts_points_and_overlay)


def the_offset_and_the_shape_per_register():
    # played cents 10, 0 (register 1), 20, 30 (register 2): offset 15; shapes -5, -15, 5, 15
    caps = [cap("G#2", detuned(44, 10)), cap("A#2", detuned(46, 0)),
            cap("C4", detuned(60, 20)), cap("D4", detuned(62, 30))]
    run = I.run_from_captures([{"captures": caps}], "Tenor")
    res = I.offset_and_shape(run)
    approx(res["offset"], 15.0)
    by = {r["note"]: r for r in res["rows"]}
    approx(by["Bb3"]["shape"], -5.0)
    approx(by["C4"]["shape"], -15.0)
    r1, r2 = res["registers"][1], res["registers"][2]
    assert r1["n"] == 2 and approx(r1["mean_shape"], -10.0) is None, "r1[\"n\"] == 2 and approx(r1[\"mean_shape\"], -10.0) is None"
    approx(r1["rms_shape"], math.sqrt((25 + 225) / 2))
    assert r1["worst_note"] == "C4" and approx(r1["worst_shape"], -15.0) is None, "r1[\"worst_note\"] == \"C4\" and approx(r1[\"worst_shape\"], -15.0) is None"
    assert r2["n"] == 2 and approx(r2["mean_shape"], 10.0) is None and r2["worst_note"] == "E5", "r2[\"n\"] == 2 and approx(r2[\"mean_shape\"], 10.0) is None and r2[\"worst_note\"] == \"E5\""
    approx(res["overall"]["rms_shape"], math.sqrt((25 + 225 + 25 + 225) / 4))
    # 10 and -10 give offset 0, rms 10
    run = I.run_from_captures([{"captures": [cap("G#2", detuned(44, 10)), cap("A#2", detuned(46, -10))]}], "Tenor")
    res = I.offset_and_shape(run)
    approx(res["offset"], 0.0)
    approx(res["registers"][1]["rms_shape"], 10.0)
    approx(I.low_register_offset(run), 0.0)
    approx(I.low_register_offset(I.run_from_captures([{"captures": caps}], "Tenor")), 5.0)   # the median of 10 and 0
    assert I.low_register_offset(I.run_from_captures([{"captures": [cap("C4", hz(60))]}], "Tenor")) is None, "I.low_register_offset(I.run_from_captures([{\"captures\": [cap(\"C4\", hz(60))]}], \"Tenor\")) is None"
check("the offset is the mean of the notes' cents, the shape each note less it, with per-register mean, rms and "
      "worst note; the low-register offset is that register's median", the_offset_and_the_shape_per_register)


def nothing_in_here_touches_tkinter():
    import ast
    tree = ast.parse(open(I.__file__, encoding="utf-8").read())
    imported = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported |= {a.name.split(".")[0] for a in node.names}
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported.add(node.module.split(".")[0])
    assert imported == {"math", "statistics", "dataclasses", "toner_engine"}, imported
check("the module is pure: it imports the maths and the engine's note helpers, no tkinter, no tab",
      nothing_in_here_touches_tkinter)



def a_take_that_is_not_its_note_is_left_out():
    """The causes found in Matt's library (2026-11-10), fabricated: labels two transpositions
    below the sounding pitch (six presets from 2026-03-23/25: tenor +28 semitones, alto +18,
    soprano +4), an octave on top of that, and a semitone slip of the detector; and the line
    itself, half a semitone either side, at the run's reference pitch."""
    caps = [
        cap("A#0", hz(22 + 28)),                 # tenor, the early storage mistake: the pitch is D3, the label two transpositions low
        cap("C2", hz(36 + 28 + 12)),             # the same, an octave higher still
        cap("C4", hz(60 + 12)),                  # an octave: the detector's sub-octave / partial slip
        cap("C4", detuned(60, 100)),             # a semitone: the label is the note next door
        cap("C4", detuned(60, 49)),              # kept: within half a semitone
        cap("C4", detuned(60, -49)),             # kept
        cap("C4", detuned(60, 51)),              # not that note any more
        cap("C4", detuned(60, -51)),             # not that note
    ]
    run = I.run_from_captures([{"captures": caps}], "Tenor")
    # the first is the March label and is corrected (below); the other five are not their note
    assert run.left_out == {"implausible": 0, "unlabelled": 0, "off_note": 5}, run.left_out
    assert run.corrected == 1 and run.captures_used == 3, "run.corrected == 1 and run.captures_used == 3"
    assert run.notes == ["E4", "D5"] and len(run.reading("D5").takes) == 2, run.notes     # D3 is written E4 on a tenor
    assert sorted(round(t.cents) for t in run.reading("D5").takes) == [-49, 49], "sorted(round(t.cents) for t in run.reading(\"D5\").takes) == [-49, 49]"
    # a detector slip on the chart's own neighbour is not corrected either way
    assert I.run_from_captures([{"captures": [cap("C4", detuned(60, 100))]}], "Tenor").corrected == 0, "I.run_from_captures([{\"captures\": [cap(\"C4\", detuned(60, 100))]}], \"Tenor\").corrected == 0"
    # the line is drawn at the run's reference pitch: 55 cents sharp of A=440 is 47 cents sharp of A=442
    at_440 = I.run_from_captures([{"captures": [cap("C4", detuned(60, 55))]}], "Tenor", 440.0)
    at_442 = I.run_from_captures([{"captures": [cap("C4", detuned(60, 55))]}], "Tenor", 442.0)
    assert at_440.left_out["off_note"] == 1 and at_442.left_out["off_note"] == 0 and at_442.notes == ["D5"], "at_440.left_out[\"off_note\"] == 1 and at_442.left_out[\"off_note\"] == 0 and at_442.notes == [\"D5\"]"
    # a wrong-note take whose wrong label is also out of the horn's range counts as off note, not as out of range
    # (the A#0 one is the March label and is corrected now, not counted off note; C1 in tune is out of range)
    both = I.run_from_captures([{"captures": [cap("A#0", hz(22 + 28)), cap("C1", 32.7)]}], "Tenor")
    assert both.left_out == {"implausible": 1, "unlabelled": 0, "off_note": 0} and both.corrected == 1, both.left_out
    assert I.OFF_NOTE_CENTS == 50.0, "I.OFF_NOTE_CENTS == 50.0"
check("a take further than half a semitone from its labelled note is left out and counted (the library's "
      "two-transpositions labels, an octave, a semitone slip; 49 c kept, 51 c out; at the reference pitch; "
      "off note before out of range)", a_take_that_is_not_its_note_is_left_out)


def the_march_labels_are_corrected_on_read():
    """Matt, 2026-11-10: "Correct for sure." A label exactly two transpositions below the
    sounding pitch, within the bound, with the corrected note inside the horn's range,
    is charted under the note it is and counted; anything else is left as it was."""
    # tenor (+28): A#0 sounding D3 (written E4), 7 cents sharp of it; the cents are against the corrected note
    run = I.run_from_captures([{"captures": [cap("A#0", detuned(22 + 28, 7))]}], "Tenor")
    assert run.corrected == 1 and run.left_out["off_note"] == 0 and run.notes == ["E4"], (run.corrected, run.notes)
    approx(run.reading("E4").cents, 7.0, 1e-6)
    assert run.reading("E4").concert_midi == 50, "run.reading(\"E4\").concert_midi == 50"
    # alto (+18) and soprano (+4), the same mistake at their own transpositions
    assert I.run_from_captures([{"captures": [cap("A2", hz(45 + 18))]}], "Alto").corrected == 1, "I.run_from_captures([{\"captures\": [cap(\"A2\", hz(45 + 18))]}], \"Alto\").corrected == 1"
    assert I.run_from_captures([{"captures": [cap("A#4", hz(70 + 4))]}], "Soprano").corrected == 1, "I.run_from_captures([{\"captures\": [cap(\"A#4\", hz(70 + 4))]}], \"Soprano\").corrected == 1"
    # the guard: the corrected note must lie in the horn's range (a tenor's A#4 + 28 = F#7 does not)
    out = I.run_from_captures([{"captures": [cap("A#4", hz(70 + 28))]}], "Tenor")
    assert out.corrected == 0 and out.left_out["off_note"] == 1, "out.corrected == 0 and out.left_out[\"off_note\"] == 1"
    # the guard: only exactly 2T, within the bound (2T + 51 cents is not it; 2T + 49 is)
    assert I.run_from_captures([{"captures": [cap("A#0", detuned(50, 51))]}], "Tenor").left_out["off_note"] == 1, "I.run_from_captures([{\"captures\": [cap(\"A#0\", detuned(50, 51))]}], \"Tenor\").left_out[\"off_note\"] == 1"
    assert I.run_from_captures([{"captures": [cap("A#0", detuned(50, 49))]}], "Tenor").corrected == 1, "I.run_from_captures([{\"captures\": [cap(\"A#0\", detuned(50, 49))]}], \"Tenor\").corrected == 1"
    # an octave above that again is the detector's slip, not the label's: left out
    assert I.run_from_captures([{"captures": [cap("A#0", hz(22 + 28 + 12))]}], "Tenor").left_out["off_note"] == 1, "I.run_from_captures([{\"captures\": [cap(\"A#0\", hz(22 + 28 + 12))]}], \"Tenor\").left_out[\"off_note\"] == 1"
    # the corrected take still goes through the fingerprint's filter (an impossible harmonic stays out)
    bad = I.run_from_captures([{"captures": [cap("A#0", hz(50), harm=(0.0, 25.0))]}], "Tenor")
    assert bad.corrected == 1 and bad.left_out["implausible"] == 1 and bad.notes == [], "bad.corrected == 1 and bad.left_out[\"implausible\"] == 1 and bad.notes == []"
    # a take within the bound of its own label is never touched, on any horn
    for sax in ("Tenor", "Alto", "Soprano", "Baritone", "C Melody", "Sopranino"):
        ok = I.run_from_captures([{"captures": [cap("D4", detuned(62, 20))]}], sax)     # D4 is in every range
        assert ok.corrected == 0 and ok.captures_used == 1, sax
    # the chart's overlay sees the corrected note like any other
    ov = I.overlay([run])
    assert ov["notes"] == ["E4"] and ov["registers"] == {"E4": 1}, "ov[\"notes\"] == [\"E4\"] and ov[\"registers\"] == {\"E4\": 1}"
check("the March 2026 labels (exactly two transpositions low, corrected note in the horn's range) are corrected "
      "on read and counted; the guards: in range, within the bound, not an octave above; the fingerprint's filter "
      "still applies; a good take is never touched", the_march_labels_are_corrected_on_read)

print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
