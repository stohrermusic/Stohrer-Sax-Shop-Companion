"""Intonation from the tone analyzer's captures: each note's cents as played.

Pure maths, no tkinter. Written 2026-11-10 for the Analyzer tab's intonation chart.

A capture (toner_engine; one held note) carries its concert note name, its averaged
fundamental and its harmonics. Here it becomes one `Take` of a `NoteReading`: its
cents against equal temperament at the reference pitch (the captures do not record
the pitch they were played at; the caller's is used and `IntonationRun.reference_hz`
says which), under the note's WRITTEN name on this saxophone. A note's cents are the
MEDIAN of its takes, so one misfiled or bad take cannot drag it. A capture the
analyzer's own fingerprint leaves out (`toner_engine._capture_is_plausible`: outside
the saxophone's range, an impossible harmonic) is left out here too, and counted.

Registers follow the written note, as the horn is played: up to C#5 the low register
(1), from D5 the octave key (2); D6 and up are the palm keys, a band of the upper
register the summary can name. The chart shows the register change between 1 and 2.

What the chart shows is cents AS PLAYED (Matt, 2026-11-10: "just as played is what
everyone will want to see"). The offset (the mean over the notes, "the player and the
cork") and the shape (each note less it, "the horn") and their per-register figures
are computed here because they cost nothing, for a later ask; nothing in the UI or the
export shows them.
"""
import math
import statistics
from dataclasses import dataclass, field

from toner_engine import (  # noqa: F401  (A4_HZ, OFF_NOTE_CENTS, cents_from_equal_temperament, corrected_label
    A4_HZ, OFF_NOTE_CENTS, SAX_TRANSPOSITIONS, _capture_is_plausible,  # and midi_to_hz lived
    cents_from_equal_temperament, corrected_capture, corrected_label, midi_to_hz, note_name_to_midi,  # here until
    player_note_name, transpose_note,  # 2026-11-11 and are still read from this module; the engine owns them now)
)

LOW_REGISTER_TOP_WRITTEN_MIDI = 73      # written C#5: the last note of the low register
PALM_KEYS_FROM_WRITTEN_MIDI = 86        # written D6: the palm keys begin
FRAME_SECONDS = 0.033                   # one analysis frame at 30 fps, for a take's length


class IntonationError(ValueError):
    """A run that cannot be made; the message says why."""


def written_midi(concert_midi, sax_type):
    """The written MIDI number of a concert pitch on this saxophone."""
    if sax_type not in SAX_TRANSPOSITIONS:
        raise IntonationError(f"unknown saxophone type {sax_type!r}")
    return concert_midi + SAX_TRANSPOSITIONS[sax_type]


def register_of_written(midi):
    """1 for the low register (written C#5 and below), 2 from written D5 up."""
    return 1 if midi <= LOW_REGISTER_TOP_WRITTEN_MIDI else 2


def written_band(midi):
    """The summary's three bands: 'low', 'octave_key', 'palm_keys'."""
    if midi <= LOW_REGISTER_TOP_WRITTEN_MIDI:
        return "low"
    return "palm_keys" if midi >= PALM_KEYS_FROM_WRITTEN_MIDI else "octave_key"


@dataclass
class Take:
    """One capture as a take of a note."""
    cents: float                        # against equal temperament at the run's reference pitch
    f0_hz: float                        # the capture's averaged fundamental
    frames: int                         # the capture's n_frames
    seconds: float                      # frames * FRAME_SECONDS
    method: str = ""                    # free / calibration / file
    level_dbfs: float | None = None     # the capture's level, where it recorded one
    session_date: str = ""              # the session the capture belongs to
    wobble_cents: float | None = None   # the spread over the held frames; captures do not record it


@dataclass
class NoteReading:
    """One written note and its takes."""
    note: str                           # written name as a player writes it (Bb, Eb; C#, F#, G#)
    written_midi: int
    concert_midi: int
    register: int                       # 1 low, 2 octave key (palm keys are register 2)
    takes: list = field(default_factory=list)

    @property
    def cents(self):
        """The note's cents: the median of its takes (None without a take)."""
        if not self.takes:
            return None
        return float(statistics.median(t.cents for t in self.takes))

    @property
    def take_spread_cents(self):
        """How far the takes' own cents differ (max - min); None under two takes."""
        if len(self.takes) < 2:
            return None
        c = [t.cents for t in self.takes]
        return float(max(c) - min(c))

    @property
    def low(self):
        return min(t.cents for t in self.takes) if self.takes else None

    @property
    def high(self):
        return max(t.cents for t in self.takes) if self.takes else None

    @property
    def wobble_cents(self):
        """The median of the takes' wobble, where any recorded one (none do today)."""
        vals = [t.wobble_cents for t in self.takes if t.wobble_cents is not None]
        return float(statistics.median(vals)) if vals else None

    @property
    def band(self):
        return written_band(self.written_midi)


@dataclass
class IntonationRun:
    """A horn as played: every note's reading, low to high, from one preset's
    captures (all of its sessions, or the ones chosen)."""
    name: str
    sax_type: str
    reference_hz: float
    readings: list = field(default_factory=list)          # NoteReading, sorted low to high
    left_out: dict = field(default_factory=dict)          # {"implausible": n, "unlabelled": n}
    sessions_used: int = 0
    captures_used: int = 0
    corrected: int = 0                  # takes charted under a corrected label (corrected_label)

    def reading(self, note):
        for r in self.readings:
            if r.note == note:
                return r
        return None

    def played(self):
        """{written note: (cents, register)} of every note with a reading."""
        return {r.note: (r.cents, r.register) for r in self.readings if r.cents is not None}

    @property
    def notes(self):
        return [r.note for r in self.readings]


def run_from_captures(sessions, sax_type, reference_hz=A4_HZ, name=""):
    """The sessions' captures (each a dict with "captures") as an IntonationRun.

    One Take a capture. A capture the fingerprint would leave out is left out and
    counted under "implausible"; one with no note name or no fundamental under
    "unlabelled"; one whose pitch is further than OFF_NOTE_CENTS from its labelled
    note (at the reference pitch) under "off_note": it is not that note, and the
    chart must never draw a point that is not its note. The one exception is a label
    that is exactly two transpositions low (`toner_engine.corrected_capture`, the
    rule the fingerprint applies too, written once): the take is charted under the
    note it is and counted in `corrected`. Read-side only: nothing on disk changes.
    Takes of the same written note are one reading."""
    if sax_type not in SAX_TRANSPOSITIONS:
        raise IntonationError(f"unknown saxophone type {sax_type!r}")
    readings = {}
    implausible = unlabelled = off_note = corrected = used = 0
    for s in sessions:
        caps = s.get("captures", []) or []
        for cap in caps:
            cap, fixed = corrected_capture(cap, sax_type, reference_hz)
            concert = note_name_to_midi(cap.get("note", ""))
            try:
                f0 = float(cap.get("fundamental_freq") or 0.0)
            except (TypeError, ValueError):
                f0 = 0.0
            if concert is None or f0 <= 0:
                unlabelled += 1
                continue
            if fixed:
                corrected += 1
            # the pitch against the label: a take that is not its note is counted as that,
            # even when its (wrong) label is also outside the horn's range
            cents = cents_from_equal_temperament(f0, concert, reference_hz)
            if abs(cents) > OFF_NOTE_CENTS:
                off_note += 1
                continue
            if not _capture_is_plausible(cap, sax_type):
                implausible += 1
                continue
            wm = written_midi(concert, sax_type)
            try:
                frames = int(cap.get("n_frames") or 0)
            except (TypeError, ValueError):
                frames = 0
            take = Take(cents=cents, f0_hz=f0,
                        frames=frames, seconds=frames * FRAME_SECONDS, method=str(cap.get("method", "")),
                        level_dbfs=cap.get("level_dbfs"), session_date=str(s.get("date", "")))
            if wm not in readings:
                readings[wm] = NoteReading(note=player_note_name(transpose_note(cap["note"], sax_type)),
                                           written_midi=wm,
                                           concert_midi=concert, register=register_of_written(wm))
            readings[wm].takes.append(take)
            used += 1
    return IntonationRun(name=name, sax_type=sax_type, reference_hz=float(reference_hz),
                         readings=[readings[k] for k in sorted(readings)],
                         left_out={"implausible": implausible, "unlabelled": unlabelled, "off_note": off_note},
                         sessions_used=len(sessions), captures_used=used, corrected=corrected)


def run_from_preset(preset, reference_hz=A4_HZ, session_indices=None, name=None):
    """A tone preset (toner_engine's dict: horn_type, sessions, ...) as a run, from all
    its sessions or the ones at `session_indices`."""
    sessions = preset.get("sessions", []) or []
    if session_indices is not None:
        sessions = [sessions[i] for i in session_indices if 0 <= i < len(sessions)]
    label = name if name is not None else str(preset.get("name", ""))
    return run_from_captures(sessions, preset.get("horn_type", ""), reference_hz, name=label)


# --------------------------------------------------------------------------
# the chart's data


@dataclass(frozen=True)
class NoteSpread:
    """One note of a run on the chart: the median of its takes, their lowest and
    highest, how many."""
    note: str
    register: int
    cents: float
    low: float
    high: float
    takes: int


def note_spreads(run):
    """Each played note, low to high: the median of its takes (one misfiled or bad
    take can't drag it), their range and their count. The chart's points."""
    out = []
    for r in run.readings:
        if r.takes:
            out.append(NoteSpread(r.note, r.register, r.cents, r.low, r.high, len(r.takes)))
    return out


def overlay(runs):
    """What the chart draws for several runs: the axis notes (the union, low to high),
    each note's register, and per run its spreads by note."""
    by_midi = {}
    for run in runs:
        for r in run.readings:
            if r.takes and r.written_midi not in by_midi:
                by_midi[r.written_midi] = (r.note, r.register)
    notes = [by_midi[m][0] for m in sorted(by_midi)]
    registers = {by_midi[m][0]: by_midi[m][1] for m in by_midi}
    lines = [{"name": run.name, "points": {s.note: s for s in note_spreads(run)}} for run in runs]
    return {"notes": notes, "registers": registers, "lines": lines}


def register_changes(notes, registers):
    """The indexes i where the register changes between notes[i-1] and notes[i]
    (the chart's dashed lines sit between them)."""
    return [i for i in range(1, len(notes)) if registers.get(notes[i]) != registers.get(notes[i - 1])]


# --------------------------------------------------------------------------
# the figures nothing shows yet: the offset and the shape


def low_register_offset(run):
    """The median of the low register's notes: where the player's tuning and the
    cork put the whole horn. None without a low-register note."""
    lows = [r.cents for r in run.readings if r.register == 1 and r.takes]
    return float(statistics.median(lows)) if lows else None


def _mean_rms(vals):
    mean = sum(vals) / len(vals)
    rms = math.sqrt(sum(v * v for v in vals) / len(vals))
    return mean, rms


def offset_and_shape(run):
    """{"offset", "rows", "registers", "overall"}: the offset is the mean of the notes'
    cents (the player and the cork), the shape each note less it (the horn); per
    register its count, mean and rms shape, worst note; None/empty with no note."""
    rows = [{"note": r.note, "register": r.register, "cents": r.cents} for r in run.readings if r.takes]
    if not rows:
        return {"offset": None, "rows": rows, "registers": {}, "overall": None}
    offset = sum(r["cents"] for r in rows) / len(rows)
    for r in rows:
        r["shape"] = r["cents"] - offset
    regs = {}
    for reg in sorted({r["register"] for r in rows}):
        rr = [r for r in rows if r["register"] == reg]
        mean, rms = _mean_rms([r["shape"] for r in rr])
        worst = max(rr, key=lambda r: abs(r["shape"]))
        regs[reg] = {"n": len(rr), "mean_shape": mean, "rms_shape": rms,
                     "worst_note": worst["note"], "worst_shape": worst["shape"]}
    _m, rms_all = _mean_rms([r["shape"] for r in rows])
    return {"offset": offset, "rows": rows, "registers": regs, "overall": {"n": len(rows), "rms_shape": rms_all}}


__all__ = [
    "A4_HZ", "FRAME_SECONDS", "LOW_REGISTER_TOP_WRITTEN_MIDI", "PALM_KEYS_FROM_WRITTEN_MIDI",
    "IntonationError", "IntonationRun", "NoteReading", "NoteSpread", "OFF_NOTE_CENTS", "Take",
    "cents_from_equal_temperament", "corrected_capture", "corrected_label", "low_register_offset", "midi_to_hz",
    "note_spreads", "offset_and_shape",
    "overlay", "register_changes", "register_of_written", "run_from_captures", "run_from_preset",
    "written_band", "written_midi",
]
