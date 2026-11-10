# Tone Analyzer (Toner)

Real-time harmonic analyzer for saxophone. Architecture mirrors the tuner:

- `toner_engine.py`: Pure audio/math — FFT with 16384-sample window (~2.7 Hz resolution), fundamental detection via peak-picking with harmonic series verification (a sub-harmonic candidate must show its own peak plus harmonics at multiples coprime to the divisor — see "Pitch reading" below), temporal hysteresis, harmonic extraction up to 20th harmonic (noise floor cutoff at -60 dB) with parabolic amplitude correction and Hann closed-form frequency, spectral centroid computation, descriptor computation (complexity, warmth). `TonerEngine` manages its own sounddevice input stream independently from the tuner. Preset storage uses `load_tone_presets()`/`save_tone_presets()` with `TONER_DATA_FILE` (`toner_data.json`, auto-migrated from old `tone_profiles.json`).
- `toner_tab.py`: All tkinter UI — `TonerTabMixin` builds the tab with spectrum canvas (left), intonation gauge + note display (right), and control strip (bottom). Preset management, comparison tool with multi-select and filtering, import/export.
- The toner auto-starts/stops when switching tabs, same as the tuner.
- Same three mechanisms as the tuner engine and tab (see CLAUDE-engines.md "Strobe Tuner"): numpy and sounddevice imported separately; `TonerEngine.synthetic_hz` feeds a Bb3-with-harmonics tone for tests and the CI tours (`test_toner_live`); `_toner_error_state` is redrawn after `_toner_build_spectrum_bars` wipes the canvas.

- There are no live descriptor gauges. Descriptors feed only the Analyze tool, where deltas between presets cancel mic and setup confounders; the right side of the live tab is the intonation gauge and note display. Why they were removed: `docs/decisions/toner.md`.

- Two comparison descriptors (shown in Analyze tool, not live gauges): Even/Odd Ratio (even vs odd harmonic balance), Rolloff Shape (nonlinearity of harmonic rolloff — spectral peaks/bumps). Even/Odd defaults to on (proven most robust across recording conditions); Rolloff Shape defaults to off. Configured in Options > Settings > Analysis tab.
- Population percentiles: The Analyze tool computes `compute_population_stats()` for all presets of the same sax type and shows percentile rankings (P0-P100) alongside descriptor values. Interpretive labels: low/below avg/mid-range/above avg/high. Out-of-range notes are flagged using `SAX_NOTE_RANGES` in toner_engine.py.
- Comparison analysis includes harmonic-range interpretation: H1-H4 shifts = bore, H7-H13 shifts = neck/mpc, broadband = mpc/player. Same-player comparisons noted as more reliable.
- Scale toggle switches between linear (default, true amplitude ratios) and dB (logarithmic).
- Spectrum overlay: a preset can be loaded as a ghost overlay on the live spectrum (via Analyze > single preset or group average > Overlay on Spectrum). Blue ghost bars update per-note as you play. All live descriptor gauges (delta and absolute) have been removed — the analyze tool (both sides averaged) is where comparison actually works.

## Pitch reading

- Frequencies come from `audio_utils.hann_peak_freq`; levels (`peak_mag` with its parabolic amplitude correction) are deliberately unchanged, so `harmonics_db` and every descriptor stay comparable with stored captures.
- Sub-harmonic verification counts only multiples coprime to the divisor (`gcd(mult, divisor) == 1`), with the requirement one lower. **Don't "simplify" it to skipping only `mult == divisor`**: measured, that reads octaves high.

Gated by Tests 20 and 21 in `test_toner_engine`. Measurements: `docs/decisions/toner.md`.

## Recording Quality

`compute_rolloff_rate()` measures harmonic rolloff (dB/harmonic). The live warning uses `get_rolloff_threshold(mic_type, sax_type)`, which depends on both (ribbon 3.5, dynamic 2.8, alto with dynamic 3.5, condenser the base `ROLLOFF_WARN_THRESHOLD`); both call sites pass sax_type. Suppressible per preset; stored per session and in fingerprints; the comparison tool warns on a mismatch. Readings per mic: `docs/decisions/toner.md`.

## Mic Type/Model

`mic_type` (condenser/ribbon/dynamic/other, required) and `mic_model` (required) are preset fields, same as mouthpiece and reed. Sessions snapshot mic from the preset at creation time. The control strip shows the loaded preset's mic info. Comparison tool filters by mic type and warns on mismatches. Migration: presets without mic fields get backfilled from their sessions' most common mic_type on first load.

## Analyze Tool

File > Analyze...: one preset, a two-preset delta or a multi-preset spread, with population percentiles (`compute_population_stats()`) and a Character Map. The window and the Group Report are non-modal and deliberately not transient (a transient loses the maximize button on Windows). Detail: `docs/decisions/toner.md`.

## Mutate Preset

Duplicates a preset with all fields pre-filled (name cleared), letting users change one variable (mouthpiece, mic, reed) and save as a new preset. Streamlines A/B testing workflows.

## Settings Dialog

Options > Settings opens a tabbed dialog (General: input device, recording, pitch; Analysis: preset fields, descriptor visibility). Capture Threshold stays as a standalone menu item due to its live level meter.

## Capture Performance

During free-mode capture, JSON saves are deferred to every 10 seconds (`_toner_schedule_save`) instead of after every micro-capture. A cached `_toner_captured_notes` set avoids per-frame iteration over captures for the note counter. Both prevent progressive lag during long sessions. `_toner_flush_save()` is called on capture stop to ensure data is persisted. The app-exit path (`_toner_save_settings()`, called from `on_exit`) also flushes any in-flight free-mode micro-capture and persists the active session, so exiting mid-capture loses nothing — and it merge-updates `toner_settings` rather than replacing it, so keys owned by the Settings dialog (like `analysis_descriptors`) survive exit.

## Capture Modes

Two selectable capture modes: "free" (0.5s stability, continuous micro-captures while playing naturally) and "calibration" (guided chromatic scale Bb3-F6, 5s per note, labels from guide not detector, stores `detected_as` field for detector accuracy analysis). A separate "file" method (import WAV via File > Presets > Import Audio File) extracts stable note segments offline. Each capture is tagged with its method.

## Auto-Transposition

SAX selector sets both the break frequency and the displayed note names. Written pitch is shown by default (alto shows A4 when concert C4 is played). "Concert" checkbox overrides to concert pitch.

## Coverage Summary

Dialog appears after stopping a capture session, showing a bar chart of note distribution colored by register (low/mid/high), gap assessment, and a "Resume Capturing" button to fill underrepresented registers.

## Fingerprinting

`compute_fingerprint(sessions, sax_type)` recomputes descriptors from raw harmonics (never uses stored descriptors), averages per-note first (equal weight per note), then across notes. This prevents register skew.

## WAV Recording

Enabled by default. On first capture, the app prompts the user to choose a recording folder (no silent default). The `recording_file` field in the session dict links the WAV to its session. Filenames use preset name + timestamp at save time for uniqueness. WAV reanalysis is automatic when recording is on — the offline `analyze_audio_file()` pipeline replaces live captures at session end. The offline analyzer uses stricter segment detection (attack skip, minimum stability frames) producing roughly 2x the accuracy of live free-capture (20 harmonics extracted vs ~12 live). Takes ~5 seconds for a 4-minute recording. "Delete WAV after analysis" optionally removes the file afterward. Progress is shown in the capture bar. Disabling WAV recording shows a warning about reduced accuracy.

## Sandbox Mode

Enabled via Options > Settings > Analysis > "Allow sandbox mode". Sandbox presets (`"sandbox": true` in JSON) relax all field requirements except preset name — no mic type, sax type, or instrument fields required. Designed for non-sax instruments, contact mics, effects chains, or experimental setups. The sandbox flag is immutable once sessions exist. Visual indicators: `[sandbox]` prefix in preset list, amber "sandbox" mic label, `[sandbox]` in capture bar, yellow banner in analyze tool. Comparison between sandbox and regular presets shows a warning. A "sandbox (concert pitch)" checkbox on the control strip appears when sandbox mode is enabled and forces concert pitch display.

## Tone Preset Data Model

Presets use nested library format: `{library_name: {preset_name: preset_data}}`. A preset saves setup details (horn + player + mouthpiece + reed + mic type + mic model) and pre-fills session metadata for quick capture start. Sessions store their own copy of the metadata at creation time.

A preset contains sessions (date + captures). **Captures store only raw measurement data** — no pre-computed descriptors:

- `harmonics_db`: list of dB values relative to fundamental (up to 20 harmonics, index 0 = fundamental)
- `harmonic_cents`: list of cents deviations from ideal harmonic positions
- `fundamental_freq`: detected fundamental frequency in Hz
- `spectral_centroid`: amplitude-weighted center frequency of the harmonic series (Hz)
- `signal_level`: RMS signal level at capture time (0.0-1.0)
- `note`: concert-pitch note name
- `method`: "free", "calibration", or "file"
- `n_frames`, `timestamp`, and other metadata

Descriptors (complexity, warmth) are always computed on the fly by `descriptors_from_harmonics()` using current formulas. This means formula improvements apply retroactively to all historical data. Old captures with baked-in `descriptors` fields still load fine — stored descriptors are ignored.

`compute_fingerprint(sessions, sax_type)` in `toner_engine.py` aggregates all sessions: averages harmonics per-note, computes descriptors from the averaged harmonics, then averages descriptors across notes with equal weight. The `per_note` dict maps note names to averaged harmonic data, enabling note-by-note comparison across horns.

**Read-side capture filter** (`_capture_is_plausible()`): `compute_fingerprint` drops a capture whose note is outside `SAX_NOTE_RANGES[sax_type]` (no altissimo) or whose `harmonics_db` holds a non-finite value or one over +20 dB. It is read-only: never modify raw captures on disk. Why: `docs/decisions/toner.md`.

Flat legacy format (presets at top level without library wrapper) is auto-migrated on load.

## Toner Calibration

Preset notes can contain subjective tone descriptions ("rich horn", "very bright", "dark and warm"). The `tools/calibrate_toner.py` script scans all annotated presets, extracts keywords, compares them against computed descriptors, and reports alignment. The `tools/analyze_horn_spread.py` script computes statistical spread of descriptors across all presets — min/max/mean/stddev per descriptor, per-note variation, gauge scaling suggestions, and grouping analysis by horn type and manufacturer. Both tools help identify where the descriptor scaling constants need adjustment. Bias sliders in the UI provide per-user visual calibration without affecting captured data.

**Descriptor calibration status**: complexity (spectral flatness) and warmth (H2 strength) are the descriptors kept. Key constants live in toner_engine.py: `RICHNESS_RAW_MIN`, `RICHNESS_RAW_RANGE`, `WARMTH_DB_FLOOR`, `WARMTH_DB_RANGE`. `tools/profile_report.py` generates a detailed single-preset report. Why brightness, darkness, resonance and fullness were removed: `docs/decisions/toner.md`.
