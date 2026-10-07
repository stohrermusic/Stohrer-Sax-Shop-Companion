# Tone Analyzer (Toner)

Real-time harmonic analyzer for saxophone. Architecture mirrors the tuner:
- `toner_engine.py`: Pure audio/math — FFT with 16384-sample window (~2.7 Hz resolution), fundamental detection via peak-picking with harmonic series verification (a sub-harmonic candidate must show its own peak plus harmonics at multiples coprime to the divisor — see "Pitch reading" below), temporal hysteresis, harmonic extraction up to 20th harmonic (noise floor cutoff at -60 dB) with parabolic amplitude correction and Hann closed-form frequency, spectral centroid computation, descriptor computation (complexity, warmth). `TonerEngine` manages its own sounddevice input stream independently from the tuner. Preset storage uses `load_tone_presets()`/`save_tone_presets()` with `TONER_DATA_FILE` (`toner_data.json`, auto-migrated from old `tone_profiles.json`).
- `toner_tab.py`: All tkinter UI — `TonerTabMixin` builds the tab with spectrum canvas (left), intonation gauge + note display (right), and control strip (bottom). Preset management, comparison tool with multi-select and filtering, import/export.
- The toner auto-starts/stops when switching tabs, same as the tuner.
- Same three mechanisms as the tuner engine and tab (see CLAUDE-engines.md "Strobe Tuner"): numpy and sounddevice imported separately; `TonerEngine.synthetic_hz` feeds a Bb3-with-harmonics tone for tests and the CI tours (`test_toner_live`); `_toner_error_state` is redrawn after `_toner_build_spectrum_bars` wipes the canvas.
- Live descriptor gauges (Pure↔Complex, Thin↔Warm) were removed 2026-04-06. Mic position alone shifted complexity 10-20% across same-horn takes, mouthpiece dominates the signal (~25% systematic offset between mpcs), and the 51-preset corpus is too thin and too varied for absolute single-preset readouts to be meaningful. Brightness/darkness/resonance/fullness were removed earlier (2026-03-27) for similar reasons. The descriptor functions still live in `toner_engine.py` and feed the Analyze tool, where deltas cancel the mic/setup confounders. The right side of the live tab now shows just the intonation gauge and note display.
- Two comparison descriptors (shown in Analyze tool, not live gauges): Even/Odd Ratio (even vs odd harmonic balance), Rolloff Shape (nonlinearity of harmonic rolloff — spectral peaks/bumps). Even/Odd defaults to on (proven most robust across recording conditions); Rolloff Shape defaults to off. Configured in Options > Settings > Analysis tab.
- Population percentiles: The Analyze tool computes `compute_population_stats()` for all presets of the same sax type and shows percentile rankings (P0-P100) alongside descriptor values. Interpretive labels: low/below avg/mid-range/above avg/high. Out-of-range notes are flagged using `SAX_NOTE_RANGES` in toner_engine.py.
- Comparison analysis includes harmonic-range interpretation: H1-H4 shifts = bore, H7-H13 shifts = neck/mpc, broadband = mpc/player. Same-player comparisons noted as more reliable.
- Scale toggle switches between linear (default, true amplitude ratios) and dB (logarithmic).
- Spectrum overlay: a preset can be loaded as a ghost overlay on the live spectrum (via Analyze > single preset or group average > Overlay on Spectrum). Blue ghost bars update per-note as you play. All live descriptor gauges (delta and absolute) have been removed — the analyze tool (both sides averaged) is where comparison actually works.

## Pitch reading (2026-10-02)

**Frequencies come from `audio_utils.hann_peak_freq`, levels don't.** `_detect_fundamental` and `_extract_harmonics` read each peak's frequency with the Hann closed form (mechanism and the rejected zero-padded alternative under "Strobe Tuner" in CLAUDE-engines.md). The old linear parabola put `fundamental_freq` up to ~4 c off (median 0.44, p95 1.9) and `harmonic_cents` up to ~5 c. `peak_mag` and its parabolic amplitude correction are **deliberately unchanged**, so `harmonics_db` — and every descriptor, none of which reads pitch — stays comparable with every capture already stored. Measured on 20-partial tones (H20 up to 9.9 kHz): every partial's level identical to the old code, every descriptor identical. A more accurate f0 does move the ±3-bin search window at high harmonic numbers, so a *noise-level* reading where a tone has no partial can change; that's noise either way. Captures stored before this carry the old pitch error in `fundamental_freq`/`harmonic_cents`; nothing in the Analyze tool reads them for a descriptor.

**Sub-harmonic verification counts only multiples coprime to the divisor.** The check "is the strongest peak harmonic *d* of a lower fundamental?" used to count every multiple in `[2, 3, 4]` / `[2..6]`, including `mult == divisor` (the strongest peak itself) and multiples of it (`/2` ×4 is the strongest's own H2; `/4` ×2 and ×6 are the `/2` candidate's H1 and H3). Those are present whether or not the candidate is real, so any noise bump at half a note passed as its fundamental — 138 octave-low frames of 14,070 in a sweep of every soprano/alto/tenor/bari note. Now only `gcd(mult, divisor) == 1` counts, and the requirement drops by one because the candidate's own peak is already checked for noise and prominence; for `/3` and `/5` that is exactly the old effective rule. Same sweep: 10 octave-low, 0 high. **Don't "simplify" it to skipping only `mult == divisor`** — that was proposed in an outside review of the ported code, and measured: with self excluded, `/2` needs both H3 and H4, so a tone with H2 a little over H1 and no H4 reads an octave HIGH — 2,199 of the same 14,070 frames, and Test 13's warm tone (H1, H2, H3 only) read 880 for 440. Also: the peak search round a sub-harmonic can reach below `MIN_FUNDAMENTAL_HZ`; such a peak is skipped. Gated by Tests 20–21 in `test_toner_engine` (a 233.08 Hz tone within 0.05 c and its harmonics within 1 dB; a hand-built spectrum per octave mechanism; H2-over-H1 tones read the right octave; a 2,144-frame sweep with no octave error either way and every frame within 0.5 c).

**Cost**: none worth measuring — the estimator reads the spectrum already computed. A 4-minute WAV through `analyze_audio_file`: best of three 3.85 s before, 4.26 s after on a noisy machine.

## Recording Quality

`compute_rolloff_rate()` measures harmonic rolloff (dB/harmonic). Live warning if rolloff exceeds the threshold returned by `get_rolloff_threshold(mic_type, sax_type)` (suppressible per-preset via "Don't show again" checkbox). Stored per session and in fingerprints. Comparison tool warns on mismatch. The threshold function takes both mic class and sax type because both matter: ribbon mics naturally read higher than dynamics, which read higher than condensers (HF rolloff in the mic itself); and higher-pitched horns naturally read higher than lower-pitched horns (their H12 reaches further into the mic's HF rolloff region). Current values: ribbon 3.5, dynamic 2.8 (alto+dynamic bumped to 3.5 after Foster's Conn 6M data showed alto reads ~1 dB/H higher than bari with the same RE20), condenser uses the base `ROLLOFF_WARN_THRESHOLD` constant. Both call sites in `toner_tab.py` pass sax_type — the live capture handler from `engine.sax_type`, the Analyze tool from `fp['_preset']['horn_type']`. Real-world rolloff readings observed: Matt AT2020 condenser 1.1-1.3 dB/H, Tyler Neumann TLM 1.8-3.1, Edinger U67 1.65-2.04, Mario AT2035 alto 1.99 / soprano 4.00, Foster RE20 dynamic bari 1.5-2.3 / alto 2.99-3.49, Grant MXL RI44 ribbon tenor 2.16-2.46.

## Mic Type/Model

`mic_type` (condenser/ribbon/dynamic/other, required) and `mic_model` (required) are preset fields, same as mouthpiece and reed. Sessions snapshot mic from the preset at creation time. The control strip shows the loaded preset's mic info. Comparison tool filters by mic type and warns on mismatches. Migration: presets without mic fields get backfilled from their sessions' most common mic_type on first load.

## Analyze Tool

Renamed from Compare. Handles single preset view, two-preset delta analysis with difference chart, and multi-preset spread analysis. Accessible via File > Analyze... or File > Presets > Analyze. Analysis descriptors (complexity, warmth, even/odd, rolloff shape, evenness) are user-configurable in Options > Settings > Analysis tab. Back button on analysis/group report windows returns to the selection picker. Clickable legend labels and chart lines show preset detail (player, mpc, mic, etc.). **Character Map**: 2D quad-chart scatter plot of every selected preset on warmth x complexity axes (Y axis was originally "H4-H5 mean" but switched to complexity 2026-04-06 because warmth and H4-H5 share a denominator and were not independent within tenors); click a dot, legend entry, or chart line and the matching items highlight together in the same color across legend, harmonic chart, and quad map. **Bars chart toggle**: switch the harmonic comparison chart between line view and side-by-side bars view. **Maximize / fullscreen mode**: the Analyze window is non-modal (intentionally NOT a transient toplevel, since transients lose the maximize button on Windows), resizes freely, and has a Maximize button that toggles between zoomed and normal window state (with a macOS Aqua fallback that uses a screen-filling geometry since Aqua doesn't support `state('zoomed')`). The whole window scrolls vertically since comparison data can exceed any viewport. Same pattern is applied to the Group Report dialog. **Cross-sax-type warning**: mixing alto + tenor + bari + soprano in one comparison is allowed but pops a warning, since lower-pitched horns intrinsically read warmer/brighter regardless of mouthpiece.

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

**Read-side capture filter** (`_capture_is_plausible()`): Before averaging, `compute_fingerprint` drops captures that fail two checks. (1) **Range filter**: the labeled note must fall inside `SAX_NOTE_RANGES[sax_type]`, which uses tight bounds — lower bound = written low Bb for most types, extended to low A only on bari (common spec) and alto (rare); upper bound = high F# written (no altissimo margin, because altissimo is a harmonic partial of a lower fingering, not a true closed-tube fundamental, so its harmonic structure isn't useful for tone analysis). (2) **Plausibility filter**: no value in `harmonics_db` may be NaN, ±inf (checked with `math.isfinite` — JSON round-trips `-Infinity` happily and one poisons every average), or > +20 dB above the labeled fundamental — that's a clear sub-octave detection error where the real fundamental landed in an upper-harmonic bin. Real low-register alto H2 tops out near +10 dB legitimately, so +20 dB is the hard sentinel. The filter is read-only — raw on-disk captures are NOT modified, so any descriptor formula or filter improvement applies retroactively to historical data on the next load. Was added 2026-04-07 after a library audit found 37 of 58 presets had at least one out-of-range capture; the worst case (Selmer VI alto GP) had 105 of 164 captures bogus.

Flat legacy format (presets at top level without library wrapper) is auto-migrated on load.

## Toner Calibration

Preset notes can contain subjective tone descriptions ("rich horn", "very bright", "dark and warm"). The `tools/calibrate_toner.py` script scans all annotated presets, extracts keywords, compares them against computed descriptors, and reports alignment. The `tools/analyze_horn_spread.py` script computes statistical spread of descriptors across all presets — min/max/mean/stddev per descriptor, per-note variation, gauge scaling suggestions, and grouping analysis by horn type and manufacturer. Both tools help identify where the descriptor scaling constants need adjustment. Bias sliders in the UI provide per-user visual calibration without affecting captured data.

**Descriptor calibration status**: Two descriptors survived data-driven validation: complexity (spectral flatness) and warmth (H2 strength). Brightness/darkness were removed because the labels conflicted with player vocabulary (data showed Yamahas read "dark" but players call them "bright"). Resonance was removed because it read 100% on all 39 presets — no differentiation. Fullness was removed because it depended on brightness/darkness. Key constants live in toner_engine.py: `RICHNESS_RAW_MIN`, `RICHNESS_RAW_RANGE`, `WARMTH_DB_FLOOR`, `WARMTH_DB_RANGE`. Development history in `.claude/projects/*/memory/descriptors.md`. `tools/profile_report.py` generates a detailed single-preset report.
