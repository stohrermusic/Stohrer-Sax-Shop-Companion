# Tone analyzer: decisions and history

This file holds the why and the history behind the rules in the CLAUDE files, and the full detail they summarise: what was tried, measured, rejected and decided, with dates and the user's own words. The paragraphs were moved here verbatim from the CLAUDE files on 2026-11-10 (HEAD 01071f5); each section names where it came from. Add new decisions to the matching section, dated.

## Tone Analyzer (Toner)

*Moved from CLAUDE-toner.md, section "Tone Analyzer (Toner)" (2026-11-10).*

- Live descriptor gauges (Pure↔Complex, Thin↔Warm) were removed 2026-04-06. Mic position alone shifted complexity 10-20% across same-horn takes, mouthpiece dominates the signal (~25% systematic offset between mpcs), and the 51-preset corpus is too thin and too varied for absolute single-preset readouts to be meaningful. Brightness/darkness/resonance/fullness were removed earlier (2026-03-27) for similar reasons. The descriptor functions still live in `toner_engine.py` and feed the Analyze tool, where deltas cancel the mic/setup confounders. The right side of the live tab now shows just the intonation gauge and note display.

## Pitch reading (2026-10-02)

*Moved from CLAUDE-toner.md, section "Tone Analyzer (Toner) > Pitch reading (2026-10-02)" (2026-11-10).*

### Frequencies come from `audio_utils.hann_peak_freq`, levels don't

**Frequencies come from `audio_utils.hann_peak_freq`, levels don't.** `_detect_fundamental` and `_extract_harmonics` read each peak's frequency with the Hann closed form (mechanism and the rejected zero-padded alternative under "Strobe Tuner" in CLAUDE-engines.md). The old linear parabola put `fundamental_freq` up to ~4 c off (median 0.44, p95 1.9) and `harmonic_cents` up to ~5 c. `peak_mag` and its parabolic amplitude correction are **deliberately unchanged**, so `harmonics_db` — and every descriptor, none of which reads pitch — stays comparable with every capture already stored. Measured on 20-partial tones (H20 up to 9.9 kHz): every partial's level identical to the old code, every descriptor identical. A more accurate f0 does move the ±3-bin search window at high harmonic numbers, so a *noise-level* reading where a tone has no partial can change; that's noise either way. Captures stored before this carry the old pitch error in `fundamental_freq`/`harmonic_cents`; nothing in the Analyze tool reads them for a descriptor.

### Sub-harmonic verification counts only multiples coprime to the divisor

**Sub-harmonic verification counts only multiples coprime to the divisor.** The check "is the strongest peak harmonic *d* of a lower fundamental?" used to count every multiple in `[2, 3, 4]` / `[2..6]`, including `mult == divisor` (the strongest peak itself) and multiples of it (`/2` ×4 is the strongest's own H2; `/4` ×2 and ×6 are the `/2` candidate's H1 and H3). Those are present whether or not the candidate is real, so any noise bump at half a note passed as its fundamental — 138 octave-low frames of 14,070 in a sweep of every soprano/alto/tenor/bari note. Now only `gcd(mult, divisor) == 1` counts, and the requirement drops by one because the candidate's own peak is already checked for noise and prominence; for `/3` and `/5` that is exactly the old effective rule. Same sweep: 10 octave-low, 0 high. **Don't "simplify" it to skipping only `mult == divisor`** — that was proposed in an outside review of the ported code, and measured: with self excluded, `/2` needs both H3 and H4, so a tone with H2 a little over H1 and no H4 reads an octave HIGH — 2,199 of the same 14,070 frames, and Test 13's warm tone (H1, H2, H3 only) read 880 for 440. Also: the peak search round a sub-harmonic can reach below `MIN_FUNDAMENTAL_HZ`; such a peak is skipped. Gated by Tests 20–21 in `test_toner_engine` (a 233.08 Hz tone within 0.05 c and its harmonics within 1 dB; a hand-built spectrum per octave mechanism; H2-over-H1 tones read the right octave; a 2,144-frame sweep with no octave error either way and every frame within 0.5 c).

**Cost**: none worth measuring — the estimator reads the spectrum already computed. A 4-minute WAV through `analyze_audio_file`: best of three 3.85 s before, 4.26 s after on a noisy machine.

## Recording Quality

*Moved from CLAUDE-toner.md, section "Tone Analyzer (Toner) > Recording Quality" (2026-11-10).*

`compute_rolloff_rate()` measures harmonic rolloff (dB/harmonic). Live warning if rolloff exceeds the threshold returned by `get_rolloff_threshold(mic_type, sax_type)` (suppressible per-preset via "Don't show again" checkbox). Stored per session and in fingerprints. Comparison tool warns on mismatch. The threshold function takes both mic class and sax type because both matter: ribbon mics naturally read higher than dynamics, which read higher than condensers (HF rolloff in the mic itself); and higher-pitched horns naturally read higher than lower-pitched horns (their H12 reaches further into the mic's HF rolloff region). Current values: ribbon 3.5, dynamic 2.8 (alto+dynamic bumped to 3.5 after Foster's Conn 6M data showed alto reads ~1 dB/H higher than bari with the same RE20), condenser uses the base `ROLLOFF_WARN_THRESHOLD` constant. Both call sites in `toner_tab.py` pass sax_type — the live capture handler from `engine.sax_type`, the Analyze tool from `fp['_preset']['horn_type']`. Real-world rolloff readings observed: Matt AT2020 condenser 1.1-1.3 dB/H, Tyler Neumann TLM 1.8-3.1, Edinger U67 1.65-2.04, Mario AT2035 alto 1.99 / soprano 4.00, Foster RE20 dynamic bari 1.5-2.3 / alto 2.99-3.49, Grant MXL RI44 ribbon tenor 2.16-2.46.

## Analyze Tool

*Moved from CLAUDE-toner.md, section "Tone Analyzer (Toner) > Analyze Tool" (2026-11-10).*

Renamed from Compare. Handles single preset view, two-preset delta analysis with difference chart, and multi-preset spread analysis. Accessible via File > Analyze... or File > Presets > Analyze. Analysis descriptors (complexity, warmth, even/odd, rolloff shape, evenness) are user-configurable in Options > Settings > Analysis tab. Back button on analysis/group report windows returns to the selection picker. Clickable legend labels and chart lines show preset detail (player, mpc, mic, etc.). **Character Map**: 2D quad-chart scatter plot of every selected preset on warmth x complexity axes (Y axis was originally "H4-H5 mean" but switched to complexity 2026-04-06 because warmth and H4-H5 share a denominator and were not independent within tenors); click a dot, legend entry, or chart line and the matching items highlight together in the same color across legend, harmonic chart, and quad map. **Bars chart toggle**: switch the harmonic comparison chart between line view and side-by-side bars view. **Maximize / fullscreen mode**: the Analyze window is non-modal (intentionally NOT a transient toplevel, since transients lose the maximize button on Windows), resizes freely, and has a Maximize button that toggles between zoomed and normal window state (with a macOS Aqua fallback that uses a screen-filling geometry since Aqua doesn't support `state('zoomed')`). The whole window scrolls vertically since comparison data can exceed any viewport. Same pattern is applied to the Group Report dialog. **Cross-sax-type warning**: mixing alto + tenor + bari + soprano in one comparison is allowed but pops a warning, since lower-pitched horns intrinsically read warmer/brighter regardless of mouthpiece.

## Tone Preset Data Model

*Moved from CLAUDE-toner.md, section "Tone Analyzer (Toner) > Tone Preset Data Model" (2026-11-10).*

### Read-side capture filter

**Read-side capture filter** (`_capture_is_plausible()`): Before averaging, `compute_fingerprint` drops captures that fail two checks. (1) **Range filter**: the labeled note must fall inside `SAX_NOTE_RANGES[sax_type]`, which uses tight bounds — lower bound = written low Bb for most types, extended to low A only on bari (common spec) and alto (rare); upper bound = high F# written (no altissimo margin, because altissimo is a harmonic partial of a lower fingering, not a true closed-tube fundamental, so its harmonic structure isn't useful for tone analysis). (2) **Plausibility filter**: no value in `harmonics_db` may be NaN, ±inf (checked with `math.isfinite` — JSON round-trips `-Infinity` happily and one poisons every average), or > +20 dB above the labeled fundamental — that's a clear sub-octave detection error where the real fundamental landed in an upper-harmonic bin. Real low-register alto H2 tops out near +10 dB legitimately, so +20 dB is the hard sentinel. The filter is read-only — raw on-disk captures are NOT modified, so any descriptor formula or filter improvement applies retroactively to historical data on the next load. Was added 2026-04-07 after a library audit found 37 of 58 presets had at least one out-of-range capture; the worst case (Selmer VI alto GP) had 105 of 164 captures bogus.

## Toner Calibration

*Moved from CLAUDE-toner.md, section "Tone Analyzer (Toner) > Toner Calibration" (2026-11-10).*

### Descriptor calibration status

**Descriptor calibration status**: Two descriptors survived data-driven validation: complexity (spectral flatness) and warmth (H2 strength). Brightness/darkness were removed because the labels conflicted with player vocabulary (data showed Yamahas read "dark" but players call them "bright"). Resonance was removed because it read 100% on all 39 presets — no differentiation. Fullness was removed because it depended on brightness/darkness. Key constants live in toner_engine.py: `RICHNESS_RAW_MIN`, `RICHNESS_RAW_RANGE`, `WARMTH_DB_FLOOR`, `WARMTH_DB_RANGE`. Development history in `.claude/projects/*/memory/descriptors.md`. `tools/profile_report.py` generates a detailed single-preset report.
