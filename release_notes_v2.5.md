# v2.5 — Camera-Driven Workflow + Direct Laser Control

Sax Shop Companion now talks to your laser directly. v2.5 adds an experimental machine-integration layer: a one-time ChArUco-card camera calibration, a "Get from camera" button that auto-detects scrap outlines, a live camera overlay so you can trace by eye, and a **Frame & Cut** button that generates G-code in memory, positions the head, traces the cut at low power for verification, then streams the actual cut to the laser. All opt-in via File > Feature Set. The rest of the app (Pad Maker SVG/G-code workflows, Tuner, Tone Analyzer) is unchanged.

## Pad Maker

- **(new) Machine Integration (experimental, opt-in)** — direct USB serial control of a Grbl-compatible laser. Enabled via File > Feature Set > "Experimental: machine integration." Off by default; off entirely if you skip the opt-in. Tested on the Creality Falcon2 Pro 40W; might also work on other Grbl 1.1+ machines.
  - **Camera Calibration** — one-time wizard. Engrave a ChArUco card on basswood, capture 12 frames; the app now knows where the camera sees vs. where the laser cuts.
  - **Get from Camera** (polygon-draw dialog) — snap a photo of a scrap piece on the bed, the app traces its outline. Camera-polygon inset margin (Options > Machine) shrinks the polygon a few millimeters to absorb edge-measurement noise.
  - **Live camera overlay** in the polygon-draw dialog — overlays the camera feed at 1:1 scale so you can trace your scrap by eye, even before a capture.
  - **Frame & Cut** — third button next to Generate SVG / Generate G-code. Generates in memory, opens the position-the-head dialog (Home Laser, jog cluster, Try Auto Locate when a camera-referenced polygon is loaded), runs a low-power framing loop until you click "Looks Good — Cut!", then streams the cut. Pause / Resume / Stop in real time.
  - **Inset Margin** — adds a safety margin to placement on scraps so you don't accidentally clip an edge. Frame still traces the actual scrap outline; cuts respect the safety boundary.
  - **Machine menu** (Options > Machine): Home Laser, Test Connection, Clear Errors ($X), Reset Falcon (soft-reset), Camera Calibration, Camera-Polygon Inset Margin.
- **(new) Polygon draw improvements**:
  - Free vertex placement (no grid snap) for precise tracing.
  - Grid auto-grows to cover your laser bed (default 17 in / 43 cm).
  - "Draw / Capture Shape" button label tracks the machine-integration toggle — "Draw Shape" alone when machine integration is off.
- **(new) Scrap mode — large-batch optimization** (secret-handshake popup at ≥ 75 pads remaining): multistart greedy nester that tries multiple disc orderings per scrap and keeps the best result. Opt-in per session. Adds ~5–30s of compute per scrap; typically fits 5–15% more pads on dense batches.
- Sizing Rules Presets — save the entire Sizing Rules dialog as a named preset, load via dropdown, import/export to share with other techs.
- Nesting preview with per-material approval, edge bias d-pad (cardinal + corner directions, smallest pads first in corners), max fill mode (`size x max`).
- Custom polygon shapes for irregular leather skins.
- Scrap mode — place pads across multiple irregular pieces, tracking remaining between sheets, with preview, edge bias, and polygons all working together.
- Filled engraving mode — scan-line raster fill using Roboto outlines, with overscan for clean character edges. Per-material engraving mode (line vs filled).
- Auto-fit engraving — text shifts toward center on small pads, scales only as a last resort.
- G-code options — air assist toggles (M8/M9) per layer, cut grouping (by layer or by pad), full-kerf compensation, configurable return speed, optional SD card eject after export (Windows).
- Last-used library memory for both pad presets and key heights.

## Tooling

- Die inserts: small (50mm OD) and large (70mm OD).
- Die holders: 85mm OD stack. Pick 5-layer or 6-layer; variant Large / Small / Both (Both nests two complete independent holders on one sheet). User-defined sheet size with a clear minimum-size error if pieces don't fit.
- Die Organizer — SVG templates for a stackable die organizer (230 × 330 mm). Cut three Uppers and one Lower, align the four 1/8″ corner holes, glue the stack together.
- Pad Press Spacers — bundled 3D-printable STL files for setting pad press depth.
- Kerf Test pattern generator for calibrating your laser.
- Speed & Power Test (beta) — generate a sheet of small test discs at different speed / power / passes combinations to dial in laser settings on a new material.
- **(new) G-code presets** — Options > Tooling Settings now exposes per-material G-code for **Acrylic AND Basswood**. The basswood preset feeds the die organizer (when you cut it in LightBurn) AND the camera-calibration card engrave. Defaults tuned for the Falcon2 Pro 40W; adjust to your machine.

## Chromatic Strobe Tuner

- 12-wheel stroboscopic chromatic tuner.
- GPU-accelerated rendering via Rust/wgpu — 60–120 fps; automatic CPU fallback if GPU unavailable.
- Per-ring octave brightness from real spectral data.
- Grouped slider panel (display, pitch, bias) and vintage backlit VU meter.
- Per-pitch-class phase tracking with temporal smoothing.
- Transposition support (Concert, Bb, Eb, F).
- Configurable frame rate (60/90/120 fps), backlight color, and faceplate color.

## Harmonic Tone Analyzer (beta)

A real-time harmonic spectrum analyzer for saxophone. Captures the fundamental and overtones of your sound and lets you compare setups (horn, mouthpiece, reed, mic, mic placement, embouchure) over time.

- Live spectrum (FFT) and Bars (per-harmonic) views, linear or dB scale.
- Detects fundamental pitch, extracts up to 20 harmonics.
- Intonation gauge with cents readout and ±4¢ "in tune" lamp.
- Auto-transposition by saxophone type with concert pitch toggle.
- Spectrum overlay: load any preset as a ghost overlay on the live spectrum.
- **Tone presets** with horn/player/mouthpiece/reed/mic metadata.
  - Free capture (continuous micro-captures while playing) and guided calibration modes.
  - WAV recording on by default, with offline reanalysis at ~2× the harmonic resolution of live capture.
  - WAV import for offline analysis (16/24/32-bit).
  - Mic type, model, and position stored per preset for reproducibility.
  - Mutate Preset for A/B testing (duplicate with one variable changed).
  - Sandbox mode for non-sax instruments and experimental setups.
  - All captures stored in concert pitch for cross-instrument comparison.
- **Analyze tool** — single preset detail, two-preset delta, multi-preset spread analysis.
  - Difference charts and harmonic-range interpretation (H1-H4 ≈ bore, H7-H13 ≈ neck/mpc, broadband ≈ mpc/player).
  - Population percentiles by sax type.
  - Configurable comparison descriptors: complexity, warmth, even/odd, rolloff shape, evenness.
  - Filter by make/model/mic type/search, multi-select, cross-player context notes.
- Recording-quality tracking (rolloff rate) with live warnings and cross-mismatch detection.
- Coverage summary after capture sessions.

## Cross-platform & General

- **(new) Machine integration available on Windows, macOS, and Linux** — pyserial works cross-platform. Off by default everywhere; opt in via File > Feature Set if you have a Grbl machine.
- **(new) Tightened user guide** for the Custom Shapes / Scrap Mode / Machine Integration sections; added Camera Calibration + Frame & Cut subsections; trimmed redundant phrasing throughout. Tone passes match the rest of the app — succinct, trusts the reader.
- **(new) Translations** — all v2.5 strings translated into es / de / fr / it. Sax-craft terminology (pad / tampón / tampon / Polster / tampone; basswood / tilo / tilleul / Lindenholz / tiglio; etc.) kept consistent across locales.
- **macOS** — dual builds: Apple Silicon (full features) and Intel (no audio features). Native dark/light mode support.
- **Linux** — GPU rendering via Vulkan with X11 display handle. Audio features require libportaudio2. **(new)** Machine integration available.
- **Windows** — auto-eject removable drives after G-code export. **(new)** Inno Setup installer; standalone .exe is the existing build, the installer adds Start Menu + uninstaller support.
- Platform-appropriate config storage with automatic migration from old locations (`%APPDATA%`, `~/Library/Application Support`, `~/.config`).
- Import Settings from Folder — manual transfer between machines.
- Tab-aware User Guide — Help > User Guide shows the section relevant to your current tab.
- Error logging — rotating log file at Help > Open Log File for diagnostics.
- Feature Set (File > Feature Set) — choose which tabs to show. Toner remains opt-in beta; Tuner is on by default.

## Known limitations

- **Machine Integration is experimental** and opt-in for a reason. The Falcon2 Pro 40W is the tested machine; other Grbl machines should work but YMMV. The "Try Auto Locate" feature drives the head to the polygon's bottom-left vertex using the camera homography — accuracy is good enough for rough positioning but you may need to fine-tune with the jog buttons before clicking Start Frame. Disabled until you home the laser in the current session.
- **Camera Calibration** requires engraving a ChArUco card on a 12×12 basswood blank. Takes a while; basswood is consumable (each engrave is permanent, so a re-calibration needs a fresh piece).
- **Tone Analyzer is marked beta.** Descriptors are still being calibrated as we gather more data from different horns, mouthpieces, mics, and players. Raw harmonic measurements are always saved, so future formula improvements apply retroactively to your historical captures.
- **macOS** — the app is not signed with an Apple Developer certificate. Right-click → Open the first time, or run `xattr -cr` on the download. Instructions in the README.

## Upgrading from v2.0 / v1.x

- Settings, presets, and libraries auto-migrate from older config locations on first run.
- Existing camera calibrations (if you used the v3.0-dev builds) remain valid — no need to recalibrate.
- Old `tone_profiles.json` is automatically renamed to `toner_data.json`.
- Pad presets in legacy flat format are migrated into a "My Presets" library.
- Tested back to v1.0 — settings loading is hardened against malformed/old config files.
- Stale settings keys from removed features (auto-frame, dot calibration) are silently ignored on load — your config file stays clean.
