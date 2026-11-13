# v2.61 — macOS Bug Fixes (Tuner Crash + Cmd-Q Settings Save)

Sax Shop Companion v2.61 is a bug-fix release aimed squarely at macOS. It turns out the strobe tuner's GPU renderer was crashing the whole app on Apple Silicon Macs the moment you opened the Tuner tab — the windowing toolkit on macOS never actually gives the renderer a real surface to draw into, and the failure is the hard kind Python can't catch. Macs now always use the canvas renderer (fully functional, just not as silky-smooth), and the GPU extension isn't even bundled into the Mac build anymore. While we were in there: Cmd-Q on a Mac now saves your settings on the way out, which it silently wasn't doing before. **If you're on an Apple Silicon Mac, please upgrade** — every Mac build from v1.95 through v2.6 has the Tuner-tab crash. Nothing changed for Windows or Linux.

## Pad Maker

- **Per-material G-code presets** — Options > G-code Settings gives every material (felt, card, leather, acrylic, basswood) its own preset bar: Load / Save / Rename / Delete. Editing one material doesn't disturb another, and the library is shared with the Tooling dialog — a felt preset you save here shows up there. A `Default` is created for each material on first run.
- **Machine Integration (experimental, opt-in)** — direct USB serial control of a Grbl-compatible laser. Enabled via File > Feature Set > "Experimental: machine integration." Off by default; off entirely if you skip the opt-in. Tested on the Creality Falcon2 Pro 40W; might also work on other Grbl 1.1+ machines.
  - **Camera Calibration** — one-time wizard. Engrave a ChArUco card on basswood, capture frames; the app then knows where the camera sees vs. where the laser cuts.
  - **Get from Camera** (polygon-draw dialog) — snap a photo of a scrap piece on the bed, the app traces its outline. Camera-polygon inset margin (Options > Machine) shrinks the polygon a few millimeters to absorb edge-measurement noise.
  - **Live camera overlay** in the polygon-draw dialog — overlays the camera feed at 1:1 scale so you can trace your scrap by eye, even before a capture.
  - **Frame & Cut** — third button next to Generate SVG / Generate G-code. Generates in memory, opens the position-the-head dialog (Home Laser, jog cluster, Try Auto Locate when a camera-referenced polygon is loaded), runs a low-power framing loop until you click "Looks Good — Cut!", then streams the cut. Pause / Resume / Stop in real time. Works in **Scrap Mode** — one scrap per click: place a partial batch, frame the captured outline, cut, then re-capture the next piece from the continue dialog, repeating until every pad is placed.
  - **Stuck-alarm recovery** — if a previous run left Grbl in an alarm state, the next stream clears it automatically (`$X` at stream start, a no-op when idle), so a hiccup no longer means power-cycling the Falcon and restarting the app.
  - **Inset Margin** — adds a safety margin to placement on scraps so you don't accidentally clip an edge. Frame still traces the actual scrap outline; cuts respect the safety boundary.
  - **Machine menu** (Options > Machine): Home Laser, Test Connection, Clear Errors ($X), Reset Falcon (soft-reset), Camera Calibration, Camera-Polygon Inset Margin.
- **Scrap mode — large-batch optimization** (opt-in popup at ≥ 75 pads remaining): multistart greedy nester that tries multiple disc orderings per scrap and keeps the best result. Per session. Adds ~5–30s of compute per scrap; typically fits 5–15% more pads on dense batches. Safe with camera scrap mode — every scrap is nested independently into its own captured shape.
- **Polygon draw**: free vertex placement (no grid snap) for precise tracing; grid auto-grows to cover your laser bed (default 17 in / 43 cm); "Draw / Capture Shape" button label tracks the machine-integration toggle.
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
- Per-material G-code presets — the Acrylic and Basswood sections in Options > Tooling Settings carry the same Load / Save / Rename / Delete preset bar as Pad Maker, sharing one library. Acrylic feeds the die holders and inserts; basswood feeds the camera-calibration card engrave. Defaults tuned for the Falcon2 Pro 40W; adjust to your machine.
- Phil Noy's pad-making method is credited at the top of the tab (with a link to noysaxophonesupplies.com) and engraved on every holder ring and die insert.

## Chromatic Strobe Tuner

- 12-wheel stroboscopic chromatic tuner.
- GPU-accelerated rendering via Rust/wgpu on Windows and Linux — 60–120 fps; automatic CPU fallback if GPU unavailable.
- **(new) macOS now always uses the canvas renderer** — and the fix matters more than it sounds: the GPU renderer was crashing the entire app on Apple Silicon Macs as soon as the Tuner tab opened (every Mac build v1.95–v2.6). Tk on macOS doesn't expose a native view the GPU renderer can draw into, so Macs are canvas-only by design from here on. Fully functional, just capped at canvas frame rates.
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
  - 2D Character Map (warmth × complexity), bars/line chart toggle, click-to-highlight across legend / chart / map.
  - Population percentiles by sax type.
  - Configurable comparison descriptors: complexity, warmth, even/odd, rolloff shape, evenness.
  - Filter by make/model/mic type/search, multi-select, cross-player context notes.
- Recording-quality tracking (rolloff rate) with live warnings and cross-mismatch detection.
- Coverage summary after capture sessions.

## Cross-platform & General

- **Machine integration available on Windows, macOS, and Linux** — pyserial works cross-platform. Off by default everywhere; opt in via File > Feature Set if you have a Grbl machine.
- **Fully translated** — the entire UI is localized into Spanish, German, French, and Italian. Sax-craft terminology (pad / tampón / tampon / Polster / tampone; basswood / tilo / tilleul / Lindenholz / tiglio; etc.) kept consistent across locales.
- **macOS** — dual builds: Apple Silicon (full features) and Intel (no audio features). Native dark/light mode support. **(new)** Cmd-Q (and the app menu's Quit) now saves your settings on the way out — previously only closing the window with the red button did, and a Cmd-Q quit silently dropped any settings you'd changed that session.
- **Linux** — GPU rendering via Vulkan with X11 display handle. Audio features require libportaudio2. Machine integration available.
- **Windows** — Inno Setup installer (Start Menu + uninstaller) plus the standalone .exe; auto-eject removable drives after G-code export.
- Platform-appropriate config storage with automatic migration from old locations (`%APPDATA%`, `~/Library/Application Support`, `~/.config`).
- Import Settings from Folder — manual transfer between machines.
- Tab-aware User Guide — Help > User Guide shows the section relevant to your current tab.
- Error logging — rotating log file at Help > Open Log File for diagnostics.
- Feature Set (File > Feature Set) — choose which tabs to show. Toner remains opt-in beta; Tuner is on by default.

## Known limitations

- **Machine Integration is experimental** and opt-in for a reason. The Falcon2 Pro 40W is the tested machine; other Grbl machines should work but YMMV. "Try Auto Locate" drives the head to the polygon's bottom-left vertex using the camera homography — good enough for rough positioning, but fine-tune with the jog buttons before clicking Start Frame. Disabled until you home the laser in the current session. Stuck-alarm recovery makes a mid-job hiccup recoverable in-app rather than requiring a power cycle, but if a cut ever stalls, re-frame and re-cut the affected scrap.
- **Camera Calibration** requires engraving a ChArUco card on a 12×12 basswood blank. Takes a while; basswood is consumable (each engrave is permanent, so a re-calibration needs a fresh piece).
- **Tone Analyzer is marked beta.** Descriptors are still being calibrated as we gather more data from different horns, mouthpieces, mics, and players. Raw harmonic measurements are always saved, so future formula improvements apply retroactively to your historical captures.
- **macOS** — the strobe tuner is not GPU-accelerated (canvas renderer only — see the Tuner section above). The app is also not signed with an Apple Developer certificate: right-click → Open the first time, or run `xattr -cr` on the download. Instructions in the README.

## Upgrading from v2.6 / earlier

- **Apple Silicon Mac users: this upgrade fixes the Tuner-tab crash** in every Mac build from v1.95 through v2.6. Settings and presets carry over untouched.
- One heads-up about the older release pages: v1.95 through v2.6 describe the tuner as GPU-accelerated with "automatic CPU fallback" on every platform. That was never true on macOS — instead of falling back, those Mac downloads crash when the Tuner tab opens. The old pages stay as they were published, so if you're grabbing (or linking someone to) a Mac build, use this release or newer.
- Settings, presets, and libraries auto-migrate from older config locations on first run.
- The per-material G-code preset library (`gcode_presets.json`) is created automatically from your current G-code settings, and any missing material is backfilled with a `Default` — your existing feeds and powers become your starting presets.
- Existing camera calibrations remain valid — no need to recalibrate.
- Old `tone_profiles.json` is automatically renamed to `toner_data.json`.
- Pad presets in legacy flat format are migrated into a library.
- Tested back to v1.0 — settings loading is hardened against malformed/old config files.
