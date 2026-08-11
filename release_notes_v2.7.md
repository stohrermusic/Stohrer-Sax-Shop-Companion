# v2.7 — Labeled Zones + Job History

Sax Shop Companion is a desktop toolkit for saxophone repair techs: it generates SVG and G-code for laser-cutting pad materials, packs dies and die holders, keeps reference libraries for key heights, serial numbers, and screw specs, and includes a chromatic strobe tuner and a harmonic tone analyzer. v2.7 solves a problem anyone who cuts octave pads knows well — **once small pads are off the laser, you can't tell them apart.** A 7.0 and a 7.5 disc look identical, and the size number engraved on the pad itself doesn't help: below about 5 mm there's no room to engrave it at all, and above that it's either too small to read on card and felt or lost in the darts on leather. Moving it inboard isn't an option either — the middle of a small leather pad is the sealing surface, and those are usually octave pads that need to stay tough. So v2.7 puts the label on the **waste** instead of the part: turn on **Labeled Zones** and each size is cut in its own bordered block with the size engraved beside it. This release also brings the **Job History** to the Pad Maker, a lid reminder before the laser fires, and completes all four translations.

## Pad Maker

- **(new) Labeled Zones** (Options > Sizing Rules) — cut each pad size in its own bordered, labeled group so you can tell small discs apart when you're picking them off the bed. Off by default; turn it on and set the size range it applies to (7.0–12.5 mm out of the box).
  - **(new)** The layout adapts to what you're cutting on. On a **rectangular sheet** each size gets a tidy grid block, packed into a band along the bottom, with everything larger nesting normally at full density above it. On a **traced scrap** each size gets its own horizontal band clipped to your outline, so the groups follow the real shape of the piece.
  - **(new)** Every group is drawn with an **engraved boundary line** and its size number, both scored rather than cut — the sheet stays in one piece and nothing drops out early. Labels and borders are engraved first, before any pad is cut loose, so the sheet is marked before parts start coming free.
  - **(new)** On camera-captured scraps the boundary line automatically lands a few millimetres inside the real edge of the material, so it's engraved on leather rather than off the edge.
  - **(new)** The nesting preview shows the zones exactly as they'll be cut, so you can see the grouping before anything moves.
- **(new) Job History** (File > Job History) — a log of every batch that reached an output stage: SVG written, G-code written, or streamed to the laser. Newest first, showing date, output type, material, pad count, and sheet size.
  - **(new)** Select any job to see its full pad list, center hole, base filename, output folder, and which preset it came from.
  - **(new) Load into Pad Maker** puts the pad list, materials, sheet size, center hole, and filename back into the form so you can re-cut a set you've done before. Your sizing rules and G-code settings are left alone.
  - **(new)** Laser runs that were stopped or errored are logged too, flagged with a `*` — so pads that never got cut have an explanation.
  - **(new)** In Scrap Mode each scrap is logged separately and numbered to match the session. Holds the last 300 jobs; delete single entries or clear the log.
- **(new) Sizing Rules names the preset you're actually in.** Open the dialog and the dropdown now shows which saved preset matches your current values, instead of sitting blank until you load something.
- **Per-material G-code presets** — Options > G-code Settings gives every material (felt, card, leather, acrylic, basswood) its own preset bar: Load / Save / Rename / Delete. Editing one material doesn't disturb another, and the library is shared with the Tooling dialog — a felt preset you save here shows up there. A `Default` is created for each material on first run.
- **Machine Integration (experimental, opt-in)** — direct USB serial control of a Grbl-compatible laser. Enabled via File > Feature Set > "Experimental: machine integration." Off by default; off entirely if you skip the opt-in. Tested on the Creality Falcon2 Pro 40W; might also work on other Grbl 1.1+ machines.
  - **Camera Calibration** — one-time wizard. Engrave a ChArUco card on basswood, capture frames; the app then knows where the camera sees vs. where the laser cuts.
  - **Get from Camera** (polygon-draw dialog) — snap a photo of a scrap piece on the bed, the app traces its outline. Camera-polygon inset margin (Options > Machine) shrinks the polygon a few millimeters to absorb edge-measurement noise.
  - **Live camera overlay** in the polygon-draw dialog — overlays the camera feed at 1:1 scale so you can trace your scrap by eye, even before a capture.
  - **Frame & Cut** — third button next to Generate SVG / Generate G-code. Generates in memory, shows the nesting preview first so you can see exactly which pads will land on this scrap (or the whole batch) before anything moves — backing out consumes nothing — then opens the position-the-head dialog (Home Laser, jog cluster, Try Auto Locate when a camera-referenced polygon is loaded), runs a low-power framing loop until you click "Looks Good — Cut!", and streams the cut. Pause / Resume / Stop in real time. Works in **Scrap Mode** — one scrap per click: preview the partial batch, frame the captured outline, cut, then re-capture the next piece from the continue dialog, repeating until every pad is placed.
  - **(new) Lid reminder before the cut.** A quick "Is the lid closed?" prompt appears once, right before the cut starts. Forgetting the lid drops the machine into a door alarm that costs you a laser power cycle *and* an app restart, so it's worth the one click. Framing doesn't prompt — it runs at low power. Stopping a job mid-cut keeps its own separate confirmation.
  - **(new) Framing power is now adjustable per material** (Options > Machine > Framing Power). The framing pass runs at low power so you can see where the cut will land without marking anything, but how visible that is depends on what you're looking at — a setting that reads clearly on card can be hard to see on dark leather. Set as a percentage, per material. Defaults are unchanged, so nothing about your framing looks different unless you go change it.
  - **Stuck-alarm recovery** — if a previous run left Grbl in an alarm state, the next stream clears it automatically (`$X` at stream start, a no-op when idle), so a hiccup no longer means power-cycling the Falcon and restarting the app.
  - **Inset Margin** — adds a safety margin to placement on scraps so you don't accidentally clip an edge. Frame still traces the actual scrap outline; cuts respect the safety boundary.
  - **Machine menu** (Options > Machine): Home Laser, Test Connection, Clear Errors ($X), Reset Falcon (soft-reset), Camera Calibration, Camera-Polygon Inset Margin, **(new)** Framing Power.
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
- **Speed & Power Test (beta)** — generate a sheet of small test discs at different speed / power / passes combinations to dial in laser settings on a new material. Set a **hole diameter** and each test disc becomes a **washer/ring** for shim stock — the inner hole is cut first so the part stays anchored to the sheet, and the disc's ID engraves in the ring. The **sheet size is a soft target**: if your sweep doesn't fit the sheet you entered, the app grows it to fit and shows you the layout rather than erroring out. Engraving feed/power are editable so the labels stay legible while the cut settings are still unknown.
- Per-material G-code presets — the Acrylic and Basswood sections in Options > Tooling Settings carry the same Load / Save / Rename / Delete preset bar as Pad Maker, sharing one library. Acrylic feeds the die holders and inserts; basswood feeds the camera-calibration card engrave. Defaults tuned for the Falcon2 Pro 40W; adjust to your machine.
- Phil Noy's pad-making method is credited at the top of the tab (with a link to noysaxophonesupplies.com) and engraved on every holder ring and die insert.

## Chromatic Strobe Tuner

- 12-wheel stroboscopic chromatic tuner.
- GPU-accelerated rendering via Rust/wgpu on Windows and Linux — 60–120 fps; automatic CPU fallback if GPU unavailable.
- macOS always uses the canvas renderer — Tk on macOS doesn't expose a native view the GPU renderer can draw into, so Macs are canvas-only by design. Fully functional, just capped at canvas frame rates.
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
- **Fully translated** — the entire UI is localized into Spanish, German, French, and Italian. Sax-craft terminology (pad / zapatilla / tampon / Polster / tampone; basswood / tilo / tilleul / Lindenholz / tiglio; etc.) kept consistent across locales. **(new)** All four catalogs are at 100%, including the new Labeled Zones controls — the last untranslated strings have been filled in, and character-encoding corruption that had crept into a handful of German, French, and Spanish messages is repaired, so accented characters display properly again.
- **macOS** — dual builds: Apple Silicon (full features) and Intel (no audio features). Native dark/light mode support. Cmd-Q (and the app menu's Quit) saves your settings on the way out. The microphone and camera permission prompts appear correctly as of v2.63 (every earlier Apple Silicon build had a code-signing packaging bug that made macOS silently deny access without ever asking — see Upgrading below if an older build already bit you). The macOS download is also ~68 MB rather than the old 258 MB.
- **Linux** — GPU rendering via Vulkan with X11 display handle. Audio features require libportaudio2. Machine integration available.
- **Windows** — Inno Setup installer (Start Menu + uninstaller) plus the standalone .exe; auto-eject removable drives after G-code export.
- Platform-appropriate config storage with automatic migration from old locations (`%APPDATA%`, `~/Library/Application Support`, `~/.config`).
- Reference libraries — Key Heights, Serial Number lookup, and Screw Specs, with one-click import of Matt's published libraries from stohrermusic.com. **(new)** The Screw Specs tab now shows a short getting-started hint when the library is empty, instead of an unexplained blank list.
- Import Settings from Folder — manual transfer between machines.
- Tab-aware User Guide — Help > User Guide shows the section relevant to your current tab.
- Error logging — rotating log file at Help > Open Log File for diagnostics.
- Feature Set (File > Feature Set) — choose which tabs to show. Toner remains opt-in beta; Tuner is on by default.

## Known limitations

- **Labeled Zones cost material, and that's the trade.** Grouping pads into blocks and bands packs less tightly than letting the nester fill every gap — expect to give up roughly 10–15% of a sheet on the sizes you've zoned. It's aimed at tiny pads, where the yield is high and the demand is low anyway; if you're cutting a sheet full of large pads, leave it off. If a size no longer fits once it's been grouped, the app tells you it couldn't fit everything rather than quietly cutting a short sheet — widen the size range, use a bigger piece, or turn zones off for that job.
- **Labeled Zones are not available in Scrap Mode.** A scrap only takes part of a size's count, so a group would have to be re-sized for every piece. Zones apply when you're cutting a whole batch on one sheet or one traced scrap, which is the case they're built for.
- **Job History records what you entered, not the settings behind it.** Loading a past job restores the pad list, materials, sheet size, center hole, and filename — it deliberately does not touch your sizing rules or G-code settings, and it does not restore a custom polygon shape (a camera-captured scrap has already been cut up, so re-using its outline would put pads in the wrong place). Draw or capture the piece you're cutting now. Jobs can't be loaded while a scrap session is running, since the session owns the pad list.
- **Machine Integration is experimental** and opt-in for a reason. The Falcon2 Pro 40W is the tested machine; other Grbl machines should work but YMMV. "Try Auto Locate" drives the head to the polygon's bottom-left vertex using the camera homography — good enough for rough positioning, but fine-tune with the jog buttons before clicking Start Frame. Disabled until you home the laser in the current session. Stuck-alarm recovery makes a mid-job hiccup recoverable in-app rather than requiring a power cycle, but if a cut ever stalls, re-frame and re-cut the affected scrap.
- **The lid reminder is a reminder, not a sensor.** The app does not detect whether your lid is actually closed — testing on the Falcon2 Pro 40W showed it never reports lid state at all, so there's nothing reliable to read. The prompt is there to make you look.
- **Camera Calibration** requires engraving a ChArUco card on a 12×12 basswood blank. Takes a while; basswood is consumable (each engrave is permanent, so a re-calibration needs a fresh piece).
- **Kerf is a property of your cut, not just the material.** A laser's real kerf shifts with focus height, lens condition, cut speed/power/passes, and even how a material chars — so a profile that cuts at a different speed or pass count than the one you measured will need its own kerf value. If parts start coming out slightly off-size, re-run the Tooling > Kerf Test and update the affected material's `kerf_width`.
- **Tone Analyzer is marked beta.** Descriptors are still being calibrated as we gather more data from different horns, mouthpieces, mics, and players. Raw harmonic measurements are always saved, so future formula improvements apply retroactively to your historical captures.
- **macOS** — the strobe tuner is not GPU-accelerated (canvas renderer only — see the Tuner section above). The app is also not signed with an Apple Developer certificate: right-click → Open the first time, or run `xattr -cr` on the download. Instructions in the README. And because the signature is ad-hoc rather than Apple-issued, macOS may re-ask for mic/camera permission after you update to a new version — that's normal.

## Upgrading from v2.63 / earlier

- **This is a drop-in upgrade** — no migrations, no recalibration. Settings, presets, libraries, and camera calibrations carry over untouched.
- **Labeled Zones are off by default.** Nothing about your existing jobs changes until you turn them on in Options > Sizing Rules. The setting is saved with your Sizing Rules presets, so you can keep a zoned preset for octave-pad work and an unzoned one for everything else.
- **Your saved Sizing Rules presets still work.** Presets saved before this release simply read as "zones off," and the dropdown will still recognise which one you're in.
- **Job History starts empty.** It logs from the moment you install v2.7 forward; there's no way to reconstruct jobs you cut before this release. The log lives in `job_history.json` in your config folder alongside your presets.
- **Mac users on v2.62 or earlier:** v2.63 was the release that made the microphone and camera permission prompts appear for the first time. Click **OK / Allow** when macOS asks. If the Tuner still can't hear anything afterward, your Mac may have cached the old silent denial — clear it with these Terminal commands, then relaunch:

  ```
  tccutil reset Microphone com.stohrer.saxshopcompanion
  tccutil reset Camera com.stohrer.saxshopcompanion
  ```
- **Apple Silicon Mac users:** the Tuner-tab crash present in every Mac build from v1.95 through v2.6 was fixed in v2.61 and remains fixed here — if you skipped v2.61, this upgrade includes that fix.
- Settings, presets, and libraries auto-migrate from older config locations on first run.
- The per-material G-code preset library (`gcode_presets.json`) is created automatically from your current G-code settings, and any missing material is backfilled with a `Default` — your existing feeds and powers become your starting presets.
- Old `tone_profiles.json` is automatically renamed to `toner_data.json`; pad presets in legacy flat format are migrated into a library.
- Tested back to v1.0 — settings loading is hardened against malformed/old config files.
