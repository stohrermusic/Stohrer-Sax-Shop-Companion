# v2.80 — Leather locator marks, tuner accuracy, Mac testing

Sax Shop Companion is a desktop toolkit for saxophone repair techs: it generates SVG and G-code for laser-cutting pad materials, packs dies and die holders, keeps reference libraries for key heights, serial numbers, and screw specs, and includes a chromatic strobe tuner and a harmonic tone analyzer. v2.80 adds engraved locator marks for centering felt on small leather pads, tidies the Sizing Rules dialog, improves tuner accuracy, fixes a handful of error paths that showed an "Unexpected Error" instead of a plain message, and has been tested on real Macs for the first time.

## Pad Maker

- **(new) Leather Locator Marks (beta)** (Options > Sizing Rules > Lesser-used settings) — engraved guides on leather pads showing where the felt sits, so a pad with no center hole can be centered by eye before pressing. Pick four short lines ending at the felt edge (equal gaps all round = centered), a circle the size of the felt, or both; optionally dashed. Applies to a pad-size range you set (7.0–16.0 mm out of the box — the sizes that get no center hole under the default sizing). Off by default.
  - The marks have to be on the side the felt touches, so **cut the leather flesh-side (fuzzy side) up** when you use them. The size number lands on the flesh side too, in the wrap zone, where it's glued against the card and hidden.
  - **SVG output** puts the marks on their own layer (Options > Layer Colors > Leather locator), so you can give them their own power in LightBurn. **G-code** engraves them at the leather engraving settings — there's no separate power for them yet — so use the SVG output if the marks need their own power.
- **(new) Lesser-used settings** — Labeled Zones, Leather Locator Marks, and Export Settings now live behind one collapsible header at the bottom of Sizing Rules, so the dialog stays about pad geometry. It opens itself when any of the three is switched on. The frame is now just "Sizing Rules" (the "(Advanced)" is gone), and the bottom button reads **Revert to Program Defaults** so it can't be mistaken for "back to the loaded preset."
- **Labeled Zones (beta)** — cut each pad size in its own bordered, labeled group so you can tell small discs apart when you're picking them off the bed. Off by default; turn it on and set the size range it applies to (7.0–12.5 mm out of the box).
  - Each group is laid out the way you'd arrange the pads by hand — six as a 3×2 block, nine as 3×3, eight as 4×2 — so you can count them at a glance and spot a missing one before you leave the machine.
  - The groups themselves get nested, like oversized pads. On a **traced scrap** they tuck in wherever they fit, following the shape of the piece; on a **rectangular sheet** they pack into a band along the bottom, with everything larger nesting normally at full density above. Either way, sizes outside the range you set are unaffected.
  - The boundary line and size number are **engraved, never cut**, so the sheet stays in one piece — and they go down first, before any pad is cut loose, so the material is marked before parts start coming free.
  - On camera-captured scraps the boundary line automatically lands a few millimetres inside the real edge of the material, so it's engraved on leather rather than off the edge.
  - The nesting preview shows the zones exactly as they'll be cut, so you can see the grouping before anything moves.
- **Job History** (File > Job History) — a log of every batch that reached an output stage: SVG written, G-code written, or streamed to the laser. Newest first, showing date, output type, material, pad count, and sheet size.
  - Select any job to see its full pad list, center hole, base filename, output folder, and which preset it came from.
  - **Load into Pad Maker** puts the pad list, materials, sheet size, center hole, and filename back into the form so you can re-cut a set you've done before. Your sizing rules and G-code settings are left alone.
  - Laser runs that were stopped or errored are logged too, flagged with a `*` — so pads that never got cut have an explanation.
  - In Scrap Mode each scrap is logged separately and numbered to match the session. Holds the last 300 jobs; delete single entries or clear the log.
- **Sizing Rules names the preset you're actually in.** Open the dialog and the dropdown shows which saved preset matches your current values, instead of sitting blank until you load something.
- **Per-material G-code presets** — Options > G-code Settings gives every material (felt, card, leather, acrylic, basswood) its own preset bar: Load / Save / Rename / Delete. Editing one material doesn't disturb another, and the library is shared with the Tooling dialog — a felt preset you save here shows up there. A `Default` is created for each material on first run. **(new)** Saving with a bad number in a field now tells you which material and field is wrong, instead of an "Unexpected Error."
- **Machine Integration (experimental, opt-in)** — direct USB serial control of a Grbl-compatible laser. Enabled via File > Feature Set > "Experimental: machine integration." Off by default; off entirely if you skip the opt-in. Tested on the Creality Falcon2 Pro 40W; might also work on other Grbl 1.1+ machines.
  - **Camera Calibration** — one-time wizard. Engrave a ChArUco card on basswood, capture frames; the app then knows where the camera sees vs. where the laser cuts. **(new)** The save step at the end of calibration no longer crashes with the OpenCV version the current builds ship (OpenCV 5 hands back the board corner IDs in a different shape than 4.x did).
  - **Get from Camera** (polygon-draw dialog) — snap a photo of a scrap piece on the bed, the app traces its outline. Camera-polygon inset margin (Options > Machine) shrinks the polygon a few millimeters to absorb edge-measurement noise.
  - **Live camera overlay** in the polygon-draw dialog — overlays the camera feed at 1:1 scale so you can trace your scrap by eye, even before a capture.
  - **Frame & Cut** — third button next to Generate SVG / Generate G-code. Generates in memory, shows the nesting preview first so you can see exactly which pads will land on this scrap (or the whole batch) before anything moves — backing out consumes nothing — then opens the position-the-head dialog (Home Laser, jog cluster, Try Auto Locate when a camera-referenced polygon is loaded), runs a low-power framing loop until you click "Looks Good — Cut!", and streams the cut. Pause / Resume / Stop in real time. Works in **Scrap Mode** — one scrap per click: preview the partial batch, frame the captured outline, cut, then re-capture the next piece from the continue dialog, repeating until every pad is placed. On a camera-captured scrap the cut lands exactly where the framing pass showed it. The G-code actually streamed for the last cut is kept as `last_frame_cut.gcode` in your config folder, next to the log, for troubleshooting.
  - **Lid reminder before framing.** A quick "Is the lid closed?" prompt, once, right before the framing pass starts — the moment the laser first comes on. Framing runs at low power, but the lid interlock doesn't care about power: an open lid there drops the machine into a door alarm that costs a laser power cycle *and* an app restart. The cut doesn't ask again; by then framing has been running, so the lid is shut.
  - **Framing power is adjustable per material** (Options > Machine > Framing Power). Framing runs at low power so you can see where the cut will land without marking anything — but a setting that reads clearly on card can be invisible on dark leather. Set it as a percentage per material, 0 to 3%; framing only ever needs a fraction of cutting power, so the range is deliberately narrow and fractional values like 1.5 are what you'll want.
  - **Stuck-alarm recovery** — if a previous run left Grbl in an alarm state, the next stream clears it automatically (`$X` at stream start, a no-op when idle), so a hiccup no longer means power-cycling the Falcon and restarting the app.
  - **Inset Margin** — adds a safety margin to placement on scraps so you don't accidentally clip an edge. Frame still traces the actual scrap outline; cuts respect the safety boundary.
  - **Machine menu** (Options > Machine): Home Laser, Test Connection, Clear Errors ($X), Reset Falcon (soft-reset), Camera Calibration, Camera-Polygon Inset Margin, Framing Power.
- **Polygon draw**: free vertex placement (no grid snap) for precise tracing; grid auto-grows to cover your laser bed (default 17 in / 43 cm); "Draw / Capture Shape" button label tracks the machine-integration toggle.
- Sizing Rules Presets — save the entire Sizing Rules dialog as a named preset, load via dropdown, import/export to share with other techs. **(new)** Presets saved before v2.80 simply read as "locator marks off."
- Nesting preview with per-material approval, edge bias d-pad (cardinal + corner directions, smallest pads first in corners), max fill mode (`size x max`). Nesting is quick at every edge-bias setting.
- Custom polygon shapes for irregular leather skins.
- Scrap mode — place pads across multiple irregular pieces, tracking remaining between sheets, with preview, edge bias, and polygons all working together. The pad list stays editable during a session: add a size, change a quantity, or delete a line between scraps and the next scrap picks it up, with everything already cut still credited.
- Filled engraving mode — scan-line raster fill using Roboto outlines, with overscan for clean character edges. Per-material engraving mode (line vs filled).
- Auto-fit engraving — text shifts toward center on small pads, scales only as a last resort.
- G-code options — air assist toggles (M8/M9) per layer, cut grouping (by layer or by pad), full-kerf compensation, configurable return speed, optional SD card eject after export (Windows).
- Last-used library memory for both pad presets and key heights.

## Tooling

- Die inserts: small (50mm OD) and large (70mm OD). **(new)** Generating die G-code no longer ends in an "Unexpected Error" dialog after the file is written — the file was always fine; the message wasn't.
- Die holders: 85mm OD stack. Pick 5-layer or 6-layer; variant Large / Small / Both (Both nests two complete independent holders on one sheet). User-defined sheet size with a clear minimum-size error if pieces don't fit. **(new)** That check now happens before you're asked for a filename, not after.
- Die Organizer — SVG templates for a stackable die organizer (230 × 330 mm). Cut three Uppers and one Lower, align the four 1/8″ corner holes, glue the stack together.
- Pad Press Spacers — bundled 3D-printable STL files for setting pad press depth.
- Kerf Test pattern generator for calibrating your laser.
- **Speed & Power Test (beta)** — generate a sheet of small test discs at different speed / power / passes combinations to dial in laser settings on a new material. Set a **hole diameter** and each test disc becomes a **washer/ring** for shim stock — the inner hole is cut first so the part stays anchored to the sheet, and the disc's ID engraves in the ring. The **sheet size is a soft target**: if your sweep doesn't fit the sheet you entered, the app grows it to fit and shows you the layout rather than erroring out. Engraving feed/power are editable so the labels stay legible while the cut settings are still unknown.
- Per-material G-code presets — the Acrylic and Basswood sections in Options > Tooling Settings carry the same Load / Save / Rename / Delete preset bar as Pad Maker, sharing one library. Acrylic feeds the die holders and inserts; basswood feeds the camera-calibration card engrave. Defaults tuned for the Falcon2 Pro 40W; adjust to your machine.
- Phil Noy's pad-making method is credited at the top of the tab (with a link to noysaxophonesupplies.com) and engraved on every holder ring and die insert.

## Chromatic Strobe Tuner

- 12-wheel stroboscopic chromatic tuner.
- **(new) Tuner accuracy improved.** Low notes used to read a few cents flat; every note now reads within a fraction of a cent, so an in-tune note stands still.
- GPU-accelerated rendering via Rust/wgpu on Windows and Linux — 60–120 fps; automatic CPU fallback if GPU unavailable. **(new)** Big or high-DPI windows no longer knock the tuner back to the slow canvas renderer (a 2048-pixel surface limit was being hit at 150%+ display scaling on a wide window); frames are presented without stalling the rest of the app on every refresh; and if the GPU path ever does fail mid-run, the tab switches to the canvas and says so instead of going dark.
- macOS always uses the canvas renderer — Tk on macOS doesn't expose a native view the GPU renderer can draw into, so Macs are canvas-only by design. Fully functional, just capped at canvas frame rates.
- Per-ring octave brightness from real spectral data.
- Grouped slider panel (display, pitch, bias) and vintage backlit VU meter.
- Per-pitch-class phase tracking with temporal smoothing.
- Transposition support (Concert, Bb, Eb, F).
- Configurable frame rate (60/90/120 fps), backlight color, and faceplate color.

## Harmonic Tone Analyzer (beta)

A real-time harmonic spectrum analyzer for saxophone. Captures the fundamental and overtones of your sound and lets you compare setups (horn, mouthpiece, reed, mic, mic placement, embouchure) over time.

- Live spectrum (FFT) and Bars (per-harmonic) views, linear or dB scale.
- Detects fundamental pitch, extracts up to 20 harmonics. **(new)** Pitch accuracy improved, and some notes that read an octave low now read correctly. Harmonic levels and stored captures are unchanged.
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
- **Analyze tool** — single preset detail, two-preset delta, multi-preset spread analysis. **(new)** Fixed: File > Analyze had failed to open with an "Unexpected Error" since the translated releases.
  - Difference charts and harmonic-range interpretation (H1-H4 ≈ bore, H7-H13 ≈ neck/mpc, broadband ≈ mpc/player).
  - 2D Character Map (warmth × complexity), bars/line chart toggle, click-to-highlight across legend / chart / map.
  - Population percentiles by sax type.
  - Configurable comparison descriptors: complexity, warmth, even/odd, rolloff shape, evenness.
  - Filter by make/model/mic type/search, multi-select, cross-player context notes.
- Recording-quality tracking (rolloff rate) with live warnings and cross-mismatch detection.
- Coverage summary after capture sessions.

## Cross-platform & General

- **(new) Tested on real Macs.** Every release now runs the full test suite on Apple Silicon and Intel macOS (plus Windows and Linux), and the built app is launched and walked through every tab and dialog. Two Mac layout bugs found that way are fixed: clipped tab labels with the Tone Analyzer enabled, and a clipped preset bar in Sizing Rules.
- **Machine integration available on Windows, macOS, and Linux** — pyserial works cross-platform. Off by default everywhere; opt in via File > Feature Set if you have a Grbl machine.
- **Fully translated** — the entire UI is localized into Spanish, German, French, and Italian. Sax-craft terminology (pad / zapatilla / tampon / Polster / tampone; basswood / tilo / tilleul / Lindenholz / tiglio; etc.) kept consistent across locales, and all four catalogs are at 100%.
- **macOS** — dual builds: Apple Silicon (full features) and Intel (no audio features). Native dark/light mode support. Cmd-Q (and the app menu's Quit) saves your settings on the way out. The microphone and camera permission prompts appear correctly as of v2.63 (every earlier Apple Silicon build had a code-signing packaging bug that made macOS silently deny access without ever asking — see Upgrading below if an older build already bit you). The macOS download is ~68 MB rather than the old 258 MB.
- **Linux** — GPU rendering via Vulkan with X11 display handle. Audio features require libportaudio2. Machine integration available.
- **Windows** — Inno Setup installer (Start Menu + uninstaller) plus the standalone .exe; auto-eject removable drives after G-code export.
- Platform-appropriate config storage with automatic migration from old locations (`%APPDATA%`, `~/Library/Application Support`, `~/.config`).
- Reference libraries — Key Heights, Serial Number lookup, and Screw Specs, with one-click import of Matt's published libraries from stohrermusic.com. The Screw Specs tab shows a short getting-started hint when the library is empty.
- Import Settings from Folder — manual transfer between machines.
- Tab-aware User Guide — Help > User Guide shows the section relevant to your current tab. **(new)** Covers Labeled Zones, the Lesser-used settings section, and Leather Locator Marks.
- Error logging — rotating log file at Help > Open Log File for diagnostics. **(new)** The tuner logs which graphics adapter it's using at every start, which helps when a report says "it's slow."
- Feature Set (File > Feature Set) — choose which tabs to show. Toner remains opt-in beta; Tuner is on by default.

## Known limitations

- **(new) Leather Locator Marks are marked beta.** The geometry is tested, but nobody has yet cut leather with them and pressed a pad, and whether a laser line reads clearly on the flesh side depends on your leather and power. Try one sheet first and report what you see. In G-code they engrave at the leather engraving settings; use the SVG output if they need their own power.
- **Labeled Zones are marked beta.** The layout and output are solid and fully tested, but few sheets have been cut with them on real material so far, so treat the first few as a check rather than a batch you're counting on. Please report anything that looks off.
- **Labeled Zones cost material, and that's the trade.** Keeping each size in its own tidy block packs less tightly than letting the nester fill every gap — expect to give up roughly 10–15% of a sheet on the sizes you've zoned. It's aimed at tiny pads, where the yield is high and the demand is low anyway; if you're cutting a sheet full of large pads, leave it off. If a size no longer fits once it's been grouped, the app tells you it couldn't fit everything rather than quietly cutting it in loose — **a size is never cut without its labeled group**, since an unmarked pile of small discs is exactly the problem this is meant to solve. Widen the size range, use a bigger piece, or turn zones off for that job.
- **Labeled Zones are not available in Scrap Mode.** A scrap only takes part of a size's count, so a group would have to be re-sized for every piece. Zones apply when you're cutting a whole batch on one sheet or one traced scrap, which is the case they're built for.
- **Job History records what you entered, not the settings behind it.** Loading a past job restores the pad list, materials, sheet size, center hole, and filename — it deliberately does not touch your sizing rules or G-code settings, and it does not restore a custom polygon shape (a camera-captured scrap has already been cut up, so re-using its outline would put pads in the wrong place). Draw or capture the piece you're cutting now. Jobs can't be loaded while a scrap session is running, since the session owns the pad list.
- **Machine Integration is experimental** and opt-in for a reason. The Falcon2 Pro 40W is the tested machine; other Grbl machines should work but YMMV. "Try Auto Locate" drives the head to the polygon's bottom-left vertex using the camera homography — good enough for rough positioning, but fine-tune with the jog buttons before clicking Start Frame. Disabled until you home the laser in the current session. Stuck-alarm recovery makes a mid-job hiccup recoverable in-app rather than requiring a power cycle, but if a cut ever stalls, re-frame and re-cut the affected scrap.
- **The lid reminder is a reminder, not a sensor.** The app does not detect whether your lid is actually closed — testing on the Falcon2 Pro 40W showed it never reports lid state at all, so there's nothing reliable to read. The prompt is there to make you look, and it appears once per run, before framing.
- **Camera Calibration** requires engraving a ChArUco card on a 12×12 basswood blank. Takes a while; basswood is consumable (each engrave is permanent, so a re-calibration needs a fresh piece).
- **Kerf is a property of your cut, not just the material.** A laser's real kerf shifts with focus height, lens condition, cut speed/power/passes, and even how a material chars — so a profile that cuts at a different speed or pass count than the one you measured will need its own kerf value. If parts start coming out slightly off-size, re-run the Tooling > Kerf Test and update the affected material's `kerf_width`.
- **Tone Analyzer is marked beta.** Descriptors are still being calibrated as we gather more data from different horns, mouthpieces, mics, and players. Raw harmonic measurements are always saved, so future formula improvements apply retroactively to your historical captures.
- **macOS** — the strobe tuner is not GPU-accelerated (canvas renderer only — see the Tuner section above). The app is also not signed with an Apple Developer certificate: right-click → Open the first time, or run `xattr -cr` on the download. Instructions in the README. And because the signature is ad-hoc rather than Apple-issued, macOS may re-ask for mic/camera permission after you update to a new version — that's normal. The automated Mac testing runs on standard-resolution displays; Retina scaling has not been checked.

## Upgrading from v2.75 / earlier

- **This is a drop-in upgrade** — no migrations, no recalibration. Settings, presets, libraries, and camera calibrations carry over untouched.
- **(new) Labeled Zones and Export Settings have moved** to the Lesser-used settings section at the bottom of Options > Sizing Rules. Same controls, same saved values; click the header to open it (it opens itself if either is already on).
- **(new) Your saved Sizing Rules presets still work.** Presets saved before v2.80 read as "locator marks off," and the dropdown still recognises which one you're in.
- **(new) Tuner:** low notes now stop where the pitch actually is; if you'd learned to aim a little sharp of "stopped" on low notes, you no longer need to.
- **(new) Tone Analyzer:** older captures keep their slightly less precise pitch reading; harmonic levels are identical, so comparisons and descriptors are unaffected.
- **(new) Mac users:** the window may open a little wider so all the tab labels fit.
- The large-batch optimization prompt in Scrap Mode is gone (v2.75). If you want the leftover space on a scrap used, add small sizes to the list or an `x max` line.
- **Job History** logs from the moment you installed v2.7 forward; there's no way to reconstruct jobs cut before that. The log lives in `job_history.json` in your config folder alongside your presets.
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
