# Engines: Pad Generation, G-code, Tuner, Tooling

## Pad Generation Data Flow

**Pad Generation**: User input → `parse_pad_list()` → `can_all_pads_fit()` check → `generate_svg()` or `generate_gcode()` → output files

**Sizing Calculations**: `get_disc_diameter()` in svg_engine.py applies material-specific offsets:

- Felt: pad_size - felt_offset
- Card: pad_size - (felt_offset + card_to_felt_offset)
- Leather: pad_size + 2*(felt_thickness + wrap) with star bonus for small pads
- Exact: pad_size unchanged
- Adhesive ring: the **nesting footprint**, not the cut (ring plus an outside number; see "Adhesive Rings"). The cut is the card disc and comes from `adhesive_ring_geometry`

**Per-Size-Range Settings**: Sizing rules, dart/star settings, engraving settings, and engraving placement each support a Universal/Range mode. In Universal mode, one set of values applies to all pad sizes. In Range mode, users define size ranges with different values. Four helper functions in config.py centralize this lookup — engines should always use these instead of reading raw `settings["dart_*"]` etc:

- `get_dart_settings_for_size(pad_size, settings)` → returns range dict or **None** (no match = no stars)
- `get_sizing_for_size(pad_size, settings)` → returns sizing dict (no match = **universal fallback**)
- `get_engraving_settings_for_size(pad_size, settings)` → returns engraving on/off + font sizes (universal fallback)
- `get_engraving_placement_for_size(pad_size, settings)` → returns placement mode/value per material (universal fallback)

Key difference: dart ranges return None on no match (opt-in), while the other three always return a dict (every pad needs values).

Engraving font ranges (`engraving_settings_ranges`) and placement ranges (`engraving_placement_ranges`) are two separate lists with their own Universal / Per Size Range switches. Do not merge them: it breaks every saved preset.

**Darted leather engraving**: Leather pads with darts use the `darted_leather` key in `engraving_location` for placement (default 2.5mm from outside), separate from regular `leather` placement. The dart config's `engraving_on` flag still controls whether darted leather is engraved at all.

## Nesting

**Nesting Algorithm**: `_nest_discs()` implements greedy circle-packing, shared by both `can_all_pads_fit()` and `generate_svg()`. Supports both rectangular sheets and custom polygon shapes. `nest_pads()` is the public API for the preview window. Edge bias controls scan direction: cardinal (N/S/E/W) scan linearly from the edge, corners (NW/NE/SW/SE) use distance-from-corner scoring with smallest-first sort order for efficient corner packing. The bias itself comes from `nest_edge_bias(material, settings)`, the only place it is read: the Pad Maker's setting for pads, a fixed corner for Tooling's die inserts (see Tooling Tab).

- The scan paths are vectorized with numpy (`_scan_radial_numpy`, `_scan_linear_numpy`, `_find_best_polygon_large_numpy`, `_find_best_polygon_small_numpy`), each with a verbatim Python reference (`*_python`) that is the parity target and the fallback when numpy is missing (`_HAS_NUMPY`). Placements must match exactly; `test_nesting_parity` (rectangle, all nine biases, a `< 5 s` bound on a 120-pad batch) and `test_polygon_parity` pin them.
- The generate flows (SVG, G-code, Frame & Cut) nest once with `nest_with_zones()`, show `NestingPreviewWindow`, then write with `generate_svg_from_placed()` / `generate_gcode_from_placed()`. The fit check is the `fixed_placed == fixed_total` pair from that call. Never call `can_all_pads_fit()` next to a `nest_with_zones()` call (it is a second full nest). Frame & Cut passes `proceed_label="Continue →"`, and a cancel there consumes nothing.
- When a "crash" leaves no traceback, check the Windows Application log for `Application Hang` before assuming the exception hook missed it.

History: `docs/decisions/nesting-and-scrap.md`.

## SVG & G-code Rendering

**Shared Rendering Helpers**: SVG rendering is centralized in `_render_svg_discs()` and `_create_svg_drawing()` in svg_engine.py. Both `generate_svg()` and `generate_svg_from_placed()` call these. Similarly, `generate_gcode()` delegates to `generate_gcode_from_placed()` after nesting, so all G-code disc rendering logic lives in one place.

**Coordinate Systems**: SVG uses Y=0 at top (Y increases downward), while G-code uses Y=0 at bottom (Y increases upward). The `gcode_engine.py` functions flip Y coordinates (`sheet_height_mm - cy`) to ensure G-code output matches the SVG preview.

- **Size-label placement** (`engraving_location[material] = {mode, value}`, `darted_leather` its own key; radios out / mid / int). `from_outside`: centre `value` mm inside the cut edge, or, with `darted_out_from_valley` on a darted leather disc, `value` mm from the dart valley (`svg_engine.dart_valley_radius`) to the top of the number. `from_inside`: `value` mm outside the hole edge (or the centre with no hole). `centered` ("mid"): `(r + hole_r) / 2`. Both engines resolve the hole the same way. A dart tip always sits over the label (`dart_wave_raw` peaks at `DART_TIP_ANGLE`, negated in the Y-up G-code). Never change what a saved mode or value means; a new meaning is a new option.
- **`compatibility_mode`** (unitless SVG): both modes share the root's `width`/`height`/`viewBox`. A new SVG element type in `_render_svg_discs()` or `_render_svg_dies()` branches on `compatibility_mode` like the others or emits only the unitless form.
- **Fonts**: `STROKE_FONT` (line mode) and `FILLED_FONT` (Roboto outlines, scan-line fill) carry digits, `.`, `-`, space and D E S I G N B Y P H L O. A new glyph goes into both (`tools/extract_font_outlines.py` for filled). Arc text (`get_text_strokes_arc`, `get_filled_text_strokes_arc`) is shared by the die SVG and G-code.
- `generate_gcode_from_placed()` is the G-code entry point (grouping "layer" or "pad", air per layer). Engraving, then holes, then outer cuts is the only order that matters.

Pinned by `test_engine_parity` and `test_pad_preview`. History: `docs/decisions/engraving-and-darts.md`.

## Polygon Shape Tool

`PolygonDrawWindow` draws custom shapes for irregular skins; vertices are floats, never snapped to the grid, and the grid grows to cover the laser bed. `_nest_discs_polygon()` nests in them. With a calibration, `_render_overlay` warps the live camera under the canvas and "Get from camera" runs `CameraCaptureDialog`. Detail: `docs/decisions/nesting-and-scrap.md`.

## Scrap Mode

Allows users to place pads across multiple irregular scrap pieces instead of requiring one large sheet. Key components:

- Session state in `self.scrap_session` tracks the scrap count, locked material, and (since 2026-09-20) a per-size `done` tally of pads already cut. `original_pads` / `remaining_pads` are refreshed from the box at every Generate, not frozen at session start (see **Live pad list** below)
- `try_nest_partial(pads, ...)` and `compute_remaining_pads()` in svg_engine.py handle partial placement
- `generate_svg_from_placed()` / `generate_gcode_from_placed()` generate output from pre-computed placements
- Progress popup window (`scrap_remaining_window`) shows three columns: Remaining (not yet placed), On this scrap (nested, waiting for the cut), Done (cut). See **Pending scrap** below
- Materials are locked (disabled) during active session to prevent switching
- Files named with `_scrap1`, `_scrap2` suffixes
- After each scrap with a polygon loaded, a dialog asks whether to unload, keep the shape, OR re-capture from camera (when calibrated)

- The pad list box stays editable and is re-read on every Generate and Frame & Cut: `_scrap_pads_for_this_scrap(pads)` subtracts the session's per-size `done` tally. `_scrap_record_result(remaining)` is the one commit point (SVG, G-code and `_frame_cut_scrap_advance` route through it, before anything overwrites `remaining_pads`).
- `pending` (what the current scrap's nest placed) is set by `_scrap_set_pending` and cleared by `_scrap_record_result`; `_scrap_abandon_pending` runs from a `finally` in both scrap generators and `on_frame_and_cut`, so every non-commit exit restores Remaining.
- No ordering search: the large-batch optimizer was removed. Do not reintroduce one without a measured win on a real list.

Tests: `test_scrap_live_list`, `test_scrap_pending`, `test_scrap_partial_nest`. History: `docs/decisions/nesting-and-scrap.md`.

- A `max` line fills the first scrap only (pinned by `test_scrap_live_list`).

## Leather Locator Marks

Opt-in engraved guides on leather showing where the felt sits. One geometry, `svg_engine.locator_mark_geometry` and `locator_circle_points`, translated by `_render_svg_discs`, `gcode_engine._collect_disc_strokes` and `PadPreviewWindow._draw_locator`. Lines start at the felt edge and stop `LOCATOR_EDGE_MARGIN_MM` inside the cut or dart valleys. SVG layer colour `leather_locator`; G-code layer `C06` at the leather engraving settings, after the engraving layer and before any hole or cut. Off must give byte-identical output. Five keys in `SIZING_PRESET_KEYS`, backfilled by `normalize_sizing_preset`.

**Lesser-used settings** (the collapsible at the bottom of Sizing Rules: Labeled Zones, locator marks, Adhesive Rings, Export Settings) opens itself when anything inside is non-default. Put rarely-touched settings there.

Tests: `test_locator_marks`. History: `docs/decisions/engraving-and-darts.md`.

## Adhesive Rings (beta)

Material `adhesive_ring`: a plain ring of tape per pad, no engraving settings, no center hole, no darts.

- `svg_engine.adhesive_ring_geometry(pad_size, settings)` is the only place a ring is computed (outer = the card disc; width = `leather_wrap_for_size` plus `adhesive_ring_trim_mm`, "Wrap clearance" in the dialog; solid up to `ADHESIVE_RING_SOLID_MAX_PAD_MM`). Never change the wrap formula to fix a ring: it moves every leather disc.
- **The radius in `placed` is the nesting footprint, not the cut.** Every renderer takes the cut radius from `adhesive_ring_geometry`; pass `settings=` to the nesting preview.
- The pad size is stroke-font polylines on the cut layer, inside the hole when it fits at 2.5 mm, else under the ring (`adhesive_ring_number_strokes(..., y_up=)` for both outputs). Order per ring: number, inner cut, outer cut.
- Zones, label placement, auto-fit, locator marks and center holes stand down for rings. The G-code settings are cut-only and their defaults are untested: no text may say otherwise.

Tests: `test_adhesive_ring`, `test_adhesive_ring_fuzz` (reads the written files back), `test_adhesive_ring_upgrade`. History: `docs/decisions/adhesive-rings.md`.

## Engraving Auto-Fit

When engraved text (pad size number) would impinge on the disc edge, both engines prefer shifting the text toward the disc center over shrinking it. Only scales down as a last resort when text can't fit even when centered. SVG engine uses bounding-box corner checks; G-code engine uses actual stroke-point measurements. Both maintain a 0.5mm clearance margin. The 80% radius check (`font_size >= r * 0.8`) disables engraving entirely before auto-fit runs.

## SD Card & Eject

- "Eject SD card after G-code export" checkbox (Windows only, below Generate buttons): auto-ejects removable drives after G-code generation. Uses `GetDriveTypeW` to detect removable drives; silently skips non-removable destinations.

There is no "Send G-code to SD Card" menu item.

## Strobe Tuner

12-wheel chromatic stroboscopic tuner. Architecture:

- `tuner_engine.py`: Pure audio/math — FFT-based pitch detection, per-pitch-class phase tracking, per-ring (octave) magnitude extraction. `TunerEngine` manages the sounddevice input stream; `TunerResult` holds magnitudes, phase offsets, ring magnitudes, and cents errors for all 12 pitch classes. `ReferencePlayer` outputs reference tones. Magnitude normalization is gated: `max_mag` must exceed `threshold * 1.5` before normalizing to 0-1, otherwise all magnitudes are zeroed. This prevents sensitive mics from showing wheel activity on room noise.

- Peak frequency comes from `audio_utils.hann_peak_freq` (Hann closed form after climbing to the lobe's top) at both call sites. Do not go back to a parabola on linear magnitudes or add a zero-padded FFT. Gated by "Pitch reading accuracy" in `test_tuner_engine`. JustATuner carries the same code.

- `tuner_tab.py`: All tkinter UI — `StrobeWheel` class renders one disc (annular sector polygons with wedge mask), `TunerTabMixin` builds the tab with control panel: three labeled slider groups (DISP: SENS/BRIGHT/FPS | PITCH: A=/KEY | BIAS: NOTE per-wheel/OCT. per-ring), flat/pilot/sharp indicator, vintage backlit VU meter. The VU needle has damped movement (lerp toward target each frame). Sensitivity gain: the VU path uses a quadratic curve (`sens**2`) so the low end of the slider has fine control; the wheel-brightness path uses a linear curve — they intentionally share the slider but respond differently. Theme walker is bypassed via `_skip_theme` and `_dark_canvas` flags on all dark widgets.
- The tuner auto-starts/stops when switching tabs (`_tuner_start`/`_tuner_stop` called from `on_tab_changed` in main.py).

- numpy and sounddevice are imported separately; `AUDIO_AVAILABLE` means both. `TunerEngine.synthetic_hz` is a test and CI hook only. `_tuner_error_state` redraws the audio-error message after every canvas rebuild.

- macOS is canvas-only: see CLAUDE.md "GPU Tuner Renderer".

- Transposition support: wheel labels and VU readout both apply the shift from TRANSPOSITION_SHIFTS, with octave correction when the shift wraps past C.
- **GPU/Canvas constant alignment**: these constants exist on both sides of the Python/Rust boundary and must stay in sync: `DIM_MULTIPLIER` (`tuner_tab.py` ↔ `tuner_renderer/src/shader.wgsl`), `BRIGHTNESS_GAMMA` (`tuner_tab.py` ↔ `tuner_renderer/src/renderer.rs`, applied in host code, not the shader), `MAGNITUDE_THRESHOLD` (`tuner_tab.py` ↔ `renderer.rs`), `CENTER_GAP_FRACTION`/`CENTER_GAP` (`tuner_tab.py` ↔ `shader.wgsl`), and `RING_SEGMENTS` (`tuner_tab.py` ↔ the `ring_segments()` switch in `shader.wgsl`). If you change any of them, update both sides. Known divergence: `RING_GAP_FRACTION` (Python, 0.05 off the ring's outer edge) vs `RING_GAP` (shader, 0.015 off each side) — the two renderers draw slightly different ring gaps; align these if anyone complains about the look differing between GPU and CPU mode.
- `ReferencePlayer` in tuner_engine.py can output reference tones, but no UI control currently calls it (the play button was removed with the "Experimental" section) — it's exercised only by `tools/test_tuner_engine.py`.

History: `docs/decisions/tuner.md`.

## Audio Stream Health

Both `tuner_engine.py` and `toner_engine.py` import `AudioRingBuffer` from `audio_utils.py` for stream health monitoring. The ring buffer tracks a write counter; if the analysis loop detects no new audio data for ~1 second, the engine automatically restarts the sounddevice stream. This recovers from silent callback death on Windows.

## Tooling Tab

Accordion-style sections. Die inserts: small up to 39.5 mm (50 mm OD), large from 40 mm (70 mm OD). Die holders: 85 mm OD, 5 or 6 layers, Small / Large / Both (two independent holders); `_holder_pieces_for`, `_pack_holder_grid` and `_min_holder_sheet` in svg_engine.py are shared with gcode_engine.py so SVG and G-code stay aligned, and a sheet too small raises `ValueError`. Canvas widgets need explicit handling in the resonance theme walker.

- **Section buttons**: one grid, three columns in one `uniform` group (`tooling-sections`), every button `sticky="ew"`, at least 22 characters or the longest label's need. The window is never narrower than the Tooling tab's content (`_fit_window_to_tabs`, chrome measured once by `_notebook_page_chrome_px()`). Gated per language in `test_release_3_0`. Do not change the open section's sunken look on Aqua without asking.
- **Dies nest in rows from the bottom-left corner**: `svg_engine.nest_edge_bias(material, settings)` is the only reader of edge bias; `die_ring` always gets `DIE_NEST_EDGE_BIAS` ("s"). Nothing writes the shared settings.
- **Zones never reach dies**: `ZONE_FREE_MATERIALS` (adhesive rings, `die_ring`) nest plainly in `nest_with_zones`. Only `compatibility_mode`, `gcode_return_speed`, the filled overscan, a pad material's `gcode_settings` when a test runs on it, and the `die_*` layer colours may move a Tooling output.

**Die Organizer**: A fourth Tooling section. The two source SVGs (`tooling_assets/die_organizer_upper.svg` and `_lower.svg`, Matt's CAD output) are bundled as data assets. Generate SVG copies one byte-for-byte to the user's chosen path, as it always has. `build.py` bundles `tooling_assets/` via `--add-data`.

- **Die Organizer G-code**: `read_die_organizer_drawing` reads only `rect`, `circle` and `ellipse` placed by `translate` or a rotation `matrix`, and raises `OrganizerDrawingError` on anything else; never let it half-read a drawing. Paths are exact parallel curves, openings first, the edge last. Kerf refusals are `OrganizerKerfError`, translated by the tab.

**Pad Press Spacers** are bundled STL files (`pad_press_spacers/*.stl`) copied byte for byte by `ToolingTabMixin._save_pad_spacer_stl`.

- **Tonehole Crescent**: all geometry in svg_engine.py (`tonehole_crescent_geometry`, `layout_tonehole_crescents`, `tonehole_crescent_toolpath`), shared by preview, SVG, STL and G-code; the percentage is of open area. The SVG has no kerf; the G-code path is the exact parallel curve with arc-wrapped horns, never a mitre. With more than one list entry, numbers are engraved under the parts (never on the STL). A thin crescent is never refused for being thin.
- **Speed & Power Test**: G-code only (per-disc settings cannot survive an SVG). Washers cut the hole first, then the outer edge; labels come from `feeds_speeds_label_geometry` for G-code and preview alike. The sheet grows to fit.

Tests: `test_tooling`, `test_tooling_e2e`, `test_die_organizer_gcode`, `test_tonehole_crescent`, `test_tonehole_crescent_gcode`, `test_feeds_speeds_tester`, `test_release_3_0`. History: `docs/decisions/tooling.md`.

## Camera Calibration

`CameraCalibrationDialog` solves intrinsics, distortion and the pixel-to-machine-mm homography in one workflow (engrave the card with the basswood preset, reference captures, "Done with references", moved captures) and saves `homography_px_to_machine_mm`. Rules: read ChArUco output only through `camera_capture.corner_points` and ids through `reshape(-1)` (OpenCV 5 changed both shapes); a per-frame loop logs its first failure with a traceback, then counts; write images with `imencode`, not `cv2.imwrite`; pack the right panel before the preview; `_calibration_inputs_ok` refuses inputs that save a wrong calibration (the RMS does not catch them); a bed-bounds check that cannot be made refuses; ports close in a `finally`; calibration G-code starts with `G92.1`. Tests: `test_calibration_dialog`, `test_camera_capture`, `test_audit_fixes`. History: `docs/decisions/machine-integration.md`.

## Frame & Cut

Frame & Cut streams a job to the laser: nesting preview (Adjust consumes nothing), the jog-only `FalconRunDialog` to put the head on the material's bottom-left corner (Try Auto Locate uses `$J=G90G53`), the lid question on Start Frame, a low-power framing loop (`generate_polygon_framing_gcode`, ending in `M5` only), then the cut behind the same `G92` as the framing.

- **The lid question** (`_remind_lid_closed`, `LidClosedDialog`, "Is the lid closed?", Yes, Start Frame / No, Cancel) is shown at exactly one boundary, Start Frame, opted into with `remind_lid_on_advance=True`. Framing energises the laser, so that is where an open lid alarms the machine; the cut needs no prompt. `test_lid_confirm` asserts one call site. Do not try to detect the lid.
- Scrap Mode runs one scrap per click through `_scrap_begin_partial()`; `_frame_cut_scrap_advance()` commits only when the cut's `_final_reason == "complete"`.

**Outline vs Inset (v2.5)**: camera-captured polygons get split into TWO storage representations sharing one coordinate origin:

- `self.custom_polygon` — the INSET shape (camera_capture.inset_polygon_mm applied, default 3mm), used for pad nesting placement so cuts land safely inside the visible scrap edge.
- `self.custom_polygon_outline` — the ORIGINAL un-inset shape, used for the framing trace so the visible trace matches the actual scrap edge.

- Flip the cut with the outline's height (`flip_height_mm=`), since framing does; flipping with the inset polygon lands the cut `inset` mm off (`test_frame_cut_alignment`). A hand-drawn polygon has no outline; fall back to `custom_polygon`.
- `_keep_last_cut_gcode` keeps the streamed lines in `<config dir>/last_frame_cut.gcode` with one app.log line. Read it before suspecting the engine.
- Between framing passes nothing streams: a Cut or Done then still waits for Idle and ends with `_on_done("complete")`; Stop and Home go through `_stop_stream()`, which finishes the run itself when nothing streams (`test_framing_loop_handoff`).
- Jog-only mode has three close paths that must stay distinct: Cancel and the close box set `_final_reason = "cancelled"`; advance (`_on_advance_clicked`) leaves it "complete". Start Frame unlocks, and Pause/Resume flip, only when `FalconSender` reports the bytes went out (`None` counts as sent).
- Framing power is `config.get_framing_power_s(material, settings)`, per material, clamped on read to `FRAMING_POWER_S_MAX` (30). Defaults stay 10.

**Framing must always use M4, never M3**, whatever the power: M4 drops the beam to zero when the head stalls at a vertex or while the user jogs, M3 holds full S and scorches the material. `tools/test_framing_power.py` asserts it for `generate_polygon_framing_gcode` and `generate_framing_gcode`.

- **Machine menu** (Options > Machine) exists only with the `experimental_machine_menu` toggle and pyserial. Laser items (`MACHINE_MENU_LASER_ITEMS`) need only the toggle; camera items (`MACHINE_MENU_CAMERA_ITEMS`) need a calibration; Camera Calibration is always on. A new item goes in one tuple. Everything that uses the camera stays behind `_camera_capture_ready()`.
- `falcon_sender._stream_loop` opens every stream with `G4 P0.01` (wait for `ok`) and `$X`. Test mocks answer the wake line with `ok`.

History: `docs/decisions/machine-integration.md`.

## The laser watch

`falcon_sender.LaserWatch` (bookkeeping only) run by `MachineControlMixin._start_laser_watch` (machine_control.py) finds a laser plugged in after launch and loses one unplugged.

- `_start_laser_watch` is called from the `__main__` block only, never from `__init__`, `_selftest` or `run_tour`. `_detect_falcon_async()` is a no-op when no watch exists, so building the app probes nothing.
- `probe_grbl_port` asks `?` first and accepts any status report; only a silent port gets Ctrl-X and the banner wait.
- **Never probe while a laser is known**, so nothing disturbs a running cut. A known port that disappears is lost: `falcon_port = None`, `_falcon_homed_this_session = False`, `_refresh_machine_ui_state()`.
- A new port is probed once, a laser-looking one again every `LIKELY_RETRY_S`, Bluetooth only in a user's rescan. A failed connect asks for a rescan and is never retried automatically. Every state change happens on the Tk thread (`_laser_watch_apply`); log once per event, never per tick.

Tests: `test_laser_watch`. History: `docs/decisions/machine-integration.md`.

## The camera's saved number

Every camera window opens its camera through `camera_capture.open_camera_or_fallback(index, saved_index, ...)`.

- If the saved number fails, the resolver is tried once more and the window runs on what opened. Nothing is saved: only a calibration save or Switch camera changes `camera_index_override`. An explicit, non-saved number never falls back.
- **With a calibration on disk the only stand-in is the camera `find_falcon_camera_index()` identifies**; a calibration belongs to one camera. An unreadable calibration file counts as present.
- A fallback is never silent: the status line shows `fallback_notice(saved, used)` and one warning is logged.
- The app never competes with itself: `MachineControlMixin._release_live_camera()` closes `LiveCameraWindow` before the shape dialog and Camera Calibration open a camera.

Tests: `test_camera_fallback`, `test_calibration_dialog`. History: `docs/decisions/machine-integration.md`.

## Tooling: Frame & Cut

Die Inserts (and its Scrap Mode), Die Holders, Die Organizer, Kerf Test, Speed & Power Test and Tonehole Crescent each have Frame & Cut beside Generate G-code.

- One routine: `MachineControlMixin.on_frame_and_cut(tooling_job=None)` (machine_control.py). A Tooling handler passes a function returning `{'gcode_text', 'material', 'label', 'count'}` (optional `'scrap_num'`, `'after_cut'`) or None. The laser check comes first.
- The machine half reads the job's own values, never the Pad Maker's state. A job not nested on the loaded shape (every Tooling job; card on "Fit card to paper") has `on_shape` False: no framing polygon, no flip override, `G92 X0 Y0`, no Auto Locate target.
- Everything after a successful `connect()` is inside the `try` whose `finally` disconnects; the sleep lock is taken after the last early return.
- **`on_frame_and_cut`'s source is read by `test_lid_confirm`, `test_frame_cut_alignment` and `test_framing_power`.** If the method is ever split, those source checks move with it.
- **The stream is the file**: each Generate G-code handler and its Frame & Cut twin share validation and generator; Frame & Cut streams a temp file it reads back and deletes (`_gcode_text_from`), and writes nothing but the Speed & Power legend.
- Only `NoCutSettingsError` earns the "no cut speed and power set" message. Map translated material names with `_tooling_material_key` / `_tooling_material_label`. Die Scrap Mode counts a sheet only when its file is written or its cut ended "complete".

Tests: `test_tooling_frame_cut`, `test_tooling_e2e`, `test_release_3_0`. History and the five older bugs fixed with this work: `docs/decisions/machine-integration.md`.

- The Speed & Power legend is asked for before the laser is touched, and a cancelled prompt cancels the run.

## Cork and "engraving not set"

`cork` is a G-code material in `GCODE_PRESET_MATERIALS` and a Tooling material only (`config.TOOLING_GCODE_MATERIALS`). Its cut default is Matt's tested value; its kerf is not measured (no text may call it tested); its engraving ships at 0 and 0 on purpose.

- A speed or power of 0 means "not set" (`config.engraving_is_set`), for any material. No generator may write `F0` or engrave at `S0`: leave the engraving out and write every cut.
- Never silently: `_engraving_not_set_notice` in the Done box; Frame & Cut asks first (`_confirm_engraving_skipped`). The Speed & Power Test accepts 0 in its engraving fields.

Tests: `test_cork_material`. History: `docs/decisions/tooling.md`.

## The G-code file line

Every box shown after G-code is written to a file ends with "Send this file to the laser as it is. For LightBurn, use Generate SVG.", written only in `ui_dialogs.gcode_file_note()`. Never on an SVG or STL box, never on Frame & Cut. The scrap continue dialog takes it as a fourth positional flag (suites stand in with `lambda *args`). Tests: `test_gcode_file_note`. History: `docs/decisions/tooling.md`.

## Phil Noy Credit (non-negotiable)

The die holders and die inserts implement Phil Noy's pad-making method, which Phil shared freely. The Tooling tab has a paragraph at the top crediting Phil with a clickable link to noysaxophonesupplies.com. Every die holder retaining ring is engraved "DESIGNED BY PHIL NOY" arced along the top of the annulus (replacing the previous redundant size-range label). Every die insert is engraved "NOY" arced along the bottom of the ring opposite the size number. Engravings respect the user's filled/line `engraving_mode` so they match the visual style of the size labels. **Do not modify, hide, shrink, or remove this credit** — it was missing from an earlier version, Phil was upset about it, and Matt re-added it 2026-04-07 to do right by him.
