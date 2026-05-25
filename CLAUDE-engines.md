# Engines: Pad Generation, G-code, Tuner, Tooling

## Pad Generation Data Flow

**Pad Generation**: User input → `parse_pad_list()` → `can_all_pads_fit()` check → `generate_svg()` or `generate_gcode()` → output files

**Sizing Calculations**: `get_disc_diameter()` in svg_engine.py applies material-specific offsets:
- Felt: pad_size - felt_offset
- Card: pad_size - (felt_offset + card_to_felt_offset)
- Leather: pad_size + 2*(felt_thickness + wrap) with star bonus for small pads
- Exact: pad_size unchanged

**Per-Size-Range Settings**: Sizing rules, dart/star settings, engraving settings, and engraving placement each support a Universal/Range mode. In Universal mode, one set of values applies to all pad sizes. In Range mode, users define size ranges with different values. Four helper functions in config.py centralize this lookup — engines should always use these instead of reading raw `settings["dart_*"]` etc:
- `get_dart_settings_for_size(pad_size, settings)` → returns range dict or **None** (no match = no stars)
- `get_sizing_for_size(pad_size, settings)` → returns sizing dict (no match = **universal fallback**)
- `get_engraving_settings_for_size(pad_size, settings)` → returns engraving on/off + font sizes (universal fallback)
- `get_engraving_placement_for_size(pad_size, settings)` → returns placement mode/value per material (universal fallback)

Key difference: dart ranges return None on no match (opt-in), while the other three always return a dict (every pad needs values).

**Darted leather engraving**: Leather pads with darts use the `darted_leather` key in `engraving_location` for placement (default 2.5mm from outside), separate from regular `leather` placement. The dart config's `engraving_on` flag still controls whether darted leather is engraved at all.

## Nesting

**Nesting Algorithm**: `_nest_discs()` implements greedy circle-packing, shared by both `can_all_pads_fit()` and `generate_svg()`. Supports both rectangular sheets and custom polygon shapes. `nest_pads()` is the public API for the preview window. Edge bias controls scan direction: cardinal (N/S/E/W) scan linearly from the edge, corners (NW/NE/SW/SE) use distance-from-corner scoring with smallest-first sort order for efficient corner packing.

**Vectorized scan paths (v2.40+)**: the radial-bias rectangle scan (`_scan_radial_numpy`) and both polygon scans (`_find_best_polygon_large_numpy`, `_find_best_polygon_small_numpy`) build a numpy candidate grid, compute distances/scores in one bulk op, mask occupied cells via per-pad bounding-box updates, then `argmin` to pick the closest valid cell. ~100× faster than the Python loop on real-world inputs; bit-identical placements (`argmin`'s first-index-of-min matches the Python reference's strict-less-than y-then-x tie-break). A Python reference (`_scan_radial_python`, `_find_best_polygon_*_python`) is preserved verbatim as the parity target and as the runtime fallback when numpy is unavailable (macOS Intel build — try-import at the top of `svg_engine.py` sets `_HAS_NUMPY`). `tools/test_nesting_parity.py` (rectangle) and `tools/test_polygon_parity.py` (polygon) assert exact-match placements across all 9 edge biases, every polygon shape, and max-quantity fills.

**Nesting Preview**: `NestingPreviewWindow` in ui_dialogs.py shows the nested layout before file generation. Works in standard mode (one material at a time) and scrap mode. The generate flow nests first via `nest_pads()`, shows the preview, then writes files via `generate_svg_from_placed()` / `generate_gcode_from_placed()` — no double-nesting.

## SVG & G-code Rendering

**Shared Rendering Helpers**: SVG rendering is centralized in `_render_svg_discs()` and `_create_svg_drawing()` in svg_engine.py. Both `generate_svg()` and `generate_svg_from_placed()` call these. Similarly, `generate_gcode()` delegates to `generate_gcode_from_placed()` after nesting, so all G-code disc rendering logic lives in one place.

**SVG Output Modes (`compatibility_mode`)**: Options > Settings > Export Settings has an "Enable Inkscape/Compatibility Mode (unitless SVG)" toggle (setting key: `compatibility_mode`, default False). Both modes emit the same SVG root (`width="Xmm" height="Ymm" viewBox="0 0 X Y"`), so physical dimensions are identical either way — the difference is inside the document. Default mode declares `profile='tiny'` and writes every child attribute with explicit `mm` suffixes (e.g. `r="12.5mm"`, `stroke_width="0.1mm"`); compatibility mode omits the profile and writes bare numbers (`r=12.5`, `stroke_width=0.1`) that inherit the viewBox's 1-unit-per-mm mapping. The flag gates four branches each in `_render_svg_discs()` and `_render_svg_dies()` (outlines, holes, engraving text, and die-specific elements). When adding new SVG element types to either renderer, remember to branch on `compatibility_mode` the same way or emit only the unitless form. Historical context: the flag exists for older tool chains that rejected per-attribute unit suffixes or SVG Tiny 1.2; modern Inkscape/LightBurn/Illustrator handle both forms fine.

**Coordinate Systems**: SVG uses Y=0 at top (Y increases downward), while G-code uses Y=0 at bottom (Y increases upward). The `gcode_engine.py` functions flip Y coordinates (`sheet_height_mm - cy`) to ensure G-code output matches the SVG preview.

## G-code Engine Details

`gcode_engine.py` contains two font systems: `STROKE_FONT` (single-stroke outlines for "line" engraving mode) and `FILLED_FONT` (Roboto glyph outlines for "filled" raster mode). Both fonts contain digits, period, hyphen, space, and 12 letters (D E S I G N B Y P H L O) — the letters cover the Phil Noy credit text engraved on tooling parts. Stroke letters are hand-designed single-stroke approximations matching the digit style; filled letters are extracted from `tools/Roboto-Regular.ttf` via `tools/extract_font_outlines.py`. To regenerate the filled letters, edit `CHARS` in the extractor and re-run; paste the output into the `FILLED_FONT` and `FILLED_CHAR_WIDTHS` blocks in gcode_engine.py.

Filled engraving uses scan-line fill with even-odd rule to convert font outlines into horizontal raster lines, with optional overscan (extending lines beyond character edges so the laser reaches full speed). `generate_gcode_from_placed()` is the main entry point — it supports two cut grouping modes ("layer": all engravings → all holes → all cuts; "pad": complete each disc before moving to the next) and per-layer air assist control (M8 on / M9 off). The internal `_collect_disc_strokes()` helper extracts per-disc stroke data for both modes.

**Arc text helpers**: `get_text_strokes_arc()` and `get_filled_text_strokes_arc()` lay text out along an arc with rigid per-character rotation (each character is a rotation+translation, not a polar warp). The filled variant runs scan-line fill IN THE CHARACTER'S LOCAL FRAME and then transforms each scan segment to world coordinates via the basis vectors, so curved text fills solid like flat text instead of using horizontal scan lines that would misalign on rotated glyphs. Both helpers take `(text, font_size_mm, cx, cy, radius, center_angle_rad, side='top'|'bottom')` where 'top' sweeps CW with character tops radially outward and 'bottom' sweeps CCW with character tops radially inward (so text reads upright when viewed from outside the circle). The svg_engine die/holder renderers call the same helpers and emit polylines so the SVG preview matches the gcode output exactly.

## Polygon Shape Tool

Users can draw custom polygon shapes for irregular leather skins. Grid size adapts to unit setting: 15x15 inches (1" squares) or 40x40 cm (1cm squares). The polygon nesting algorithm (`_nest_discs_polygon()`) uses ray-casting for point-in-polygon checks and distance-to-edge calculations for circle fitting. The rectangle algorithm remains the fast path when no custom shape is defined.

## Scrap Mode

Allows users to place pads across multiple irregular scrap pieces instead of requiring one large sheet. Key components:
- Session state in `self.scrap_session` tracks original pads, remaining pads, scrap count, locked material
- `try_nest_partial()` and `compute_remaining_pads()` in svg_engine.py handle partial placement
- `generate_svg_from_placed()` / `generate_gcode_from_placed()` generate output from pre-computed placements
- Progress popup window (`scrap_remaining_window`) shows two columns: Remaining and Done
- Materials are locked (disabled) during active session to prevent switching
- Files named with `_scrap1`, `_scrap2` suffixes
- After each scrap with a polygon loaded, a dialog asks whether to unload or keep the shape for the next piece

## Engraving Auto-Fit

When engraved text (pad size number) would impinge on the disc edge, both engines prefer shifting the text toward the disc center over shrinking it. Only scales down as a last resort when text can't fit even when centered. SVG engine uses bounding-box corner checks; G-code engine uses actual stroke-point measurements. Both maintain a 0.5mm clearance margin. The 80% radius check (`font_size >= r * 0.8`) disables engraving entirely before auto-fit runs.

## SD Card & Eject

Two mechanisms:
- "Eject SD card after G-code export" checkbox (Windows only, below Generate buttons): auto-ejects removable drives after G-code generation. Uses `GetDriveTypeW` to detect removable drives; silently skips non-removable destinations.
- File → Send G-code to SD Card: guided workflow that copies a .gcode file, optionally clears old files, and ejects (Windows only via PowerShell COM object).

## Strobe Tuner

12-wheel chromatic stroboscopic tuner. Architecture:
- `tuner_engine.py`: Pure audio/math — FFT-based pitch detection, per-pitch-class phase tracking, per-ring (octave) magnitude extraction. `TunerEngine` manages the sounddevice input stream; `TunerResult` holds magnitudes, phase offsets, ring magnitudes, and cents errors for all 12 pitch classes. `ReferencePlayer` outputs reference tones. Magnitude normalization is gated: `max_mag` must exceed `threshold * 1.5` before normalizing to 0-1, otherwise all magnitudes are zeroed. This prevents sensitive mics from showing wheel activity on room noise.
- `tuner_tab.py`: All tkinter UI — `StrobeWheel` class renders one disc (annular sector polygons with wedge mask), `TunerTabMixin` builds the tab with control panel: three labeled slider groups (DISP: SENS/BRIGHT/FPS | PITCH: A=/KEY | BIAS: NOTE per-wheel/OCT. per-ring), flat/pilot/sharp indicator, vintage backlit VU meter. The VU needle has damped movement (lerp toward target each frame). Sensitivity uses a quadratic gain curve (`sens**2`) so the low end of the slider has fine control. Theme walker is bypassed via `_skip_theme` and `_dark_canvas` flags on all dark widgets.
- The tuner auto-starts/stops when switching tabs (`_tuner_start`/`_tuner_stop` called from `on_tab_changed` in main.py).
- Transposition support: wheel labels and VU readout both apply the shift from TRANSPOSITION_SHIFTS, with octave correction when the shift wraps past C.
- **GPU/Canvas constant alignment**: `DIM_MULTIPLIER` exists in two places that must stay in sync: `tuner_tab.py` (Python canvas path) and `tuner_renderer/src/shader.wgsl` (GPU shader). `BRIGHTNESS_GAMMA` exists in two places: `tuner_tab.py` and `tuner_renderer/src/renderer.rs` (applied in host code, not the shader). If you change either constant, update all locations.

## Audio Stream Health

Both `tuner_engine.py` and `toner_engine.py` import `AudioRingBuffer` from `audio_utils.py` for stream health monitoring. The ring buffer tracks a write counter; if the analysis loop detects no new audio data for ~1 second, the engine automatically restarts the sounddevice stream. This recovers from silent callback death on Windows.

## Tooling Tab

Accordion-style UI with die insert and die holder SVG/G-code generation. Small dies ≤39.5mm (50mm OD), large ≥40mm (70mm OD). Die holders 85mm OD; user picks 5-layer (solid + magnet + 2× pin + ring) or 6-layer (solid + magnet + 3× pin + ring), with magnet hole 6.5mm, pin holes 3.5mm. Variant is Small / Large / Both, where Both = two complete independent holders (one of each size) nested onto the same sheet — the layers are cemented together permanently, so there's no convertible-holder concept. User defines sheet size (W × H + in/mm) like Die Inserts; engine raises `ValueError` if pieces don't fit. Helpers `_holder_pieces_for(variant, layer_count)`, `_pack_holder_grid(...)`, and `_min_holder_sheet(...)` live in `svg_engine.py` and are imported by `gcode_engine.py` so SVG and G-code stay aligned. Canvas widgets need explicit handling in the resonance theme walker.

**Die Organizer (SVG-only)**: A fourth Tooling section. The two source SVGs (`tooling_assets/die_organizer_upper.svg` and `_lower.svg`, Matt's CAD output) are bundled as data assets and copied byte-for-byte to the user's chosen path on Generate. No G-code path — instructions in the section direct the user to LightBurn or similar (LightBurn handles kerf comp, layer mapping, and material profiles better than we would, and the organizer is a one-and-done part). `generate_die_organizer_svg(variant, filename, settings)` in `svg_engine.py` resolves the asset via `sys._MEIPASS` when frozen and `__file__` otherwise. `build.py` bundles `tooling_assets/` via `--add-data`. To update the design, just drop a replacement SVG into `tooling_assets/` — no parser to update.

**Pad Press Spacers**: bundled 3D-printable STL files (`pad_press_spacers/*.stl`) copied byte-for-byte to the user's chosen path via `ToolingTabMixin._save_pad_spacer_stl`. Three sets: half-step (3.0/3.5/4.0/4.5 mm × 4 each, 16 total), quarter-step (3.25/3.75/4.25 mm × 4 each, 12 total), and an organizer rack with 7 compartments. No engine code — they're static assets bundled via `build.py --add-data`.

**Camera Calibration (v3.0+)**: integrated single-pass calibration that solves camera intrinsics + lens distortion + pixel-to-MACHINE-mm homography in one workflow. `CameraCalibrationDialog` (ui_dialogs.py) opens with a horizontal-split layout (preview left, controls right) showing two phases. Phase 1 (engrave): live camera + jog cluster + MPos display + Home Laser button + Frame (continuous-loop bbox tracer, follows jogs, no Pause/Cut buttons — those break absolute-coord framing) + Engrave button (commits the ~45-60 min 10×10×25mm card engrave at machine `(head_x − card_w/2, head_y − card_h/2)`). Bed-bounds safety check refuses to engrave/frame if the card would extend off the bed. Phase 2 (capture): live preview + multi-pose ChArUco captures with explicit reference / intrinsics substates — the user takes 1+ reference captures (card untouched at engraved position, pooled in `calibrate_from_frames` for the machine-mm homography) then clicks "Done with references" and moves the card for the remaining intrinsics captures. Retake-last + Reset-all recovery buttons let users iterate on bad detections without losing other captures. After save, calibration JSON has `homography_px_to_machine_mm` (replaces the old `homography_px_to_mm` which was board-frame). `pixels_to_mm` returns machine coords directly. `is_legacy_calibration` detects pre-v3.0 files; `load_calibration` returns None on those (forces recalibration).

**Frame & Cut (v3.0+)**: Pad Maker > Frame & Cut button. Only appears when both a Falcon is detected AND a saved calibration exists (`_detect_falcon_async` checks `camera_capture.load_calibration`). On click: auto-homes the laser once per session (`_falcon_homed_this_session` flag, reset by Machine > Reset Falcon). If the loaded polygon has a saved `custom_polygon_machine_offset` (camera-captured with calibration), user picks Auto (head drives to scrap's known machine position) or Manual (user jogs head to material's bottom-left). Without offset, Manual is the only path. Manual uses `G92 X0 Y0` at current head position; Auto uses `G0 X{offset_x} Y{offset_y} + G92 X0 Y0`. Framing pass uses `generate_polygon_framing_gcode` (polygon outline) for irregular scrap; falls back to `generate_framing_gcode` (bbox rectangle) for rectangular jobs. Framing supports continuous-loop mode (FalconRunDialog `loop=True` + optional `gcode_provider` callback that regenerates the trace from current MPos each iteration so jogs follow the head). Live camera preview checkbox spawns a non-modal `LiveCameraWindow` alongside the cut.

**Machine menu (Pad Maker > Options > Machine)**: groups Falcon control and camera settings — Home Laser, Test Connection, Clear Errors (`$X`), Reset Falcon (Ctrl-X soft-reset), Camera Calibration, Auto-Framing Inset Margin. The inset margin (`camera_polygon_inset_mm`, default 3mm) shrinks camera-captured polygons before nesting as a safety buffer against camera measurement error at the edges of its view.

**Scrap mode + camera capture**: the scrap-continue dialog (`_show_scrap_continue_dialog`) offers a "Re-capture from camera" shortcut alongside Unload/Keep Shape when a calibration exists, streamlining the per-scrap workflow (each scrap is a different piece, so the user typically wants a fresh polygon for each). Camera capture is consolidated into the single "Draw / Capture Shape" dialog — the duplicate main-pane "Get from camera" button was removed.

**Speed & Power Test (v2.40+, beta)**: laser-settings calibration tool. Generates a sheet of small test discs at user-defined sweeps of speed / power / passes / air-assist; each disc gets a 2-digit ID and a `legend.txt` mapping IDs to parameters. Engine entry point `generate_feeds_speeds_test_gcode` in `gcode_engine.py`. Matrix expansion via `build_feeds_speeds_matrix` and grid packing via `_grid_pack_discs` (both in `svg_engine.py`), wired through `ToolingTabMixin._prepare_feeds_speeds_pieces`. Preview window `SpeedPowerTestPreview` in `ui_dialogs.py` shows the layout (color-coded by air state when "Also test with air off" is on) before the file dialog; toggleable via the "Preview before saving" checkbox. G-code only — per-disc parameters can't survive an SVG round-trip, so the SVG path was deliberately not added. Engraving feed/power are exposed as editable fields so the labels stay legible when the cut settings are still unknown (the whole point of the tool).

## Phil Noy Credit (non-negotiable)

The die holders and die inserts implement Phil Noy's pad-making method, which Phil shared freely. The Tooling tab has a paragraph at the top crediting Phil with a clickable link to noysaxophonesupplies.com. Every die holder retaining ring is engraved "DESIGNED BY PHIL NOY" arced along the top of the annulus (replacing the previous redundant size-range label). Every die insert is engraved "NOY" arced along the bottom of the ring opposite the size number. Engravings respect the user's filled/line `engraving_mode` so they match the visual style of the size labels. **Do not modify, hide, shrink, or remove this credit** — it was missing from an earlier version, Phil was upset about it, and Matt re-added it 2026-04-07 to do right by him.
