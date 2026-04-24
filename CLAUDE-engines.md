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

Accordion-style UI with die insert and die holder SVG/G-code generation. Small dies ≤39.5mm (50mm OD), large ≥40mm (70mm OD). Die holders 85mm OD, 6-layer stack (solid base, magnet disc with 6.5mm hole, 3x pin discs with 3.5mm holes, retaining ring). Canvas widgets need explicit handling in the resonance theme walker.

## Phil Noy Credit (non-negotiable)

The die holders and die inserts implement Phil Noy's pad-making method, which Phil shared freely. The Tooling tab has a paragraph at the top crediting Phil with a clickable link to noysaxophonesupplies.com. Every die holder retaining ring is engraved "DESIGNED BY PHIL NOY" arced along the top of the annulus (replacing the previous redundant size-range label). Every die insert is engraved "NOY" arced along the bottom of the ring opposite the size number. Engravings respect the user's filled/line `engraving_mode` so they match the visual style of the size labels. **Do not modify, hide, shrink, or remove this credit** — it was missing from an earlier version, Phil was upset about it, and Matt re-added it 2026-04-07 to do right by him.
