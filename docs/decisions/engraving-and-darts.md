# Engraving, darts and locator marks: decisions and history

This file holds the why and the history behind the rules in the CLAUDE files, and the full detail they summarise: what was tried, measured, rejected and decided, with dates and the user's own words. The paragraphs were moved here verbatim from the CLAUDE files on 2026-11-10 (HEAD 01071f5); each section names where it came from. Add new decisions to the matching section, dated.

## Dart Shape Spectrum

*Moved from CLAUDE.md, section "CLAUDE.md > Dart Shape Spectrum" (2026-11-10).*

The dart wave shape (formerly "Star/Dart"; now just "Darts" in the UI) is a single 0.0–1.0 slider that smoothly interpolates between three primitive shapes:

- 0.0 = Triangle (linear ramps between peaks/valleys)

- 0.5 = Sine (raw cosine)

- 1.0 = Square (saturating sign function via `|c|^p`, `p = _SQUARE_POWER = 0.01`)

Math lives in `_wave_value` in `svg_engine.py`; `calculate_star_path` (SVG) and `_generate_star_points` (G-code) both call it per sample — `tools/test_engine_parity.py` pins them together. The default value is now `0.5` (sine). Legacy configs used a 0.0=sine, 1.0=square scale; `load_settings` migrates them once via `0.5 + 0.5 * old` and sets `dart_shape_v2: True` so the migration doesn't run again. `dart_ranges[*].shape_factor` is migrated alongside the universal value. When changing dart-shape behavior, update `tools/test_dart_shapes.py` (anchor + smoothness + migration coverage).

Internal variable names retain the `dart_` prefix (`dart_shape_factor`, `dart_threshold`, etc.); only the user-facing labels were renamed to "Darts".

## Pad Generation Data Flow

*Moved from CLAUDE-engines.md, section "Engines: Pad Generation, G-code, Tuner, Tooling > Pad Generation Data Flow" (2026-11-10).*

### Engraving font ranges and engraving placement ranges are two separate lists

**Engraving font ranges and engraving placement ranges are two separate lists** (`engraving_settings_ranges` vs `engraving_placement_ranges`), each with its own Universal / Per Size Range switch in the Sizing Rules dialog, one section apart. That split is deliberate (the boundaries need not match, and both are in `SIZING_PRESET_KEYS`), but it hides the second one: a user set up font-size ranges, found no placement fields there, and asked for placement to be added to ranges on 2026-09-18 when it already existed a section down. A one-line pointer in the font-size range editor was added and removed again two days later: Matt looked at the dialog and judged the placement section obvious as it stands. Don't merge the two lists to "fix" this — it's a preset-schema change that breaks every saved preset.

## SVG & G-code Rendering

*Moved from CLAUDE-engines.md, section "Engines: Pad Generation, G-code, Tuner, Tooling > SVG & G-code Rendering" (2026-11-10).*

### Size-label placement modes

**Size-label placement modes** (`engraving_location[material] = {mode, value}`, per material, with `darted_leather` as its own key; Sizing Rules dialog radios **out / mid / int**, outer edge to centre, reordered from out / int / mid on 2026-10-28): `from_outside` puts the label's centre `value` mm inside the cut edge, or, on a darted leather disc with the **"Darted leather: measure 'out' from the dart valley"** box ticked (`settings["darted_out_from_valley"]`, in `SIZING_PRESET_KEYS`, off by default), `value` mm between the **dart valley** (the star's inner radius, `svg_engine.dart_valley_radius`, now the one place both engines get it from) and the **top of the number**, i.e. the label's centre sits `value + font_size / 2` inside the valley and 0 means touching. Matt, 2026-10-26: the valley deepens with pad size, so measured from the cut edge no single value kept the label just inside it and that placement had been a running nuisance; it is a checkbox rather than a silent redefinition so nobody's existing presets move, and his own preset (`presets/Matt mode.json`, 0.5 in every range) turns it on. It first measured to the label's middle like every other mode, so 0 put the number half into the darts; Matt asked for the near edge the same day, for this option only, because it was unreleased (changing `from_outside` from the cut edge or `from_inside` would move every saved preset's labels by half a font height). The half-height is the nominal text box (`font_size / 2`), so the burned digits sit a hair further in: 0.05 × font for the stroke font, ~0.01 for filled, more for LightBurn's own font in the SVG. **A dart tip always sits over the label**: `svg_engine.dart_wave_raw` peaks at `DART_TIP_ANGLE` (12 o'clock, Y-down) in the SVG path and the pad preview, and at its negation in the Y-up G-code, so the number gets a tip's room above it. Before, the wave peaked at 3 o'clock, which put a notch over the label on every size whose dart count is 2 mod 4 (10 and 12 under Matt mode); the cut is otherwise the same shape, turned half a dart. `test_engine_parity` pins the tip on both engines across sizes of both counts, plus the G-code star as the SVG star's mirror image; `test_pad_preview` pins the preview; `from_inside` puts it `value` mm outside the hole edge (outside the centre when the pad gets no hole); `centered` — labelled **mid** since 2026-10-26, previously "ctr" — puts it halfway between the hole edge and the cut edge, `(r + hole_r) / 2` from the centre, or halfway between the centre and the cut edge when the pad gets no hole (`hole_r = 0`); its mm value is ignored, and the dialog greys that row's entry out while mid is selected (`grey_value_when_mid`, a trace on the mode var, so preset loads and range switches follow; Matt, 2026-10-28). Mid measures from the cut edge on darted pads too (the dart tips), whether or not the valley box is ticked: that box only changes "out". The settings key stays `centered` so saved presets keep working. Matt asked for exactly that midpoint on 2026-10-26, not knowing "ctr" was nearly it: until then a no-hole pad used a phantom 1.75 mm hole radius, and the G-code engine placed int/mid labels from the *preset* hole size even on pads below `min_hole_size`, so small no-hole pads got different label positions in SVG and G-code. Both engines now resolve the disc's hole the same way (`should_have_center_hole` in SVG, `disc_hole_dia` in `_collect_disc_strokes`) before placing the label; `tools/test_engine_parity.py` pins the exact offsets for hole present / hole option None / pad below the hole size.

### SVG Output Modes (`compatibility_mode`)

**SVG Output Modes (`compatibility_mode`)**: Options > Settings > Export Settings has an "Enable Inkscape/Compatibility Mode (unitless SVG)" toggle (setting key: `compatibility_mode`, default False). Both modes emit the same SVG root (`width="Xmm" height="Ymm" viewBox="0 0 X Y"`), so physical dimensions are identical either way — the difference is inside the document. Default mode declares `profile='tiny'` and writes every child attribute with explicit `mm` suffixes (e.g. `r="12.5mm"`, `stroke_width="0.1mm"`); compatibility mode omits the profile and writes bare numbers (`r=12.5`, `stroke_width=0.1`) that inherit the viewBox's 1-unit-per-mm mapping. The flag gates a branch per unit-suffixed element type in `_render_svg_discs()` (outline, hole, engraving text — the dart path is always unitless, which is fine) and `_render_svg_dies()` (plus die-specific elements). When adding new SVG element types to either renderer, remember to branch on `compatibility_mode` the same way or emit only the unitless form. Historical context: the flag exists for older tool chains that rejected per-attribute unit suffixes or SVG Tiny 1.2; modern Inkscape/LightBurn/Illustrator handle both forms fine.

## G-code Engine Details

*Moved from CLAUDE-engines.md, section "Engines: Pad Generation, G-code, Tuner, Tooling > G-code Engine Details" (2026-11-10).*

`gcode_engine.py` contains two font systems: `STROKE_FONT` (single-stroke outlines for "line" engraving mode) and `FILLED_FONT` (Roboto glyph outlines for "filled" raster mode). Both fonts contain digits, period, hyphen, space, and 12 letters (D E S I G N B Y P H L O) — the letters cover the Phil Noy credit text engraved on tooling parts. Stroke letters are hand-designed single-stroke approximations matching the digit style; filled letters are extracted from `tools/Roboto-Regular.ttf` via `tools/extract_font_outlines.py`. To regenerate the filled letters, edit `CHARS` in the extractor and re-run; paste the output into the `FILLED_FONT` and `FILLED_CHAR_WIDTHS` blocks in gcode_engine.py.

Filled engraving uses scan-line fill with even-odd rule to convert font outlines into horizontal raster lines, with optional overscan (extending lines beyond character edges so the laser reaches full speed). `generate_gcode_from_placed()` is the main entry point — it supports two cut grouping modes ("layer": all engravings → all holes → all cuts; "pad": complete each disc before moving to the next) and per-layer air assist control (M8 on / M9 off). The internal `_collect_disc_strokes()` helper extracts per-disc stroke data for both modes.

### Arc text helpers

**Arc text helpers**: `get_text_strokes_arc()` and `get_filled_text_strokes_arc()` lay text out along an arc with rigid per-character rotation (each character is a rotation+translation, not a polar warp). The filled variant runs scan-line fill IN THE CHARACTER'S LOCAL FRAME and then transforms each scan segment to world coordinates via the basis vectors, so curved text fills solid like flat text instead of using horizontal scan lines that would misalign on rotated glyphs. Both helpers take `(text, font_size_mm, cx, cy, radius, center_angle_rad, side='top'|'bottom')` where 'top' sweeps CW with character tops radially outward and 'bottom' sweeps CCW with character tops radially inward (so text reads upright when viewed from outside the circle). The svg_engine die/holder renderers call the same helpers and emit polylines so the SVG preview matches the gcode output exactly.

## Leather Locator Marks

*Moved from CLAUDE-engines.md, section "Engines: Pad Generation, G-code, Tuner, Tooling > Leather Locator Marks" (2026-11-10).*

Sizing Rules > **Lesser-used settings** > Leather Locator Marks (2026-10-06, a feature request from another tech; Matt doesn't use it himself). Engraved guides on leather pads showing where the felt sits, so a pad with **no center hole** can be centered by eye before pressing. Opt-in, pad-size range (inclusive, default 7.0–16.0 = the no-hole sizes under default sizing), leather only. Styles: four cardinal lines ending at the felt edge (default — equal gaps all round reads as centered, where a circle exactly at the felt edge just disappears under the felt), a circle the size of the felt disc, or both; optional dashes.

### The marks are on the flesh side, so the leather is cut flesh-side up

**The marks are on the flesh side, so the leather is cut flesh-side up.** The felt sits on the flesh side and the laser marks the face-up side; a mark on the grain side is either on the sealing surface or on the visible rim. With the leather flipped, the size number lands on the flesh side too, in the wrap zone, glued against the card and hidden. Whether a laser line reads well on fuzzy flesh side is for the requesting users to find out — it's in the tooltip and the user guide, not enforced.

### One geometry, three renderers

**One geometry, three renderers.** `svg_engine.locator_mark_geometry(pad_size, outer_r, settings)` returns primitives relative to the disc center (`circles` radii + line `segments`), or None; `locator_circle_points` samples a circle (solid or dashed arcs). `_render_svg_discs`, `gcode_engine._collect_disc_strokes` and `PadPreviewWindow._draw_locator` all call them and only translate. Everything is symmetric about both axes, so the SVG↔G-code Y-flip is a no-op — still pinned by `test_svg_and_gcode_draw_the_same_marks`. Felt edge = `get_disc_diameter(pad, 'felt') / 2`, the same number as the felt disc that gets cut. Lines run from the felt edge outward and stop `LOCATOR_EDGE_MARGIN_MM` (1.0) inside the cut, or inside the dart valleys on a darted pad, so nothing runs off material; a line shorter than `LOCATOR_MIN_LINE_MM` (0.8) is dropped rather than left as a speck. Dashes are laid from the felt edge outward so the inner end — the one that locates — is always a full dash on `felt_r`. No kerf: they're engraved.

### Layers and power — Matt's call, 2026-10-06

**Layers and power — Matt's call, 2026-10-06.** SVG: own layer color `leather_locator` (LightBurn 06 Cyan, in Layer Colors), so LightBurn users set their own power. G-code: own layer comment `C06` but the **leather engraving speed/power/passes/air** — there is no separate locator power in G-code yet. The marks go down right after the engraving layer and before any hole or cut, in both grouping modes, with no overscan, and the file carries a comment saying where the power came from. Matt wants separate power "eventually", only if users really ask — it means four new keys in `gcode_settings.leather` and `GCODE_PRESET_KEYS`, hence a backfill for every saved leather preset. Until then the UI note and user guide recommend SVG output for this feature.

### Lesser-used settings

**Lesser-used settings** is a collapsible at the bottom of the Sizing Rules dialog holding, in order, Labeled Zones, the locator marks, Adhesive Rings (beta, 2026-11-05) and Export Settings (zones and export moved there 2026-10-06, Matt's call: both are opt-in features most users never open). Collapsed by default; opens itself on dialog open when anything inside is non-default, so an active setting is never out of sight. Put future rarely-touched settings there rather than growing the main dialog.

**Preset schema**: five keys (`locator_marks_enabled/min_size/max_size/style/dashed`) in `SIZING_PRESET_KEYS`, backfilled by `normalize_sizing_preset` like the zone keys. Universal only — no per-range variant.

Tests: `tools/test_locator_marks.py` (16) — off/absent is byte-identical output, leather-only and inclusive range, line endpoints on plain and darted pads, the too-short drop, styles and dash anchoring, SVG↔G-code point parity, compatibility-mode unitless marks, G-code layer settings and ordering in both grouping modes, preset capture/backfill, layer color validity, and three GUI cases (self-skip without a display): collapsible default/auto-open and form round trip, Layer Colors row, pad preview draws the marks.
