# Labeled Zones: decisions and history

This file holds the why and the history behind the rules in the CLAUDE files, and the full detail they summarise: what was tried, measured, rejected and decided, with dates and the user's own words. The paragraphs were moved here verbatim from the CLAUDE files on 2026-11-10 (HEAD 01071f5); each section names where it came from. Add new decisions to the matching section, dated.

## Labeled Zones

*Moved from CLAUDE.md, section "CLAUDE.md > Labeled Zones" (2026-11-10).*

Options > Sizing Rules > **Lesser-used settings** > **Labeled Zones** (moved into the collapsible 2026-10-06): an opt-in toggle plus an editable pad-size range (default 7.0–12.5mm). Pads in that range are cut in bordered blocks — one block per size, grid-packed, with the size engraved along the block's top edge. Everything outside the range nests normally at full density.

### Why it exists

**Why it exists**: small discs are indistinguishable once they're off the laser — a 7.0 and a 7.5 look the same. Their own engraved number doesn't solve it: the font gate (`font_size >= radius * 0.8`) drops the engraving entirely below ~5mm, and above that the number is often unreadable — too small on card/felt, and buried in the darts on leather (every leather pad under `dart_threshold` gets darts). **The label cannot move to the middle of a small leather pad — that's the sealing surface, and those are usually octave pads.** So the label goes on the waste instead of the part, which is the only place it can go.

### Cost

**Cost**: zones trade sheet area for legibility, hence opt-in and the "may increase material wastage" note in the dialog. Measured on 7–12.5mm pads with a 1mm gutter, rectangular zones run 53–69% full (leather ~68%, near the `π/4 × (d/(d+g))²` ceiling for the gutter). Roughly 10–15% more sheet for a zoned batch.

### One model, two packers

**One model, two packers.** A group is a compact grid of one size with a rectangle round it and the size engraved on it, and groups are nested **as units** — like oversized pads. `nest_with_zones` dispatches on sheet shape: rectangular sheets shelf-pack groups into a band along the bottom (`_shelf_pack_zones`), traced polygons first-fit them into the outline (`_nest_polygon_groups`). Both emit identical `shape: 'rect'` zone dicts, so all three renderers share one path.

### Grid shape

**Grid shape** (`zone_grid_candidates`) scores `aspect + empty_slots`: squareness matters, but a grid with holes doesn't read as a block, so a gap costs about as much as one step of elongation. That gives 6→3×2, 9→3×3, 8→4×2 (exact, not 3×3-with-a-hole), and a prime like 7→4×2 with one gap rather than a row of seven. Callers walk the list in order, so an awkward scrap degrades to a flatter grid instead of refusing the size.

**Rejected alternatives, all measured:**

- *Circular groups* — no new packing code at all (a group is just a big disc to `_nest_discs`), but roughly half as dense (28–36% vs 53–69% fill), and far too big for real material: 8× 7mm **leather** needs a 70mm circle, 48% of a 100×80 scrap; 8× 12.5mm needs 97mm.

- *Sequential per-size placement with an inflated inter-size collision radius* — the cheapest possible "grouping", and it **doesn't group**: the 7mm set still spread 83–108mm across a 150mm scrap (baseline 84.5mm), and widening the gap made it worse.

- *One full-width horizontal band per size, clipped to the outline* — shipped briefly and reverted. Bands are fine while each size is a single row, but once pads are gridded the rows stack: two grids ate a 146mm scrap whole and the remaining sizes had nowhere to go (**15 of 25 pads placed** on a real piece). The same four groups packed as units occupy 207×54mm and all 25 fit. `_clip_polygon_y` survives from that work as a tested general utility — nothing in the app calls it anymore (the leftover region uses cover circles over the full outline instead), only its tests do.

**Key invariants**:

- `placed` keeps its plain `(pad_size, cx, cy, r)` shape. Zones travel as a **separate parallel list**, so `compute_remaining_pads`, the preview window, and job history all work untouched. Don't fold zone data into `placed`.

- Every zone is a `shape: 'rect'` dict (`x/y/w/h` + `cols/rows/label`). A `'poly'` variant existed for the band layout and was removed with it — nothing emitted it, so it was untested dead code. If a future shape (a circular group, say) is added, all three renderers — SVG `_render_svg_zones`, the G-code zone-stroke block, and the preview canvas — must learn it together, plus a parity test.

- Nothing touches the parity-pinned scan functions (`_scan_radial_*`, `_find_best_polygon_*`). Both layouts work by handing the nester a *smaller region*, never by adding obstacles. Scattering zones would leave the free area with *holes*, needing obstacle support in all six scan paths — deliberately not done.

- **A `max`-quantity size can't be grouped, and the app refuses the combination.** A group is a fixed grid, so it needs a fixed count. `_prepare_generation()` in main.py errors out when zones are on and a `max` pad's size falls inside the zone range, rather than silently cutting that size loose and unlabeled. Using `max` on a size OUTSIDE the range still works and is the intended way to fill the rest of a piece.

- **Leftover (unzoned) pads nest over the whole outline, not just below the groups.** Each placed group's footprint is reserved via `_cover_rect_with_circles` and passed to the nester as `preplaced` — a seeded `placed` list that the scan functions collision-check unchanged, so no scan math is touched and the parity contracts hold (`preplaced` defaults to `None`). Seeding the group's *discs* alone is not enough: loose pads settle into the gaps between them, inside someone else's labeled box. The nesters strip the seeds from their return value. Before this, leftovers were clipped to the strip under the lowest group — a hangover from the stacked-band layout that stranded most of a scrap, placing **zero** max-fill pads above or beside the groups.

- **Zones stand down in scrap mode** (`main.py` passes `zones=[]`): a scrap takes only part of a size's count, so a group would have to be re-sized per piece. Revisit only on real demand — the honest use case is ~10 each of a few neighbouring octave sizes, which fits one piece.

- Discs are checked with `spacing_mm + zone_edge_margin_mm` clearance while the boundary is drawn at the group edge, so the engraved line doesn't crowd the outer discs (2.5mm measured, vs 1.0mm without the margin).

- On camera-captured scraps the group rectangles land ~3mm inside the physical edge for free, because placement tests against `custom_polygon` (already inset by `camera_polygon_inset_mm`) rather than `custom_polygon_outline`. Don't "fix" this by switching to the outline. `camera_capture.inset_polygon_mm` is not usable here — it requires OpenCV, and `svg_engine` must stay dependency-light.

- **A zoned size is never cut without its group, on either path.** If a size can't get one, its pads are left unplaced (but still counted in `fixed_total`) so `can_all_pads_fit` reports the shortfall. A too-wide block first degrades through flatter grid shapes before being given up on. The tempting fallback — nest them anyway so nothing is "lost" — was tried and reverted on 2026-08-11: it dropped an unlabeled group of small discs right next to the labeled ones (exactly the confusion zones prevent) while the caller saw a full placed count and reported success. A visible refusal beats a silent unlabeled pile. Out-of-range sizes still nest normally in the leftover region.

- **Groups are placed biggest-first** (`disc_d × qty`) so the roomy parts of a scrap are claimed before small groups fill in around them. The polygon first-fit scans top-left to bottom-right at `ZONE_GRID_SEARCH_STEP_MM`, rejecting cheaply (overlap, then rectangle-in-polygon) before the expensive per-disc `_circle_fits_in_polygon` test — that ordering is what keeps it instant rather than seconds.

- **Both the discs and the drawn rectangle must be on material.** `_rect_fits_in_polygon` samples along each edge, not just the corners, so a concave notch biting into the middle of an edge is caught — otherwise the engraved boundary runs off the scrap and marks nothing.

- Grouping genuinely costs yield — 30 leather discs fit a 150×110 scrap free-nested but need 210×140 grouped. That trade was accepted deliberately: this targets tiny pads, where yield is already high and demand low.

- Borders and labels are **engraved, never cut** — a cut border would drop the zone tile through the bed slats. They're emitted as one engraving pass before any disc is cut, so the sheet is labeled before parts come loose.

- The label is digits and `.` only. `STROKE_FONT`/`FILLED_FONT` in gcode_engine carry no `x`, so a "×10" count suffix would need a new glyph in **both** fonts.

### One shape rule, both paths

**One shape rule, both paths**: `zone_grid_candidates` is the single source of truth; `_zone_box` computes dimensions from it and accepts a `shape=` override so a caller can walk the list. The rectangular path used to score zone area independently, which silently disagreed with the polygon path — 6 came out 2×3, 8 as 2×4, and a prime like 7 as a **1×7 column** — so the same job produced different blocks depending on sheet type. Caught by the 2026-08-11 initbig audit. If you touch grid selection, `test_both_sheet_types_pick_the_same_grid` and `test_no_quantity_produces_a_single_file_strip` are the guards.

### One box geometry, both paths

**One box geometry, both paths**: same rule, same failure mode, found the same way. `_group_metrics(settings)` returns `(gutter, border, font, group_gap)` and is the **only** place those keys are read — neither packer reads them directly, because when they did they disagreed: the border was `zone_border_mm` 1.0 on rectangular sheets but `zone_edge_margin_mm` 1.5 on polygons, and the inter-box gap was a hardcoded `ZONE_GAP_MM = 2.0` versus a `zone_band_gap_mm` setting of 6.0. `zone_edge_margin_mm` still exists but now has exactly one job — extra clearance when testing a group's discs against a traced outline — and no longer doubles as the box border. Guard: `test_both_sheet_types_space_boxes_the_same`.

### The gap between boxes is not a moat

**The gap between boxes is not a moat** (`zone_group_gap_mm`, default 1.0 = the disc gutter). It was 6.0 while each size got a full-width *band*, where the gap had to separate strips. Once a size became a box with its own `zone_border_mm` of white space inside it, 6.0 put ~7mm between neighbouring boxes and ~15mm between discs in adjacent groups, against 1mm between discs inside one — Matt read it as leftover moat on real material, correctly. Tightening it reclaims that: on the test scrap the four-group span went 123×61mm → 111×57mm, and on **leather** — where a 7mm pad is a 17mm disc — it's the difference between one group fitting and two. Don't take it to 0: two abutting engraved lines read as one box. `test_group_gap_is_not_a_moat` pins both ends (gap ≤ 2×border, and > 0) plus "tighter never places fewer pads".

### Preset schema

**Preset schema**: only the three user-facing keys (`zone_labels_enabled`, `zone_label_min_size`, `zone_label_max_size`) are in `SIZING_PRESET_KEYS`; gutter/border/font/group-gap are global tuning constants with no UI. Adding keys to that tuple breaks `_detect_active_preset()` for every already-saved preset, so `config.normalize_sizing_preset()` backfills missing keys from `DEFAULT_SETTINGS` before comparing — use it whenever the schema grows.

### Keep test fixtures to real material and real pad sizes

**Keep test fixtures to real material and real pad sizes.** The two `x max` leftover tests originally filled a 302×158mm scrap with a `4.0mm` pad — a 1.2mm card disc — and placed ~2900 of them, which cost **88s and 71s** and made the suite unusable at ~170s. Matt makes nothing under 7.0mm and works on offcuts never bigger than 14×14in; the fixture is now a 254×162mm offcut filled with `MAX_FILL = 18.0`, and the suite runs in **6.5s**. The old fixture also silently distorted what was being tested: only a sub-millimetre disc fits *above* the groups (they're placed biggest-first from the top and claim the topmost material), so "fills above the groups" was an artifact of the unrealistic size, not a property worth pinning. The real signal is "fills *beside* them" plus a ≥85% comparison against the same fill with zones off.

Tests: `tools/test_zone_labels.py` (47) — opt-in/regression safety (zones off must reproduce the old nester placement-for-placement), range bounds, grid shapes and their fallbacks, containment, group non-overlap and visible separation, groups-and-boundaries-land-on-material, never-cut-unlabeled, the clip helper on concave shapes, SVG↔G-code Y-flip agreement, and three preview-canvas cases that self-skip without a display.
