# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Stohrer Sax Shop Companion is a cross-platform desktop GUI application for saxophone repair technicians. It provides SVG/G-code generation for laser-cutting pad materials, reference databases for key heights, serial numbers, and screw specifications, a tooling tab for die inserts and holders, a chromatic strobe tuner, and a harmonic tone analyzer.

## Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

External dependencies: `svgwrite`, `numpy`, `sounddevice` (tuner/toner). The GUI uses Python's built-in `tkinter`. Requires Python 3.11+.

### Testing

Test suites live in `tools/` and run non-interactively (no GUI). Run individual suites directly:

```bash
python tools/test_toner_engine.py
python tools/test_toner_full.py
python tools/test_tuner_engine.py
python tools/test_bugfixes.py
python tools/test_config.py
```

All test suites (48 files): `test_audio_utils`, `test_autofit_shift`, `test_bugfixes`, `test_camera_capture`, `test_card_paper_size`, `test_compare_filters`, `test_concert_pitch`, `test_config`, `test_dart_ranges`, `test_dart_shapes`, `test_descriptor_validity`, `test_detection_fix`, `test_edge_bias`, `test_engine_parity`, `test_falcon_sender`, `test_feeds_speeds_tester`, `test_fingerprint_filtering`, `test_frame_cut_scrap`, `test_framing_power`, `test_gcode_passes`, `test_gcode_presets_workflow`, `test_goodson_import`, `test_gpu_tuner`, `test_i18n`, `test_job_history`, `test_large_batch_optimization`, `test_lid_confirm`, `test_nesting_parity`, `test_pad_notes`, `test_pad_preview`, `test_polygon_parity`, `test_release_1_9`, `test_serial_lookup`, `test_sizing_presets_workflow`, `test_sizing_ranges`, `test_smoke_ui`, `test_tooling`, `test_toner_display`, `test_toner_engine`, `test_toner_full`, `test_tooltips`, `test_tuner_engine`, `test_tuner_updates`, `test_v161_compat`, `test_wav_import`, `test_wav_recording`, `test_web_pad_import`, `test_zone_labels`.

**SVG↔G-code parity**: `test_engine_parity` pins the contract that the SVG/preview output and the G-code output describe the same physical object (dart wave shape, engraving label placement, engine purity). The two engines render independently and have drifted before — when touching shared geometry (wave math, placement formulas, Y-flip), run this suite and extend it for any new shared shape.

**Portability notes**:
- `test_descriptor_validity` hardcodes a local WAV corpus path (`C:\sax shop companion\recordings`) and only runs on Matt's workstation. Skip it in CI and clean checkouts.
- `test_smoke_ui` constructs the full `PadSVGGeneratorApp` in a withdrawn Tk root — requires a display, so works on Windows/macOS and GitHub Actions Windows runners. On headless Linux it self-skips with a "no display" message.
- `test_zone_labels` is headless for everything except its three `preview_*` cases, which build a real `NestingPreviewWindow` and inspect its canvas. Those self-skip without a display; the other 31 always run. Note the window calls `wait_window()` in `__init__`, so the canvas inspection must be scheduled on the parent via `after()` *before* constructing it — see `_probe_preview`.

Before committing, run the suites affected by your changes. For releases, run all (minus `test_descriptor_validity` unless the WAV corpus is available). If adding new functionality, write a test script in `tools/` that exercises affected code paths. Test engine/logic functions directly. Print PASS/FAIL per test with a summary.

### Linting

```bash
pip install ruff
ruff check .          # zero errors expected on beta/main
ruff check --fix .    # auto-fix safe issues (unused imports, empty f-strings, etc.)
```

Config lives in `ruff.toml` (py311 target, 120-char lines, default E+F rules with E501/E701/E731/E741 relaxed, tools/ exempts E402 for `sys.path.insert` patterns). The CI `lint` job runs `ruff check .` and fails the workflow on any violation — keep the tree clean.

**Ruff gotcha for test scripts**: `ruff --fix` will strip "unused" imports even when they're the whole point (e.g. a test that verifies names import cleanly). Reference the imported names afterward (`assert all([Class1, Class2, ...])`) so ruff sees them as used. See `tools/test_smoke_ui.py` for the pattern.

## Job History

File > Job History (Pad Maker) opens `JobHistoryWindow` (`ui_dialogs.py`) — a log of every batch that reached an output stage. Storage is `job_history.json` via `load_job_history()` / `save_job_history()` / `append_job_history()` in config.py (newest first, trimmed to `JOB_HISTORY_LIMIT` = 300).

**Recording**: `PadSVGGeneratorApp._record_job(output, materials, pads, params, ...)` in main.py is called from exactly five places, always *after* real output exists:

| Call site | `output` | Notes |
|---|---|---|
| `on_generate_svg` | `"svg"` | after the per-material write loop |
| `_generate_svg_scrap_mode` | `"svg"` | per scrap, with `scrap_num` |
| `on_generate_gcode` | `"gcode"` | after the write loop, before the working popup closes |
| `_generate_gcode_scrap_mode` | `"gcode"` | per scrap, with `scrap_num` |
| `on_frame_and_cut` | `"laser"` | after the cut dialog; `status` carries `_final_reason`, so stopped/errored runs are logged too |

`_record_job` swallows every exception and logs it. Nothing in the app reads the history back except the dialog, so a history failure must never surface as a generation error — keep it that way when adding call sites.

**Reload**: `_load_job_into_form(job)` restores only what the user typed — pad text, materials, sheet size, center hole, base filename. It deliberately does NOT restore sizing rules, G-code settings, or `custom_polygon` (a camera-captured scrap is gone by then; the entry records `polygon_vertices` for display only). Sheet size is stored both as-typed and in mm, so a job saved in inches reloads correctly when the app is now in mm. Loading is refused while a scrap session is active (the session owns the pad list and locks the material checkboxes).

**List columns are measured, not fixed**: `_column_widths()` sizes each column from the widest header/value at refresh time. Translated headers vary a lot ("Pads" is "Zapatillas" in Spanish), and hardcoded widths truncated them. If you add a column, add it to `_headers()` and `_cells()` together — they're zipped positionally.

Tests: `tools/test_job_history.py` (storage round-trip + corruption handling, column alignment including a simulated long-translation case, and a record→reload round trip through a real form).

## Labeled Zones

Options > Sizing Rules > **Labeled Zones**: an opt-in toggle plus an editable pad-size range (default 7.0–12.5mm). Pads in that range are cut in bordered blocks — one block per size, grid-packed, with the size engraved along the block's top edge. Everything outside the range nests normally at full density.

**Why it exists**: small discs are indistinguishable once they're off the laser — a 7.0 and a 7.5 look the same. Their own engraved number doesn't solve it: the font gate (`font_size >= radius * 0.8`) drops the engraving entirely below ~5mm, and above that the number is often unreadable — too small on card/felt, and buried in the darts on leather (every leather pad under `dart_threshold` gets darts). **The label cannot move to the middle of a small leather pad — that's the sealing surface, and those are usually octave pads.** So the label goes on the waste instead of the part, which is the only place it can go.

**Cost**: zones trade sheet area for legibility, hence opt-in and the "may increase material wastage" note in the dialog. Measured on 7–12.5mm pads with a 1mm gutter, rectangular zones run 53–69% full (leather ~68%, near the `π/4 × (d/(d+g))²` ceiling for the gutter). Roughly 10–15% more sheet for a zoned batch.

**One model, two packers.** A group is a compact grid of one size with a rectangle round it and the size engraved on it, and groups are nested **as units** — like oversized pads. `nest_with_zones` dispatches on sheet shape: rectangular sheets shelf-pack groups into a band along the bottom (`_shelf_pack_zones`), traced polygons first-fit them into the outline (`_nest_polygon_groups`). Both emit identical `shape: 'rect'` zone dicts, so all three renderers share one path.

**Grid shape** (`zone_grid_candidates`) scores `aspect + empty_slots`: squareness matters, but a grid with holes doesn't read as a block, so a gap costs about as much as one step of elongation. That gives 6→3×2, 9→3×3, 8→4×2 (exact, not 3×3-with-a-hole), and a prime like 7→4×2 with one gap rather than a row of seven. Callers walk the list in order, so an awkward scrap degrades to a flatter grid instead of refusing the size.

**Rejected alternatives, all measured:**
- *Circular groups* — no new packing code at all (a group is just a big disc to `_nest_discs`), but roughly half as dense (28–36% vs 53–69% fill), and far too big for real material: 8× 7mm **leather** needs a 70mm circle, 48% of a 100×80 scrap; 8× 12.5mm needs 97mm.
- *Sequential per-size placement with an inflated inter-size collision radius* — the cheapest possible "grouping", and it **doesn't group**: the 7mm set still spread 83–108mm across a 150mm scrap (baseline 84.5mm), and widening the gap made it worse.
- *One full-width horizontal band per size, clipped to the outline* — shipped briefly and reverted. Bands are fine while each size is a single row, but once pads are gridded the rows stack: two grids ate a 146mm scrap whole and the remaining sizes had nowhere to go (**15 of 25 pads placed** on a real piece). The same four groups packed as units occupy 207×54mm and all 25 fit. `_clip_polygon_y` survives from that work and is still used for the leftover region.

**Key invariants**:
- `placed` keeps its plain `(pad_size, cx, cy, r)` shape. Zones travel as a **separate parallel list**, so `compute_remaining_pads`, the preview window, and job history all work untouched. Don't fold zone data into `placed`.
- Every zone is a `shape: 'rect'` dict (`x/y/w/h` + `cols/rows/label`). A `'poly'` variant existed for the band layout and was removed with it — nothing emitted it, so it was untested dead code. If a future shape (a circular group, say) is added, all three renderers — SVG `_render_svg_zones`, the G-code zone-stroke block, and the preview canvas — must learn it together, plus a parity test.
- Nothing touches the parity-pinned scan functions (`_scan_radial_*`, `_find_best_polygon_*`). Both layouts work by handing the nester a *smaller region*, never by adding obstacles. Scattering zones would leave the free area with *holes*, needing obstacle support in all six scan paths — deliberately not done.
- **Zones stand down in scrap mode** (`main.py` passes `zones=[]`): a scrap takes only part of a size's count, so a group would have to be re-sized per piece. Revisit only on real demand — the honest use case is ~10 each of a few neighbouring octave sizes, which fits one piece.
- Discs are checked with `spacing_mm + zone_edge_margin_mm` clearance while the boundary is drawn at the group edge, so the engraved line doesn't crowd the outer discs (2.5mm measured, vs 1.0mm without the margin).
- On camera-captured scraps the group rectangles land ~3mm inside the physical edge for free, because placement tests against `custom_polygon` (already inset by `camera_polygon_inset_mm`) rather than `custom_polygon_outline`. Don't "fix" this by switching to the outline. `camera_capture.inset_polygon_mm` is not usable here — it requires OpenCV, and `svg_engine` must stay dependency-light.
- **A zoned size is never cut without its group.** If a size can't get one, its pads are left unplaced so `can_all_pads_fit` reports the shortfall. The tempting fallback — nest them anyway so nothing is "lost" — was tried and reverted on 2026-08-11: it dropped an unlabeled group of small discs right next to the labeled ones (exactly the confusion zones prevent) while the caller saw a full placed count and reported success. A visible refusal beats a silent unlabeled pile. Out-of-range sizes still nest normally in the leftover region.
- **Groups are placed biggest-first** (`disc_d × qty`) so the roomy parts of a scrap are claimed before small groups fill in around them. The polygon first-fit scans top-left to bottom-right at `ZONE_GRID_SEARCH_STEP_MM`, rejecting cheaply (overlap, then rectangle-in-polygon) before the expensive per-disc `_circle_fits_in_polygon` test — that ordering is what keeps it instant rather than seconds.
- **Both the discs and the drawn rectangle must be on material.** `_rect_fits_in_polygon` samples along each edge, not just the corners, so a concave notch biting into the middle of an edge is caught — otherwise the engraved boundary runs off the scrap and marks nothing.
- Grouping genuinely costs yield — 30 leather discs fit a 150×110 scrap free-nested but need 210×140 grouped. That trade was accepted deliberately: this targets tiny pads, where yield is already high and demand low.
- Borders and labels are **engraved, never cut** — a cut border would drop the zone tile through the bed slats. They're emitted as one engraving pass before any disc is cut, so the sheet is labeled before parts come loose.
- The label is digits and `.` only. `STROKE_FONT`/`FILLED_FONT` in gcode_engine carry no `x`, so a "×10" count suffix would need a new glyph in **both** fonts.

**Grid shape gotcha**: `_zone_box` scores the *full* zone including border and label strip. Scoring the bare inner grid always picks a degenerate 1×N strip — one column has the fewest gutters and so the smallest inner area — which shelf-packs badly and reads poorly. Among shapes within `ZONE_ASPECT_SLACK` of the smallest area, the squarest wins.

**Preset schema**: only the three user-facing keys (`zone_labels_enabled`, `zone_label_min_size`, `zone_label_max_size`) are in `SIZING_PRESET_KEYS`; gutter/border/font are global tuning constants with no UI. Adding keys to that tuple breaks `_detect_active_preset()` for every already-saved preset, so `config.normalize_sizing_preset()` backfills missing keys from `DEFAULT_SETTINGS` before comparing — use it whenever the schema grows.

Tests: `tools/test_zone_labels.py` (38) — opt-in/regression safety (zones off must reproduce the old nester placement-for-placement), range bounds, grid shapes and their fallbacks, containment, group non-overlap and visible separation, groups-and-boundaries-land-on-material, never-cut-unlabeled, the clip helper on concave shapes, SVG↔G-code Y-flip agreement, and three preview-canvas cases that self-skip without a display.

## Pad Preview Window

The Sizing Rules dialog has an opt-in live preview (`PadPreviewWindow` in `ui_dialogs.py`). A "Show live pad preview" checkbox just below the preset section opens a resizable Toplevel that renders the selected pad with the parent form's current sizing rules applied. Controls: pad size (mm), per-material checkboxes (leather / felt / card / exact size), layout radio (layered concentric vs side-by-side).

Geometry comes from the same helpers as the SVG output: `svg_engine.get_disc_diameter`, `svg_engine.get_felt_thickness_mm`, and `svg_engine._wave_value` (for the dart wave). That means what you see in the preview is the exact shape that will be cut. Drawing happens on a `tk.Canvas` — fast, dependency-free, redraws on `<Configure>`.

Live updates: the window polls `parent_options._capture_form_to_dict()` every 200 ms and re-renders if the snapshot changed. No tk-var traces are used because some form state (sizing/dart range lists) lives in plain Python lists that aren't trace-able. Polling tolerates mid-edit invalid states (catches `TclError` / `ValueError`) and falls back to a placeholder message.

The preview tears down whenever the OptionsWindow is destroyed (any path) via a `<Destroy>` bind on `self.top`. Closing the preview directly resets the parent's `show_preview_var` so the checkbox stays in sync.

## Sizing Rules Presets Workflow

The Sizing Rules dialog (`OptionsWindow`) is preset-first — the preset section sits at the *top* of the form, and the bottom button is **Apply** (not "Save"). Workflow:

- **Preset dropdown + Load** at the top: explicit Load click, with a "discard unsaved changes?" confirm if the form is dirty.
- **Save Preset** opens `SaveSizingPresetDialog`, a small radio-choice modal: *Overwrite existing* (combobox of saved names) or *Save as new preset* (text entry). Defaults to overwrite when an active preset is loaded, to "new" when nothing is loaded or the library is empty.
- **Rename** prompts for a new name; refuses empty / duplicate names.
- **Delete** refuses to wipe the last preset — at least one must remain.
- **Apply** (bottom button): if the form is dirty, prompts the user to either save as a preset first or back out to keep editing — there is no path to commit unsaved-as-a-preset changes silently.
- **Cancel / window-close X**: if dirty, three-way prompt — save as preset / discard / keep editing.

Dirty detection is a `_capture_form_to_dict()` snapshot vs `self._baseline`; the baseline resets on dialog open, after Load, and after a successful Save Preset. `active_preset_name` tracks which preset's values currently sit in the form (used for the Save Preset overwrite default).

**Naming the loaded preset**: on open, `_detect_active_preset()` matches the form's snapshot against each saved preset and selects the one that fits, so the dropdown names what's loaded instead of sitting blank. There is deliberately no stored "active preset" settings key — the applied values themselves identify the preset, which can't go stale after an edit and stays honest when a config is hand-edited (the dropdown just stays blank). This relies on Apply refusing to commit changes that aren't captured in a preset, so applied settings always correspond to a saved one. `GcodeSettingsWindow` does the same per material.

**Bootstrap**: `main.py` auto-creates a `Default` preset from current settings on first run if `sizing_presets` is empty (via `config.settings_to_sizing_preset`). The app guarantees at least one preset always exists.

## G-code Settings Presets Workflow

`GcodeSettingsWindow` (Options > G-code Settings... on Pad Maker, Options > Settings... on Tooling) is preset-aware **per material**. Each material section (felt / card / leather / acrylic / basswood) has its own preset bar at the top with Load / Save / Rename / Delete, and its own active-preset name + dirty baseline. Editing felt does not dirty card. The preset library is shared across the two dialogs — saving a felt preset from Pad Maker shows up in any future dialog that includes felt.

- **Per-material storage**: `gcode_presets.json` shape is `{material: {preset_name: data}}`. Top-level keys are the 5 materials in `config.GCODE_PRESET_MATERIALS`. Inner data captures only `config.GCODE_PRESET_KEYS` (the 19 keys per material — engraving mode, line/filled speed+power+passes, fill density, hole/cut speed+power+passes, kerf, four air toggles).
- **Cross-material isolation by design**: a felt preset will not load into the acrylic slot. Materials have characteristic settings ranges and mixing them silently is dangerous; users who want to cross-apply must Save As under the target material.
- **Apply (bottom button)**: if any material is dirty, a three-way prompt (Yes / No / Cancel) — save dirty materials as preset(s) before applying, apply anyway, or keep editing.
- **Cancel / window-close X**: same three-way prompt, but the "apply anyway" branch becomes "discard and close."
- **Save**: opens `SaveSizingPresetDialog` (generalized — accepts `title`/`intro` kwargs) with material-specific copy ("Save Felt Preset" etc.). Overwrite defaults to the active preset when one is loaded.
- **Delete refuses to wipe the last preset** for that material. **Rename refuses empty / duplicate names.**

Dirty tracking uses per-material `_capture_material_to_dict(mat)` snapshots compared against `self.material_baseline[mat]`. Baselines reset on dialog open, after Load, and after a successful Save Preset. `active_preset_name[mat]` tracks which preset's values currently sit in each material's fields.

**Bootstrap**: `main.py` loads `gcode_presets.json` and, on first run *or* if any material is missing, backfills a `Default` preset for that material from the current `gcode_settings[material]` (via `config.settings_to_gcode_presets`). The app guarantees at least one preset per material always exists.

**Backward compatibility**: `GcodeSettingsWindow(..., gcode_presets=None)` (the default) disables the preset bar entirely — used by `test_smoke_ui` and any future caller that wants the bare settings dialog.

## Settings-Dialog Tooltips

`ui_dialogs.py` exposes a `Tooltip` helper plus `add_tooltip(widget, text)` and `add_tooltips(text, *widgets)` convenience functions. All settings dialogs (Sizing Rules, G-code Settings, Layer Colors, Key Layout, Tuner Settings, Toner Settings) hover-explain their fields. When adding a new setting widget, attach a tooltip alongside it — short sentence, plain English, focused on *why* the user would change it. Attach to both the label and the input so users can hover either. The helper popup is overrideredirect + topmost so it appears above modal dialogs without stealing focus.

## Dart Shape Spectrum

The dart wave shape (formerly "Star/Dart"; now just "Darts" in the UI) is a single 0.0–1.0 slider that smoothly interpolates between three primitive shapes:
- 0.0 = Triangle (linear ramps between peaks/valleys)
- 0.5 = Sine (raw cosine)
- 1.0 = Square (saturating sign function via `|c|^p`, `p = _SQUARE_POWER = 0.01`)

Math lives in `_wave_value` in `svg_engine.py`; `calculate_star_path` (SVG) and `_generate_star_points` (G-code) both call it per sample — `tools/test_engine_parity.py` pins them together. The default value is now `0.5` (sine). Legacy configs used a 0.0=sine, 1.0=square scale; `load_settings` migrates them once via `0.5 + 0.5 * old` and sets `dart_shape_v2: True` so the migration doesn't run again. `dart_ranges[*].shape_factor` is migrated alongside the universal value. When changing dart-shape behavior, update `tools/test_dart_shapes.py` (anchor + smoothness + migration coverage).

Internal variable names retain the `dart_` prefix (`dart_shape_factor`, `dart_threshold`, etc.); only the user-facing labels were renamed to "Darts".

## Building Executables

The app uses PyInstaller to create standalone executables. Each platform must build its own executable (no cross-compilation).

```bash
# Install build dependencies
pip install -r requirements.txt

# Build for current platform
python build.py

# Clean and rebuild
python build.py --clean

# (macOS only) Build and create .dmg disk image
python build.py --dmg

# Output locations:
#   Windows: dist/SaxShopCompanion.exe
#   macOS:   dist/SaxShopCompanion.app (or .dmg with --dmg flag)
#   Linux:   dist/SaxShopCompanion
```

### GPU Tuner Renderer (Rust/wgpu)

The strobe tuner has an optional GPU-accelerated renderer in `tuner_renderer/` (Rust crate using pyo3 + wgpu) — **Windows and Linux only**. CI builds this on the Windows and Linux runners, but a local checkout of `python main.py` will silently fall back to the slower canvas renderer unless you build the extension yourself:

```bash
# Requires Rust toolchain (rustup). Windows/Linux only — do NOT do this on
# a Mac (see the macOS warning below).
pip install maturin
python -m maturin build --release --manifest-path tuner_renderer/Cargo.toml
pip install --find-links tuner_renderer/target/wheels tuner_render
```

After this, `import tuner_render` succeeds and the tuner uses GPU rendering at 60-120 fps. If the import fails for any reason the tuner falls back to canvas rendering on Windows/Linux — the app still runs, just slower. See the `_skip_theme` / `_dark_canvas` flag notes in the Strobe Tuner architecture section for integration details, and the GPU/Canvas constant alignment warning (`DIM_MULTIPLIER`, `BRIGHTNESS_GAMMA`) for what has to stay in sync between the Python path and the shader.

**macOS is canvas-only — never load `tuner_render` on darwin.** Tk Aqua draws all widgets into a single NSView per toplevel, and `winfo_id()` returns a pointer to Tk's internal `MacDrawable` struct, not an NSView ("the value has no meaning outside Tk" — Tk docs). `tuner_renderer/src/platform.rs` would wrap that handle as an NSView, so wgpu's Metal backend segfaults in `objc_msgSend` during surface creation — a native crash the Python init-failure `except` in `tuner_tab.py` can never catch. Three layers enforce the gate: `tuner_tab.py` skips the `tuner_render` import on darwin, `build.py` skips the `--hidden-import` on darwin, and CI skips the Rust/maturin steps on macOS runners. Even a real NSView wouldn't fix it — a CAMetalLayer on the shared per-window view would paint over the entire UI (and would still need Retina scale handling), so macOS GPU rendering is off the table by design. **History**: the macOS Apple Silicon zips for v1.95 through v2.6 shipped with the renderer bundled — on those builds, opening the Tuner tab crashes the app outright.

On Windows, Rust needs the MSVC linker. The one-time machine setup is `winget install Rustlang.Rustup` followed by `winget install Microsoft.VisualStudio.2022.BuildTools --override "--add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"`.

### Windows Installer (Inno Setup)

CI wraps `dist\SaxShopCompanion.exe` into a versioned `SaxShopCompanion-Windows-Setup-{version}.exe` via `installer.iss`. The installer creates Start Menu + optional desktop shortcuts, registers an uninstaller, and installs to `{autopf}\SaxShopCompanion` (requires admin UAC). User data in `%APPDATA%\StohrerSaxShopCompanion\` is untouched on uninstall. Local build:

```bash
python build.py
iscc /DAppVersion=2.65 installer.iss    # requires Inno Setup 6
```

**Do not change the `AppId` GUID** in `installer.iss` — Windows uses it to recognize upgrades. Changing it produces a parallel install instead of an in-place upgrade.

## Bundled Runtime Assets

Files that need to be reachable at runtime in both source and frozen builds (icons, SVG templates, etc.) follow this pattern:

1. Place the asset at the repo root (single files) or in a subfolder (collections).
2. Extend `build.py`'s PyInstaller `cmd` via `--add-data`. Single files use `'<file>{os.pathsep}.'`; folders use `'<folder>{os.pathsep}<folder>'`.
3. Resolve at runtime with `base = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(__file__)`, then `os.path.join(base, ...)`.

Current examples: `icon.ico` (loaded by `main.py` for the title bar / taskbar icon), `tooling_assets/die_organizer_{upper,lower}.svg` (copied by `generate_die_organizer_svg` in `svg_engine.py`), `pad_press_spacers/*.stl` (copied by `ToolingTabMixin._save_pad_spacer_stl` in `tooling_tab.py`), and `locale/` (compiled translation catalogs resolved by `i18n._locale_dir()`).

## Internationalization (i18n)

User-facing strings are wrapped with `_("...")` and translated via GNU gettext. The catalog is initialized in `main.py` **before** any UI module is imported (so module-level `_()` calls resolve against the active language).

**Pattern in source**:
```python
# `_` and `ngettext` are installed into builtins by i18n.init_translation().
# No explicit import needed. ruff.toml whitelists them in `builtins`.
label = _("Cancel")
msg = _("Imported {n} captures").format(n=n)
plural = ngettext("{n} pad", "{n} pads", n).format(n=n)
```

**Files**:
- `i18n.py` — `init_translation(lang)`, `available_languages()`, module-level `_` and `ngettext` handles for tests
- `babel.cfg` — extraction config for pybabel
- `locale/saxshop.pot` — generated template (commit it; regenerate after adding strings)
- `locale/<lang>/LC_MESSAGES/saxshop.po` + `saxshop.mo` — per-language catalogs (commit both)

**Workflow when adding or changing strings**:
```bash
python tools/extract_strings.py        # regenerate saxshop.pot
python tools/update_translations.py    # merge .pot into existing .po files (marks new entries fuzzy)
# Edit each locale/<lang>/LC_MESSAGES/saxshop.po — translate the fuzzy entries
python tools/compile_translations.py   # rebuild all .mo files
python tools/test_i18n.py              # verify
```

**Languages**: v1 ships English (source) + Spanish / German / French / Italian. Native-name display in `i18n.LANGUAGE_NAMES`. Language switching is restart-required, picked via File > Feature Set > Language.

**Gotchas**:
- Module-level `_` shadowing: Python's idiomatic throwaway variable name is `_`. We use `gettext.install()` instead of `from i18n import _` so local `_` shadows the builtin within the for-loop scope only — no rename needed for existing `for _, x in items` patterns.
- Module-level constants with translatable strings: define as a function (e.g. `get_resonance_messages()`) so each call resolves against the *current* catalog. Lists evaluated at module import time bake the source-language values.
- Don't translate data — pad sizes, material keys (`"felt"`, `"card"`), settings keys are code. Only translate at the display layer.
- f-strings with embedded variables: `_("Imported {n} captures").format(n=n)`. Don't put the f-prefix on the gettext string itself or the placeholder gets baked.

## Branching Strategy

**ALWAYS sync with the remote before doing anything else.** Matt develops this app from several
different computers, so a local checkout is frequently behind `origin/beta` — and may also hold
unpushed local commits, making the branch *diverged* rather than simply stale. Before reading code,
answering questions about what the app does, or making any edit:

```bash
git fetch --all --prune
git status --short --branch                       # ahead / behind counts
git log --oneline beta..origin/beta               # what this machine is missing
git log --oneline origin/beta..beta               # what this machine hasn't pushed
```

- **Behind only** → `git pull --rebase` and continue.
- **Diverged (ahead *and* behind)** → `git pull --rebase` so local work replays on top; check the
  local commits for conflicts with what landed upstream before pushing.
- Never start work off a stale checkout. `APP_VERSION` in the local `config.py`, the local release
  notes, and memory files are all unreliable until this sync has run — check
  `git show origin/beta:config.py | grep APP_VERSION` and `gh release list --limit 5` for the real
  current version.

The `.claude/settings.json` `SessionStart` hook runs the fetch + divergence report automatically at
the start of each session, but it only *reports* — reconciling is still a deliberate step.

- **`main`**: Stable release branch. Merges from `beta` when features are tested and ready.
- **`beta`**: Active development branch. New features land here first (e.g. filled engraving, air assist toggles, cut grouping). Always work on `beta` unless told otherwise.
- CI builds trigger on push to `main` or `beta`.

## Versioning

`APP_VERSION` in `config.py` is the manual source of truth — bump it when preparing a release. `APP_BUILD_DATE` is auto-derived from `sys.executable`'s mtime in frozen builds (installer/zip copies preserve mtime) and falls back to the manual constant when running from source. The About dialog reads both via `ui_dialogs.py`.

## Release Notes Style

Every GitHub release uses the same body template: a short single-paragraph lead describing the app, then the **full feature overview** organized tab-by-tab (Pad Maker, Tooling, Chromatic Strobe Tuner, Harmonic Tone Analyzer, Cross-platform & General), followed by **Known limitations** and **Upgrading from v1.x**. Every release is the complete picture of the app, so a user landing on any release page cold gets the full story — no "What's new since" sections or fix logs.

Items new *in that specific release* get a plain `**(new)**` marker prefixed before the relevant bullet (or interpolated into a longer bullet that has both old and new content). Items from prior releases carry no marker. v2.1 is the canonical example — see [the v2.1 release page](https://github.com/stohrermusic/Stohrer-Sax-Shop-Companion/releases/tag/v2.1) for the exact format. Don't mix in `(new in v2.X.Y)` cross-version tags or "Fixes & polish" sections — both got tried and rejected.

## CI/CD (GitHub Actions)

The `.github/workflows/build.yml` workflow has two jobs:
- **`lint`** (ubuntu-latest, ~10s): runs `ruff check .` — fails the workflow on any violation
- **`build`** (4-platform matrix): Windows Inno Setup installer (the bare PyInstaller .exe is built but not published — only the installer ships), macOS Apple Silicon .app, macOS Intel .app, and Linux binary

Triggers on push to `main` or `beta`, on release creation, or manually.

- macOS Intel build (`macos-15-intel` runner) installs only svgwrite+pyinstaller (no numpy/sounddevice) — tuner and toner are unavailable
- `full_build: true/false` matrix flag controls whether Rust toolchain + maturin are installed for the GPU tuner renderer; the macOS runners additionally skip the Rust steps unconditionally — macOS is canvas-only (see GPU Tuner Renderer section)
- The Windows job also installs Inno Setup 6 via Chocolatey and builds the installer; version is extracted from `config.py`'s `APP_VERSION` via PowerShell regex
- macOS jobs package the `.app` with `ditto -c -k --keepParent` (never `zip -r` — it materializes bundle symlinks and breaks the code-signature resource seal) and run three guards: `plutil -extract` for the mic + camera usage keys, `codesign --verify --deep --strict` on the built .app, and the same verify on an unzipped copy of the final artifact (see "macOS Build" in CLAUDE-architecture.md)
- Uploads artifacts to the workflow run; on Windows that's the installer only (`SaxShopCompanion-Windows-Setup-*.exe`). All published artifacts also attach to GitHub Releases on release events.

**Action pins**: `actions/checkout@v5`, `actions/setup-python@v6`, `actions/upload-artifact@v6` — all on Node 24. Don't downgrade; GitHub removes Node 20 from runners in September 2026.

## Config File Locations

The app stores settings and presets in platform-appropriate locations:

| Platform | Location |
|----------|----------|
| Windows | `%APPDATA%\StohrerSaxShopCompanion\` |
| macOS | `~/Library/Application Support/StohrerSaxShopCompanion/` |
| Linux | `~/.config/StohrerSaxShopCompanion/` (respects `XDG_CONFIG_HOME`) |

**Backward compatibility**: On first run, existing config files in the old location (current working directory) are automatically migrated to the new location.

**Manual import**: Users can also manually import settings from a previous installation via File > "Import Settings from Folder..." which copies config files from a selected directory.

## Detailed Documentation

Architecture and domain-specific guidance is split across companion files imported via `@`-statements below. Claude Code loads them as part of CLAUDE.md's context.

- **CLAUDE-architecture.md** — Module structure, design patterns, settings/presets, error logging, feature set
- **CLAUDE-engines.md** — Pad generation, G-code, SVG rendering, nesting, strobe tuner, tooling, Phil Noy credit
- **CLAUDE-toner.md** — Tone analyzer engine, data model, capture modes, analyze tool, WAV recording, calibration
- **CLAUDE-web.md** — Web data sync, screw specs submission form, related repository

---

## Subsystem imports

@CLAUDE-architecture.md
@CLAUDE-engines.md
@CLAUDE-toner.md
@CLAUDE-web.md
