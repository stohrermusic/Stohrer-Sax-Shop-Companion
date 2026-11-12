# User interface: decisions and history

This file holds the why and the history behind the rules in the CLAUDE files, and the full detail they summarise: what was tried, measured, rejected and decided, with dates and the user's own words. The paragraphs were moved here verbatim from the CLAUDE files on 2026-11-10 (HEAD 01071f5); each section names where it came from. Add new decisions to the matching section, dated.

## Job History

*Moved from CLAUDE.md, section "CLAUDE.md > Job History" (2026-11-10).*

File > Job History (Pad Maker) opens `JobHistoryWindow` (`ui_dialogs.py`) — a log of every batch that reached an output stage. Storage is `job_history.json` via `load_job_history()` / `save_job_history()` / `append_job_history()` in config.py (newest first, trimmed to `JOB_HISTORY_LIMIT` = 300).

**Recording**: `PadSVGGeneratorApp._record_job(output, materials, pads, params, ...)` in main.py is called from exactly five places, always *after* real output exists:

`_record_job` swallows every exception and logs it. Nothing in the app reads the history back except the dialog, so a history failure must never surface as a generation error — keep it that way when adding call sites.

### Tooling runs are in it too (2026-11-09; Matt: "tooling runs go into job history if easy")

**Tooling runs are in it too (2026-11-09; Matt: "tooling runs go into job history if easy").** `PadSVGGeneratorApp._record_tooling_job(job, output, status, save_dir)` is `_record_job`'s sibling: the same entry keys, plus `"tool"` (the Tooling section key: `die_inserts`, `die_holders`, `die_organizer`, `kerf_test`, `feeds_speeds_tester`, `tonehole_crescent`), with a one-line description in the user's terms where `pad_text` goes (die sizes with counts; holder variant and layers; organizer plate; the kerf test's cut speed and power; the Speed & Power matrix; the crescent list as the tab shows it), `pads` empty, `placed_count` = `requested_count` = pieces, the sheet in mm only, `materials` = `[material]` (empty for the organizer SVG and the crescent SVG/STL, which know no material), `output` `"svg"`, `"gcode"`, `"stl"` or `"laser"`. It is called from **one** place for Frame & Cut — `on_frame_and_cut`'s Tooling branch, after the cut dialog closes, with `_final_reason` as the status, before `after_cut`; the seven Tooling job dicts carry the fields via `ToolingTabMixin._tooling_history(...)`, the Die Scrap job with its sheet number — and from every Tooling file export through `ToolingTabMixin._record_tooling_export(output, save_path, **fields)`, one line after each write (fourteen sites: die SVG/G-code standard and scrap, holder SVG/G-code, organizer SVG/G-code, kerf SVG/G-code, Speed & Power G-code, crescent SVG/STL/G-code). Nothing about what is written or streamed changed: the 48-click Tooling export recording and the 28-click Pad Maker recording are identical before and after. The dialog shows a Tooling row with the tool's name in the Pads column (`ui_dialogs.tooling_tool_name`), says "pieces" not "pads", prints the description as it is, keeps Load disabled for it and answers a double-click with one sentence (`tooling_job_not_loadable`, the one place it is written); `_load_job_into_form` returns at once for a `tool` entry as the backstop. Delete and Clear work as for any row. Before this date Tooling runs were deliberately left out ("that log is a pad list that can be loaded back into the form"); the brief allowed them in only if the entry kept its shape with a field or two added, which it does.

### Reload

**Reload**: `_load_job_into_form(job)` restores only what the user typed — pad text, materials, sheet size, center hole, base filename. It deliberately does NOT restore sizing rules, G-code settings, or `custom_polygon` (a camera-captured scrap is gone by then; the entry records `polygon_vertices` for display only). Sheet size is stored both as-typed and in mm, so a job saved in inches reloads correctly when the app is now in mm. Loading is refused while a scrap session is active (the session owns the pad list and locks the material checkboxes).

### List columns are measured, not fixed

**List columns are measured, not fixed**: `_column_widths()` sizes each column from the widest header/value at refresh time. Translated headers vary a lot ("Pads" is "Zapatillas" in Spanish), and hardcoded widths truncated them. If you add a column, add it to `_headers()` and `_cells()` together — they're zipped positionally.

Tests: `tools/test_job_history.py` (storage round-trip + corruption handling, column alignment including a simulated long-translation case, a record→reload round trip through a real form, and the Tooling rows: the entry shape, the tool name and "pieces" in the real dialog, Load off for them and on for a pad row, the double-click sentence with no question, delete, the loader backstop); `tools/test_tooling_frame_cut.py` requires the export's and the run's entries for every tool, described the same way, stopped and errored cuts with their status, and the Die Scrap sheet numbers on runs and exports.

## Direct machine control everywhere (the 3.0 work, 2026-11-09)

*Moved from CLAUDE.md, section "CLAUDE.md > Direct machine control everywhere (the 3.0 work, 2026-11-09)" (2026-11-10).*

- **The User Guide upgraded (2026-11-09, the last piece before release prep).** Matt: "keep it concise, but have it be a little more of a guide, vs just blocks of description. have a TOC that is clickable on the right, and have it open to full screen. put example graphics where appropriate." The text was rewritten task-led and cut from 7,513 words to about 4,150 (every paragraph its own string, so an edit retranslates one string), with a clickable contents list, opening maximized and not transient, always the whole guide (Help > User Guide opens at the current tab's section), and seven figures drawn from the app's own geometry. Detail in CLAUDE-architecture.md ("The User Guide"); gated by `tools/test_user_guide.py`.

## Pad Preview Window

*Moved from CLAUDE.md, section "CLAUDE.md > Pad Preview Window" (2026-11-10).*

The Sizing Rules dialog has an opt-in live preview (`PadPreviewWindow` in `ui_dialogs.py`). A "Show live pad preview" checkbox just below the preset section opens a resizable Toplevel that renders the selected pad with the parent form's current sizing rules applied. Controls: pad size (mm), per-material checkboxes (leather / felt / card / exact size), layout radio (layered concentric vs side-by-side).

Geometry comes from the same helpers as the SVG output: `svg_engine.get_disc_diameter`, `svg_engine.get_felt_thickness_mm`, and `svg_engine._wave_value` (for the dart wave). That means what you see in the preview is the exact shape that will be cut. Drawing happens on a `tk.Canvas` — fast, dependency-free, redraws on `<Configure>`.

Live updates: the window polls `parent_options._capture_form_to_dict()` every 200 ms and re-renders if the snapshot changed. No tk-var traces are used because some form state (sizing/dart range lists) lives in plain Python lists that aren't trace-able. Polling tolerates mid-edit invalid states (catches `TclError` / `ValueError`) and falls back to a placeholder message.

The preview tears down whenever the OptionsWindow is destroyed (any path) via a `<Destroy>` bind on `self.top`. Closing the preview directly resets the parent's `show_preview_var` so the checkbox stays in sync.

## Sizing Rules Presets Workflow

*Moved from CLAUDE.md, section "CLAUDE.md > Sizing Rules Presets Workflow" (2026-11-10).*

The Sizing Rules dialog (`OptionsWindow`) is preset-first — the preset section sits at the *top* of the form, and the bottom button is **Apply** (not "Save"). Workflow:

- **Preset dropdown + Load** at the top: explicit Load click, with a "discard unsaved changes?" confirm if the form is dirty.

- **Save Preset** opens `SaveSizingPresetDialog`, a small radio-choice modal: *Overwrite existing* (combobox of saved names) or *Save as new preset* (text entry). Defaults to overwrite when an active preset is loaded, to "new" when nothing is loaded or the library is empty.

- **Rename** prompts for a new name; refuses empty / duplicate names.

- **Delete** refuses to wipe the last preset — at least one must remain.

- **Apply** (bottom button): if the form is dirty, prompts the user to either save as a preset first or back out to keep editing — there is no path to commit unsaved-as-a-preset changes silently.

- **Cancel / window-close X**: if dirty, three-way prompt — save as preset / discard / keep editing.

Dirty detection is a `_capture_form_to_dict()` snapshot vs `self._baseline`; the baseline resets on dialog open, after Load, and after a successful Save Preset. `active_preset_name` tracks which preset's values currently sit in the form (used for the Save Preset overwrite default).

### Naming the loaded preset

**Naming the loaded preset**: on open, `_detect_active_preset()` matches the form's snapshot against each saved preset and selects the one that fits, so the dropdown names what's loaded instead of sitting blank. There is deliberately no stored "active preset" settings key — the applied values themselves identify the preset, which can't go stale after an edit and stays honest when a config is hand-edited (the dropdown just stays blank). This relies on Apply refusing to commit changes that aren't captured in a preset, so applied settings always correspond to a saved one. `GcodeSettingsWindow` does the same per material.

**Bootstrap**: `main.py` auto-creates a `Default` preset from current settings on first run if `sizing_presets` is empty (via `config.settings_to_sizing_preset`). The app guarantees at least one preset always exists.

## G-code Settings Presets Workflow

*Moved from CLAUDE.md, section "CLAUDE.md > G-code Settings Presets Workflow" (2026-11-10).*

`GcodeSettingsWindow` (Options > G-code Settings... on Pad Maker, Options > Settings... on Tooling) is preset-aware **per material**. Each material section (felt / card / leather on the Pad Maker, plus adhesive rings while that feature is on; acrylic / basswood / cork on Tooling) has its own preset bar at the top with Load / Save / Rename / Delete, and its own active-preset name + dirty baseline. Editing felt does not dirty card. The preset library is shared across the two dialogs — saving a felt preset from Pad Maker shows up in any future dialog that includes felt.

- **Per-material storage**: `gcode_presets.json` shape is `{material: {preset_name: data}}`. Top-level keys are the 7 materials in `config.GCODE_PRESET_MATERIALS` (cork joined on 2026-11-09). Inner data captures only `config.gcode_preset_keys_for(material)`: `config.GCODE_PRESET_KEYS` (the 19 keys per material — engraving mode, line/filled speed+power+passes, fill density, hole/cut speed+power+passes, kerf, four air toggles), or `GCODE_CUT_ONLY_KEYS` (cut speed / power / passes, kerf, cut air) for the materials in `GCODE_CUT_ONLY_MATERIALS`: adhesive rings have no engraving and no center hole, so their block, their presets and their dialog section carry the cut keys only.

- **Cross-material isolation by design**: a felt preset will not load into the acrylic slot. Materials have characteristic settings ranges and mixing them silently is dangerous; users who want to cross-apply must Save As under the target material.

- **Apply (bottom button)**: if any material is dirty, a three-way prompt (Yes / No / Cancel) — save dirty materials as preset(s) before applying, apply anyway, or keep editing.

- **Cancel / window-close X**: same three-way prompt, but the "apply anyway" branch becomes "discard and close."

- **Save**: opens `SaveSizingPresetDialog` (generalized — accepts `title`/`intro` kwargs) with material-specific copy ("Save Felt Preset" etc.). Overwrite defaults to the active preset when one is loaded.

- **Delete refuses to wipe the last preset** for that material. **Rename refuses empty / duplicate names.**

Dirty tracking uses per-material `_capture_material_to_dict(mat)` snapshots compared against `self.material_baseline[mat]`. Baselines reset on dialog open, after Load, and after a successful Save Preset. `active_preset_name[mat]` tracks which preset's values currently sit in each material's fields.

**Bootstrap**: `main.py` loads `gcode_presets.json` and, on first run *or* if any material is missing, backfills a `Default` preset for that material from the current `gcode_settings[material]` (via `config.settings_to_gcode_presets`). The app guarantees at least one preset per material always exists.

**Backward compatibility**: `GcodeSettingsWindow(..., gcode_presets=None)` (the default) disables the preset bar entirely — used by `test_smoke_ui` and any future caller that wants the bare settings dialog.

## Key Design Patterns

*Moved from CLAUDE-architecture.md, section "Architecture > Key Design Patterns" (2026-11-10).*

### The User Guide (upgraded 2026-11-09)

**The User Guide (upgraded 2026-11-09).** Matt: "let's give the user guide a bit of an upgrade- keep it concise, but have it be a little more of a guide, vs just blocks of description. have a TOC that is clickable on the right, and have it open to full screen. put example graphics where appropriate." `UserGuideWindow` (ui_dialogs.py) is now always the WHOLE guide: a clickable contents list down the right side (a `ttk.Treeview`: sections and sub-sections; a click scrolls the heading to the top of the view; the entry for the part being read is highlighted as the user scrolls, by the last heading at or above the top of the view), a search box (Ctrl+F / Cmd+F), and drawn figures. `section` is the tab's key to open AT (`open_user_guide` maps the current tab to it through `tab_sections`; `scroll_to(key)`), so Help > User Guide opens at the tab the user is on; the one-tab filtering and its "Show Full Guide" button are gone, and search always covers everything. It opens maximized (`state("zoomed")`; macOS Aqua has no such state, so there the window is made as large as Tk says a window may be, `wm maxsize`, which on Aqua is the visible frame less the menu bar, the Dock and the frame decorations, asked for at +0+0; the earlier screen-sized geometry was cut down by the window manager on CI's Macs, which lose 122 to 138 px of height, and the suite's gate is written against `wm maxsize` for the same reason, not against the screen) and is NOT transient (a transient window has no maximize button on Windows); on a wide screen the text keeps to `TEXT_MAX_PX` (760) through a right margin on every line (`_fit_column`, the `col` tag). **The contents list is driven by a click binding, not `<<TreeviewSelect>>`**: `selection_set` fires that event too, after any guard flag has been reset, so the scroll sync and a selection handler would select each other for ever (the first version hung; found in the first smoke run). The text is ten `_section_*` methods in the guide's order, each opening with `self._section(key, title)` (a top contents entry) and using `_sub` (a child entry), `_body`, `_bullet`, `_steps` (numbered; the numbers are added by the app, each step one plain string), `_blank`, `_link` and `_figure(name)`. Every paragraph is its own `_()` string, so an edit by Matt retranslates one string. **Figures** are canvases drawn from the app's own geometry helpers (`_fig_*`: the sizing rules and the dart wave through `PadPreviewWindow`'s own drawing methods, lent by `_FigureCanvasHost`; `nest_pads_with_zones` on a small traced scrap; `adhesive_ring_geometry`; the die engraving's arc text from `gcode_engine`; the crescent geometry; a three-lane diagram of the ways to the laser), so they are always right, need no image files and carry no baked-in English: every label goes through `_()`, every colour is explicit (macOS dark mode), each figure is placed once, and they draw at construction into fixed-size canvases (nothing scheduled). **When adding a tab or a feature**: add a `_section_*` method (or a `_sub` in the right section) and call it from `_insert_content` in the guide's order; add the tab's key to `SECTION_KEYS` and to `tab_sections` in `open_user_guide()`; add a figure only where it explains something words do worse, through `_figure` with a `_fig_<name>` drawer and a height in `FIGURE_HEIGHTS`. Gated by `tools/test_user_guide.py` (window, contents list, figures, search, Help, the Quick Start link, the text's rules, and a German child process); its gates are written in Tk's terms, not the platform's: the maximize gate accepts the zoomed state or a window within 90 % / 80 % of `wm maxsize`, and the cursor-in-the-search-box gate reads `focus_lastfor` and a recorded `focus_search` call, never `focus_get`, which is None on every window under xvfb, where no window manager hands out keyboard focus (both failed on CI's first run of e2293e4, Macs and Linux, with the window right). The Quick Start Guide's last line ("The rest is in Help > User Guide.") opens the whole guide (`QuickStartWindow(on_user_guide=...)`, from `main.open_quick_start`).

## Welcome window, Quick Start Guide, User Guide search (2026-11-09)

*Moved from CLAUDE-architecture.md, section "Architecture > Welcome window, Quick Start Guide, User Guide search (2026-11-09)" (2026-11-10).*

Matt: "each install shows a quick welcome, a version number, and thank you for installing, be sure to search the user guide with any questions, and says click here for the quick start guide", and of the guide: "whatever guide we write should be SUPER concise". He rejected captions under the Generate buttons as too busy, so nothing was added to the main tabs.

### Welcome window

**Welcome window** (`ui_dialogs.WelcomeWindow`). Shown once for each version: `main.maybe_show_welcome()` compares `settings["welcomed_version"]` (in `DEFAULT_SETTINGS`, "" for never) with `APP_VERSION`, records the version **when the window is shown** (a crash must not bring it back every launch) and opens it. So a new install sees it once, and an existing user once after each update. Help > Welcome opens it any time (`open_welcome`, one window at a time). **`maybe_show_welcome` is called from exactly one place: the `__main__` block of main.py, under `if not tour:`, on an `after(600)`** once the main window is up. Building the app never calls it, so `--selftest`, `--tour` and every suite that constructs `PadSVGGeneratorApp` never see a Welcome window they did not ask for, and startup is no slower. Keep it that way: a call in `__init__` would pop a window in dozens of test runs.

- *Content, and nothing more*: the app name and version, "Thank you for installing.", "Questions? Search the User Guide." (a link: `open_user_guide_search`, the full guide with the cursor in its search box), "Click here for the quick start guide" (a link: `open_quick_start`), and Close.

- *Look: plain, on purpose.* The app's ordinary dialog look (`DIALOG_BG`, the usual fonts, links as in the About dialog). Not modal: it takes no grab. **Do not style it.** The first version was a nod to an old dial-up sign-on screen (three framed panels that lit up one after another, then the text); the same day Matt first asked for no animation ("just a quick popup that doesn't get in the way") and then for no styling at all: "let's have the welcome screen be simple. if I have a detailed ui request I will give it." It has no canvas, no timers and no colours of its own, and the suite checks that none come back.

### Quick Start Guide

**Quick Start Guide** (`ui_dialogs.QuickStartWindow`, Help > Quick Start Guide, above User Guide). One small page of labels: three ways to cut (SD card, LightBurn, Frame & Cut), "Do not mix them", and what to check before cutting. Each line is its own translatable string (`QuickStartWindow.lines`), because the wording is Matt's to change. The Frame & Cut line is only there when `machine_supported()`. No images, no scrolling (the window is as big as its text), not modal. It never opens by itself.

### User Guide search

**User Guide search.** The Welcome window tells people to search the guide, so `UserGuideWindow` has a search box: typing highlights every match (`search_hit`) and jumps to the first (`search_current`), Enter or Next goes to the next and wraps round, the count reads "3 of 12" or "No matches", Ctrl+F (Cmd+F on macOS) puts the cursor there. Case-insensitive, plain text. The whole guide is always loaded (see "The User Guide" under Key Design Patterns), so it always searches everything.

**The tour has a stop for each** ("welcome" and "quick-start", after "about"; 20 stops now), so both are in the Mac screenshots.

Tests: `tools/test_welcome.py` (9): once per version from a fresh folder and from another stored version; the single call site and `--selftest` recording nothing; the window being a plain dialog (five lines and Close, the dialog background, no canvas, nothing scheduled, no grab, none of the animated version's attributes); both links; the guide's eleven lines; Help on every tab; search (count, highlight, wrap, case, no matches, a one-tab guide expanding); and, in a German child process, both windows fitting.

## SD Card & Eject

*Moved from CLAUDE-engines.md, section "Engines: Pad Generation, G-code, Tuner, Tooling > SD Card & Eject" (2026-11-10).*

Two mechanisms:

- (There is no "Send G-code to SD Card" menu item: this line used to describe one, and no such command exists in `main.py` — found while rewriting the User Guide, 2026-11-09. The guide describes only the eject tick.)

## ui_dialogs becomes a package (2026-11-10, structural step 1)

Matt approved the structural work before 3.0 ("Yeah go for it"; no rush to release, nothing to ship for a while after), one step per round, each a move only. Step 1: the 11,330-line `ui_dialogs.py` (32 classes, 12 functions, 9 constants) became the package `ui_dialogs/`, cut **by what the dialogs share, not by size**: `common.py` (what 28 of the 32 classes use: `DIALOG_BG`, the tooltip helper, `bind_mousewheel`, `ConfirmationDialog`, `material_display_name`, `gcode_file_note`, the loggers, `_PREVIEW_COLORS`), `sizing.py` (OptionsWindow with the two dialogs only it and GcodeSettingsWindow open, the pad preview the guide borrows, Layer Colors), `gcode_settings.py`, `presets_io.py`, `previews.py` (the two that share `_PREVIEW_COLORS`), `history.py`, `guide.py` (the four help windows), `machine.py` (the six that share `mlog`, `FalconRunDialog`, the Grbl tables and the port lock), `resonance.py`. Line counts after the cut (ruff having dropped the unused header imports): machine 4,193, sizing 2,843, guide 1,427, gcode_settings 1,125, presets_io 556, previews 480, history 421, common 285, resonance 120, __init__ 143; 11,593 in all for 11,330 before (the module docstrings, the sibling imports, the re-exports).

**How it was cut.** A script took every top-level statement of the old module as a chunk with the blank and comment lines above it (a banner stays with the item it introduces), wrote each chunk to its module in the old order, gave each module the old import header (ruff then dropped what a module does not use) and the names it reads from sibling modules, and generated `__init__.py` re-exporting all 52 names plus `tk`, `ttk`, `messagebox`, `filedialog`, `simpledialog` (suites patch `ui_dialogs.messagebox.<fn>`; that is the tkinter module object, so it still reaches every submodule). A second script compared every old top-level object's source with its new home, line for line: 47 of 52 identical; the 5 that differ do so only by the six call-site lines below. ruff's import fix reached inside two method bodies (a redundant local `import sys` in PolygonDrawWindow, a local `from tkinter import simpledialog` in CameraCalibrationDialog) and both lines were put back byte for byte; the module-level imports they shadowed were dropped from machine.py's header instead, where ruff is free to act.

**The one design decision: names a suite stands in for.** Suites and the recordings replace five dialogs on the package by attribute (`ui_dialogs.FalconRunDialog = Fake`, `LidClosedDialog`, `SaveSizingPresetDialog`, `CameraCaptureDialog`, `CameraCalibrationDialog`), and in one module that worked because the constructing dialog looked the name up in the same namespace at call time. In a package the lookup would hit the submodule's own global and miss the stand-in, so the six constructions of those five (OptionsWindow and GcodeSettingsWindow -> SaveSizingPresetDialog; PolygonDrawWindow -> CameraCaptureDialog; CameraCalibrationDialog -> FalconRunDialog, twice; FalconRunDialog -> LidClosedDialog) read `_pkg.Name(...)` with `import ui_dialogs as _pkg` in sizing, gcode_settings and machine. Every other cross-module name is a plain `from .common import ...` (or `from .sizing import PadPreviewWindow` in guide). main.py and the tabs were not touched: `from ui_dialogs import (...)` at import, and `from ui_dialogs import FalconRunDialog` inside `on_frame_and_cut` at call time, both resolve on the package. The suites were the proof and were not edited for stand-ins; seven were re-pointed where they read the module as a file or scanned the root: `test_adhesive_ring` (the package's files), `test_tuner_updates` (`guide.py`), `test_camera_fallback` (every submodule's source), `test_gettext_shadowing`, `test_exception_handling` (allow-list keys by the new file, 58 handlers in 46 functions unchanged), `test_threads` (the package in their scans) and `test_gcode_file_note` (the sentence written once: it walks the root and every package under it).

**Proof.** The catalog template regenerated from the package has the same 1,726 msgids (2,526 source references, 1,173 of them now in `ui_dialogs/`); `babel.cfg` already scanned `**.py`. Both recordings identical to bdf05cb's; `python main.py --selftest` and the tour (`test_mac_paths`) construct everything; the full run green. Two things to know: the module logger `log` in `common.py` is now named `ui_dialogs.common` (it was `ui_dialogs`; nothing configures or reads that name), and `git log --follow` on a new module does not reach back into `ui_dialogs.py` (a split, not a rename; the old file's history stays on its path).

## The Quick Start Guide in Matt's words (2026-11-11)

After the laser walk-through Matt asked for the page's text and sent it back rewritten; the page is his text now: a welcome line, the User Guide line (the link, second on the page instead of last), "Quick start", the three ways (SD card, LightBurn, direct laser control) each as a short sequence, "Before you cut!", where the speeds and powers are with "Defaults are for a 40 W diode laser.", and "Test first. Don't start a fire." Only names were set to the app's (Pad Maker, Frame & Cut, 40 W). `QuickStartWindow.lines` has a third kind, `'link'`, for the User Guide line; a build that cannot drive a laser drops the direct laser control line and reads "Two ways to do it:". His text left out "Test on the material you will really cut." and "Die inserts and holders use the Acrylic settings."; they are gone from the page (both are still in the User Guide).

The same evening, from his read of the User Guide: the darts line says "in the size ranges you set" (it said "under a size"); the Edge Bias line says the center dot "packs from the middle outward" (it said "clears it"); the Labeled Zones figure is drawn with a corner bias and captioned "the rest nested as usual" (see `labeled-zones.md`, "The groups follow the Edge Bias").
