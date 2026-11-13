# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Stohrer Sax Shop Companion is a cross-platform desktop GUI application for saxophone repair technicians. It provides SVG/G-code generation for laser-cutting pad materials, reference databases for key heights, serial numbers, and screw specifications, a tooling tab for die inserts and holders, a chromatic strobe tuner, and an analyzer (an intonation chart first; the harmonic analyzer behind a switch, experimental).

## Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

External dependencies: `svgwrite`, `numpy`, `sounddevice` (tuner/toner). The GUI uses Python's built-in `tkinter`. Requires Python 3.11+.

### Testing

Test suites live in `tools/` and run non-interactively (no GUI). **Run everything** with `python tools/run_tests.py` (one line per suite, failures' output at the end, non-zero exit on any failure; `python tools/run_tests.py zone svg` filters by name). That is what CI runs. Run individual suites directly:

```bash
python tools/test_toner_engine.py
python tools/test_toner_full.py
python tools/test_tuner_engine.py
python tools/test_bugfixes.py
python tools/test_config.py
```

All test suites (89 files): `test_adhesive_ring`, `test_adhesive_ring_fuzz`, `test_adhesive_ring_upgrade`, `test_analyzer_tab`, `test_audio_utils`, `test_audit_fixes`, `test_autofit_shift`, `test_bugfixes`, `test_calibration_dialog`, `test_camera_capture`, `test_camera_fallback`, `test_card_paper_size`, `test_compare_filters`, `test_concert_pitch`, `test_config`, `test_config_isolation`, `test_cork_material`, `test_dart_ranges`, `test_dart_shapes`, `test_descriptor_validity`, `test_detection_fix`, `test_die_organizer_gcode`, `test_edge_bias`, `test_engine_parity`, `test_exception_handling`, `test_falcon_sender`, `test_feeds_speeds_tester`, `test_fingerprint_filtering`, `test_frame_cut_alignment`, `test_frame_cut_scrap`, `test_frame_cut_steps`, `test_framing_loop_handoff`, `test_framing_power`, `test_gcode_file_note`, `test_gcode_passes`, `test_gcode_presets_workflow`, `test_gettext_shadowing`, `test_goodson_import`, `test_gpu_tuner`, `test_i18n`, `test_intonation`, `test_intonation_chart`, `test_job_history`, `test_laser_watch`, `test_lid_confirm`, `test_library_tabs_e2e`, `test_locator_marks`, `test_mac_paths`, `test_nesting_parity`, `test_pad_maker_e2e`, `test_pad_notes`, `test_pad_preview`, `test_pitch_octaves`, `test_polygon_parity`, `test_release_1_9`, `test_release_2_86`, `test_release_3_0`, `test_scrap_live_list`, `test_scrap_partial_nest`, `test_scrap_pending`, `test_serial_lookup`, `test_settings_schema`, `test_sizing_presets_workflow`, `test_sizing_ranges`, `test_smoke_ui`, `test_svg_wellformed`, `test_threads`, `test_tls_context`, `test_tonehole_crescent`, `test_tonehole_crescent_gcode`, `test_tooling`, `test_tooling_e2e`, `test_tooling_frame_cut`, `test_toner_data`, `test_toner_display`, `test_toner_engine`, `test_toner_full`, `test_toner_live`, `test_tooltips`, `test_tuner_canvas`, `test_tuner_engine`, `test_tuner_updates`, `test_user_guide`, `test_v161_compat`, `test_wav_import`, `test_wav_recording`, `test_web_pad_import`, `test_welcome`, `test_zone_labels`.

**What CI runs**: `lint` (ruff), `test` (every suite via `run_tests.py` on Windows, macOS Apple Silicon, macOS Intel with the Intel build's dependency set and `--allow-missing sounddevice`, and Linux under `xvfb-run`), and `build`, which `needs: [lint, test]` and runs `--selftest` on every frozen bundle. The macOS test job is the only place the macOS code paths run: nobody on the project owns a Mac.

**Run the suites as CI does before reporting them green**: `python tools/run_tests_as_ci.py [filters]` builds `<repo>/.venv/ci-py3.11` from a Python 3.11 (the `py` launcher, PATH, or uv's managed interpreters; `--install` fetches one with uv) with `locks/full.txt` installed `--require-hashes` as the workflow does, and runs `run_tests.py` in it; its first lines name the interpreter and numpy, which every report states. Why: five Analyzer commits passed here on Python 3.14 and failed on all four CI platforms on 3.11 (CI's locked numpy 2.4.6 against this machine's unlocked 2.4.2: a float one last place off a golden literal; a chart read that raced a 100 ms debounce; a window that opened wider than a check assumed), 2026-11-11, `docs/decisions/testing-and-ci.md`. A computed float is never compared exactly in a suite (`near`, `near_list`); a check never reads the canvas after a plain wait (draw it, or wait for the drawing, bounded).

**On Windows the suites run on a private desktop** (Matt, 2026-11-11: "I can't use my pc for a while when they are running"): `tools/run_tests.py` re-runs itself through `tools/hidden_desktop.py` on a desktop of its own, every suite, tour and `--selftest` child inherits it, and nothing shows on or takes focus from the user's desktop; `run_tests_as_ci.py` inherits it; `--visible` (or `SSC_VISIBLE=1`) keeps them on yours; a no-op on macOS, Linux and in CI (`GITHUB_ACTIONS`). The runner's first line says where they are. Run a probe, a tour or a planted-break tree the same way: `python tools/hidden_desktop.py <command...>`. The screen size read there is the same, the 300 s kill takes the whole tree there too, and a native box nobody stood in for hangs unseen until that kill, so the dialog backstop matters more, not less.

**Standing rules for tests** (why: `docs/decisions/testing-and-ci.md`):

- **Never touch the real settings folder.** Every suite that can reach it sets `SAXSHOP_CONFIG_DIR` to a temp folder at module level, before any project module is imported. `run_tests.py` gives each suite its own. Gated by `test_config_isolation`.
- **No suite may open a native message box.** A suite that builds the app, a dialog or a tab does `import box_guard` and `box_guard.install()` at module level, calls `box_guard.settle()` in its `check()` helper, and names each question it causes with `guard.expect("<title>", answer=...)`. Never write `messagebox.askyesno = lambda *a, **k: True`. List the suite in `GUARDED_SUITES` (`tools/test_release_3_0.py`).
- **Never run a deliberately broken tree without a backstop under the native dialog layer.** A check of the guard asserts a function is the guard's stand-in before calling it.
- **No suite may open or enumerate real serial ports.** Stand in for enumeration and the probe.
- **Writing a check**: `check(name, fn)` fails only on an exception or a literal `False`. Assert inside the function; never `lambda: a == b or a`. Run every new gate against the old code before trusting it. Build a modal dialog whose threads call `after()` inside `root.mainloop()`. Drive frames with `mainloop` and a quit timer, never a `root.update()` loop.
- **No new silent `except Exception: pass`, no bare `except:`** (allow-list in `test_exception_handling`; see CLAUDE-architecture.md "Error Logging"). **TLS is always verified** (`test_tls_context`).
- **Every worker thread starts through `threads.start_worker`; no `threading.Thread(` outside threads.py; the locks a worker holds while the Tk thread may wait are `threads.GcHeldRLock`** (`test_threads`; see CLAUDE-architecture.md "Worker threads").
- **Bound what can hang**: audio-device calls under a time bound, threads joined with a timeout. `run_tests.py` kills a suite after 300 s (`SSC_SUITE_TIMEOUT_S`) with its process group and arms `faulthandler`.
- **SVG and G-code parity**: when touching shared geometry (wave math, placement formulas, Y-flip), run `test_engine_parity` and extend it for any new shared shape.
- **Real fixtures**: real material and real pad sizes, nothing under 7.0 mm.
- **Self-skips are explicit**: `test_descriptor_validity` (WAV corpus), `test_goodson_import` (website checkout), `test_gpu_tuner` without the wheel (`SSC_REQUIRE_GPU=1` forces it), audio suites without `AUDIO_AVAILABLE`.
- **`--selftest` and `--tour` stay side-effect free.** `_selftest()` builds the app withdrawn and exits 0 or 1. `run_tour` schedules each stop's finish with `after()` before calling its `open()`. Synthetic audio: `TunerEngine.synthetic_hz`, `TonerEngine.synthetic_hz`.
- **A new UI feature gets a construct-and-poke test** (see `test_locator_marks`).

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

File > Job History opens `JobHistoryWindow` (ui_dialogs/history.py); storage is `job_history.json` via `load_job_history()` / `save_job_history()` / `append_job_history()` (newest first, `JOB_HISTORY_LIMIT` = 300). `_record_job(...)` is called from exactly these places, always after real output exists:

| Call site | `output` | Notes |
|---|---|---|
| `on_generate_svg` | `"svg"` | after the per-material write loop |
| `_generate_svg_scrap_mode` | `"svg"` | per scrap, with `scrap_num` |
| `on_generate_gcode` | `"gcode"` | after the write loop, before the working popup closes |
| `_generate_gcode_scrap_mode` | `"gcode"` | per scrap, with `scrap_num` |
| `on_frame_and_cut` | `"laser"` | after the cut dialog; `status` carries `_final_reason`, so stopped/errored runs are logged too |

- `_record_job` and `_record_tooling_job` swallow and log every exception: a history failure must never surface as a generation error.
- Tooling runs are recorded by `_record_tooling_job` (Frame & Cut) and `ToolingTabMixin._record_tooling_export` (after each file write); their entries carry `"tool"`, cannot be loaded, and `_load_job_into_form` returns at once for them.
- `_load_job_into_form` restores only what the user typed, never sizing rules, G-code settings or `custom_polygon`, and refuses during a scrap session.
- A new column goes into `_headers()` and `_cells()` together; widths are measured.

Tests: `test_job_history`. Detail: `docs/decisions/user-interface.md`.

## Labeled Zones

Opt-in (Sizing Rules > Lesser-used settings): pads in a size range (default 7.0 to 12.5 mm) are cut in bordered grid blocks, one per size, with the size engraved on the block, because small discs cannot be told apart once cut. Never move a label to the middle of a small leather pad (the sealing surface). Code: `nest_with_zones`, `_shelf_pack_zones`, `_nest_polygon_groups`, `zone_grid_candidates`, `_group_metrics` (svg_engine.py). Invariants:

- `placed` keeps its `(pad_size, cx, cy, r)` shape; zones travel as a separate list of `shape: 'rect'` dicts. A new zone shape goes into all three renderers (SVG, G-code, preview) together, with a parity test.
- Nothing touches the parity-pinned scan functions. Zones work by giving the nester a smaller region or `preplaced` seeds, never obstacles.
- A zoned size is never cut without its group, on either path; `max` on a zoned size is refused; zones stand down in scrap mode and for `ZONE_FREE_MATERIALS`.
- The groups follow the Edge Bias (2026-11-11): on a traced shape the first-fit scan starts from the bias side (`_zone_scan_order`; "n" is the old top-left scan, "center" starts at the centroid); on a sheet the band is along the top for n / nw / ne, else the bottom, and filled from the right for e / ne / se, else the left (mirrored in `nest_with_zones`; the nester still only sees the smaller rectangle).
- `zone_grid_candidates` and `_group_metrics` are the single source for both sheet types. Borders and labels are engraved, never cut, before any disc.
- On a camera-captured scrap, groups are placed against the inset `custom_polygon`, not the outline. Do not switch it.
- When `SIZING_PRESET_KEYS` grows, backfill with `config.normalize_sizing_preset()`.

Tests: `test_zone_labels`. Why and what was rejected: `docs/decisions/labeled-zones.md`.

## Adhesive Rings (beta)

An opt-in fifth Pad Maker material, `adhesive_ring` (Sizing Rules > Lesser-used settings > Adhesive Rings (beta)): a ring of double-sided tape per pad, for sticking a thin silicone wrap to the back of the card. While `adhesive_ring_enabled` is off the material does not exist in the UI: `main._refresh_adhesive_ring_ui` adds and removes its `material_vars` entry, Materials checkbox, Layer Colors row, G-code Settings section and Framing Power row, so no flow can pick up a material nobody can see. Geometry and invariants: CLAUDE-engines.md "Adhesive Rings". Tests: `test_adhesive_ring`, `test_adhesive_ring_fuzz`, `test_adhesive_ring_upgrade`. History: `docs/decisions/adhesive-rings.md`.

## Direct machine control everywhere (the 3.0 work)

Frame & Cut and G-code for every Tooling tool, cork, one gating rule, the G-code file line, the Welcome window, Quick Start and User Guide, and a laser or camera plugged in after launch: one body of work meant to become version 3.0. Do not bump the version or write release notes until Matt asks for release prep. Rules are in the companion files. History: `docs/decisions/machine-integration.md`.

## Pad Preview Window

`PadPreviewWindow` (opened from Sizing Rules) draws the selected pad with the form's current rules. Draw only through the output's own helpers (`get_disc_diameter`, `_wave_value`, `locator_mark_geometry`, `adhesive_ring_geometry`) so it shows what will be cut. It polls `_capture_form_to_dict()` every 200 ms and closes with the OptionsWindow. Tests: `test_pad_preview`.

## Sizing Rules and G-code Settings presets

Both dialogs are preset-first (Load, Save via `SaveSizingPresetDialog`, Rename, Delete). `GcodeSettingsWindow` has one preset bar per material; its library (`gcode_presets.json`, `{material: {name: data}}`, keys from `config.gcode_preset_keys_for(material)`) is shared by the Pad Maker and Tooling dialogs. Rules:

- Apply never commits unsaved changes silently; Cancel on a dirty form asks save / discard / keep editing.
- Delete never removes the last preset (per material); Rename refuses empty and duplicate names. Startup creates a `Default` when none exists, and backfills a G-code `Default` for a missing material and saves at once (keep that).
- A G-code preset never loads into another material.
- There is no stored "active preset" key: `_detect_active_preset()` names the preset whose values match the form.
- When `SIZING_PRESET_KEYS` grows, use `config.normalize_sizing_preset()`.

Tests: `test_sizing_presets_workflow`, `test_gcode_presets_workflow`. Detail: `docs/decisions/user-interface.md`.

## Settings-Dialog Tooltips

`ui_dialogs/common.py` exposes a `Tooltip` helper plus `add_tooltip(widget, text)` and `add_tooltips(text, *widgets)` convenience functions. All settings dialogs (Sizing Rules, G-code Settings, Layer Colors, Key Layout, Tuner Settings, Toner Settings) hover-explain their fields. When adding a new setting widget, attach a tooltip alongside it — short sentence, plain English, focused on *why* the user would change it. Attach to both the label and the input so users can hover either. The helper popup is overrideredirect + topmost so it appears above modal dialogs without stealing focus.

## Dart Shape Spectrum

The dart shape ("Darts" in the UI; internal names keep `dart_`) is one 0.0 to 1.0 slider: triangle, sine (0.5, the default), square. The math is `_wave_value` in svg_engine.py, used by both `calculate_star_path` and `_generate_star_points` (pinned by `test_engine_parity`). `load_settings` migrates the legacy scale once (`dart_shape_v2`). Update `tools/test_dart_shapes.py` when changing it. Detail: `docs/decisions/engraving-and-darts.md`.

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

After this, `import tuner_render` succeeds and the tuner uses GPU rendering at 60-120 fps.

Rules for the renderer (history: `docs/decisions/tuner.md`):

- Limits are `downlevel_defaults().using_resolution(adapter.limits())`; `clamp_surface()` clamps to `max_texture_dimension_2d` in `Renderer::new` and `resize()`.
- A Rust panic is a `BaseException` (`PanicException`). `tuner_tab.py` catches `BaseException` (re-raising `KeyboardInterrupt`, `SystemExit`) at every surface call and falls back through `_tuner_gpu_fallback(reason)`; `GPU_RENDER_FAIL_LIMIT` consecutive failures switch to the canvas. Never retry the GPU on the same frame.
- Present with Mailbox where offered, else Fifo. A software adapter (`adapter_info()` type `"Cpu"`) counts as no GPU.

**macOS is canvas-only: never load `tuner_render` on darwin.** Tk Aqua's `winfo_id()` is not an NSView, and wgpu's Metal backend segfaults on it natively, which no `except` can catch. Three places enforce it; keep all three: `tuner_tab.py` skips the import on darwin, `build.py` skips the `--hidden-import` on darwin, and CI skips the Rust/maturin steps on macOS runners.

On Windows, Rust needs the MSVC linker. The one-time machine setup is `winget install Rustlang.Rustup` followed by `winget install Microsoft.VisualStudio.2022.BuildTools --override "--add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"`.

### Windows Installer (Inno Setup)

CI wraps `dist\SaxShopCompanion.exe` into a versioned `SaxShopCompanion-Windows-Setup-{version}.exe` via `installer.iss`. The installer creates Start Menu + optional desktop shortcuts, registers an uninstaller, and installs to `{autopf}\SaxShopCompanion` (requires admin UAC). User data in `%APPDATA%\StohrerSaxShopCompanion\` is untouched on uninstall. Local build:

```bash
python build.py
iscc /DAppVersion=2.80 installer.iss    # requires Inno Setup 6
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

- **Never assign to `_` in a function that also calls `_()`.** `_` is gettext, installed into builtins. Python makes a name local for the whole function body if it is assigned anywhere in it, so one `for _, x in items` or `a, _, _ = f()` breaks every `_("...")` in that function. Use `_unused`, `_name`, `_label`. Nested defs, lambdas and comprehensions have their own `_`. Gated by `tools/test_gettext_shadowing.py`. History: `docs/decisions/i18n.md`.

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

`APP_VERSION` in `config.py` is the manual source of truth — bump it when preparing a release. `APP_BUILD_DATE` is auto-derived from `sys.executable`'s mtime in frozen builds (installer/zip copies preserve mtime) and falls back to the manual constant when running from source. The About dialog (ui_dialogs/guide.py) reads both.

## Release Notes Style

A release note tells someone who uses the program what changed for them, and is short (Matt, 2026-10-26): "Bug fixes." can be the whole changelog; a major fix or new feature gets one line. Title `# vX.YZ`. No background, history or engineering detail. Files: `docs/release-notes/release_notes_vX.YZ.md`, published with `gh release create ... --notes-file`; `release_notes_v2.9.md` is the model. History: `docs/decisions/build-and-release.md`.

## CI/CD (GitHub Actions)

`.github/workflows/build.yml`: `lint`, `test` (four platforms), and `build` (`needs: [lint, test]`, `fail-fast: false`: Windows installer, macOS Apple Silicon and Intel .app, Linux binary, each probed with `--selftest`). Triggers: push to `main` or `beta`, release creation, manual. Rules:

- Every CI install is `pip install --require-hashes -r locks/<set>.txt`. Edit `requirements*.txt` or `locks/*.in`, then run `python tools/update_locks.py`; the lint job fails on stale locks.
- The Intel Mac build ships svgwrite and numpy only (no sounddevice, no pyserial); `main.audio_supported()` and `main.machine_supported()` hide what it cannot run.
- Rust and maturin only where `full_build` is true, never on macOS. macOS packaging is `ditto -c -k --keepParent`, never `zip -r`.
- Tours never fail the build. A job failing at "Set up job" with no log is runner infrastructure: re-run it.

Detail: `docs/decisions/build-and-release.md`.

**Action pins**: `actions/checkout@v5`, `actions/setup-python@v6`, `actions/upload-artifact@v6` — all on Node 24. Don't downgrade; GitHub removes Node 20 from runners in September 2026.

## Config File Locations

The app stores settings and presets in platform-appropriate locations:

| Platform | Location |
|----------|----------|
| Windows | `%APPDATA%\StohrerSaxShopCompanion\` |
| macOS | `~/Library/Application Support/StohrerSaxShopCompanion/` |
| Linux | `~/.config/StohrerSaxShopCompanion/` (respects `XDG_CONFIG_HOME`) |

**Override for tests and CI**: `SAXSHOP_CONFIG_DIR` replaces the directory above when it is set before `config` is imported (`config.get_config_dir()`). Every suite that can reach the folder sets it, `tools/run_tests.py` sets one per suite, and the CI tours use it, so nothing they do touches a user's real settings. Gated by `tools/test_config_isolation.py`. History: `docs/decisions/testing-and-ci.md`.

**Backward compatibility**: On first run, existing config files in the old location (current working directory) are automatically migrated to the new location.

**Manual import**: Users can also manually import settings from a previous installation via File > "Import Settings from Folder..." which copies config files from a selected directory.

## Detailed Documentation

Architecture and domain-specific guidance is split across companion files imported via `@`-statements below. Claude Code loads them as part of CLAUDE.md's context.

- **CLAUDE-architecture.md** — Module structure, design patterns, settings/presets, error logging, feature set
- **CLAUDE-engines.md** — Pad generation, G-code, SVG rendering, nesting, strobe tuner, tooling, Phil Noy credit
- **CLAUDE-toner.md** — Tone analyzer engine, data model, capture modes, analyze tool, WAV recording, calibration
- **CLAUDE-web.md** — Web data sync, screw specs submission form, related repository

**`docs/decisions/` holds the why and the history**: what was tried and rejected, measured results, bugs and how they were found, and what Matt said, by topic (index: `docs/decisions/README.md`). These files are not imported; read the relevant one before changing an area. When you make a decision worth keeping (a rejected alternative, a measured result, something the user said), add it to the matching file, dated, and keep the CLAUDE files to rules and pointers.

---

## Subsystem imports

@CLAUDE-architecture.md
@CLAUDE-engines.md
@CLAUDE-toner.md
@CLAUDE-web.md
