# Build, CI jobs and release notes: decisions and history

This file holds the why and the history behind the rules in the CLAUDE files, and the full detail they summarise: what was tried, measured, rejected and decided, with dates and the user's own words. The paragraphs were moved here verbatim from the CLAUDE files on 2026-11-10 (HEAD 01071f5); each section names where it came from. Add new decisions to the matching section, dated.

## Release Notes Style

*Moved from CLAUDE.md, section "CLAUDE.md > Release Notes Style" (2026-11-10).*

### Written from the user's perspective, and short (Matt, 2026-10-26)

**Written from the user's perspective, and short (Matt, 2026-10-26).** A release note tells someone who uses the program what changed for them. "Bug fixes." can be the entire changelog. A specific major fix or a new feature gets its own one-line bullet saying what it is and, when it matters, where to find it. Title is `# vX.YZ`, optionally with a few words after a dash when there is a headline feature. No background, no history, no explanation of how it works or why the old behaviour was wrong, no engineering detail: that belongs in the commit messages and CLAUDE-engines.md. `docs/release-notes/release_notes_vX.YZ.md` is the GitHub release body as published (the twelve files moved there from the repo root on 2026-11-10 with `git mv`, history kept; `gh release create` takes it with `--notes-file`); `docs/release-notes/release_notes_v2.9.md` is the model (prepared as v2.87, released as v2.9: Matt judged it more than a .01).

This replaces the full-feature-overview template used from v2.1 to v2.86 (a paragraph describing the app, then every feature of every tab with `**(new)**` markers). Those notes grew with every release and spent most of their length explaining things to people who already use the program. Don't bring back the overview, the `(new in v2.X.Y)` tags, or "Fixes & polish" sections.

## CI/CD (GitHub Actions)

*Moved from CLAUDE.md, section "CLAUDE.md > CI/CD (GitHub Actions)" (2026-11-10).*

The `.github/workflows/build.yml` workflow has three jobs. `build` has `needs: [lint, test]` and `fail-fast: false`, so nothing is built or attached to a release unless every test job passed, and one platform's build failure does not cancel the others:

- **`lint`** (ubuntu-latest, ~10s): runs `ruff check .` — fails the workflow on any violation

- **`test`** (windows-latest, macos-latest, macos-15-intel, ubuntu-latest): `python tools/run_tests.py` with `requirements.txt` + `requirements-dev.txt`; Linux under `xvfb-run`; the Intel Mac with the shipped dependency set (numpy, no sounddevice) and `--allow-missing sounddevice`

- **`build`** (4-platform matrix): Windows Inno Setup installer (the bare PyInstaller .exe is built but not published — only the installer ships), macOS Apple Silicon .app, macOS Intel .app, and Linux binary. After PyInstaller each job runs the frozen-build probe (`--selftest` on the built binary); Windows also runs the installer round-trip (silent install → installed copy's selftest → Start Menu shortcut → silent uninstall → `%APPDATA%` preserved) and the German screenshot tour (`win-tour-de`); Windows and Linux run the tuner suites with the freshly built GPU wheel; both Macs run the screenshot tour in light and dark (`mac-tour-*`). Tours never fail the build.

Triggers on push to `main` or `beta`, on release creation, or manually.

- macOS Intel build (`macos-15-intel` runner) installs svgwrite+numpy+pyinstaller (no sounddevice) — tuner and toner are unavailable; nesting runs at full numpy speed (numpy added 2026-10-06 — before that the pure-Python scan took 13 minutes on one test suite there)

- `full_build: true/false` matrix flag controls whether Rust toolchain + maturin are installed for the GPU tuner renderer; the macOS runners additionally skip the Rust steps unconditionally — macOS is canvas-only (see GPU Tuner Renderer section)

- The Windows job also installs Inno Setup 6 via Chocolatey and builds the installer; version is extracted from `config.py`'s `APP_VERSION` via PowerShell regex

- macOS jobs package the `.app` with `ditto -c -k --keepParent` (never `zip -r` — it materializes bundle symlinks and breaks the code-signature resource seal) and run three guards: `plutil -extract` for the mic + camera usage keys, `codesign --verify --deep --strict` on the built .app, and the same verify on an unzipped copy of the final artifact (see "macOS Build" in CLAUDE-architecture.md)

- Uploads artifacts to the workflow run; on Windows that's the installer only (`SaxShopCompanion-Windows-Setup-*.exe`). All published artifacts also attach to GitHub Releases on release events.

## macOS Build

*Moved from CLAUDE-architecture.md, section "Architecture > macOS Build" (2026-11-10).*

The keys alone are **not sufficient**: PyInstaller ad-hoc signs the bundle during build, and Info.plist is sealed into that signature. Patching the plist afterwards breaks the seal, and macOS TCC refuses to show a permission prompt for any app whose signature doesn't validate — same silent-denial symptom, keys present. Every shipped Apple Silicon build through v2.62 had a broken seal (verified against the released zip: sealed plist hash ≠ actual plist hash). `_resign_macos_app()` in build.py therefore re-signs ad-hoc (`codesign --force --deep --sign -`) immediately after the patch and verifies with `--strict`; any future work on the macOS build must keep the re-sign as the last step that touches the bundle.

Packaging must preserve symlinks too: CI zips the `.app` with `ditto -c -k --keepParent`, because `zip -r` follows the bundle's internal Frameworks↔Resources symlinks and stores them as duplicate real files (767 of them in the v2.62 zip, roughly doubling the download), which breaks the signature's resource seal on extraction. The `--dmg` path is fine as-is — hdiutil preserves symlinks. CI verifies all of it on both macOS runners: `plutil -extract` for both usage keys, `codesign --verify --deep --strict` on the built .app, and again on an unzipped copy of the final artifact. Users upgrading from a broken build may need `tccutil reset Microphone com.stohrer.saxshopcompanion` (and `Camera`) if macOS cached the old silent denial — the README documents this.
