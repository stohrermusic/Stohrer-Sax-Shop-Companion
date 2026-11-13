# Decisions and history

The CLAUDE files at the repo root hold the rules and where things are. This folder holds the why: what was tried and rejected, what was measured, how bugs were found, and what Matt said, by topic. These files are not loaded into every session. Read the relevant one before changing an area, and add new decisions to it, dated.

| File | What it holds |
|---|---|
| [testing-and-ci.md](testing-and-ci.md) | CI test jobs, locked dependencies, the Intel Mac, Mac and Windows coverage, the frozen-build probe, the screenshot tour, settings isolation, writing a check, the message-box guard, serial ports, hangs, the exception audit. |
| [build-and-release.md](build-and-release.md) | The CI build job in detail, macOS signing and packaging history, the release notes style and the template it replaced. |
| [labeled-zones.md](labeled-zones.md) | Why zones exist, their cost, the packers, grid shape, rejected alternatives, invariants with their history, test fixtures. |
| [adhesive-rings.md](adhesive-rings.md) | Ring geometry, the Wrap clearance naming, the number rule, the footprint trap, laser defaults, machine control, the three test suites and their deliberate breaks. |
| [machine-integration.md](machine-integration.md) | The 3.0 work, gating history, camera calibration hardening, Frame & Cut, the lid question's two corrections, framing, the laser watch, the camera fallback, Tooling Frame & Cut and the five older bugs. The machine half of main.py as a mixin (structural step 2). |
| [tooling.md](tooling.md) | Section buttons, die nesting by rows, zones and dies, the Die Organizer G-code, the Tonehole Crescent, Speed & Power, Tooling material names and legends, cork, the G-code file line. |
| [nesting-and-scrap.md](nesting-and-scrap.md) | Vectorized scans, the linear-scan hang, the nesting preview, the polygon tool, Scrap Mode's live list and pending scrap, the removed large-batch optimizer. |
| [engraving-and-darts.md](engraving-and-darts.md) | Dart shape, engraving range lists, size-label placement modes, SVG output modes, fonts and arc text, leather locator marks. |
| [tuner.md](tuner.md) | The GPU renderer traps, macOS canvas-only history, the Hann peak-frequency fix, tuner test hooks. |
| [toner.md](toner.md) | Removed descriptor gauges, pitch reading measurements, rolloff readings per mic, the Analyze tool in full, the capture filter audit, descriptor calibration status. |
| [user-interface.md](user-interface.md) | Job History, the User Guide upgrade, the Welcome window and Quick Start Guide, pad preview, both presets workflows, SD card. The ui_dialogs package (structural step 1). |
| [i18n.md](i18n.md) | The gettext `_` shadowing bugs. |
| [settings.md](settings.md) | The settings schema and validation at load (structural step 4): where the ranges came from, what the first scan found, what the scan cannot see, presets. |
| [web-data.md](web-data.md) | Web data imports and the TLS fix. |
