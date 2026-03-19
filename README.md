# Stohrer Sax Shop Companion

A cross-platform desktop app for saxophone repair technicians. Generates SVG and G-code files for laser-cutting pad materials, provides reference databases for key heights, serial numbers, and screw specifications, and includes tooling generation, a chromatic strobe tuner, and a harmonic tone analyzer.

## Features

### Pad Generator

#### Output Formats
- SVG for laser cutters and plotters (LightBurn color layers)
- G-code for Grbl-based lasers with per-material speed, power, and passes
- Generate felt, card, leather, and exact size patterns

#### Nesting & Layout
- Smart circle-packing algorithm to maximize material usage
- Edge bias d-pad to direct packing toward a specific edge or corner
  - Cardinal directions scan from the edge inward
  - Corner directions radiate outward — small pads nestle in the corner, larger ones fan out
- Nesting preview — see the layout before generating, adjust and retry
- Custom polygon shapes for irregular leather skins and scrap pieces
- Scrap mode — place pads across multiple pieces, tracking remaining between sheets
  - Preview, edge bias, and custom polygons all work together in scrap mode
- Max fill mode — use `18.0 x max` to fill remaining space with a size

#### Engraving
- Pad size labels on each disc, with per-material font size and position
- Line mode (single-stroke) or filled mode (scan-line raster fill with overscan)
- Auto-fit — shifts text toward center on small pads, scales only as last resort

#### G-code Options
- Air assist toggles (M8/M9) per layer
- Cut grouping: by layer or by pad
- Kerf compensation (enter full kerf, app splits in half automatically)
- Auto-eject SD card after export (Windows)

#### Presets & Data
- Save and load pad size lists organized into libraries
- Notes field for annotating presets
- Import Matt's Pad Sets from [stohrermusic.com](https://www.stohrermusic.com/articles/pad-sets-library/)
- Import/export for sharing with colleagues

### Key Height Library
- Store and organize key height measurements by instrument
- Import Matt's Key Heights from [stohrermusic.com](https://www.stohrermusic.com/articles/key-heights-library/)
- Import/export libraries for backup or sharing

### Serial Number Lookup
- Reference database for saxophone serial numbers by manufacturer

### Screw Specifications
- OEM screw and rod specifications database
- Import Matt's Specs from [stohrermusic.com](https://www.stohrermusic.com/articles/screw-specs-library/)
- Import/export for sharing specs

### Tooling — Die Inserts & Die Holders
- Generate SVG/G-code for laser-cut acrylic pad die inserts (small: 50mm OD, large: 70mm OD)
- Die holders: 85mm OD, 4-layer stack with magnet and alignment holes
- Enter individual sizes, ranges, or generate full sets
- Kerf test pattern for calibrating your laser
- Scrap mode for spreading dies across multiple sheets

### Chromatic Strobe Tuner (Experimental)
- 12-wheel stroboscopic tuner
- Per-ring octave brightness from real spectral data
- Transposition support (Concert, Bb, Eb, F)
- Reference tone player, analog VU meter, configurable colors and frame rate
- A quality microphone is recommended (e.g. Audio-Technica AT2020 USB)

### Harmonic Tone Analyzer (Experimental)
- Real-time harmonic spectrum analyzer for saxophone
- Detects fundamental pitch, extracts up to 12 harmonics
- Spectrum view (full FFT) and Bars view (per-harmonic), with linear or dB scale
- Five VU-style gauges: intonation, resonance, richness, brightness, darkness
- Brightness/darkness adapted per saxophone type using Benade break frequencies
- Tone profiles: capture harmonic fingerprints of individual horns
  - Four capture modes: structured, free, calibration (guided chromatic scale), file import (WAV)
  - Profiles are a fixed setup: horn + player + mouthpiece + reed
  - Organized into libraries with import/export
- Comparison tool with filtering, per-note and horn-average views
- Auto-transposition by saxophone type with concert pitch toggle
- A quality microphone is essential — laptop mics are not recommended for tone analysis

### General
- Feature Set (File > Feature Set) — choose which tabs to show
- Runs on Windows, macOS, and Linux
- macOS dark mode support
- Platform-appropriate config storage with automatic migration
- Import Settings from Folder for moving between machines

## Installation

### From Release (Recommended)
Download the latest release for your platform from the [Releases](https://github.com/stohrermusic/Stohrer-Sax-Shop-Companion/releases) page.

**macOS users**: The app is not signed with an Apple Developer certificate, so macOS will block it on first launch. To open it, right-click (or Control-click) the app and select **Open**, then click **Open** in the dialog. You only need to do this once — after that it launches normally.

If the DMG shows as "damaged," open Terminal and run:
```
xattr -cr ~/Downloads/StohrerSaxShopCompanion.dmg
```
Then open the DMG again. This strips the quarantine flag that macOS adds to downloaded files.

### From Source
```bash
pip install -r requirements.txt
python main.py
```

## Building

```bash
# Build for current platform
python build.py

# Clean and rebuild
python build.py --clean

# macOS: create .dmg
python build.py --dmg
```

## Config Location

Settings and presets are stored in platform-appropriate locations:
- **Windows**: `%APPDATA%\StohrerSaxShopCompanion\`
- **macOS**: `~/Library/Application Support/StohrerSaxShopCompanion/`
- **Linux**: `~/.config/StohrerSaxShopCompanion/`

---

Made for saxophone techs, by a saxophone tech.
