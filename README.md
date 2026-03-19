# Stohrer Sax Shop Companion

A cross-platform desktop app for saxophone repair technicians. Generates SVG and G-code files for laser-cutting pad materials, provides reference databases for key heights, serial numbers, and screw specifications, and includes tooling generation, a chromatic strobe tuner, and a harmonic tone analyzer.

## Features

### Pad Generator
- Generate cut patterns for **felt**, **card**, **leather**, and **exact size** pads
- **SVG output** for laser cutters and plotters
- **G-code output** for Grbl-based lasers with per-layer power/speed/passes settings
- Smart nesting algorithm to maximize material usage
- **Star/dart patterns** for small leather pads (configurable threshold)
- Configurable sizing rules, center holes, and engraving labels
- **Filled engraving mode** — scan-line raster fill using Roboto font outlines, with overscan optimization
- **Auto-fit engraving** — shifts text toward center on small pads, scales only as last resort
- **Air assist toggles** (M8/M9) per layer
- **Cut grouping**: "by layer" or "by pad" ordering
- **Max fill mode**: use `18.0 x max` to fill remaining space with a pad size
- **Edge bias**: d-pad control to direct circle packing toward a specific edge or corner

### Custom Polygon Shapes
- Draw irregular shapes for leather skins and scrap pieces (up to 8 points)
- Grid adapts to unit setting (15x15 inches or 40x40 cm)
- Smart nesting: large pads go center, small pads fill edges/corners
- Two fill strategies: center-out or longest-edge priority

### Scrap Mode
- Place pads across multiple irregular scrap pieces instead of requiring one large sheet
- Progress popup tracks remaining and completed pads
- Files named with `_scrap1`, `_scrap2` suffixes

### Pad Presets
- Save and load pad size lists organized into libraries
- **Notes field** for annotating presets with context (source, instrument details)
- **Import Matt's Pad Sets** — download reference pad sets from the [Pad Sets Library](https://www.stohrermusic.com/articles/pad-sets-library/)
- Import/export preset files for sharing with colleagues

### SD Card Workflow (Windows)
- **Eject SD card** checkbox — auto-ejects removable drive after G-code export
- **Send G-code to SD Card** — guided workflow to copy, clean, and eject

### Key Height Library
- Store and organize key height measurements by instrument
- **Import Matt's Key Heights** — download reference data from the [Key Heights Library](https://www.stohrermusic.com/articles/key-heights-library/)
- Import/export libraries for backup or sharing

### Serial Number Lookup
- Reference database for saxophone serial numbers by manufacturer

### Screw Specifications
- OEM screw and rod specifications database
- **Import Matt's Specs** — download reference data from the [Screw Specs Library](https://www.stohrermusic.com/articles/screw-specs-library/)
- Import/export for sharing specs

### Tooling — Die Inserts & Die Holders
- Generate SVG/G-code for laser-cut **acrylic pad die inserts** (small: 50mm OD, large: 70mm OD)
- **Die holders**: 85mm OD, 4-layer stack with magnet and alignment holes
- Enter individual sizes, ranges, or generate full sets
- **Kerf test pattern** — calibrate your laser's kerf width with a simple cut-and-measure workflow
- Scrap mode for spreading dies across multiple sheets of acrylic

### Chromatic Strobe Tuner (Experimental)
- 12-wheel strobe tuner modeled after the **Peterson Stroboconn 6T-5**
- Per-ring octave brightness from real spectral data
- Transposition support (Concert, Bb, Eb, F)
- Reference tone player (pure or rich)
- Analog VU meter with damped needle
- Configurable stripe color, faceplate color, frame rate

### Harmonic Tone Analyzer (Experimental)
- Real-time **harmonic spectrum analyzer** for saxophone
- Detects fundamental pitch, extracts up to 12 harmonics
- **Spectrum view** (full FFT) and **Bars view** (per-harmonic), with linear or dB scale
- Five VU-style gauges: intonation, resonance, richness, brightness, darkness
- Brightness/darkness uses **Benade break frequencies** adapted per saxophone type
- **Tone profiles**: capture harmonic fingerprints of individual horns
  - Three capture modes: structured (held notes), free (natural playing), file import (WAV)
  - Profiles are a fixed setup: horn + player + mouthpiece + reed
  - Organized into libraries with import/export
- **Comparison tool**: multi-select profiles, filter by type/player/mouthpiece, side-by-side analysis with per-note and horn-average views
- Auto-transposition by saxophone type with concert pitch toggle
- Bias sliders for subjective gauge calibration

### Feature Set
- **File > Feature Set** — choose which tabs to show
- Experimental features (Tuner, Toner) hidden by default; opt in when ready

### Cross-Platform
- Runs on **Windows**, **macOS**, and **Linux**
- **macOS dark mode** support — uses native system colors
- Platform-appropriate config storage with automatic migration
- **Import Settings from Folder** for moving between machines

## Installation

### From Release (Recommended)
Download the latest release for your platform from the [Releases](https://github.com/stohrermusic/Stohrer-Sax-Shop-Companion/releases) page.

**macOS users**: The app is not signed with an Apple Developer certificate, so macOS will block it on first launch. To open it, right-click (or Control-click) the app and select **Open**, then click **Open** in the dialog. You only need to do this once — after that it launches normally.

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
