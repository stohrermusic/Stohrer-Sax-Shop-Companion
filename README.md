# Stohrer Sax Shop Companion

A cross-platform desktop app for saxophone repair technicians. Generates SVG and G-code files for laser-cutting pad materials and provides reference databases for key heights, serial numbers, and screw specifications.

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
- **Import Matt's Pad Sets** — download reference pad sets from stohrermusic.com
- Import/export preset files for sharing with colleagues

### SD Card Workflow (Windows)
- **Eject SD card** checkbox — auto-ejects removable drive after G-code export
- **Send G-code to SD Card** — guided workflow to copy, clean, and eject

### Key Height Library
- Store and organize key height measurements by instrument
- **Import Matt's Key Heights** — download reference data from stohrermusic.com
- Import/export libraries for backup or sharing

### Serial Number Lookup
- Reference database for saxophone serial numbers by manufacturer

### Screw Specifications
- OEM screw and rod specifications database
- **Import Matt's Specs** — download reference data from stohrermusic.com
- Import/export for sharing specs

### Cross-Platform
- Runs on **Windows**, **macOS** (universal binary: Intel + Apple Silicon), and **Linux**
- **macOS dark mode** support — uses native system colors
- Platform-appropriate config storage with automatic migration
- **Import Settings from Folder** for moving between machines

## Installation

### From Release (Recommended)
Download the latest release for your platform from the [Releases](https://github.com/stohrermusic/Stohrer-Sax-Shop-Companion/releases) page.

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
