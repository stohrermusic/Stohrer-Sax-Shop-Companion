# Stohrer Sax Shop Companion

A cross-platform desktop app for saxophone repair technicians. Generates SVG files for laser-cutting pad materials and provides reference databases for key heights, serial numbers, and screw specifications.

## Features

### Pad SVG Generator
- Generate cut patterns for **felt**, **card**, **leather**, and **exact size** pads
- Smart nesting algorithm to maximize material usage
- **Star/dart patterns** for small leather pads (configurable threshold)
- Configurable sizing rules, center holes, and engraving labels
- Save and load pad presets organized by library

### Custom Polygon Shapes (New!)
- Draw irregular shapes for leather skins and scrap pieces
- 15x15 grid for precise shape definition (up to 8 points)
- Smart nesting: large pads go center, small pads fill edges/corners
- **Max fill mode**: use `18.0 x max` to fill remaining space with a pad size
- Two fill strategies: center-out or longest-edge priority

### Key Height Library
- Store and organize key height measurements by instrument
- Import/export libraries for backup or sharing

### Serial Number Lookup
- Reference database for saxophone serial numbers by manufacturer

### Screw Specifications
- OEM screw and rod specifications database
- Import/export for sharing specs

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
