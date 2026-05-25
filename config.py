import os
import sys
import json
import copy
import shutil
import ssl
import logging
import logging.handlers
from tkinter import messagebox


def get_dart_settings_for_size(pad_size, settings):
    """Return dart settings dict for the given pad size, or None if no stars.

    In universal mode: returns global dart settings if darts_enabled and pad_size < threshold.
    In range mode: finds matching range (min_size <= pad_size <= max_size), returns it or None.
    """
    if not settings.get("darts_enabled", True):
        return None

    mode = settings.get("dart_range_mode", "universal")

    if mode == "range":
        for r in settings.get("dart_ranges", []):
            if r["min_size"] <= pad_size <= r["max_size"]:
                return r
        return None
    else:  # universal
        threshold = settings.get("dart_threshold", 18.0)
        if pad_size < threshold:
            return {
                "overwrap": settings.get("dart_overwrap", 0.5),
                "wrap_bonus": settings.get("dart_wrap_bonus", 0.75),
                "frequency_multiplier": settings.get("dart_frequency_multiplier", 1.0),
                "shape_factor": settings.get("dart_shape_factor", 0.5),
                "engraving_on": settings.get("dart_engraving_on", True),
                "engraving_loc": settings.get("dart_engraving_loc", {"mode": "from_outside", "value": 2.5}),
            }
        return None


def get_sizing_for_size(pad_size, settings):
    """Return sizing dict for the given pad size.

    In range mode: finds matching range, or falls back to universal.
    In universal mode: returns the global sizing settings.
    """
    if settings.get("sizing_range_mode", "universal") == "range":
        for r in settings.get("sizing_ranges", []):
            if r["min_size"] <= pad_size <= r["max_size"]:
                return r
    return {
        "felt_offset": settings.get("felt_offset", 0.75),
        "card_to_felt_offset": settings.get("card_to_felt_offset", 0.5),
        "leather_wrap_multiplier": settings.get("leather_wrap_multiplier", 1.0),
        "min_hole_size": settings.get("min_hole_size", 16.5),
        "felt_thickness": settings.get("felt_thickness", 3.175),
        "felt_thickness_unit": settings.get("felt_thickness_unit", "mm"),
    }


def get_engraving_settings_for_size(pad_size, settings):
    """Return engraving on/off and font sizes for the given pad size.

    In range mode: finds matching range, or falls back to universal.
    In universal mode: returns the global engraving settings.
    """
    if settings.get("engraving_settings_range_mode", "universal") == "range":
        for r in settings.get("engraving_settings_ranges", []):
            if r["min_size"] <= pad_size <= r["max_size"]:
                return r
    return {
        "engraving_on": settings.get("engraving_on", True),
        "engraving_font_size": settings.get("engraving_font_size", {
            "felt": 3.0, "card": 3.0, "leather": 3.0, "exact_size": 3.0
        }),
    }


def get_engraving_placement_for_size(pad_size, settings):
    """Return engraving placement (location per material) for the given pad size.

    In range mode: finds matching range, or falls back to universal.
    In universal mode: returns the global engraving location settings.
    """
    if settings.get("engraving_placement_range_mode", "universal") == "range":
        for r in settings.get("engraving_placement_ranges", []):
            if r["min_size"] <= pad_size <= r["max_size"]:
                return r
    return {
        "engraving_location": settings.get("engraving_location", DEFAULT_SETTINGS.get("engraving_location")),
    }


def get_ssl_context():
    """Get an SSL context that works on macOS (and everywhere else).

    macOS Python often can't find system certificates. This tries:
    1. certifi package (if installed)
    2. Default system context
    3. Unverified fallback (last resort)
    """
    # Try certifi first (most reliable on macOS)
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        pass

    # Try default context
    ctx = ssl.create_default_context()
    try:
        # Test if it can actually verify
        ctx.load_default_certs()
        return ctx
    except Exception:
        pass

    # Fallback: unverified (still encrypted, just no cert check)
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx

# ==========================================
# PLATFORM-SPECIFIC CONFIG DIRECTORY
# ==========================================

APP_NAME = "StohrerSaxShopCompanion"


def get_input_devices():
    """Return list of (device_index, device_name) for usable audio input devices.

    Filters out:
    - Bluetooth headsets (too low sample rate for audio analysis)
    - Windows sound mapper duplicates
    - Devices that can't do 44100 Hz
    """
    try:
        import sounddevice as sd
        devices = []
        seen_names = set()
        for i, d in enumerate(sd.query_devices()):
            if d['max_input_channels'] <= 0:
                continue

            name = d['name']
            name_lower = name.lower()

            # Skip Bluetooth headsets (HSP/HFP = 16kHz, useless for analysis)
            if 'bluetooth' in name_lower or 'hands-free' in name_lower:
                continue
            if 'bthhfenum' in name_lower:
                continue

            # Skip Windows sound mapper duplicates
            if 'sound mapper' in name_lower:
                continue
            if 'primary sound' in name_lower:
                continue

            # Skip if sample rate too low
            if d.get('default_samplerate', 0) < 44100:
                continue

            # Deduplicate (same device from different host APIs)
            # Normalize by taking just the meaningful part of the name
            clean_name = name.strip()
            # Windows often appends truncated driver names — dedupe on first 30 chars
            dedup_key = clean_name[:30].strip()
            if dedup_key in seen_names:
                continue
            seen_names.add(dedup_key)

            devices.append((i, clean_name))
        return devices
    except Exception:
        return []
APP_VERSION = "3.0"

def _detect_build_date():
    # In a PyInstaller-frozen build, the exe's mtime is the build time —
    # preserved across zip/installer copies on all three platforms.
    # Falls back to the manual date when running from source.
    manual = "2026-05-10"
    if getattr(sys, 'frozen', False):
        try:
            import datetime
            mtime = os.path.getmtime(sys.executable)
            return datetime.date.fromtimestamp(mtime).isoformat()
        except Exception:
            pass
    return manual

APP_BUILD_DATE = _detect_build_date()

def get_config_dir():
    """
    Returns the platform-appropriate config directory.
    - Windows: %APPDATA%/StohrerSaxShopCompanion/
    - macOS: ~/Library/Application Support/StohrerSaxShopCompanion/
    - Linux: ~/.config/StohrerSaxShopCompanion/ (respects XDG_CONFIG_HOME)
    """
    if sys.platform == 'win32':
        base = os.environ.get('APPDATA', os.path.expanduser('~'))
        config_dir = os.path.join(base, APP_NAME)
    elif sys.platform == 'darwin':
        config_dir = os.path.join(os.path.expanduser('~'), 'Library', 'Application Support', APP_NAME)
    else:  # Linux and other Unix-like
        base = os.environ.get('XDG_CONFIG_HOME', os.path.join(os.path.expanduser('~'), '.config'))
        config_dir = os.path.join(base, APP_NAME)

    return config_dir


def ensure_config_dir():
    """Create the config directory if it doesn't exist."""
    config_dir = get_config_dir()
    if not os.path.exists(config_dir):
        os.makedirs(config_dir, exist_ok=True)
    return config_dir


LOG_FILE = None  # Set by setup_logging()

def setup_logging():
    """Set up file logging for error diagnostics.

    Log file goes in the config directory. Rotates at 500KB, keeps 1 backup.
    Returns the log file path.
    """
    global LOG_FILE
    config_dir = ensure_config_dir()
    LOG_FILE = os.path.join(config_dir, "app.log")

    logger = logging.getLogger()
    logger.setLevel(logging.WARNING)

    # Rotating file handler: 500KB max, 1 backup
    handler = logging.handlers.RotatingFileHandler(
        LOG_FILE, maxBytes=500_000, backupCount=1, encoding='utf-8')
    handler.setFormatter(logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S'))
    logger.addHandler(handler)

    # Log app startup at WARNING level so it always appears
    logging.warning("App starting — version %s", APP_VERSION)
    return LOG_FILE


def get_log_file():
    """Return the path to the log file."""
    return LOG_FILE


def migrate_legacy_files():
    """
    Migrate config files from old location (CWD) to new platform-specific location.
    This preserves backward compatibility for existing users.
    """
    config_dir = ensure_config_dir()
    legacy_files = [
        "pad_presets.json",
        "key_height_library.json",
        "app_settings.json",
        "screw_specs.json",
        "tone_profiles.json",
        "toner_data.json",
    ]

    migrated = []
    for filename in legacy_files:
        legacy_path = os.path.join(os.getcwd(), filename)
        new_path = os.path.join(config_dir, filename)

        # Only migrate if legacy file exists and new file doesn't
        if os.path.exists(legacy_path) and not os.path.exists(new_path):
            try:
                shutil.copy2(legacy_path, new_path)
                migrated.append(filename)
            except Exception as e:
                print(f"Warning: Could not migrate {filename}: {e}")

    if migrated:
        print(f"Migrated config files to {config_dir}: {', '.join(migrated)}")

    return migrated


# Run migration on module load
_migrated_files = migrate_legacy_files()


def find_config_files_in_directory(directory):
    """Returns list of config filenames found in the specified directory."""
    config_files = [
        "app_settings.json",
        "pad_presets.json",
        "key_height_library.json",
        "screw_specs.json",
        "toner_data.json",
        "tone_profiles.json",
    ]
    found = []
    for filename in config_files:
        if os.path.exists(os.path.join(directory, filename)):
            found.append(filename)
    return found


def import_config_files(source_dir, filenames):
    """Copy specified config files from source_dir to the app's config directory."""
    config_dir = ensure_config_dir()
    for filename in filenames:
        src = os.path.join(source_dir, filename)
        dst = os.path.join(config_dir, filename)
        shutil.copy2(src, dst)


# ==========================================
# CONSTANTS & FILE PATHS
# ==========================================

_CONFIG_DIR = get_config_dir()
PAD_PRESET_FILE = os.path.join(_CONFIG_DIR, "pad_presets.json")
KEY_PRESET_FILE = os.path.join(_CONFIG_DIR, "key_height_library.json")
SETTINGS_FILE = os.path.join(_CONFIG_DIR, "app_settings.json")
SCREW_SPECS_FILE = os.path.join(_CONFIG_DIR, "screw_specs.json")
TONER_DATA_FILE = os.path.join(_CONFIG_DIR, "toner_data.json")
SIZING_PRESET_FILE = os.path.join(_CONFIG_DIR, "sizing_presets.json")

# Settings keys captured by a sizing-rules preset (everything in the
# Options > Sizing Rules dialog).
SIZING_PRESET_KEYS = (
    # Sizing
    "units", "felt_offset", "card_to_felt_offset", "leather_wrap_multiplier",
    "min_hole_size", "felt_thickness", "felt_thickness_unit",
    "sizing_range_mode", "sizing_ranges",
    # Darts
    "darts_enabled", "dart_range_mode", "dart_threshold", "dart_overwrap",
    "dart_wrap_bonus", "dart_frequency_multiplier", "dart_shape_factor",
    "dart_ranges", "dart_engraving_on",
    # Engraving
    "engraving_on", "engraving_font_size",
    "engraving_settings_range_mode", "engraving_settings_ranges",
    "engraving_location",
    "engraving_placement_range_mode", "engraving_placement_ranges",
    # Export compatibility (lives in this dialog)
    "compatibility_mode",
)


def settings_to_sizing_preset(settings):
    """Pull just the sizing-preset keys out of a full settings dict.

    Used to bootstrap a Default preset on first run when the preset
    library is empty, so we can guarantee at least one preset exists.
    """
    out = {}
    for key in SIZING_PRESET_KEYS:
        if key in settings:
            out[key] = copy.deepcopy(settings[key])
    return out

# Auto-migrate old filename → new
_old_toner_file = os.path.join(_CONFIG_DIR, "tone_profiles.json")
if os.path.exists(_old_toner_file) and not os.path.exists(TONER_DATA_FILE):
    try:
        os.rename(_old_toner_file, TONER_DATA_FILE)
    except OSError:
        pass  # fallback: load_tone_presets will try old path

COOL_BLUE = "#E0F7FA"
COOL_GREEN = "#E8F5E9"

LIGHTBURN_COLORS = [
    ("00 - Black", "#000000"), ("01 - Blue", "#0000FF"), ("02 - Red", "#FF0000"),
    ("03 - Green", "#00E000"), ("04 - Yellow", "#D0D000"), ("05 - Orange", "#FF8000"),
    ("06 - Cyan", "#00E0E0"), ("07 - Magenta", "#FF00FF"), ("08 - Light Gray", "#B4B4B4"),
    ("09 - Dark Blue", "#0000A0"), ("10 - Dark Red", "#A00000"), ("11 - Dark Green", "#00A000"),
    ("12 - Dark Yellow", "#A0A000"), ("13 - Brown", "#C08000"), ("14 - Light Blue", "#00A0FF"),
    ("15 - Dark Magenta", "#A000A0"), ("16 - Gray", "#808080"), ("17 - Periwinkle", "#7D87B9"),
    ("18 - Rose", "#BB7784"), ("19 - Cornflower", "#4A6FE3"), ("20 - Cerise", "#D33F6A"),
    ("21 - Light Green", "#8CD78C"), ("22 - Tan", "#F0B98D"), ("23 - Pink", "#F6C4E1"),
    ("24 - Lavender", "#FA9ED4"), ("25 - Purple", "#500A78"), ("26 - Ochre", "#B45A00"),
    ("27 - Teal", "#004754"), ("28 - Mint", "#86FA88"), ("29 - Pale Yellow", "#FFDB66")
]

ALL_KEY_HEIGHT_FIELDS = [
    "B", "F", "Palm F", "Palm E", "Palm Eb", "Palm D", 
    "G", "D", "Low C", "Low B", "Low Bb"
]

def get_resonance_messages():
    """Return the list of motivational easter-egg strings.

    Returned by a function (not a module-level list) so each call resolves
    against the *current* translation catalog. This matters if/when the
    catalog ever gets reloaded after import — module-level constants would
    bake the source-language strings on first import.
    """
    return [
        _("Resonance added!"), _("Pad resonance increased!"), _("More resonance now!"),
        _("Timbral focus enhanced!"), _("Harmonic alignment optimized!"), _("Acoustic reflection matrix calibrated!"),
        _("Core vibrations synchronized!"), _("Nodal points stabilized!"), _("Overtone series enriched!"),
        _("Sonic clarity has been improved!"), _("Relacquer devaluation reversed!"), _("Heavy mass screws ain't SHIT!"),
        _("Now you don't even have to fit the neck!"), _("Let's call this the ULTRAhaul!"), _("Now safe to use hot glue!"),
        _("Look at me! I am the resonator now!")
    ]

# ==========================================
# DEFAULT SETTINGS
# ==========================================

DEFAULT_SETTINGS = {
    "language": "en",  # UI language code: en, es, de, fr, it (see i18n.LANGUAGE_NAMES)
    "units": "in",
    "felt_offset": 0.75,
    "card_to_felt_offset": 2.0,
    "leather_wrap_multiplier": 1.00,
    "sheet_width": "13.5",
    "sheet_height": "10",
    "hole_option": "3.5mm",
    "custom_hole_size": "4.0",
    "min_hole_size": 16.5,
    "felt_thickness": 3.175,
    "felt_thickness_unit": "mm",
    "engraving_on": True,
    "show_engraving_warning": True,
    "last_output_dir": "",
    "resonance_clicks": 0, 
    "compatibility_mode": False,
    
    # DART / STAR SETTINGS
    "darts_enabled": True,
    "dart_threshold": 18.0,
    "dart_overwrap": 0.5,
    "dart_wrap_bonus": 0.75,
    "dart_frequency_multiplier": 1.0,
    "dart_shape_factor": 0.5,  # 0.0 = Triangle, 0.5 = Sine, 1.0 = Square
    "dart_shape_v2": True,  # marks the new triangle/sine/square spectrum (legacy=False/missing)

    "dart_engraving_on": True,
    "dart_engraving_loc": {"mode": "from_outside", "value": 2.5},

    # DART RANGE MODE: "universal" uses global settings above, "range" uses per-size ranges
    "dart_range_mode": "universal",
    "dart_ranges": [],  # list of {"min_size", "max_size", "overwrap", "wrap_bonus", "frequency_multiplier", "shape_factor", "engraving_on", "engraving_loc"}

    # SIZING RANGE MODE
    "sizing_range_mode": "universal",
    "sizing_ranges": [],  # list of {"min_size", "max_size", "felt_offset", "card_to_felt_offset", "leather_wrap_multiplier", "min_hole_size", "felt_thickness", "felt_thickness_unit"}

    # ENGRAVING SETTINGS RANGE MODE
    "engraving_settings_range_mode": "universal",
    "engraving_settings_ranges": [],  # list of {"min_size", "max_size", "engraving_on", "engraving_font_size": {material: size}}

    # ENGRAVING PLACEMENT RANGE MODE
    "engraving_placement_range_mode": "universal",
    "engraving_placement_ranges": [],  # list of {"min_size", "max_size", "engraving_location": {material: {mode, value}}}

    # EDGE BIAS - direction to bias circle packing toward
    # "center" (no bias), "n", "ne", "e", "se", "s", "sw", "w", "nw"
    "edge_bias": "center",

    # CARD PAPER SIZE SETTINGS
    "card_use_paper_size": False,  # When True, card uses standard paper size instead of sheet dimensions
    "card_paper_size": "letter",   # "letter" (8.5x11 in) or "a4" (210x297 mm)

    # G-CODE OUTPUT SETTINGS
    "gcode_output_enabled": False,  # When True, generate G-code files alongside SVGs
    "gcode_settings": {
        "felt": {
            "engraving_mode": "line",          # "line" or "filled"
            "engraving_speed": 1200,           # mm/min (line mode)
            "engraving_power": 8,              # percent (line mode)
            "engraving_passes": 1,             # repeat count (line mode)
            "filled_engraving_speed": 1200,    # mm/min ("filled" mode)
            "filled_engraving_power": 8,       # percent ("filled" mode)
            "filled_engraving_passes": 1,      # repeat count ("filled" mode)
            "filled_line_spacing": 0.15,       # mm between scan lines
            "hole_speed": 300,
            "hole_power": 35,
            "hole_passes": 1,
            "cut_speed": 600,
            "cut_power": 60,
            "cut_passes": 1,
            "kerf_width": 0.5,
            "air_assist_engraving": True,
            "air_assist_filled_engraving": True,
            "air_assist_hole": True,
            "air_assist_cut": True,
        },
        "card": {
            "engraving_mode": "line",
            "engraving_speed": 1500,
            "engraving_power": 10,
            "engraving_passes": 1,
            "filled_engraving_speed": 1200,
            "filled_engraving_power": 15,
            "filled_engraving_passes": 1,
            "filled_line_spacing": 0.15,
            "hole_speed": 400,
            "hole_power": 22.5,
            "hole_passes": 1,
            "cut_speed": 1500,
            "cut_power": 50,
            "cut_passes": 1,
            "kerf_width": 0.2,
            "air_assist_engraving": True,
            "air_assist_filled_engraving": True,
            "air_assist_hole": True,
            "air_assist_cut": True,
        },
        "leather": {
            "engraving_mode": "line",
            "engraving_speed": 2200,
            "engraving_power": 5,
            "engraving_passes": 1,
            "filled_engraving_speed": 1800,
            "filled_engraving_power": 8,
            "filled_engraving_passes": 1,
            "filled_line_spacing": 0.15,
            "hole_speed": 300,
            "hole_power": 30,
            "hole_passes": 1,
            "cut_speed": 1200,
            "cut_power": 75,
            "cut_passes": 1,
            "kerf_width": 0.3,
            "air_assist_engraving": True,
            "air_assist_filled_engraving": True,
            "air_assist_hole": True,
            "air_assist_cut": True,
        },
        "acrylic": {
            "engraving_mode": "filled",
            "engraving_speed": 4000,
            "engraving_power": 26,
            "engraving_passes": 1,
            "filled_engraving_speed": 4000,
            "filled_engraving_power": 26,
            "filled_engraving_passes": 1,
            "filled_line_spacing": 0.15,
            "hole_speed": 180,
            "hole_power": 100,
            "hole_passes": 1,
            "cut_speed": 180,
            "cut_power": 100,
            "cut_passes": 1,
            "kerf_width": 0.15,
            "air_assist_engraving": True,
            "air_assist_filled_engraving": True,
            "air_assist_hole": True,
            "air_assist_cut": True,
        },
    },

    # TUNER SETTINGS
    "tuner_settings": {
        "stripe_color": "#00FF00",
        "reference_pitch": 440.0,
        "transposition": "C",        # C, Eb, Bb, F
        "sensitivity": 50,           # 0-100
        "waveform": "pure",          # pure or rich
        "fps": "60",                 # 60, 90, or 120
        "ring_brightness": 100,      # Per-ring brightness effect 0-100 (0=uniform, 100=full)
        "overall_brightness": 80,    # Overall brightness 0-100
        "octave_boost": 50,          # Dominant octave ring boost 0-100 (0=off, 50=1.5x, 100=2.0x)
        "faceplate_color": "#1A1A1A",  # Background/faceplate color
        "show_fps": False,             # Show FPS counter on tuner display
    },

    # TONER SETTINGS (Tone Analyzer)
    "toner_settings": {
        "reference_pitch": 440.0,
        "sensitivity": 50,
        "fps": "30",
        "view_mode": "spectrum",  # "spectrum" or "bars"
        "scale_mode": "linear",   # "linear" or "db"
        "gauge_bias": {"richness": 0, "warmth": 0},
        "sax_type": "Alto",
        "concert_pitch": False,
        "analysis_descriptors": {
            "richness": True,
            "warmth": True,
            "even_odd": False,
            "rolloff_shape": False,
        },
    },
    "visible_preset_fields": {
        "serial": False,
        "reed": True,
        "ligature": False,
        "room": False,
        "preamp": False,
        "notes": True,
    },

    # TOOLING SETTINGS (Die Inserts & Die Holders)
    "tooling_settings": {
        "sheet_width": "12",
        "sheet_height": "12",
        "engrave_ring": True,
        "engrave_cutout": True,
        "engraving_mode": "filled",
        "step_size": "0.5",
        "ring_font_size": 3.5,
        "cutout_font_size": 3.5,
        "ring_engraving_location": "centered",   # "centered" or "from_outside"
        "ring_engraving_offset": 0.0,             # mm offset (used with from_outside)
        "holder_layer_count": 6,                  # 5 (2x pin) or 6 (3x pin)
        "holder_sheet_width": "12",
        "holder_sheet_height": "12",
        # Feeds & Speeds Tester
        "feeds_speeds_tester": {
            "material": "Felt",
            "disc_diameter_mm": 20.0,
            "speed_value": 180,
            "speed_sweep": True,
            "speed_start": 80, "speed_end": 280, "speed_stops": 4,
            "power_value": 60,
            "power_sweep": True,
            "power_start": 30, "power_end": 90, "power_stops": 4,
            "passes_value": 1,
            "passes_sweep": False,
            "passes_start": 1, "passes_end": 3, "passes_stops": 3,
            "engraving_on": True,
            # Engraving feed/power for the disc ID labels. Pre-filled from
            # the selected material's defaults when "Apply material defaults"
            # is clicked; user can edit freely. The whole point of the tool
            # is the user may not know their cut settings yet — engraving
            # is more forgiving, but they should still see and tweak it.
            "eng_speed_value": 1200,
            "eng_power_value": 10,
            "air_assist": True,
            # When True, every disc in the matrix is duplicated — one with
            # air on, one with air off — so the user can compare edge quality
            # side-by-side. Doubles the disc count.
            "also_test_no_air": False,
            # Show the layout preview before the save-file dialog opens.
            # Defaults on because the disc count can balloon with many
            # sweep stops and users want to eyeball it before committing.
            "show_preview": True,
            "sheet_w": "4", "sheet_h": "6", "sheet_unit": "in",
            "filename": "speed_power_test",
        },
    },

    # FILLED ENGRAVING OPTIONS
    "filled_overscan_enabled": False,  # Extend scan lines beyond character edges for consistent power
    "filled_overscan_mm": 1.5,        # Distance in mm to extend on each side

    # GLOBAL G-CODE OPTIONS
    "gcode_return_speed": 1000,  # mm/min for return-to-home move (0 = rapid G0)
    "gcode_cut_grouping": "layer",  # "layer" (all engravings, then holes, then cuts) or "pad" (all ops per pad)

    # SD CARD SETTINGS
    "sd_card_path": "",  # Last used SD card path for "Send to SD Card" feature
    "eject_sd_after_gcode": False,  # Auto-eject removable drive after G-code export

    # LASER DIRECT-CUT SETTINGS (v3.0+)
    # Streamed directly to a Grbl controller over USB. Falcon-specific
    # bits (auto-detect heuristics, ChArUco card defaults) live in
    # falcon_sender.py / camera_capture.py; the settings below are
    # generic to any Grbl-compatible laser.
    #
    # Serial-port override is still "falcon_*" — auto-detection targets
    # the Falcon's VID/PID; if a future user runs another Grbl laser,
    # they'll set the port directly here and bypass detection.
    "falcon_serial_port_override": None,
    "laser_framing_power_s": 10,    # Grbl S value during framing (0-1000)
    "laser_framing_feed": 2000,     # mm/min during framing
    # Bed dimensions in machine-mm — defaults match the Creality Falcon2
    # Pro 40W. Override here for other Grbl-compatible lasers.
    "laser_bed_x_max": 400.0,
    "laser_bed_y_max": 415.0,
    # Seed-move feedrate for AUTO Frame & Cut (drives head from home
    # corner to scrap's known machine position). G1 at this rate
    # instead of G0 rapid so a stray object in the path doesn't crash
    # at full speed. 4000 mm/min ≈ 67 mm/s — fast enough not to feel
    # sluggish, slow enough to stop on contact.
    "laser_seed_feed": 4000,
    # Safety margin: shrink camera-captured polygons by this many mm
    # on every edge before nesting. Camera measurement is least
    # accurate at the edges of its view, so insetting prevents pads
    # from being placed where a chunk of leather might actually be
    # missing relative to what the camera reported.
    "camera_polygon_inset_mm": 3.0,
    # Bias applied to Otsu's auto-threshold in scrap-outline detection
    # (CameraCaptureDialog slider). 0 = use Otsu as-is. Positive =
    # less sensitive (raise threshold), negative = more sensitive.
    # Persisted so the user's lighting preference survives restarts.
    "camera_detection_threshold_bias": 0,
    # Invert-colors flag for scrap detection. False (default) assumes
    # scrap is brighter than the bed (leather on honeycomb). True
    # selects THRESH_BINARY_INV — for dark scrap on a light surface.
    "camera_detection_invert": False,
    # Last successful calibration-card engrave offset (machine-mm,
    # [X, Y]). Persisted so a closed/reopened calibration dialog skips
    # straight to the capture phase against the existing engraved card
    # instead of re-engraving. Cleared on Calibrate & Save (calibration
    # done) or on "Engrave new card" (user starts fresh).
    "camera_calibration_engrave_offset_mm": None,
    # Frame & Cut "Try auto-frame" checkbox state. When True AND a
    # polygon with a saved machine offset is loaded AND camera
    # calibration is present, Frame & Cut drives the head to the
    # scrap's known bed position automatically. When False (default),
    # MANUAL mode — user jogs the head to the material's bottom-left
    # before framing/cutting.
    "frame_cut_try_auto": False,
    # Persisted camera index — set after a successful camera
    # calibration so future runs skip the enumeration / find-Falcon
    # heuristic and open the known-working camera directly. None =
    # auto-detect on each open (find_falcon_camera_index → fall back
    # to last enumerated).
    "camera_index_override": None,
    # Draw / Capture Shape grid size, in the corresponding unit.
    # Defaults match a Falcon2 Pro 40W's 15.75 × 15.75 in / 400 × 415
    # mm bed with a small margin. Override here for larger machines —
    # the grid will display at the chosen size and camera-captured
    # polygons larger than the grid would otherwise be silently
    # vertex-clamped to the grid bounds.
    "polygon_draw_grid_size_in": 15,
    "polygon_draw_grid_size_cm": 40,

    # TUTORIAL FLAGS
    "seen_polygon_tutorial": False,
    "seen_kerf_test_tutorial": False,
    "seen_calibration_tutorial": False,

    # AUDIO INPUT
    "audio_input_device": None,  # None = system default, or device index
    "capture_threshold": 50,     # 0-100, how loud a signal must be to trigger capture
    "toner_record_wav": True,    # Record WAV audio during toner capture sessions
    "toner_recording_dir": "",   # Directory for WAV recordings (empty = Music/StohrerSaxShopCompanion)
    "toner_wav_reanalyze": True,   # Reprocess WAV offline for max analysis accuracy
    "toner_wav_auto_delete": False,  # Delete WAV after reanalysis

    # VISIBLE SAX TYPES (for toner sax selector)
    "visible_sax_types": None,  # None = all types, or list of type names

    # TONER ACCESS
    "toner_unlocked": False,
    "toner_sandbox_enabled": False,  # Show sandbox checkbox when creating presets

    # NESTING PREVIEW
    "show_preview": False,
    "seen_preview_tutorial": False,

    # LAST USED LIBRARIES
    "last_pad_library": "My Presets",
    "last_key_library": "My Presets",

    # SETTINGS-DIALOG TOOLTIPS (Feature Set toggle, default on)
    "tooltips_enabled": True,

    # VISIBLE TABS (Feature Set)
    "visible_tabs": {
        "Key Height Library": True,
        "Serial Lookup": True,
        "Screw Specs": True,
        "Tooling": True,
        "Tuner": True,
        "Toner": False,
    },

    "key_layout": {
        "show_serial": False,
        "large_notes": False,
        "show_B": True,
        "show_F": True,
        "show_Palm_F": False,
        "show_Palm_E": False,
        "show_Palm_Eb": False,
        "show_Palm_D": False,
        "show_G": False,
        "show_D": False,
        "show_Low_C": True,
        "show_Low_B": False,
        "show_Low_Bb": False
    },

    "engraving_font_size": {
        "felt": 3.0,
        "card": 3.0,
        "leather": 3.0,
        "exact_size": 3.0
    },
    "engraving_location": {
        "felt": {"mode": "from_inside", "value": 4.0},
        "card": {"mode": "from_inside", "value": 4.0},
        "leather": {"mode": "from_outside", "value": 1.0},
        "darted_leather": {"mode": "from_outside", "value": 2.5},
        "exact_size": {"mode": "from_inside", "value": 4.0}
    },
    "layer_colors": {
        'felt_outline': '#000000',
        'felt_center_hole': '#0000A0',
        'felt_engraving': '#A00000',
        'card_outline': '#0000FF',
        'card_center_hole': '#00A0FF',
        'card_engraving': '#A000A0',
        'leather_outline': '#FF0000',
        'leather_center_hole': '#00E000',
        'leather_engraving': '#FF8000',
        'exact_size_outline': '#D0D000',
        'exact_size_center_hole': '#A0A000',
        'exact_size_engraving': '#BB7784',
        'die_outer_cut': '#FF0000',
        'die_inner_cut': '#0000FF',
        'die_engraving': '#00E000',
        'die_cutout_engraving': '#FF8000',
        'die_holder_cut': '#FF0000',
        'die_holder_hole': '#0000FF',
    }
}

# ==========================================
# IO FUNCTIONS
# ==========================================

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                loaded_settings = json.load(f)
                settings = copy.deepcopy(DEFAULT_SETTINGS)

                # Deep copy key_layout to avoid shared references
                if "key_layout" not in loaded_settings:
                    loaded_settings["key_layout"] = settings["key_layout"].copy()

                # Merge loaded settings into default structure
                for key, default_value in DEFAULT_SETTINGS.items():
                    if key in loaded_settings:
                        if isinstance(default_value, dict):
                            # Only merge if loaded value is also a dict; otherwise keep default
                            if not isinstance(loaded_settings[key], dict):
                                continue
                            settings[key] = {}
                            for sub_key, sub_default in default_value.items():
                                if isinstance(sub_default, dict) and isinstance(loaded_settings[key].get(sub_key), dict):
                                    # Two-level deep merge (e.g. gcode_settings.felt)
                                    settings[key][sub_key] = sub_default.copy()
                                    settings[key][sub_key].update(loaded_settings[key][sub_key])
                                elif sub_key in loaded_settings[key]:
                                    settings[key][sub_key] = loaded_settings[key][sub_key]
                                else:
                                    settings[key][sub_key] = sub_default if not isinstance(sub_default, dict) else sub_default.copy()
                        else:
                            # Don't replace defaults with None — old configs may have nulls
                            if loaded_settings[key] is not None:
                                settings[key] = loaded_settings[key]

                # Validate layer_colors: old versions used LightBurn codes (C10, C15)
                # instead of hex colors. Replace any non-hex values with defaults.
                if "layer_colors" in settings and isinstance(settings["layer_colors"], dict):
                    default_colors = DEFAULT_SETTINGS["layer_colors"]
                    for key, val in settings["layer_colors"].items():
                        if not isinstance(val, str) or not val.startswith("#"):
                            settings["layer_colors"][key] = default_colors.get(key, "#000000")

                # Migrate dart_shape_factor from the old sine→square spectrum
                # (0.0=sine, 1.0=square) to the new triangle→sine→square one
                # (0.0=triangle, 0.5=sine, 1.0=square). Old value v maps to
                # 0.5 + 0.5*v. Run once, gated on dart_shape_v2.
                if not loaded_settings.get("dart_shape_v2"):
                    if "dart_shape_factor" in loaded_settings:
                        try:
                            old_v = float(loaded_settings["dart_shape_factor"])
                            settings["dart_shape_factor"] = 0.5 + 0.5 * max(0.0, min(1.0, old_v))
                        except (TypeError, ValueError):
                            pass
                    for r in settings.get("dart_ranges", []):
                        if "shape_factor" in r:
                            try:
                                old_v = float(r["shape_factor"])
                                r["shape_factor"] = 0.5 + 0.5 * max(0.0, min(1.0, old_v))
                            except (TypeError, ValueError):
                                pass
                    settings["dart_shape_v2"] = True

                return settings
        except (json.JSONDecodeError, TypeError, KeyError, AttributeError, ValueError):
            return copy.deepcopy(DEFAULT_SETTINGS)
    return copy.deepcopy(DEFAULT_SETTINGS)

def save_settings(settings):
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=2)
    except Exception as e:
        messagebox.showerror(_("Error Saving Settings"), _("Could not save settings:\n{e}").format(e=e))

def load_presets(file_path, preset_type_name="Preset"):
    data = {}
    if os.path.exists(file_path):
        try:
            with open(file_path, 'r') as f:
                data = json.load(f)
            if not isinstance(data, dict):
                data = {}
        except (json.JSONDecodeError, TypeError):
            data = {}
    
    # Migration logic for old flat files (legacy support)
    if data and not any(isinstance(v, dict) for v in data.values()):
        print(f"Migrating old {preset_type_name} file...")
        new_data = {"My Presets": data}
        if save_presets(new_data, file_path):
            messagebox.showinfo(
                _("Library Updated"),
                _("Your existing {preset_type_name} sets have been moved into a new library called 'My Presets'.").format(preset_type_name=preset_type_name),
            )
            return new_data
        else:
            return {}
    
    return data if data else {}

def save_presets(presets, file_path):
    try:
        with open(file_path, 'w') as f:
            json.dump(presets, f, indent=2)
        return True
    except Exception as e:
        messagebox.showerror(_("Error Saving Preset"), str(e))
        return False
