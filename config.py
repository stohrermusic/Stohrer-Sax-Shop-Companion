import os
import sys
import json
import shutil
import ssl
from tkinter import messagebox


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
APP_VERSION = "1.9"
APP_BUILD_DATE = "2026-03-19"

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
        "screw_specs.json"
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
        "screw_specs.json"
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
TONE_PROFILES_FILE = os.path.join(_CONFIG_DIR, "tone_profiles.json")

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

RESONANCE_MESSAGES = [
    "Resonance added!", "Pad resonance increased!", "More resonance now!",
    "Timbral focus enhanced!", "Harmonic alignment optimized!", "Acoustic reflection matrix calibrated!",
    "Core vibrations synchronized!", "Nodal points stabilized!", "Overtone series enriched!",
    "Sonic clarity has been improved!", "Relacquer devaluation reversed!", "Heavy mass screws ain't SHIT!",
    "Now you don't even have to fit the neck!", "Let's call this the ULTRAhaul!", "Now safe to use hot glue!",
    "Look at me! I am the resonator now!"
]

# ==========================================
# DEFAULT SETTINGS
# ==========================================

DEFAULT_SETTINGS = {
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
    "dart_shape_factor": 0.0,
    
    "dart_engraving_on": True,
    "dart_engraving_loc": {"mode": "from_outside", "value": 2.5},

    # MAX FILL SETTINGS
    "max_fill_style": "center_out",  # "center_out" or "longest_edge"

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
            "filled_engraving_speed": 1200,    # mm/min ("filled" mode)
            "filled_engraving_power": 8,       # percent ("filled" mode)
            "filled_line_spacing": 0.15,       # mm between scan lines
            "hole_speed": 300,
            "hole_power": 35,
            "cut_speed": 600,
            "cut_power": 60,
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
            "filled_engraving_speed": 1200,
            "filled_engraving_power": 15,
            "filled_line_spacing": 0.15,
            "hole_speed": 400,
            "hole_power": 22.5,
            "cut_speed": 1500,
            "cut_power": 50,
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
            "filled_engraving_speed": 1800,
            "filled_engraving_power": 8,
            "filled_line_spacing": 0.15,
            "hole_speed": 300,
            "hole_power": 30,
            "cut_speed": 1200,
            "cut_power": 75,
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
            "filled_engraving_speed": 4000,
            "filled_engraving_power": 26,
            "filled_line_spacing": 0.15,
            "hole_speed": 180,
            "hole_power": 100,
            "cut_speed": 180,
            "cut_power": 100,
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
        "faceplate_color": "#1A1A1A",  # Background/faceplate color
    },

    # TONER SETTINGS (Tone Analyzer)
    "toner_settings": {
        "reference_pitch": 440.0,
        "sensitivity": 50,
        "fps": "30",
        "view_mode": "spectrum",  # "spectrum" or "bars"
        "scale_mode": "linear",   # "linear" or "db"
        "gauge_bias": {"resonance": 0, "richness": 0, "brightness": 0, "darkness": 0},
        "sax_type": "Alto",
        "concert_pitch": False,
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

    # TUTORIAL FLAGS
    "seen_polygon_tutorial": False,
    "seen_sdcard_tutorial": False,
    "seen_kerf_test_tutorial": False,
    "seen_calibration_tutorial": False,

    # AUDIO INPUT
    "audio_input_device": None,  # None = system default, or device index
    "capture_threshold": 50,     # 0-100, how loud a signal must be to trigger capture

    # TONER ACCESS
    "toner_unlocked": False,

    # NESTING PREVIEW
    "show_preview": False,

    # LAST USED LIBRARIES
    "last_pad_library": "My Presets",
    "last_key_library": "My Presets",

    # VISIBLE TABS (Feature Set)
    "visible_tabs": {
        "Key Height Library": True,
        "Serial Lookup": True,
        "Screw Specs": True,
        "Tooling": True,
        "Tuner": False,
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
                settings = DEFAULT_SETTINGS.copy()
                
                # Deep copy key_layout to avoid shared references
                if "key_layout" not in loaded_settings:
                    loaded_settings["key_layout"] = settings["key_layout"].copy()

                # Merge loaded settings into default structure
                for key, default_value in DEFAULT_SETTINGS.items():
                    if key in loaded_settings:
                        if isinstance(default_value, dict):
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
                            settings[key] = loaded_settings[key]
                
                return settings
        except (json.JSONDecodeError, TypeError):
            return DEFAULT_SETTINGS.copy()
    return DEFAULT_SETTINGS.copy()

def save_settings(settings):
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=2)
    except Exception as e:
        messagebox.showerror("Error Saving Settings", f"Could not save settings:\n{e}")

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
            messagebox.showinfo("Library Updated", f"Your existing {preset_type_name} sets have been moved into a new library called 'My Presets'.")
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
        messagebox.showerror("Error Saving Preset", str(e))
        return False
