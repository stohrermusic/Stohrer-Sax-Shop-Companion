import os
import sys
import json
import shutil
from tkinter import messagebox

# ==========================================
# PLATFORM-SPECIFIC CONFIG DIRECTORY
# ==========================================

APP_NAME = "StohrerSaxShopCompanion"

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

# ==========================================
# CONSTANTS & FILE PATHS
# ==========================================

_CONFIG_DIR = get_config_dir()
PAD_PRESET_FILE = os.path.join(_CONFIG_DIR, "pad_presets.json")
KEY_PRESET_FILE = os.path.join(_CONFIG_DIR, "key_height_library.json")
SETTINGS_FILE = os.path.join(_CONFIG_DIR, "app_settings.json")
SCREW_SPECS_FILE = os.path.join(_CONFIG_DIR, "screw_specs.json")

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
    
    "key_layout": {
        "show_serial": False,
        "large_notes": False,
        "show_B": True,
        "show_F": True,
        "show_Palm F": False,
        "show_Palm E": False,
        "show_Palm Eb": False,
        "show_Palm D": False,
        "show_G": False,
        "show_D": False,
        "show_Low C": True,
        "show_Low B": False,
        "show_Low Bb": False
    },

    "engraving_font_size": {
        "felt": 2.0,
        "card": 2.0,
        "leather": 2.0,
        "exact_size": 2.0
    },
    "engraving_location": {
        "felt": {"mode": "centered", "value": 0.0},
        "card": {"mode": "centered", "value": 0.0},
        "leather": {"mode": "from_outside", "value": 1.0},
        "exact_size": {"mode": "centered", "value": 0.0}
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
        'exact_size_engraving': '#BB7784'
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
                            settings[key] = default_value.copy()
                            settings[key].update(loaded_settings[key])
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
