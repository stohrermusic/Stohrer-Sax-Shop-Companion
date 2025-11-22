import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import json
import random
import math
import svgwrite
import re # Added for serial number parsing

# --- Import Serial Data ---
try:
    import serials
    SERIAL_DATA = serials.SERIAL_DATA
except ImportError:
    SERIAL_DATA = {} # Fallback if file is missing

# ==========================================
# SECTION 1: CONFIGURATION & DATA
# ==========================================

# --- Lightburn Color Palette ---
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

# --- Default Key Height Fields ---
ALL_KEY_HEIGHT_FIELDS = [
    "B", "F", "Palm F", "Palm E", "Palm Eb", "Palm D", 
    "G", "D", "Low C", "Low B", "Low Bb"
]

# --- Default Configuration ---
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
    
    # NEW SETTINGS FOR v2.1
    "darts_enabled": True,    
    "dart_threshold": 18.0,   
    "dart_overwrap": 0.5,     
    "dart_wrap_bonus": 0.75, 
    "dart_frequency_multiplier": 1.0,
    "dart_shape_factor": 0.0,
    
    # DART SPECIFIC ENGRAVING DEFAULTS
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

# --- Filenames ---
PAD_PRESET_FILE = "pad_presets.json"
KEY_PRESET_FILE = "key_height_library.json"
SETTINGS_FILE = "app_settings.json"

# --- Constants & Themes ---
RESONANCE_MESSAGES = [
    "Resonance added!", "Pad resonance increased!", "More resonance now!",
    "Timbral focus enhanced!", "Harmonic alignment optimized!", "Acoustic reflection matrix calibrated!",
    "Core vibrations synchronized!", "Nodal points stabilized!", "Overtone series enriched!",
    "Sonic clarity has been improved!", "Relacquer devaluation reversed!", "Heavy mass screws ain't SHIT!",
    "Now you don't even have to fit the neck!", "Let's call this the ULTRAhaul!", "Now safe to use hot glue!",
    "Look at me! I am the resonator now!"
]
COOL_BLUE = "#E0F7FA"
COOL_GREEN = "#E8F5E9"

# --- IO Functions ---

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                loaded_settings = json.load(f)
                settings = DEFAULT_SETTINGS.copy()
                
                if "key_layout" not in loaded_settings:
                    loaded_settings["key_layout"] = settings["key_layout"].copy()

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
    
    if data and not any(isinstance(v, dict) for v in data.values()):
        print(f"Migrating old {preset_type_name} file...")
        new_data = {"My Presets": data}
        if save_presets(new_data, file_path):
            messagebox.showinfo("Library Updated", f"Your existing {preset_type_name} sets have been moved into a new library called 'My Presets'.")
            return new_data
        else:
            return {"My Presets": {}}
    
    return data if data else {"My Presets": {}}

def save_presets(presets, file_path):
    try:
        with open(file_path, 'w') as f:
            json.dump(presets, f, indent=2)
        return True
    except Exception as e:
        messagebox.showerror("Error Saving Preset", str(e))
        return False


# ==========================================
# SECTION 2: LOGIC & MATH
# ==========================================

def calculate_star_path(cx, cy, outer_r, inner_r, num_points=12, shape_factor=0.0):
    """
    Generates an SVG path string for a smooth Sine Wave (Flower) shape.
    shape_factor: 0.0 = Sine, 1.0 = Flattened (Square-ish)
    """
    path_data = []
    
    avg_r = (outer_r + inner_r) / 2.0
    amplitude = (outer_r - inner_r) / 2.0
    
    steps = int(num_points * 8) 
    if steps < 64: steps = 64
    
    angle_step = (2 * math.pi) / steps

    # Calculate power for shaping. 
    power = 1.0 - (0.9 * shape_factor)

    for i in range(steps + 1):
        theta = i * angle_step
        
        # Raw Sine Wave (-1 to 1)
        raw_wave = math.cos(num_points * theta)
        
        # Apply Shaping: sign * |raw|^power
        shaped_wave = (1 if raw_wave >= 0 else -1) * (abs(raw_wave) ** power)
        
        r = avg_r + amplitude * shaped_wave
        
        x = cx + r * math.cos(theta)
        y = cy + r * math.sin(theta)
        
        command = "M" if i == 0 else "L"
        path_data.append(f"{command} {x:.3f} {y:.3f}")
        
    path_data.append("Z") 
    return " ".join(path_data)

def get_disc_diameter(pad_size, material, settings):
    if material == 'felt': return pad_size - settings["felt_offset"]
    if material == 'card': return pad_size - (settings["felt_offset"] + settings["card_to_felt_offset"])
    if material == 'exact_size': return pad_size
    
    if material == 'leather':
        threshold = settings.get("dart_threshold", 18.0)
        darts_enabled = settings.get("darts_enabled", True)
        
        if darts_enabled and pad_size < threshold:
            # DART BOOST: Add extra wrap for stars
            bonus = settings.get("dart_wrap_bonus", 0.75)
            wrap = leather_back_wrap(pad_size, settings["leather_wrap_multiplier"], extra_base=bonus)
        else:
            # Standard Wrap
            wrap = leather_back_wrap(pad_size, settings["leather_wrap_multiplier"])
            
        felt_thickness_mm = get_felt_thickness_mm(settings)
        diameter = pad_size + 2 * (felt_thickness_mm + wrap)
        return round(diameter * 2) / 2
        
    return 0

def check_for_oversized_engravings(pads, material_vars, settings):
    oversized = {}
    for material, var in material_vars.items():
        if not var.get(): continue
        
        font_size = settings["engraving_font_size"].get(material, 2.0)
        oversized_sizes = set()

        for pad in pads:
            pad_size = pad['size']
            diameter = get_disc_diameter(pad_size, material, settings)
            radius = diameter / 2
            if font_size >= radius * 0.8:
                oversized_sizes.add(pad_size)
        
        if oversized_sizes:
            oversized[material] = oversized_sizes
    return oversized

def leather_back_wrap(pad_size, multiplier, extra_base=0.0):
    base_wrap = 0
    if pad_size >= 45:
        base_wrap = 3.2
    elif pad_size >= 12:
        base_wrap = 1.2 + (pad_size - 12) * (2.0 / 33.0)
    elif pad_size >= 6:
        base_wrap = 1.0 + (pad_size - 6) * (0.2 / 6.0)
    else:
        base_wrap = 1.0
        
    # Apply the Dart Bonus (if any) before multiplier
    total_base = base_wrap + extra_base
    return total_base * multiplier

def should_have_center_hole(pad_size, hole_dia, settings):
    min_size = settings.get("min_hole_size", 16.5)
    return hole_dia > 0 and pad_size >= min_size

def get_felt_thickness_mm(settings):
    thickness = settings.get("felt_thickness", 3.175)
    if settings.get("felt_thickness_unit") == "in":
        return thickness * 25.4
    return thickness

def can_all_pads_fit(pads, material, width_mm, height_mm, settings):
    spacing_mm = 1.0
    discs = []
    
    for pad in pads:
        pad_size, qty = pad['size'], pad['qty']
        diameter = get_disc_diameter(pad_size, material, settings)
        for _ in range(qty): discs.append((pad_size, diameter))

    discs.sort(key=lambda x: -x[1])
    placed = []
    for _, dia in discs:
        r = dia / 2
        placed_successfully = False
        y = spacing_mm
        while y + dia + spacing_mm <= height_mm and not placed_successfully:
            x = spacing_mm
            while x + dia + spacing_mm <= width_mm:
                cx, cy = x + r, y + r
                is_collision = any((cx - px)**2 + (cy - py)**2 < (r + pr + spacing_mm)**2 for _, px, py, pr in placed)
                if not is_collision:
                    placed.append((None, cx, cy, r))
                    placed_successfully = True
                    break
                x += 1
            y += 1
    
    return len(placed) == len(discs)

def generate_svg(pads, material, width_mm, height_mm, filename, hole_dia_preset, settings):
    spacing_mm = 1.0
    discs = []

    for pad in pads:
        pad_size, qty = pad['size'], pad['qty']
        diameter = get_disc_diameter(pad_size, material, settings)
        for _ in range(qty): discs.append((pad_size, diameter))

    discs.sort(key=lambda x: -x[1])
    placed = []
    for pad_size, dia in discs:
        r = dia / 2
        placed_successfully = False
        y = spacing_mm
        while y + dia + spacing_mm <= height_mm and not placed_successfully:
            x = spacing_mm
            while x + dia + spacing_mm <= width_mm:
                cx, cy = x + r, y + r
                # Check collision using standard circle r
                is_collision = any((cx - px)**2 + (cy - py)**2 < (r + pr + spacing_mm)**2 for _, px, py, pr in placed)
                if not is_collision:
                    placed.append((pad_size, cx, cy, r))
                    placed_successfully = True
                    break
                x += 1
            y += 1

    compatibility_mode = settings.get("compatibility_mode", False)
    
    if compatibility_mode:
        dwg = svgwrite.Drawing(filename, size=(f"{width_mm}mm", f"{height_mm}mm"), viewBox=f"0 0 {width_mm} {height_mm}")
        stroke_w = 0.1
    else:
        dwg = svgwrite.Drawing(filename, size=(f"{width_mm}mm", f"{height_mm}mm"), profile='tiny')
        stroke_w = '0.1mm'

    layer_colors = settings.get("layer_colors", DEFAULT_SETTINGS["layer_colors"])

    for pad_size, cx, cy, r in placed:
        
        threshold = settings.get("dart_threshold", 18.0)
        darts_enabled = settings.get("darts_enabled", True)
        
        is_dart_pad = (material == 'leather' and darts_enabled and pad_size < threshold)
        
        if is_dart_pad:
            # --- STAR LOGIC ---
            felt_thick = get_felt_thickness_mm(settings)
            overwrap = settings.get("dart_overwrap", 0.5)
            
            # 1. Inner Radius (Valley) - Safe Zone
            felt_r = (pad_size - settings["felt_offset"]) / 2
            inner_r = felt_r + felt_thick + overwrap
            
            # 2. Outer Radius (Tip) - The Boosted Wrap
            # 'r' is the full Boosted radius from get_disc_diameter
            outer_r = r
            
            # Safety Check
            if inner_r >= outer_r:
                 inner_r = outer_r - 0.2 
            
            # 3. Dynamic Points & Shape
            circumference = 2 * math.pi * inner_r
            freq_mult = settings.get("dart_frequency_multiplier", 1.0)
            num_points = int((circumference / 3.5) * freq_mult)
            if num_points < 12: num_points = 12 
            if num_points % 2 != 0: num_points += 1 
            
            shape_factor = settings.get("dart_shape_factor", 0.0)
            
            path_d = calculate_star_path(cx, cy, outer_r, inner_r, num_points=num_points, shape_factor=shape_factor)
            
            dwg.add(dwg.path(d=path_d, stroke=layer_colors[f'{material}_outline'], fill='none', stroke_width=stroke_w))
        else:
            # --- STANDARD CIRCLE LOGIC ---
            if compatibility_mode:
                dwg.add(dwg.circle(center=(cx, cy), r=r, stroke=layer_colors[f'{material}_outline'], fill='none', stroke_width=stroke_w))
            else:
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{r}mm", stroke=layer_colors[f'{material}_outline'], fill='none', stroke_width=stroke_w))

        hole_dia = 0
        if should_have_center_hole(pad_size, hole_dia_preset, settings):
            hole_dia = hole_dia_preset

        if hole_dia > 0:
            if compatibility_mode:
                dwg.add(dwg.circle(center=(cx, cy), r=hole_dia / 2, stroke=layer_colors[f'{material}_center_hole'], fill='none', stroke_width=stroke_w))
            else:
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{hole_dia / 2}mm", stroke=layer_colors[f'{material}_center_hole'], fill='none', stroke_width=stroke_w))

        font_size = settings.get("engraving_font_size", {}).get(material, 2.0)
        
        # --- Determine Engraving Settings (Standard vs Star) ---
        should_engrave = False
        
        if is_dart_pad:
            # Use Star Specific Settings
            if settings.get("dart_engraving_on", True):
                engraving_settings = settings.get("dart_engraving_loc", {"mode": "from_outside", "value": 2.5})
                should_engrave = True
        else:
            # Use Standard Settings
            if settings.get("engraving_on", True):
                engraving_settings = settings["engraving_location"][material]
                should_engrave = True

        # Safety Check: Don't engrave if text is wider than the pad radius
        if should_engrave and (font_size >= r * 0.8):
            should_engrave = False
        
        if should_engrave:
            mode = engraving_settings['mode']
            value = engraving_settings['value']
            
            engraving_y = 0
            if mode == 'from_outside':
                engraving_y = cy - (r - value)
            elif mode == 'from_inside':
                hole_r = hole_dia / 2 if hole_dia > 0 else 0
                engraving_y = cy - (hole_r + value)
            else: # centered
                hole_r = hole_dia / 2 if hole_dia > 0 else 1.75
                offset_from_center = (r + hole_r) / 2
                engraving_y = cy - offset_from_center

            vertical_adjust = font_size * 0.35
            text_content = f"{pad_size:.1f}".rstrip('0').rstrip('.')
            
            if compatibility_mode:
                dwg.add(dwg.text(text_content,
                                 insert=(cx, engraving_y + vertical_adjust),
                                 text_anchor="middle",
                                 font_size=font_size,
                                 fill=layer_colors[f'{material}_engraving']))
            else:
                dwg.add(dwg.text(text_content,
                                 insert=(f"{cx}mm", f"{engraving_y + vertical_adjust}mm"),
                                 text_anchor="middle",
                                 font_size=f"{font_size}mm",
                                 fill=layer_colors[f'{material}_engraving']))
        
    dwg.save()

# --- New Serial Logic ---
def lookup_serial_year(maker, serial_str):
    if not maker or not serial_str:
        return ""
    
    if maker not in SERIAL_DATA:
        return "Manufacturer data not found."

    # Extract numbers only for comparison
    clean_serial = "".join(filter(str.isdigit, serial_str))
    if not clean_serial:
        return "Invalid Serial Number"
    
    try:
        serial_num = int(clean_serial)
    except ValueError:
        return "Invalid Serial Number"

    data = SERIAL_DATA[maker]
    # Data is list of tuples: (Start_Serial, Year)
    # We want to find the largest Start_Serial <= serial_num
    
    found_year = None
    
    # Iterate to find the range (Since lists are small, linear scan is fine)
    for start_serial, year in data:
        if serial_num >= start_serial:
            found_year = year
        else:
            break # We passed the range
            
    if found_year:
        return str(found_year)
    else:
        return "Too old / Unknown"


# ==========================================
# SECTION 3: GUI DIALOGS
# ==========================================

class ConfirmationDialog(tk.Toplevel):
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title(title)
        self.geometry("450x150")
        self.configure(bg="#F0EAD6")
        self.transient(parent)
        self.grab_set()

        self.result = False
        self.dont_show_again = tk.BooleanVar()

        tk.Label(self, text=message, wraplength=430, bg="#F0EAD6", justify="left").pack(padx=10, pady=10)
        
        checkbox_frame = tk.Frame(self, bg="#F0EAD6")
        checkbox_frame.pack(pady=5)
        tk.Checkbutton(checkbox_frame, text="Don't show this message again", variable=self.dont_show_again, bg="#F0EAD6").pack()
        
        button_frame = tk.Frame(self, bg="#F0EAD6")
        button_frame.pack(pady=10)
        tk.Button(button_frame, text="Yes, Proceed", command=self.on_yes).pack(side="left", padx=10)
        tk.Button(button_frame, text="No, Cancel", command=self.on_no).pack(side="left", padx=10)

        self.protocol("WM_DELETE_WINDOW", self.on_no)
        self.wait_window(self)

    def on_yes(self):
        self.result = True
        self.destroy()

    def on_no(self):
        self.result = False
        self.destroy()

class OptionsWindow:
    def __init__(self, parent, app, settings, update_callback, save_callback):
        self.app = app
        self.settings = settings
        self.update_callback = update_callback
        self.save_callback = save_callback
        
        self.top = tk.Toplevel(parent)
        self.top.title("Sizing Rules")
        self.top.geometry("500x750") 
        self.top.configure(bg="#F0EAD6")
        self.top.transient(parent)
        self.top.grab_set()

        # --- Main Layout Frames ---
        bottom_button_frame = tk.Frame(self.top, bg="#F0EAD6")
        bottom_button_frame.pack(side="bottom", fill="x", pady=10, padx=10)
        
        tk.Button(bottom_button_frame, text="Save", command=self.save_options).pack(side="left", padx=5)
        tk.Button(bottom_button_frame, text="Cancel", command=self.top.destroy).pack(side="left", padx=5)
        
        tk.Button(bottom_button_frame, text="Advanced", command=self.app.open_resonance_window).pack(side="right", padx=5)
        tk.Button(bottom_button_frame, text="Revert to Defaults", command=self.revert_to_defaults).pack(side="right", padx=5)
        
        main_canvas_frame = tk.Frame(self.top)
        main_canvas_frame.pack(side="top", fill="both", expand=True)

        self.canvas = tk.Canvas(main_canvas_frame, bg="#F0EAD6", highlightthickness=0)
        self.scrollbar = tk.Scrollbar(main_canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#F0EAD6", padx=10, pady=10)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.top.bind('<MouseWheel>', self._on_mousewheel)

        # --- Sizing variables ---
        self.unit_var = tk.StringVar(value=self.settings["units"])
        self.felt_offset_var = tk.DoubleVar(value=self.settings["felt_offset"])
        self.card_offset_var = tk.DoubleVar(value=self.settings["card_to_felt_offset"])
        self.leather_mult_var = tk.DoubleVar(value=self.settings["leather_wrap_multiplier"])
        self.min_hole_size_var = tk.DoubleVar(value=self.settings["min_hole_size"])
        self.felt_thickness_var = tk.DoubleVar(value=self.settings["felt_thickness"])
        self.felt_thickness_unit_var = tk.StringVar(value=self.settings["felt_thickness_unit"])
        
        # NEW VARS FOR v2.1
        self.darts_enabled_var = tk.BooleanVar(value=self.settings.get("darts_enabled", True))
        self.dart_threshold_var = tk.DoubleVar(value=self.settings.get("dart_threshold", 18.0))
        self.dart_overwrap_var = tk.DoubleVar(value=self.settings.get("dart_overwrap", 0.5))
        self.dart_wrap_bonus_var = tk.DoubleVar(value=self.settings.get("dart_wrap_bonus", 0.75))
        self.dart_frequency_multiplier_var = tk.DoubleVar(value=self.settings.get("dart_frequency_multiplier", 1.0))
        self.dart_shape_factor_var = tk.DoubleVar(value=self.settings.get("dart_shape_factor", 0.0))
        
        self.engraving_on_var = tk.BooleanVar(value=self.settings["engraving_on"])
        self.compatibility_mode_var = tk.BooleanVar(value=self.settings.get("compatibility_mode", False))
        self.engraving_font_size_vars = {}
        self.engraving_loc_vars = {}
        
        # --- NEW: Dart Engraving Vars ---
        self.dart_engraving_on_var = tk.BooleanVar(value=self.settings.get("dart_engraving_on", True))
        self.dart_engraving_mode_var = tk.StringVar(value=self.settings.get("dart_engraving_loc", {}).get("mode", "from_outside"))
        self.dart_engraving_val_var = tk.DoubleVar(value=self.settings.get("dart_engraving_loc", {}).get("value", 2.5))

        self.create_option_widgets()
    
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def create_option_widgets(self):
        main_frame = self.scrollable_frame
        
        unit_frame = tk.LabelFrame(main_frame, text="Sheet Units", bg="#F0EAD6", padx=5, pady=5)
        unit_frame.pack(fill="x", pady=5)
        tk.Radiobutton(unit_frame, text="Inches (in)", variable=self.unit_var, value="in", bg="#F0EAD6").pack(side="left", padx=5)
        tk.Radiobutton(unit_frame, text="Centimeters (cm)", variable=self.unit_var, value="cm", bg="#F0EAD6").pack(side="left", padx=5)
        tk.Radiobutton(unit_frame, text="Millimeters (mm)", variable=self.unit_var, value="mm", bg="#F0EAD6").pack(side="left", padx=5)

        rules_frame = tk.LabelFrame(main_frame, text="Sizing Rules (Advanced)", bg="#F0EAD6", padx=5, pady=5)
        rules_frame.pack(fill="x", pady=5)
        rules_frame.columnconfigure(1, weight=1)

        tk.Label(rules_frame, text="Felt Diameter Reduction (mm):", bg="#F0EAD6").grid(row=0, column=0, sticky='w', pady=2)
        tk.Entry(rules_frame, textvariable=self.felt_offset_var, width=10).grid(row=0, column=1, sticky='w', pady=2)

        tk.Label(rules_frame, text="Card Additional Reduction (mm):", bg="#F0EAD6").grid(row=1, column=0, sticky='w', pady=2)
        tk.Entry(rules_frame, textvariable=self.card_offset_var, width=10).grid(row=1, column=1, sticky='w', pady=2)

        tk.Label(rules_frame, text="Leather Wrap Multiplier (1.00=default):", bg="#F0EAD6").grid(row=2, column=0, sticky='w', pady=2)
        tk.Entry(rules_frame, textvariable=self.leather_mult_var, width=10).grid(row=2, column=1, sticky='w', pady=2)

        tk.Label(rules_frame, text="Min. Pad Size for Hole (mm):", bg="#F0EAD6").grid(row=3, column=0, sticky='w', pady=2)
        tk.Entry(rules_frame, textvariable=self.min_hole_size_var, width=10).grid(row=3, column=1, sticky='w', pady=2)
        
        felt_thickness_frame = tk.Frame(rules_frame, bg="#F0EAD6")
        felt_thickness_frame.grid(row=4, column=0, columnspan=2, sticky='w', pady=2)
        tk.Label(felt_thickness_frame, text="Felt Thickness:", bg="#F0EAD6").pack(side="left")
        tk.Entry(felt_thickness_frame, textvariable=self.felt_thickness_var, width=10).pack(side="left", padx=5)
        tk.Radiobutton(felt_thickness_frame, text="in", variable=self.felt_thickness_unit_var, value="in", bg="#F0EAD6").pack(side="left")
        tk.Radiobutton(felt_thickness_frame, text="mm", variable=self.felt_thickness_unit_var, value="mm", bg="#F0EAD6").pack(side="left")

        # --- NEW DART SETTINGS FRAME ---
        darts_frame = tk.LabelFrame(main_frame, text="Star / Dart Settings", bg="#F0EAD6", padx=5, pady=5)
        darts_frame.pack(fill="x", pady=5)
        darts_frame.columnconfigure(1, weight=1)
        
        tk.Checkbutton(darts_frame, text="Enable Star / Dart Pattern", variable=self.darts_enabled_var, bg="#F0EAD6").grid(row=0, column=0, columnspan=2, sticky='w', pady=2)
        
        tk.Label(darts_frame, text="Use Star Pattern below (mm):", bg="#F0EAD6").grid(row=1, column=0, sticky='w', pady=2)
        tk.Entry(darts_frame, textvariable=self.dart_threshold_var, width=10).grid(row=1, column=1, sticky='w', pady=2)

        tk.Label(darts_frame, text="Star Safe Overwrap (Valley) (mm):", bg="#F0EAD6").grid(row=2, column=0, sticky='w', pady=2)
        tk.Entry(darts_frame, textvariable=self.dart_overwrap_var, width=10).grid(row=2, column=1, sticky='w', pady=2)

        tk.Label(darts_frame, text="Star Wrap Bonus (Adds to Tip) (mm):", bg="#F0EAD6").grid(row=3, column=0, sticky='w', pady=2)
        tk.Entry(darts_frame, textvariable=self.dart_wrap_bonus_var, width=10).grid(row=3, column=1, sticky='w', pady=2)

        tk.Label(darts_frame, text="Star Frequency Multiplier (1.0=Default):", bg="#F0EAD6").grid(row=4, column=0, sticky='w', pady=2)
        tk.Entry(darts_frame, textvariable=self.dart_frequency_multiplier_var, width=10).grid(row=4, column=1, sticky='w', pady=2)

        # Row 5: Shape Slider
        shape_frame = tk.Frame(darts_frame, bg="#F0EAD6")
        shape_frame.grid(row=5, column=0, columnspan=2, sticky='ew', pady=5)
        
        tk.Label(shape_frame, text="Shape:", bg="#F0EAD6").pack(side="left")
        tk.Label(shape_frame, text="Sine", bg="#F0EAD6", font=("Arial", 8)).pack(side="left", padx=(5, 0))
        scale = tk.Scale(shape_frame, from_=0.0, to=1.0, orient=tk.HORIZONTAL, 
                         variable=self.dart_shape_factor_var, showvalue=0, 
                         bg="#F0EAD6", highlightthickness=0, length=150, resolution=0.01)
        scale.pack(side="left", fill="x", expand=True, padx=5)
        tk.Label(shape_frame, text="Square", bg="#F0EAD6", font=("Arial", 8)).pack(side="left")

        # Star Engraving Section (Nested Here)
        tk.Label(darts_frame, text="-------------------------", bg="#F0EAD6").grid(row=6, column=0, columnspan=2, pady=5)
        tk.Checkbutton(darts_frame, text="Show Label on Star Pads", variable=self.dart_engraving_on_var, bg="#F0EAD6").grid(row=7, column=0, columnspan=2, sticky='w', pady=2)
        
        star_loc_frame = tk.Frame(darts_frame, bg="#F0EAD6")
        star_loc_frame.grid(row=8, column=0, columnspan=2, sticky='ew', pady=2)
        tk.Radiobutton(star_loc_frame, text="outside", variable=self.dart_engraving_mode_var, value="from_outside", bg="#F0EAD6").pack(side="left")
        tk.Radiobutton(star_loc_frame, text="inside", variable=self.dart_engraving_mode_var, value="from_inside", bg="#F0EAD6").pack(side="left")
        tk.Radiobutton(star_loc_frame, text="center", variable=self.dart_engraving_mode_var, value="centered", bg="#F0EAD6").pack(side="left")
        tk.Entry(star_loc_frame, textvariable=self.dart_engraving_val_var, width=5).pack(side="left", padx=5)
        tk.Label(star_loc_frame, text="mm", bg="#F0EAD6").pack(side="left")

        engraving_frame = tk.LabelFrame(main_frame, text="Engraving Settings (Standard Pads)", bg="#F0EAD6", padx=5, pady=5)
        engraving_frame.pack(fill="x", pady=5)
        
        tk.Checkbutton(engraving_frame, text="Show Size Label", variable=self.engraving_on_var, bg="#F0EAD6").pack(anchor='w')

        font_size_frame = tk.LabelFrame(engraving_frame, text="Font Sizes (mm)", bg="#F0EAD6", padx=5, pady=5)
        font_size_frame.pack(fill='x', pady=5)
        
        materials = ['felt', 'card', 'leather', 'exact_size']
        for i, material in enumerate(materials):
            tk.Label(font_size_frame, text=f"{material.replace('_', ' ').capitalize()}:", bg="#F0EAD6").grid(row=i, column=0, sticky='w', padx=5, pady=2)
            font_size_var = tk.DoubleVar(value=self.settings["engraving_font_size"].get(material, 2.0))
            self.engraving_font_size_vars[material] = font_size_var
            tk.Entry(font_size_frame, textvariable=font_size_var, width=8).grid(row=i, column=1, sticky='w', padx=5, pady=2)


        engraving_loc_frame = tk.LabelFrame(engraving_frame, text="Placement", bg="#F0EAD6", padx=5, pady=5)
        engraving_loc_frame.pack(fill="x", pady=5)
        
        for material in materials:
            frame = tk.Frame(engraving_loc_frame, bg="#F0EAD6")
            frame.pack(fill='x', pady=2)
            tk.Label(frame, text=material.replace('_', ' ').capitalize() + ":", bg="#F0EAD6", width=10, anchor='w').pack(side="left")

            mode_var = tk.StringVar(value=self.settings["engraving_location"][material]['mode'])
            val_var = tk.DoubleVar(value=self.settings["engraving_location"][material]['value'])
            self.engraving_loc_vars[material] = {'mode': mode_var, 'value': val_var}

            tk.Radiobutton(frame, text="out", variable=mode_var, value="from_outside", bg="#F0EAD6").pack(side="left")
            tk.Radiobutton(frame, text="in", variable=mode_var, value="from_inside", bg="#F0EAD6").pack(side="left")
            tk.Radiobutton(frame, text="ctr", variable=mode_var, value="centered", bg="#F0EAD6").pack(side="left")
            
            tk.Entry(frame, textvariable=val_var, width=5).pack(side="left", padx=5)
            tk.Label(frame, text="mm", bg="#F0EAD6").pack(side="left")

        export_frame = tk.LabelFrame(main_frame, text="Export Settings", bg="#F0EAD6", padx=5, pady=5)
        export_frame.pack(fill="x", pady=5)
        tk.Checkbutton(export_frame, text="Enable Inkscape/Compatibility Mode (unitless SVG)", variable=self.compatibility_mode_var, bg="#F0EAD6").pack(anchor='w')


    def save_options(self):
        # Sizing
        self.settings["units"] = self.unit_var.get()
        self.settings["felt_offset"] = self.felt_offset_var.get()
        self.settings["card_to_felt_offset"] = self.card_offset_var.get()
        self.settings["leather_wrap_multiplier"] = self.leather_mult_var.get()
        self.settings["min_hole_size"] = self.min_hole_size_var.get()
        self.settings["felt_thickness"] = self.felt_thickness_var.get()
        self.settings["felt_thickness_unit"] = self.felt_thickness_unit_var.get()
        
        # NEW SAVE LOGIC
        self.settings["darts_enabled"] = self.darts_enabled_var.get()
        self.settings["dart_threshold"] = self.dart_threshold_var.get()
        self.settings["dart_overwrap"] = self.dart_overwrap_var.get()
        self.settings["dart_wrap_bonus"] = self.dart_wrap_bonus_var.get()
        self.settings["dart_frequency_multiplier"] = self.dart_frequency_multiplier_var.get()
        self.settings["dart_shape_factor"] = self.dart_shape_factor_var.get()
        
        # Engraving
        self.settings["engraving_on"] = self.engraving_on_var.get()
        for material, var in self.engraving_font_size_vars.items():
            self.settings["engraving_font_size"][material] = var.get()

        for material, vars in self.engraving_loc_vars.items():
            self.settings["engraving_location"][material]['mode'] = vars['mode'].get()
            self.settings["engraving_location"][material]['value'] = vars['value'].get()
            
        # NEW STAR ENGRAVING SAVE
        self.settings["dart_engraving_on"] = self.dart_engraving_on_var.get()
        self.settings["dart_engraving_loc"] = {
            "mode": self.dart_engraving_mode_var.get(),
            "value": self.dart_engraving_val_var.get()
        }
            
        # Export
        self.settings["compatibility_mode"] = self.compatibility_mode_var.get()
        
        save_settings(self.settings)
        self.update_callback()
        self.top.destroy()

    def revert_to_defaults(self):
        if messagebox.askyesno("Revert to Defaults", "Are you sure you want to revert all settings to their original defaults?"):
            # Sizing
            self.unit_var.set(DEFAULT_SETTINGS["units"])
            self.felt_offset_var.set(DEFAULT_SETTINGS["felt_offset"])
            self.card_offset_var.set(DEFAULT_SETTINGS["card_to_felt_offset"])
            self.leather_mult_var.set(DEFAULT_SETTINGS["leather_wrap_multiplier"])
            self.min_hole_size_var.set(DEFAULT_SETTINGS["min_hole_size"])
            self.felt_thickness_var.set(DEFAULT_SETTINGS["felt_thickness"])
            self.felt_thickness_unit_var.set(DEFAULT_SETTINGS["felt_thickness_unit"])
            
            # NEW REVERT LOGIC
            self.darts_enabled_var.set(DEFAULT_SETTINGS.get("darts_enabled", True))
            self.dart_threshold_var.set(DEFAULT_SETTINGS.get("dart_threshold", 18.0))
            self.dart_overwrap_var.set(DEFAULT_SETTINGS.get("dart_overwrap", 0.5))
            self.dart_wrap_bonus_var.set(DEFAULT_SETTINGS.get("dart_wrap_bonus", 0.75))
            self.dart_frequency_multiplier_var.set(DEFAULT_SETTINGS.get("dart_frequency_multiplier", 1.0))
            self.dart_shape_factor_var.set(DEFAULT_SETTINGS.get("dart_shape_factor", 0.0))
            
            # Engraving
            self.engraving_on_var.set(DEFAULT_SETTINGS["engraving_on"])
            for material, var in self.engraving_font_size_vars.items():
                 var.set(DEFAULT_SETTINGS["engraving_font_size"][material])
            
            for material, vars in self.engraving_loc_vars.items():
                 vars['mode'].set(DEFAULT_SETTINGS["engraving_location"][material]['mode'])
                 vars['value'].set(DEFAULT_SETTINGS["engraving_location"][material]['value'])
                 
            # Revert Star Engraving
            self.dart_engraving_on_var.set(True)
            self.dart_engraving_mode_var.set("from_outside")
            self.dart_engraving_val_var.set(2.5)

            # Export
            self.compatibility_mode_var.set(DEFAULT_SETTINGS.get("compatibility_mode", False))


class LayerColorWindow:
    def __init__(self, parent, settings, save_callback):
        self.settings = settings
        self.save_callback = save_callback
        
        self.top = tk.Toplevel(parent)
        self.top.title("Layer Color Mapping")
        self.top.geometry("450x420")
        self.top.configure(bg="#F0EAD6")
        self.top.transient(parent)
        self.top.grab_set()

        self.color_map = {name: hex_val for name, hex_val in LIGHTBURN_COLORS}
        color_names = list(self.color_map.keys())
        self.hex_to_name_map = {hex_val: name for name, hex_val in LIGHTBURN_COLORS}

        self.color_vars = {}

        main_frame = tk.Frame(self.top, bg="#F0EAD6", padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(1, weight=1)

        layer_map_keys = [
            'felt_outline', 'felt_center_hole', 'felt_engraving',
            'card_outline', 'card_center_hole', 'card_engraving',
            'leather_outline', 'leather_center_hole', 'leather_engraving',
            'exact_size_outline', 'exact_size_center_hole', 'exact_size_engraving'
        ]
        
        for i, key in enumerate(layer_map_keys):
            label_text = key.replace('_', ' ').capitalize() + ":"
            tk.Label(main_frame, text=label_text, bg="#F0EAD6").grid(row=i, column=0, sticky='w', pady=3)
            
            var = tk.StringVar()
            current_hex = self.settings["layer_colors"].get(key, "#000000")
            current_name = self.hex_to_name_map.get(current_hex, color_names[0])
            var.set(current_name)
            
            combo = ttk.Combobox(main_frame, textvariable=var, values=color_names, state="readonly")
            combo.grid(row=i, column=1, sticky='ew', padx=5)
            self.color_vars[key] = var

        button_frame = tk.Frame(self.top, bg="#F0EAD6")
        button_frame.grid(row=len(layer_map_keys), column=0, columnspan=2, pady=20)
        tk.Button(button_frame, text="Save", command=self.save_colors).pack(side="left", padx=10)
        tk.Button(button_frame, text="Cancel", command=self.top.destroy).pack(side="left", padx=10)

    def save_colors(self):
        for key, var in self.color_vars.items():
            selected_name = var.get()
            self.settings["layer_colors"][key] = self.color_map[selected_name]
        
        self.save_callback()
        self.top.destroy()

class KeyLayoutWindow:
    def __init__(self, parent, settings, update_callback, save_callback):
        self.settings = settings
        self.update_callback = update_callback # This is rebuild_key_tab
        self.save_callback = save_callback
        
        self.top = tk.Toplevel(parent)
        self.top.title("Key Height Layout Options")
        self.top.configure(bg="#F0EAD6")
        self.top.transient(parent)
        self.top.grab_set()
        self.top.geometry("350x450") # Give it a reasonable default size

        # --- Main Layout Frames ---
        bottom_button_frame = tk.Frame(self.top, bg="#F0EAD6")
        bottom_button_frame.pack(side="bottom", fill="x", pady=10, padx=10)
        
        tk.Button(bottom_button_frame, text="Save", command=self.save_options).pack(side="left", padx=5)
        tk.Button(bottom_button_frame, text="Cancel", command=self.top.destroy).pack(side="left", padx=5)

        # Use a scrollable frame like in OptionsWindow, as the key list is long
        main_canvas_frame = tk.Frame(self.top)
        main_canvas_frame.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(main_canvas_frame, bg="#F0EAD6", highlightthickness=0)
        self.scrollbar = tk.Scrollbar(main_canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#F0EAD6", padx=10, pady=10)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.top.bind('<MouseWheel>', self._on_mousewheel)

        self.key_layout_vars = {}
        # Get a copy to modify, matching the settings load logic
        self.key_layout_settings = DEFAULT_SETTINGS["key_layout"].copy()
        self.key_layout_settings.update(self.settings.get("key_layout", {}))


        # --- Create Widgets ---
        info_frame = tk.LabelFrame(self.scrollable_frame, text="Horn Info Layout", bg="#F0EAD6", padx=5, pady=5)
        info_frame.pack(fill="x", pady=5)

        # Serial Number Checkbox
        var = tk.BooleanVar(value=self.key_layout_settings.get("show_serial", False))
        self.key_layout_vars["show_serial"] = var
        tk.Checkbutton(info_frame, text="Show 'Serial' field", variable=var, bg="#F0EAD6").pack(anchor='w')

        # Large Notes Checkbox
        var = tk.BooleanVar(value=self.key_layout_settings.get("large_notes", False))
        self.key_layout_vars["large_notes"] = var
        tk.Checkbutton(info_frame, text="Use large 'Notes' field", variable=var, bg="#F0EAD6").pack(anchor='w')

        keys_frame = tk.LabelFrame(self.scrollable_frame, text="Visible Key Heights", bg="#F0EAD6", padx=5, pady=5)
        keys_frame.pack(fill="x", pady=5, expand=True)

        # Use ALL_KEY_HEIGHT_FIELDS to build the checkboxes
        for key_name in ALL_KEY_HEIGHT_FIELDS:
            setting_key = f"show_{key_name.replace(' ', '_')}"
            # Default to True if the setting is somehow missing from the config
            var = tk.BooleanVar(value=self.key_layout_settings.get(setting_key, True)) 
            self.key_layout_vars[setting_key] = var
            tk.Checkbutton(keys_frame, text=f"Show '{key_name}' field", variable=var, bg="#F0EAD6").pack(anchor='w')

        
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def save_options(self):
        # Update the settings dict from the vars
        for key, var in self.key_layout_vars.items():
            self.key_layout_settings[key] = var.get()
        
        self.settings["key_layout"] = self.key_layout_settings
        
        save_settings(self.settings)
        self.update_callback() # This will rebuild the key tab
        self.top.destroy()
        
class ResonanceWindow(tk.Toplevel):
    def __init__(self, parent, settings, save_callback, theme_callback):
        super().__init__(parent)
        self.settings = settings
        self.save_callback = save_callback
        self.theme_callback = theme_callback
        self.parent = parent
        
        self.title("Resonance Chamber")
        self.geometry("400x200")
        self.configure(bg="#F0EAD6")
        self.transient(parent)
        self.grab_set()

        main_frame = tk.Frame(self, bg="#F0EAD6")
        main_frame.pack(expand=True)

        res_button = tk.Button(main_frame, text="Add Resonance", command=self.start_resonance, font=("Helvetica", 14, "bold"))
        res_button.pack(pady=20, padx=40, ipadx=10, ipady=10)

    def start_resonance(self):
        self.withdraw()
        ResonanceProgressDialog(self.parent, self.settings, self.save_callback, self.theme_callback)
        self.destroy()

class ResonanceProgressDialog(tk.Toplevel):
    def __init__(self, parent, settings, save_callback, theme_callback):
        super().__init__(parent)
        self.settings = settings
        self.save_callback = save_callback
        self.theme_callback = theme_callback
        self.parent_app = parent
        
        self.title("Optimizing...")
        self.geometry("300x100")
        self.configure(bg="#F0EAD6")
        self.transient(parent)
        self.grab_set()
        
        tk.Label(self, text="Applying resonance...", bg="#F0EAD6").pack(pady=10)
        self.progress = ttk.Progressbar(self, orient="horizontal", length=250, mode="determinate")
        self.progress.pack(pady=5)
        
        self.update_progress(0)

    def update_progress(self, val):
        self.progress['value'] = val
        if val < 100:
            self.after(70, self.update_progress, val + 1)
        else:
            self.after(200, self.finish_resonance)
            
    def finish_resonance(self):
        clicks = self.settings.get("resonance_clicks", 0) + 1
        self.settings["resonance_clicks"] = clicks
        
        if clicks >= 100:
            messagebox.showinfo("Power Overwhelming", "You have become too powerful.")
            self.destroy() # Close this window
            # Start the "uninstall" process
            UninstallResonanceDialog(self.parent_app, self.settings, self.save_callback, self.theme_callback)
        else:
            self.save_callback()
            messagebox.showinfo("Success", random.choice(RESONANCE_MESSAGES))
            self.theme_callback()
            self.destroy()

class UninstallResonanceDialog(tk.Toplevel):
    def __init__(self, parent, settings, save_callback, theme_callback):
        super().__init__(parent)
        self.settings = settings
        self.save_callback = save_callback
        self.theme_callback = theme_callback
        
        self.title("Resetting...")
        self.geometry("300x100")
        self.configure(bg="#F0EAD6")
        self.transient(parent)
        self.grab_set() # This makes the main window unusable

        tk.Label(self, text="Uninstalling resonance...", bg="#F0EAD6").pack(pady=10)
        self.progress = ttk.Progressbar(self, orient="horizontal", length=250, mode="determinate")
        self.progress.pack(pady=5)
        self.update_progress(0)

    def update_progress(self, val):
        self.progress['value'] = val
        if val < 100:
            # 2 seconds total duration
            self.after(20, self.update_progress, val + 1)
        else:
            self.after(200, self.finish_uninstall)

    def finish_uninstall(self):
        self.settings["resonance_clicks"] = 0
        self.save_callback()
        self.theme_callback()
        self.destroy()

class ExportPresetsWindow(tk.Toplevel):
    def __init__(self, parent, presets, title, default_filename, ask_provenance=False):
        super().__init__(parent)
        self.presets = presets # For Pads, this is flat. For Keys, this is nested.
        self.title(title)
        self.default_filename = default_filename
        self.ask_provenance = ask_provenance
        self.geometry("400x500")
        self.configure(bg="#F0EAD6")
        self.transient(parent)
        self.grab_set()

        self.vars = {}

        tk.Label(self, text="Select sets to export:", bg="#F0EAD6", font=("Helvetica", 12)).pack(pady=10)

        button_frame = tk.Frame(self, bg="#F0EAD6")
        button_frame.pack(pady=5)
        tk.Button(button_frame, text="Select All", command=self.select_all).pack(side="left", padx=5)
        tk.Button(button_frame, text="Select None", command=self.select_none).pack(side="left", padx=5)

        list_frame = tk.Frame(self, bg="#F0EAD6")
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.canvas = tk.Canvas(list_frame, bg="#F0EAD6", highlightthickness=0)
        self.scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#F0EAD6")

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        if not presets:
             tk.Label(self.scrollable_frame, text="No local sets found.", bg="#F0EAD6").pack(pady=10)
        else:
            # Check if this is a nested dictionary (Key Libraries)
            if any(isinstance(v, dict) for v in presets.values()):
                for lib_name in sorted(self.presets.keys()):
                    tk.Label(self.scrollable_frame, text=f"[{lib_name}]", bg="#F0EAD6", font=("Helvetica", 10, "bold")).pack(anchor='w', pady=(5,0))
                    for preset_name in sorted(self.presets[lib_name].keys()):
                        var = tk.BooleanVar()
                        full_name = f"{lib_name}::{preset_name}" # Internal delimiter
                        cb = tk.Checkbutton(self.scrollable_frame, text=f"  {preset_name}", variable=var, bg="#F0EAD6")
                        cb.pack(anchor='w')
                        self.vars[full_name] = var
            else: # Flat dictionary (Pad Presets) - This is for legacy support, should be nested now
                for name in sorted(self.presets.keys()):
                    var = tk.BooleanVar()
                    cb = tk.Checkbutton(self.scrollable_frame, text=name, variable=var, bg="#F0EAD6")
                    cb.pack(anchor='w')
                    self.vars[name] = var

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.bind('<MouseWheel>', self._on_mousewheel)

        export_button = tk.Button(self, text="Export Selected", command=self.export_selected, font=("Helvetica", 10, "bold"))
        export_button.pack(pady=10)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def select_all(self):
        for var in self.vars.values():
            var.set(True)

    def select_none(self):
        for var in self.vars.values():
            var.set(False)

    def export_selected(self):
        to_export = {}
        selected_count = 0
        last_selected_data = None
        
        is_nested = any("::" in k for k in self.vars.keys())

        for name, var in self.vars.items():
            if var.get():
                if is_nested:
                    lib_name, preset_name = name.split("::", 1)
                    preset_data = self.presets[lib_name][preset_name]
                    to_export[f"[{lib_name}] {preset_name}"] = preset_data
                    selected_count += 1
                    last_selected_data = preset_data
                else:
                    to_export[name] = self.presets[name]
                    selected_count += 1

        
        if not to_export:
            messagebox.showwarning("No Selection", "Please select at least one set to export.")
            return

        initialfile = self.default_filename
        
        if self.ask_provenance:
            user_name = simpledialog.askstring("Provenance", "Enter your name (for filename):")
            if not user_name:
                user_name = "Export" # Default if cancelled
            user_name = user_name.replace(" ", "_")
            
            if selected_count == 1 and last_selected_data:
                try:
                    make = last_selected_data.get("make", "UnknownMake").replace(" ", "_")
                    model = last_selected_data.get("model", "UnknownModel").replace(" ", "_")
                    size = last_selected_data.get("size", "UnknownSize").replace(" ", "_")
                    initialfile = f"{make}_{model}_{size}_{user_name}.json"
                except Exception:
                    initialfile = f"key_height_export_{user_name}.json"
            else:
                initialfile = f"key_height_export_{user_name}.json"

        filepath = filedialog.asksaveasfilename(
            title=f"Save {self.title} As...",
            defaultextension=".json",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
            initialfile=initialfile
        )
        
        if not filepath:
            return

        try:
            with open(filepath, 'w') as f:
                json.dump(to_export, f, indent=2)
            messagebox.showinfo("Export Successful", f"Successfully exported {len(to_export)} sets.")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export presets:\n{e}")

class ImportPresetsWindow(tk.Toplevel):
    def __init__(self, parent, local_presets_lib, imported_presets, file_path, menu_widget, app_instance, preset_type_name="Preset", save_data=None):
        super().__init__(parent)
        self.parent_app = app_instance
        self.local_presets_lib = local_presets_lib # This is the specific library dict for key heights, or all presets for pads
        self.imported_presets = imported_presets # This is the dict of presets from the file
        self.file_path = file_path
        self.menu_widget = menu_widget
        self.preset_type_name = preset_type_name
        # This is the *entire* preset object (e.g., self.key_presets, which is a dict of dicts) for saving
        self.save_data = save_data if save_data is not None else local_presets_lib
        
        self.title(f"Import {preset_type_name}s")
        self.geometry("450x500")
        self.configure(bg="#F0EAD6")
        self.transient(parent)
        self.grab_set()

        self.vars = {}

        tk.Label(self, text=f"Select {preset_type_name}s to import:", bg="#F0EAD6", font=("Helvetica", 12)).pack(pady=10)

        button_frame = tk.Frame(self, bg="#F0EAD6")
        button_frame.pack(pady=5)
        tk.Button(button_frame, text="Select All", command=self.select_all).pack(side="left", padx=5)
        tk.Button(button_frame, text="Select None", command=self.select_none).pack(side="left", padx=5)

        list_frame = tk.Frame(self, bg="#F0EAD6")
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.canvas = tk.Canvas(list_frame, bg="#F0EAD6", highlightthickness=0)
        self.scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#F0EAD6")

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        if not imported_presets:
             tk.Label(self.scrollable_frame, text="No presets found in file.", bg="#F0EAD6").pack(pady=10)
        else:
            for name in sorted(self.imported_presets.keys()):
                var = tk.BooleanVar(value=True) # Default to selected
                cb = tk.Checkbutton(self.scrollable_frame, text=name, variable=var, bg="#F0EAD6")
                cb.pack(anchor='w')
                self.vars[name] = var

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.bind('<MouseWheel>', self._on_mousewheel)

        import_button = tk.Button(self, text="Import Selected", command=self.import_selected, font=("Helvetica", 10, "bold"))
        import_button.pack(pady=10)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def select_all(self):
        for var in self.vars.values():
            var.set(True)

    def select_none(self):
        for var in self.vars.values():
            var.set(False)

    def import_selected(self):
        added_count = 0
        renamed_count = 0
        
        for name, var in self.vars.items():
            if var.get():
                preset_data = self.imported_presets[name]
                new_name = name
                
                # Handle bracketed library names from key set exports
                if new_name.startswith("[") and "] " in new_name:
                    try:
                        new_name = new_name.split("] ", 1)[1]
                    except Exception:
                        pass # Keep original name if split fails
                        
                while new_name in self.local_presets_lib:
                    new_name += "*"
                
                if new_name != name:
                    renamed_count += 1
                
                self.local_presets_lib[new_name] = preset_data
                added_count += 1
        
        if added_count > 0:
            # Use the imported save_presets function from config
            if self.parent_app.save_presets(self.save_data, self.file_path):
                # Special refresh for key height library
                if self.preset_type_name == "Key Height Set":
                    self.parent_app.update_key_library_dropdown()
                else:
                    self.parent_app.update_pad_library_dropdown()
                
                messagebox.showinfo("Import Successful", 
                                  f"Import complete.\n\n"
                                  f"Added: {added_count} presets\n"
                                  f"Renamed due to conflicts: {renamed_count} presets")
            else:
                messagebox.showerror("Import Error", "Could not save new presets to file.")
        else:
            messagebox.showinfo("Import Complete", "No new presets were imported.")
            
        self.destroy()

class ImportTargetWindow(tk.Toplevel):
    def __init__(self, parent, existing_libraries):
        super().__init__(parent)
        self.parent = parent
        self.existing_libraries = existing_libraries
        self.target_library = None

        self.title("Select Import Library")
        self.geometry("350x150")
        self.configure(bg="#F0EAD6")
        self.transient(parent)
        self.grab_set()

        self.mode = tk.StringVar(value="existing")
        
        tk.Label(self, text="Where do you want to add these sets?", bg="#F0EAD6").pack(pady=10)

        existing_frame = tk.Frame(self, bg="#F0EAD6")
        existing_frame.pack(fill='x', padx=10)
        tk.Radiobutton(existing_frame, text="Add to existing library:", variable=self.mode, value="existing", bg="#F0EAD6", command=self.toggle_widgets).pack(side="left")
        self.library_dropdown = ttk.Combobox(existing_frame, values=self.existing_libraries, state="readonly", width=15)
        self.library_dropdown.pack(side="left", padx=5)
        if self.existing_libraries:
            self.library_dropdown.set(self.existing_libraries[0])

        new_frame = tk.Frame(self, bg="#F0EAD6")
        new_frame.pack(fill='x', padx=10, pady=5)
        tk.Radiobutton(new_frame, text="Create new library:", variable=self.mode, value="new", bg="#F0EAD6", command=self.toggle_widgets).pack(side="left")
        self.new_lib_entry = tk.Entry(new_frame, width=18)
        self.new_lib_entry.pack(side="left", padx=5)
        
        button_frame = tk.Frame(self, bg="#F0EAD6")
        button_frame.pack(pady=15)
        tk.Button(button_frame, text="Import", command=self.on_import).pack(side="left", padx=10)
        tk.Button(button_frame, text="Cancel", command=self.on_cancel).pack(side="left", padx=10)
        
        self.toggle_widgets()
        self.wait_window(self)

    def toggle_widgets(self):
        if self.mode.get() == "existing":
            self.library_dropdown.config(state="readonly")
            self.new_lib_entry.config(state="disabled")
        else: # "new"
            self.library_dropdown.config(state="disabled")
            self.new_lib_entry.config(state="normal")
            
    def on_import(self):
        if self.mode.get() == "existing":
            self.target_library = self.library_dropdown.get()
            if not self.target_library:
                messagebox.showwarning("No Library", "Please select a library.", parent=self)
                return
        else: # "new"
            self.target_library = self.new_lib_entry.get().strip()
            if not self.target_library:
                messagebox.showwarning("No Name", "Please enter a name for the new library.", parent=self)
                return
        
        self.destroy()
        
    def on_cancel(self):
        self.target_library = None
        self.destroy()

    def get_target_library(self):
        return self.target_library


# ==========================================
# SECTION 4: MAIN APP CLASS (main)
# ==========================================

class PadSVGGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Stohrer Sax Pad SVG Generator v2.1")
        self.root.geometry("640x720")
        self.default_bg = "#FFFDD0"
        self.root.configure(bg=self.default_bg)

        self.settings = load_settings()
        self.pad_presets = load_presets(PAD_PRESET_FILE, preset_type_name="Pad Preset")
        self.key_presets = load_presets(KEY_PRESET_FILE, preset_type_name="Key Height")
        
        self.create_menus()
        self.create_widgets() 
        
        self.apply_resonance_theme() 
        
        self.root.config(menu=self.pad_menu) 
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)

    def on_exit(self):
        # Save settings from pad generator tab
        self.settings["sheet_width"] = self.width_entry.get()
        self.settings["sheet_height"] = self.height_entry.get()
        self.settings["hole_option"] = self.hole_var.get()
        if self.hole_var.get() == "Custom":
            self.settings["custom_hole_size"] = self.custom_hole_entry.get()
        
        save_settings(self.settings)
        self.root.destroy()

    def apply_resonance_theme(self):
        clicks = self.settings.get("resonance_clicks", 0)
        color = self.default_bg
        if 10 <= clicks < 50:
            color = "#E0F7FA" # COOL_BLUE
        elif 50 <= clicks < 100:
            color = "#E8F5E9" # COOL_GREEN

        self.set_background_color(self.root, color)
        if clicks < 100:
            self.root.attributes('-alpha', 1.0)

    def set_background_color(self, parent, color):
        try:
            parent.configure(bg=color)
        except tk.TclError:
            pass
        
        style = ttk.Style()
        style.configure('App.TFrame', background=color)
        style.map('TNotebook.Tab', background=[('selected', color), ('!selected', color)], foreground=[('selected', 'black')])
        style.configure('TNotebook', background=color)

        for widget in parent.winfo_children():
            widget_class = widget.winfo_class()
            
            if widget_class in ('Frame', 'Label', 'Radiobutton', 'Checkbutton', 'LabelFrame'):
                try:
                    widget.configure(bg=color)
                except tk.TclError:
                    pass
            elif widget_class in ('TFrame', 'TLabel', 'TRadiobutton', 'TCheckbutton', 'TLabelframe', 'TNotebook'):
                try:
                    style_name = f"{widget_class}.{color.upper()}"
                    style.configure(style_name, background=color)
                    widget.configure(style=style_name)
                except tk.TclError:
                    pass

            if isinstance(widget, (tk.Frame, tk.LabelFrame, ttk.Frame, ttk.LabelFrame, ttk.Notebook)):
                self.set_background_color(widget, color)

    def create_menus(self):
        # --- Pad Generator Menu ---
        self.pad_menu = tk.Menu(self.root)
        
        pad_file_menu = tk.Menu(self.pad_menu, tearoff=0)
        self.pad_menu.add_cascade(label="File", menu=pad_file_menu)
        pad_file_menu.add_command(label="Import Pad Presets...", command=self.on_import_pad_presets)
        pad_file_menu.add_command(label="Export Pad Presets...", command=self.on_export_pad_presets)
        pad_file_menu.add_separator()
        pad_file_menu.add_command(label="Exit", command=self.on_exit)

        pad_options_menu = tk.Menu(self.pad_menu, tearoff=0)
        self.pad_menu.add_cascade(label="Options", menu=pad_options_menu)
        pad_options_menu.add_command(label="Sizing Rules...", command=self.open_options_window)
        pad_options_menu.add_command(label="Layer Colors...", command=self.open_color_window)

        # --- Key Height Library Menu ---
        self.key_menu = tk.Menu(self.root)
        
        key_file_menu = tk.Menu(self.key_menu, tearoff=0)
        self.key_menu.add_cascade(label="File", menu=key_file_menu)
        key_file_menu.add_command(label="Import Key Sets...", command=self.on_import_key_sets)
        key_file_menu.add_command(label="Export Key Sets...", command=self.on_export_key_sets)
        key_file_menu.add_separator()
        key_file_menu.add_command(label="Exit", command=self.on_exit)
        
        key_options_menu = tk.Menu(self.key_menu, tearoff=0)
        self.key_menu.add_cascade(label="Options", menu=key_options_menu)
        key_options_menu.add_command(label="Layout Options...", command=self.open_key_layout_window)

    def on_tab_changed(self, event):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0:
            self.root.config(menu=self.pad_menu)
        elif current_tab == 1:
            self.root.config(menu=self.key_menu)
        # Tab 2 is Serial Lookup, no special menu yet

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.root)
        
        # --- Create Tab 1: Pad SVG Generator ---
        self.pad_tab = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.pad_tab, text='Pad SVG Generator')
        self.create_pad_generator_tab(self.pad_tab)

        # --- Create Tab 2: Key Height Library ---
        self.key_tab = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.key_tab, text='Key Height Library')
        self.create_key_library_tab(self.key_tab)
        
        # --- Create Tab 3: Serial Lookup ---
        self.serial_tab = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.serial_tab, text='Serial Lookup')
        self.create_serial_lookup_tab(self.serial_tab)
        
        self.notebook.pack(expand=True, fill="both", padx=5, pady=5)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # Apply theme colors to the new notebook tabs
        style = ttk.Style()
        style.configure('App.TFrame', background=self.root.cget('bg'))
        style.map('TNotebook.Tab', background=[('selected', self.default_bg), ('!selected', self.default_bg)], foreground=[('selected', 'black')])
        style.configure('TNotebook', background=self.root.cget('bg'))
        
        self.apply_resonance_theme() 

    def create_serial_lookup_tab(self, parent):
        # Main layout
        frame = tk.Frame(parent, bg=self.root.cget('bg'), padx=20, pady=20)
        frame.pack(expand=True, fill='both')
        
        # Title
        tk.Label(frame, text="Saxophone Serial Number Lookup", font=("Helvetica", 16, "bold"), bg=self.root.cget('bg')).pack(pady=(0, 20))
        
        # Controls Frame
        controls_frame = tk.Frame(frame, bg=self.root.cget('bg'))
        controls_frame.pack(fill='x', pady=10)
        
        # Maker Dropdown
        tk.Label(controls_frame, text="Manufacturer:", font=("Helvetica", 12), bg=self.root.cget('bg')).grid(row=0, column=0, sticky='e', padx=10, pady=10)
        
        self.serial_maker_var = tk.StringVar()
        makers = sorted(list(SERIAL_DATA.keys())) if SERIAL_DATA else ["No Data Found"]
        self.serial_maker_dropdown = ttk.Combobox(controls_frame, textvariable=self.serial_maker_var, values=makers, state="readonly", width=25, font=("Helvetica", 12))
        if makers:
            self.serial_maker_dropdown.current(0)
        self.serial_maker_dropdown.grid(row=0, column=1, sticky='w', padx=10, pady=10)
        
        # Serial Entry
        tk.Label(controls_frame, text="Serial Number:", font=("Helvetica", 12), bg=self.root.cget('bg')).grid(row=1, column=0, sticky='e', padx=10, pady=10)
        
        self.serial_entry_var = tk.StringVar()
        self.serial_entry_var.trace("w", self.on_serial_change) # Auto-update on type
        entry = tk.Entry(controls_frame, textvariable=self.serial_entry_var, width=25, font=("Helvetica", 12))
        entry.grid(row=1, column=1, sticky='w', padx=10, pady=10)
        
        # Result Display
        self.serial_result_label = tk.Label(frame, text="Enter a serial number...", font=("Helvetica", 24, "bold"), bg=self.root.cget('bg'), fg="#0000A0")
        self.serial_result_label.pack(pady=40)
        
        # Disclaimer
        disclaimer = "Note: Dates are approximate based on available charts. Ranges represent the start of that production year."
        tk.Label(frame, text=disclaimer, font=("Helvetica", 9, "italic"), bg=self.root.cget('bg'), wraplength=400).pack(side='bottom', pady=20)

    def on_serial_change(self, *args):
        maker = self.serial_maker_var.get()
        serial = self.serial_entry_var.get()
        
        if not serial:
            self.serial_result_label.config(text="...")
            return
            
        year = lookup_serial_year(maker, serial)
        self.serial_result_label.config(text=year)

    def create_pad_generator_tab(self, parent):
        tk.Label(parent, text="Enter pad sizes (e.g. 42.0x3):", bg=self.root.cget('bg')).pack(pady=5)
        self.pad_entry = tk.Text(parent, height=10)
        self.pad_entry.pack(fill="x", padx=10)

        preset_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        preset_frame.pack(pady=10)
        
        tk.Button(preset_frame, text="Save as Preset", command=self.on_save_pad_preset).pack(side="left", padx=5)
        
        tk.Label(preset_frame, text="Library:", bg=self.root.cget('bg')).pack(side="left", padx=(10, 2))
        self.pad_library_var = tk.StringVar()
        self.pad_library_dropdown = ttk.Combobox(preset_frame, textvariable=self.pad_library_var, state="readonly", width=15)
        self.pad_library_dropdown.pack(side="left")
        self.pad_library_dropdown.bind("<<ComboboxSelected>>", self.on_pad_library_selected)
        
        preset_names = [] 
        self.pad_preset_var = tk.StringVar()
        self.pad_preset_menu = ttk.Combobox(preset_frame, textvariable=self.pad_preset_var, values=preset_names, state="readonly", width=40) 
        self.pad_preset_menu.set("Load Pad Preset")
        self.pad_preset_menu.pack(side="left", padx=5)
        self.pad_preset_menu.bind("<<ComboboxSelected>>", lambda e: self.on_load_pad_preset(self.pad_preset_var.get()))
        
        tk.Button(preset_frame, text="Delete Preset", command=self.on_delete_pad_preset).pack(side="left", padx=5)

        self.update_pad_library_dropdown() 

        tk.Label(parent, text="Select materials:", bg=self.root.cget('bg')).pack(pady=5)
        self.material_vars = {
            'felt': tk.BooleanVar(value=True), 
            'card': tk.BooleanVar(value=True), 
            'leather': tk.BooleanVar(value=True),
            'exact_size': tk.BooleanVar(value=False)
        }
        for m in self.material_vars:
            tk.Checkbutton(parent, text=m.replace('_', ' ').capitalize(), variable=self.material_vars[m], bg=self.root.cget('bg')).pack(anchor='w', padx=20)

        options_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        options_frame.pack(pady=10, fill='x', padx=10)

        hole_frame = tk.LabelFrame(options_frame, text="Center Hole", bg=self.root.cget('bg'), padx=5, pady=5)
        hole_frame.pack(fill="x")
        self.hole_var = tk.StringVar(value=self.settings["hole_option"])
        
        tk.Radiobutton(hole_frame, text="None", variable=self.hole_var, value="No center holes", bg=self.root.cget('bg'), command=self.toggle_custom_hole_entry).pack(side="left")
        tk.Radiobutton(hole_frame, text="3.0mm", variable=self.hole_var, value="3.0mm", bg=self.root.cget('bg'), command=self.toggle_custom_hole_entry).pack(side="left")
        tk.Radiobutton(hole_frame, text="3.5mm", variable=self.hole_var, value="3.5mm", bg=self.root.cget('bg'), command=self.toggle_custom_hole_entry).pack(side="left")
        tk.Radiobutton(hole_frame, text="Custom:", variable=self.hole_var, value="Custom", bg=self.root.cget('bg'), command=self.toggle_custom_hole_entry).pack(side="left")
        
        self.custom_hole_entry = tk.Entry(hole_frame, width=6)
        self.custom_hole_entry.insert(0, self.settings.get("custom_hole_size", "4.0"))
        self.custom_hole_entry.pack(side="left", padx=2)
        tk.Label(hole_frame, text="mm", bg=self.root.cget('bg')).pack(side="left")
        self.toggle_custom_hole_entry()

        sheet_frame = tk.LabelFrame(options_frame, text="Sheet Size", bg=self.root.cget('bg'), padx=5, pady=5)
        sheet_frame.pack(fill="x", pady=(10,0))

        self.unit_label = tk.Label(sheet_frame, text=f"Width ({self.settings['units']}):", bg=self.root.cget('bg'))
        self.unit_label.grid(row=0, column=0, sticky='w', padx=5)
        self.width_entry = tk.Entry(sheet_frame)
        self.width_entry.insert(0, self.settings["sheet_width"])
        self.width_entry.grid(row=0, column=1, sticky='w')

        self.height_label = tk.Label(sheet_frame, text=f"Height ({self.settings['units']}):", bg=self.root.cget('bg'))
        self.height_label.grid(row=1, column=0, sticky='w', padx=5)
        self.height_entry = tk.Entry(sheet_frame)
        self.height_entry.insert(0, self.settings["sheet_height"])
        self.height_entry.grid(row=1, column=1, sticky='w')

        tk.Label(parent, text="Output filename base (no extension):", bg=self.root.cget('bg')).pack(pady=5)
        self.filename_entry = tk.Entry(parent)
        self.filename_entry.insert(0, "my_pad_job")
        self.filename_entry.pack(padx=10) 

        tk.Button(parent, text="Generate SVGs", command=self.on_generate, font=('Helvetica', 10, 'bold')).pack(pady=15)
        
    def create_key_library_tab(self, parent):
        self.key_field_vars = {} 
        self.key_info_widgets = {} 
        self.key_height_widgets = {} 
        
        preset_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        preset_frame.pack(pady=10)
        
        tk.Button(preset_frame, text="Save as Set", command=self.on_save_key_preset).pack(side="left", padx=5)
        
        tk.Label(preset_frame, text="Library:", bg=self.root.cget('bg')).pack(side="left", padx=(10, 2))
        self.key_library_var = tk.StringVar()
        self.key_library_dropdown = ttk.Combobox(preset_frame, textvariable=self.key_library_var, state="readonly", width=15)
        self.key_library_dropdown.pack(side="left")
        self.key_library_dropdown.bind("<<ComboboxSelected>>", self.on_key_library_selected)
        
        self.key_preset_var = tk.StringVar()
        self.key_preset_menu = ttk.Combobox(preset_frame, textvariable=self.key_preset_var, state="readonly", width=40) 
        self.key_preset_menu.set("Load Key Set")
        self.key_preset_menu.pack(side="left", padx=5)
        self.key_preset_menu.bind("<<ComboboxSelected>>", lambda e: self.on_load_key_preset(self.key_preset_var.get()))
        
        tk.Button(preset_frame, text="Delete Set", command=self.on_delete_key_preset).pack(side="left", padx=5)
        
        self.update_key_library_dropdown() 

        data_frame = tk.Frame(parent, bg=self.root.cget('bg'), padx=10)
        data_frame.pack(fill="both", expand=True)

        self.horn_info_frame = tk.LabelFrame(data_frame, text="Horn Info", bg=self.root.cget('bg'), padx=5, pady=5)
        self.horn_info_frame.pack(fill="x", pady=5)
        self.horn_info_frame.columnconfigure(1, weight=1)
        
        self.key_height_frame = tk.LabelFrame(data_frame, text="Key Heights", bg=self.root.cget('bg'), padx=5, pady=5)
        self.key_height_frame.pack(fill="x", pady=5)
        self.key_height_frame.columnconfigure(1, weight=1)
        self.key_height_frame.columnconfigure(3, weight=1)
        
        self.create_key_info_widgets()
        self.create_key_height_widgets()
        self.rebuild_key_tab()


    def create_key_info_widgets(self):
        frame = self.horn_info_frame
        self.key_info_widgets = {} 
        
        default_fields = ["Make", "Model", "Size"]
        for i, field in enumerate(default_fields):
            label = tk.Label(frame, text=f"{field}:", bg=self.root.cget('bg'))
            var = tk.StringVar()
            entry = tk.Entry(frame, textvariable=var)
            self.key_field_vars[field.lower()] = var
            self.key_info_widgets[field.lower()] = (label, entry)

        label = tk.Label(frame, text="Serial:", bg=self.root.cget('bg'))
        var = tk.StringVar()
        entry = tk.Entry(frame, textvariable=var)
        self.key_field_vars["serial"] = var
        self.key_info_widgets["serial"] = (label, entry)

        label = tk.Label(frame, text="Notes:", bg=self.root.cget('bg'))
        entry = tk.Text(frame, height=3)
        self.key_field_vars['notes'] = entry
        self.key_info_widgets['notes'] = (label, entry)

    def create_key_height_widgets(self):
        frame = self.key_height_frame
        self.key_height_vars = {} 
        self.key_height_widgets = {}
        
        self.key_unit_var = tk.StringVar(value="mm")
        self.previous_key_unit = "mm"
        unit_frame = tk.Frame(frame, bg=self.root.cget('bg'))
        tk.Label(unit_frame, text="Units:", bg=self.root.cget('bg')).pack(side="left")
        tk.Radiobutton(unit_frame, text="mm", variable=self.key_unit_var, value="mm", bg=self.root.cget('bg'), command=self.on_unit_convert).pack(side="left")
        tk.Radiobutton(unit_frame, text="inches", variable=self.key_unit_var, value="in", bg=self.root.cget('bg'), command=self.on_unit_convert).pack(side="left")
        self.key_height_widgets['units'] = unit_frame 

        for key in ALL_KEY_HEIGHT_FIELDS:
            label = tk.Label(frame, text=f"{key}:", bg=self.root.cget('bg'))
            var = tk.StringVar()
            entry = tk.Entry(frame, textvariable=var, width=10)
            self.key_height_vars[key] = var
            self.key_height_widgets[key] = (label, entry)

    def rebuild_key_tab(self):
        layout_settings = self.settings.get("key_layout", DEFAULT_SETTINGS["key_layout"])

        for widget in self.horn_info_frame.winfo_children():
            widget.grid_remove()

        row = 0
        for field in ["make", "model", "size"]:
            label, entry = self.key_info_widgets[field]
            label.grid(row=row, column=0, sticky='w', padx=5, pady=2)
            entry.grid(row=row, column=1, sticky='ew', padx=5)
            row += 1
        
        if layout_settings.get("show_serial", False):
            label, entry = self.key_info_widgets["serial"]
            label.grid(row=row, column=0, sticky='w', padx=5, pady=2)
            entry.grid(row=row, column=1, sticky='ew', padx=5)
            row += 1
            
        notes_label, notes_entry = self.key_info_widgets["notes"]
        notes_height = 6 if layout_settings.get("large_notes", False) else 3
        notes_entry.config(height=notes_height)
        notes_label.grid(row=row, column=0, sticky='nw', padx=5, pady=2)
        notes_entry.grid(row=row, column=1, sticky='ew', padx=5)

        for widget in self.key_height_frame.winfo_children():
            widget.grid_remove()
            
        row = 0
        self.key_height_widgets['units'].grid(row=row, column=0, columnspan=2, sticky='w', pady=5)
        row += 1
        
        col = 0
        for key in ALL_KEY_HEIGHT_FIELDS:
            show_key = f"show_{key.replace(' ', '_')}"
            if layout_settings.get(show_key, True): 
                label, entry = self.key_height_widgets[key]
                label.grid(row=row, column=col*2, sticky='w', padx=5, pady=2)
                entry.grid(row=row, column=col*2 + 1, sticky='w', padx=5)
                
                col += 1
                if col > 1: 
                    col = 0
                    row += 1

    def on_unit_convert(self):
        new_unit = self.key_unit_var.get()
        old_unit = self.previous_key_unit

        if new_unit == old_unit:
            return

        for var in self.key_height_vars.values():
            try:
                val = float(var.get())
                if new_unit == "in" and old_unit == "mm":
                    new_val = val / 25.4
                    var.set(f"{new_val:.4f}") 
                elif new_unit == "mm" and old_unit == "in":
                    new_val = val * 25.4
                    var.set(f"{new_val:.2f}")
            except (ValueError, TypeError):
                continue 
        
        self.previous_key_unit = new_unit

    def on_save_key_preset(self):
        name = simpledialog.askstring("Save Key Height Set", "Enter a name for this set:")
        if not name:
            return
            
        active_library = self.key_library_var.get()
        if not active_library or active_library == "All Libraries":
            messagebox.showwarning("Save Error", "Please select a specific library to save to.")
            return

        make = self.key_field_vars['make'].get()
        model = self.key_field_vars['model'].get()
        size = self.key_field_vars['size'].get()
        
        if not all([make, model, size]):
            messagebox.showwarning("Missing Info", "Please fill in at least Make, Model, and Size before saving.")
            return
            
        data = {
            "make": make,
            "model": model,
            "size": size,
            "serial": self.key_field_vars['serial'].get(),
            "notes": self.key_field_vars['notes'].get("1.0", tk.END).strip(),
            "units": self.key_unit_var.get(),
            "heights": {key: var.get() for key, var in self.key_height_vars.items()}
        }
        
        if name in self.key_presets[active_library]:
            if not messagebox.askyesno("Overwrite", f"A set named '{name}' already exists in this library. Overwrite it?"):
                return
        
        self.key_presets[active_library][name] = data
        if save_presets(self.key_presets, KEY_PRESET_FILE):
            self.on_key_library_selected() 
            messagebox.showinfo("Preset Saved", f"Preset '{name}' saved successfully to '{active_library}'.")

    def on_load_key_preset(self, selected_name):
        if not selected_name or selected_name == "Load Key Set":
            return
            
        lib_name = self.key_library_var.get()
        data = None
        
        if lib_name == "All Libraries":
            try:
                lib_name, preset_name = selected_name.split("] ", 1)
                lib_name = lib_name[1:]
                if lib_name in self.key_presets and preset_name in self.key_presets[lib_name]:
                    data = self.key_presets[lib_name][preset_name]
            except ValueError:
                return
        else:
            if lib_name in self.key_presets and selected_name in self.key_presets[lib_name]:
                data = self.key_presets[lib_name][selected_name]

        if data:
            self.key_field_vars['make'].set(data.get("make", ""))
            self.key_field_vars['model'].set(data.get("model", ""))
            self.key_field_vars['size'].set(data.get("size", ""))
            
            if 'serial' in self.key_field_vars:
                self.key_field_vars['serial'].set(data.get("serial", ""))
            
            self.key_field_vars['notes'].delete("1.0", tk.END)
            self.key_field_vars['notes'].insert(tk.END, data.get("notes", ""))
            
            unit = data.get("units", "mm")
            self.key_unit_var.set(unit)
            self.previous_key_unit = unit
            
            for key, var in self.key_height_vars.items():
                var.set(data.get("heights", {}).get(key, ""))
            
    def on_delete_key_preset(self):
        selected_lib = self.key_library_var.get()
        selected_preset = self.key_preset_var.get()

        if not selected_preset or selected_preset == "Load Key Set":
            messagebox.showwarning("Delete Error", "Please load a set to delete.")
            return

        if selected_lib == "All Libraries":
            try:
                selected_lib, selected_preset = selected_preset.split("] ", 1)
                selected_lib = selected_lib[1:]
            except ValueError:
                messagebox.showerror("Delete Error", "Cannot delete from 'All Libraries' view. Please select the specific library first.")
                return

        if messagebox.askyesno("Delete Key Height Set", f"Are you sure you want to delete the set '{selected_preset}' from the '{selected_lib}' library?"):
            del self.key_presets[selected_lib][selected_preset]
            if save_presets(self.key_presets, KEY_PRESET_FILE):
                self.on_key_library_selected() 
                # Clear the form
                for var in self.key_field_vars.values():
                    if isinstance(var, tk.StringVar):
                        var.set("")
                self.key_field_vars['notes'].delete("1.0", tk.END)
                for var in self.key_height_vars.values():
                    var.set("")
                messagebox.showinfo("Preset Deleted", f"Preset '{selected_preset}' deleted.")

    def on_pad_library_selected(self, event=None):
        lib_name = self.pad_library_var.get()
        preset_list = []
        if lib_name == "All Libraries":
            for library, presets in sorted(self.pad_presets.items()):
                for name in sorted(presets.keys()):
                    preset_list.append(f"[{library}] {name}")
        else:
            preset_list = sorted(self.pad_presets.get(lib_name, {}).keys())
        
        self.pad_preset_menu['values'] = preset_list
        self.pad_preset_menu.set("Load Pad Preset")

    def update_pad_library_dropdown(self):
        lib_names = ["All Libraries"] + sorted(self.pad_presets.keys())
        self.pad_library_dropdown['values'] = lib_names
        self.pad_library_var.set("All Libraries")
        self.on_pad_library_selected()
        
    def on_key_library_selected(self, event=None):
        lib_name = self.key_library_var.get()
        preset_list = []
        if lib_name == "All Libraries":
            for library, presets in sorted(self.key_presets.items()):
                for name in sorted(presets.keys()):
                    preset_list.append(f"[{library}] {name}")
        else:
            preset_list = sorted(self.key_presets.get(lib_name, {}).keys())
        
        self.key_preset_menu['values'] = preset_list
        self.key_preset_menu.set("Load Key Set")

    def update_key_library_dropdown(self):
        lib_names = ["All Libraries"] + sorted(self.key_presets.keys())
        self.key_library_dropdown['values'] = lib_names
        self.key_library_var.set("All Libraries")
        self.on_key_library_selected()

    # --- Pad Preset Import/Export Wrappers ---
    def on_import_pad_presets(self):
        filepath = filedialog.askopenfilename(
            title="Import Pad Presets",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
            initialdir=self.settings.get("last_output_dir", "")
        )
        if not filepath:
            return
        try:
            with open(filepath, 'r') as f:
                imported_presets = json.load(f)
            if not isinstance(imported_presets, dict):
                raise TypeError("File is not a valid preset dictionary.")

            target_lib = ImportTargetWindow(self.root, list(self.pad_presets.keys())).get_target_library()
            if not target_lib:
                return 

            if target_lib not in self.pad_presets:
                self.pad_presets[target_lib] = {}

            ImportPresetsWindow(self.root, self.pad_presets[target_lib], imported_presets, PAD_PRESET_FILE, self.pad_preset_menu, self, "Pad Preset", save_data=self.pad_presets)
        except Exception as e:
            messagebox.showerror("Import Error", f"Could not import pad presets:\n{e}")

    def on_export_pad_presets(self):
        ExportPresetsWindow(self.root, self.pad_presets, "Pad Presets", "pad_preset_export.json", False)

    # --- Key Height Import/Export Wrappers ---
    def on_import_key_sets(self):
        filepath = filedialog.askopenfilename(
            title="Import Key Height Sets",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
            initialdir=self.settings.get("last_output_dir", "")
        )
        if not filepath:
            return
        try:
            with open(filepath, 'r') as f:
                imported_presets = json.load(f)
            if not isinstance(imported_presets, dict):
                raise TypeError("File is not a valid key height set file.")
            
            target_lib = ImportTargetWindow(self.root, list(self.key_presets.keys())).get_target_library()
            if not target_lib:
                return 

            if target_lib not in self.key_presets:
                self.key_presets[target_lib] = {}

            ImportPresetsWindow(self.root, self.key_presets[target_lib], imported_presets, KEY_PRESET_FILE, self.key_preset_menu, self, "Key Height Set", save_data=self.key_presets)

        except Exception as e:
            messagebox.showerror("Import Error", f"Could not import key sets:\n{e}")

    def on_export_key_sets(self):
        ExportPresetsWindow(self.root, self.key_presets, "Key Height Sets", "key_height_export.json", True)
        
    # --- Misc Wrappers ---
    def toggle_custom_hole_entry(self):
        if self.hole_var.get() == "Custom":
            self.custom_hole_entry.config(state='normal')
        else:
            self.custom_hole_entry.config(state='disabled')

    def open_options_window(self):
        OptionsWindow(self.root, self, self.settings, self.update_ui_from_settings, lambda: save_settings(self.settings))
        
    def open_key_layout_window(self):
        KeyLayoutWindow(self.root, self.settings, self.rebuild_key_tab, lambda: save_settings(self.settings))

    def open_color_window(self):
        LayerColorWindow(self.root, self.settings, lambda: save_settings(self.settings))
        
    def open_resonance_window(self):
        ResonanceWindow(self.root, self.settings, lambda: save_settings(self.settings), self.apply_resonance_theme)

    def update_ui_from_settings(self):
        self.unit_label.config(text=f"Width ({self.settings['units']}):")
        self.height_label.config(text=f"Height ({self.settings['units']}):")

    def get_hole_dia(self):
        hole_option = self.hole_var.get()
        if hole_option == "3.5mm": return 3.5
        if hole_option == "3.0mm": return 3.0
        if hole_option == "Custom":
            try:
                return float(self.custom_hole_entry.get())
            except (ValueError, TypeError):
                messagebox.showerror("Invalid Input", "Custom hole size must be a valid number.")
                return None
        return 0

    def on_generate(self):
        try:
            hole_dia = self.get_hole_dia()
            if hole_dia is None: return

            pads = self.parse_pad_list(self.pad_entry.get("1.0", tk.END))
            if not pads:
                messagebox.showerror("Error", "No valid pad sizes entered.")
                return

            if self.settings.get("engraving_on", True):
                oversized_engravings = check_for_oversized_engravings(pads, self.material_vars, self.settings)
                if oversized_engravings and self.settings.get("show_engraving_warning", True):
                    message = "Warning: The current font size is too large for some pads and the engraving will be skipped:\n\n"
                    for mat, sizes in oversized_engravings.items():
                        message += f"- {mat.replace('_', ' ').capitalize()}: {', '.join(map(str, sorted(sizes)))}\n"
                    message += "\nDo you want to proceed?"

                    dialog = ConfirmationDialog(self.root, "Engraving Size Warning", message)
                    if not dialog.result:
                        return
                    if dialog.dont_show_again.get():
                        self.settings["show_engraving_warning"] = False

            width_val = float(self.width_entry.get())
            height_val = float(self.height_entry.get())
            
            if self.settings['units'] == 'in':
                width_mm, height_mm = width_val * 25.4, height_val * 25.4
            elif self.settings['units'] == 'cm':
                width_mm, height_mm = width_val * 10, height_val * 10
            elif self.settings['units'] == 'mm':
                width_mm, height_mm = width_val, height_val
            else:
                messagebox.showerror("Error", f"Unknown unit '{self.settings['units']}' in settings.")
                return


            base = self.filename_entry.get().strip()
            if not base:
                messagebox.showerror("Error", "Please enter a base filename.")
                return
            
            for material, var in self.material_vars.items():
                if var.get() and not can_all_pads_fit(pads, material, width_mm, height_mm, self.settings):
                    messagebox.showerror("Nesting Error", f"Could not fit all '{material.replace('_',' ')}' pieces on the specified sheet size.")
                    return

            save_dir = filedialog.askdirectory(title="Select Folder to Save SVGs", initialdir=self.settings.get("last_output_dir", ""))
            if not save_dir:
                return
            
            self.settings["last_output_dir"] = save_dir 

            files_generated = False
            for material, var in self.material_vars.items():
                if var.get():
                    filename = os.path.join(save_dir, f"{base}_{material}.svg")
                    generate_svg(pads, material, width_mm, height_mm, filename, hole_dia, self.settings)
                    files_generated = True
            
            if files_generated:
                save_settings(self.settings)
                messagebox.showinfo("Done", "SVGs generated successfully.")
            else:
                messagebox.showwarning("No Materials Selected", "Please select at least one material.")

        except Exception as e:
            print(f"An error occurred during SVG generation: {e}")
            messagebox.showerror("An Error Occurred", f"Something went wrong during generation:\n\n{e}")

    def parse_pad_list(self, pad_input):
        pad_list = []
        for line in pad_input.strip().splitlines():
            try:
                size, qty = map(float, line.strip().lower().split('x'))
                pad_list.append({'size': size, 'qty': int(qty)})
            except ValueError:
                continue
        return pad_list
    
    # --- Wrapper Methods for Presets ---
    def on_save_pad_preset(self):
        self.on_save_preset(self.pad_presets, PAD_PRESET_FILE, self.pad_entry, self.pad_preset_menu, "Pad", self.pad_library_var)
        
    def on_delete_pad_preset(self):
        self.on_delete_preset(self.pad_presets, PAD_PRESET_FILE, self.pad_preset_var, self.pad_preset_menu, self.pad_entry, "Pad", self.pad_library_var, self.on_pad_library_selected)
    
    def on_load_pad_preset(self, selected_name):
        self.on_load_preset(selected_name, self.pad_presets, self.pad_entry, self.pad_library_var, "Load Pad Preset")

    # Using the shared preset logic from the class
    def on_save_preset(self, presets, file_path, entry_widget, menu_widget, preset_type_name, library_var):
        active_library = library_var.get()
        if not active_library or active_library == "All Libraries":
            messagebox.showwarning("Save Error", "Please select a specific library to save to.")
            return

        name = simpledialog.askstring(f"Save {preset_type_name} Preset", "Enter a name for this preset:")
        if name:
            text_data = entry_widget.get("1.0", tk.END)
            if not text_data.strip():
                messagebox.showwarning(f"Save {preset_type_name} Preset", "Cannot save an empty list.")
                return
            
            if active_library not in presets:
                presets[active_library] = {}

            if name in presets[active_library]:
                if not messagebox.askyesno("Overwrite", f"A set named '{name}' already exists in this library. Overwrite it?"):
                    return
            
            presets[active_library][name] = text_data
            
            if save_presets(presets, file_path):
                if preset_type_name == "Pad":
                    self.on_pad_library_selected()
                else:
                    self.on_key_library_selected()
                messagebox.showinfo("Preset Saved", f"Preset '{name}' saved successfully.")

    def on_load_preset(self, selected_name, presets, entry_widget, library_var, load_label):
        if not selected_name or selected_name == load_label:
            return
            
        lib_name = library_var.get()
        data = None
        
        if lib_name == "All Libraries":
            try:
                lib_name, preset_name = selected_name.split("] ", 1)
                lib_name = lib_name[1:] 
                if lib_name in presets and preset_name in presets[lib_name]:
                    data = presets[lib_name][preset_name]
            except ValueError:
                return 
        else:
            if lib_name in presets and selected_name in presets[lib_name]:
                data = presets[lib_name][selected_name]

        if data:
            entry_widget.delete("1.0", tk.END)
            entry_widget.insert(tk.END, data)

    def on_delete_preset(self, presets, file_path, preset_var, menu_widget, entry_widget, preset_type_name, library_var, library_refresh_func):
        selected_lib = library_var.get()
        selected_preset = preset_var.get()

        if not selected_preset or selected_preset.startswith("Load"):
            messagebox.showwarning("Delete Error", "Please load a set to delete.")
            return

        if selected_lib == "All Libraries":
            try:
                selected_lib, selected_preset = selected_preset.split("] ", 1)
                selected_lib = selected_lib[1:]
            except ValueError:
                messagebox.showerror("Delete Error", "Cannot delete from 'All Libraries' view. Please select the specific library first.")
                return
        
        if messagebox.askyesno(f"Delete {preset_type_name} Preset", f"Are you sure you want to delete the preset '{selected_preset}' from the '{selected_lib}' library?"):
            if selected_lib in presets and selected_preset in presets[selected_lib]:
                del presets[selected_lib][selected_preset]
                if save_presets(presets, file_path):
                    library_refresh_func() 
                    if isinstance(entry_widget, tk.Text):
                        entry_widget.delete("1.0", tk.END)
                    messagebox.showinfo("Preset Deleted", f"Preset '{selected_preset}' deleted.")
            else:
                messagebox.showerror("Delete Error", "Could not find the preset to delete.")

if __name__ == '__main__':
    root = tk.Tk()
    app = PadSVGGeneratorApp(root)
    root.mainloop()
