import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import random
import json
import sys

from config import (
    DEFAULT_SETTINGS, LIGHTBURN_COLORS, RESONANCE_MESSAGES,
    ALL_KEY_HEIGHT_FIELDS, save_settings, save_presets
)

# On macOS, use native system colors for dark/light mode support.
# On Windows/Linux, use our custom cream color for dialogs.
IS_MACOS = sys.platform == 'darwin'
DIALOG_BG = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"

# ==========================================
# CROSS-PLATFORM SCROLL HELPER
# ==========================================

def bind_mousewheel(widget, canvas):
    """Bind mousewheel scrolling to a canvas, cross-platform."""
    def _on_mousewheel(event):
        if sys.platform == 'darwin':
            # macOS: delta is usually 1 or -1
            canvas.yview_scroll(int(-1 * event.delta), "units")
        else:
            # Windows: delta is usually 120 or -120
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(event):
        if event.num == 4:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")

    if sys.platform == 'linux':
        widget.bind('<Button-4>', _on_mousewheel_linux)
        widget.bind('<Button-5>', _on_mousewheel_linux)
    else:
        widget.bind('<MouseWheel>', _on_mousewheel)

# ==========================================
# HELPER UTILS
# ==========================================

def get_unique_name(name, existing_keys):
    """
    Returns 'name' if it doesn't exist in existing_keys.
    Otherwise returns 'name (2)', 'name (3)', etc.
    """
    if name not in existing_keys:
        return name
    
    i = 2
    new_name = f"{name} ({i})"
    while new_name in existing_keys:
        i += 1
        new_name = f"{name} ({i})"
    return new_name

# ==========================================
# DIALOG CLASSES
# ==========================================

class ConfirmationDialog(tk.Toplevel):
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title(title)
        self.geometry("450x150")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.result = False
        self.dont_show_again = tk.BooleanVar()

        tk.Label(self, text=message, wraplength=430, bg=DIALOG_BG, justify="left").pack(padx=10, pady=10)
        
        checkbox_frame = tk.Frame(self, bg=DIALOG_BG)
        checkbox_frame.pack(pady=5)
        tk.Checkbutton(checkbox_frame, text="Don't show this message again", variable=self.dont_show_again, bg=DIALOG_BG).pack()
        
        button_frame = tk.Frame(self, bg=DIALOG_BG)
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
        self.top.configure(bg=DIALOG_BG)
        self.top.transient(parent)
        self.top.grab_set()

        # --- Main Layout Frames ---
        bottom_button_frame = tk.Frame(self.top, bg=DIALOG_BG)
        bottom_button_frame.pack(side="bottom", fill="x", pady=10, padx=10)
        
        tk.Button(bottom_button_frame, text="Save", command=self.save_options).pack(side="left", padx=5)
        tk.Button(bottom_button_frame, text="Cancel", command=self.top.destroy).pack(side="left", padx=5)
        
        if not IS_MACOS:
            tk.Button(bottom_button_frame, text="Advanced", command=self.app.open_resonance_window).pack(side="right", padx=5)
        tk.Button(bottom_button_frame, text="Revert to Defaults", command=self.revert_to_defaults).pack(side="right", padx=5)
        
        main_canvas_frame = tk.Frame(self.top)
        main_canvas_frame.pack(side="top", fill="both", expand=True)

        self.canvas = tk.Canvas(main_canvas_frame, bg=DIALOG_BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(main_canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=DIALOG_BG, padx=10, pady=10)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        bind_mousewheel(self.top, self.canvas)

        # --- Sizing variables ---
        self.unit_var = tk.StringVar(value=self.settings["units"])
        self.felt_offset_var = tk.DoubleVar(value=self.settings["felt_offset"])
        self.card_offset_var = tk.DoubleVar(value=self.settings["card_to_felt_offset"])
        self.leather_mult_var = tk.DoubleVar(value=self.settings["leather_wrap_multiplier"])
        self.min_hole_size_var = tk.DoubleVar(value=self.settings["min_hole_size"])
        self.felt_thickness_var = tk.DoubleVar(value=self.settings["felt_thickness"])
        self.felt_thickness_unit_var = tk.StringVar(value=self.settings["felt_thickness_unit"])
        
        # DART VARS
        self.darts_enabled_var = tk.BooleanVar(value=self.settings.get("darts_enabled", True))
        self.dart_threshold_var = tk.DoubleVar(value=self.settings.get("dart_threshold", 18.0))
        self.dart_overwrap_var = tk.DoubleVar(value=self.settings.get("dart_overwrap", 0.5))
        self.dart_wrap_bonus_var = tk.DoubleVar(value=self.settings.get("dart_wrap_bonus", 0.75))
        self.dart_frequency_multiplier_var = tk.DoubleVar(value=self.settings.get("dart_frequency_multiplier", 1.0))
        self.dart_shape_factor_var = tk.DoubleVar(value=self.settings.get("dart_shape_factor", 0.0))
        
        self.engraving_on_var = tk.BooleanVar(value=self.settings["engraving_on"])
        self.compatibility_mode_var = tk.BooleanVar(value=self.settings.get("compatibility_mode", False))
        self.max_fill_style_var = tk.StringVar(value=self.settings.get("max_fill_style", "center_out"))
        self.engraving_font_size_vars = {}
        self.engraving_loc_vars = {}
        
        # Dart Engraving Vars
        self.dart_engraving_on_var = tk.BooleanVar(value=self.settings.get("dart_engraving_on", True))
        self.dart_engraving_mode_var = tk.StringVar(value=self.settings.get("dart_engraving_loc", {}).get("mode", "from_outside"))
        self.dart_engraving_val_var = tk.DoubleVar(value=self.settings.get("dart_engraving_loc", {}).get("value", 2.5))

        self.create_option_widgets()
    
    def create_option_widgets(self):
        main_frame = self.scrollable_frame
        
        unit_frame = tk.LabelFrame(main_frame, text="Sheet Units", bg=DIALOG_BG, padx=5, pady=5)
        unit_frame.pack(fill="x", pady=5)
        tk.Radiobutton(unit_frame, text="Inches (in)", variable=self.unit_var, value="in", bg=DIALOG_BG).pack(side="left", padx=5)
        tk.Radiobutton(unit_frame, text="Centimeters (cm)", variable=self.unit_var, value="cm", bg=DIALOG_BG).pack(side="left", padx=5)
        tk.Radiobutton(unit_frame, text="Millimeters (mm)", variable=self.unit_var, value="mm", bg=DIALOG_BG).pack(side="left", padx=5)

        rules_frame = tk.LabelFrame(main_frame, text="Sizing Rules (Advanced)", bg=DIALOG_BG, padx=5, pady=5)
        rules_frame.pack(fill="x", pady=5)
        rules_frame.columnconfigure(1, weight=1)

        tk.Label(rules_frame, text="Felt Diameter Reduction (mm):", bg=DIALOG_BG).grid(row=0, column=0, sticky='w', pady=2)
        tk.Entry(rules_frame, textvariable=self.felt_offset_var, width=10).grid(row=0, column=1, sticky='w', pady=2)

        tk.Label(rules_frame, text="Card Additional Reduction (mm):", bg=DIALOG_BG).grid(row=1, column=0, sticky='w', pady=2)
        tk.Entry(rules_frame, textvariable=self.card_offset_var, width=10).grid(row=1, column=1, sticky='w', pady=2)

        tk.Label(rules_frame, text="Leather Wrap Multiplier (1.00=default):", bg=DIALOG_BG).grid(row=2, column=0, sticky='w', pady=2)
        tk.Entry(rules_frame, textvariable=self.leather_mult_var, width=10).grid(row=2, column=1, sticky='w', pady=2)

        tk.Label(rules_frame, text="Min. Pad Size for Hole (mm):", bg=DIALOG_BG).grid(row=3, column=0, sticky='w', pady=2)
        tk.Entry(rules_frame, textvariable=self.min_hole_size_var, width=10).grid(row=3, column=1, sticky='w', pady=2)
        
        felt_thickness_frame = tk.Frame(rules_frame, bg=DIALOG_BG)
        felt_thickness_frame.grid(row=4, column=0, columnspan=2, sticky='w', pady=2)
        tk.Label(felt_thickness_frame, text="Felt Thickness:", bg=DIALOG_BG).pack(side="left")
        tk.Entry(felt_thickness_frame, textvariable=self.felt_thickness_var, width=10).pack(side="left", padx=5)
        tk.Radiobutton(felt_thickness_frame, text="in", variable=self.felt_thickness_unit_var, value="in", bg=DIALOG_BG).pack(side="left")
        tk.Radiobutton(felt_thickness_frame, text="mm", variable=self.felt_thickness_unit_var, value="mm", bg=DIALOG_BG).pack(side="left")

        # --- DART SETTINGS FRAME ---
        darts_frame = tk.LabelFrame(main_frame, text="Star / Dart Settings", bg=DIALOG_BG, padx=5, pady=5)
        darts_frame.pack(fill="x", pady=5)
        darts_frame.columnconfigure(1, weight=1)
        
        tk.Checkbutton(darts_frame, text="Enable Star / Dart Pattern", variable=self.darts_enabled_var, bg=DIALOG_BG).grid(row=0, column=0, columnspan=2, sticky='w', pady=2)
        
        tk.Label(darts_frame, text="Use Star Pattern below (mm):", bg=DIALOG_BG).grid(row=1, column=0, sticky='w', pady=2)
        tk.Entry(darts_frame, textvariable=self.dart_threshold_var, width=10).grid(row=1, column=1, sticky='w', pady=2)

        tk.Label(darts_frame, text="Star Safe Overwrap (Valley) (mm):", bg=DIALOG_BG).grid(row=2, column=0, sticky='w', pady=2)
        tk.Entry(darts_frame, textvariable=self.dart_overwrap_var, width=10).grid(row=2, column=1, sticky='w', pady=2)

        tk.Label(darts_frame, text="Star Wrap Bonus (Adds to Tip) (mm):", bg=DIALOG_BG).grid(row=3, column=0, sticky='w', pady=2)
        tk.Entry(darts_frame, textvariable=self.dart_wrap_bonus_var, width=10).grid(row=3, column=1, sticky='w', pady=2)

        tk.Label(darts_frame, text="Star Frequency Multiplier (1.0=Default):", bg=DIALOG_BG).grid(row=4, column=0, sticky='w', pady=2)
        tk.Entry(darts_frame, textvariable=self.dart_frequency_multiplier_var, width=10).grid(row=4, column=1, sticky='w', pady=2)

        # Row 5: Shape Slider
        shape_frame = tk.Frame(darts_frame, bg=DIALOG_BG)
        shape_frame.grid(row=5, column=0, columnspan=2, sticky='ew', pady=5)
        
        tk.Label(shape_frame, text="Shape:", bg=DIALOG_BG).pack(side="left")
        tk.Label(shape_frame, text="Sine", bg=DIALOG_BG, font=("Arial", 8)).pack(side="left", padx=(5, 0))
        scale = tk.Scale(shape_frame, from_=0.0, to=1.0, orient=tk.HORIZONTAL, 
                         variable=self.dart_shape_factor_var, showvalue=0, 
                         bg=DIALOG_BG, highlightthickness=0, length=150, resolution=0.01)
        scale.pack(side="left", fill="x", expand=True, padx=5)
        tk.Label(shape_frame, text="Square", bg=DIALOG_BG, font=("Arial", 8)).pack(side="left")

        # Star Engraving Section (Nested Here)
        tk.Label(darts_frame, text="-------------------------", bg=DIALOG_BG).grid(row=6, column=0, columnspan=2, pady=5)
        tk.Checkbutton(darts_frame, text="Show Label on Star Pads", variable=self.dart_engraving_on_var, bg=DIALOG_BG).grid(row=7, column=0, columnspan=2, sticky='w', pady=2)
        
        star_loc_frame = tk.Frame(darts_frame, bg=DIALOG_BG)
        star_loc_frame.grid(row=8, column=0, columnspan=2, sticky='ew', pady=2)
        tk.Radiobutton(star_loc_frame, text="outside", variable=self.dart_engraving_mode_var, value="from_outside", bg=DIALOG_BG).pack(side="left")
        tk.Radiobutton(star_loc_frame, text="inside", variable=self.dart_engraving_mode_var, value="from_inside", bg=DIALOG_BG).pack(side="left")
        tk.Radiobutton(star_loc_frame, text="center", variable=self.dart_engraving_mode_var, value="centered", bg=DIALOG_BG).pack(side="left")
        tk.Entry(star_loc_frame, textvariable=self.dart_engraving_val_var, width=5).pack(side="left", padx=5)
        tk.Label(star_loc_frame, text="mm", bg=DIALOG_BG).pack(side="left")

        engraving_frame = tk.LabelFrame(main_frame, text="Engraving Settings (Standard Pads)", bg=DIALOG_BG, padx=5, pady=5)
        engraving_frame.pack(fill="x", pady=5)
        
        tk.Checkbutton(engraving_frame, text="Show Size Label", variable=self.engraving_on_var, bg=DIALOG_BG).pack(anchor='w')

        font_size_frame = tk.LabelFrame(engraving_frame, text="Font Sizes (mm)", bg=DIALOG_BG, padx=5, pady=5)
        font_size_frame.pack(fill='x', pady=5)
        
        materials = ['felt', 'card', 'leather', 'exact_size']
        for i, material in enumerate(materials):
            tk.Label(font_size_frame, text=f"{material.replace('_', ' ').capitalize()}:", bg=DIALOG_BG).grid(row=i, column=0, sticky='w', padx=5, pady=2)
            font_size_var = tk.DoubleVar(value=self.settings["engraving_font_size"].get(material, 2.0))
            self.engraving_font_size_vars[material] = font_size_var
            tk.Entry(font_size_frame, textvariable=font_size_var, width=8).grid(row=i, column=1, sticky='w', padx=5, pady=2)

        engraving_loc_frame = tk.LabelFrame(engraving_frame, text="Placement", bg=DIALOG_BG, padx=5, pady=5)
        engraving_loc_frame.pack(fill="x", pady=5)
        
        for material in materials:
            frame = tk.Frame(engraving_loc_frame, bg=DIALOG_BG)
            frame.pack(fill='x', pady=2)
            tk.Label(frame, text=material.replace('_', ' ').capitalize() + ":", bg=DIALOG_BG, width=10, anchor='w').pack(side="left")

            mode_var = tk.StringVar(value=self.settings["engraving_location"][material]['mode'])
            val_var = tk.DoubleVar(value=self.settings["engraving_location"][material]['value'])
            self.engraving_loc_vars[material] = {'mode': mode_var, 'value': val_var}

            tk.Radiobutton(frame, text="out", variable=mode_var, value="from_outside", bg=DIALOG_BG).pack(side="left")
            tk.Radiobutton(frame, text="in", variable=mode_var, value="from_inside", bg=DIALOG_BG).pack(side="left")
            tk.Radiobutton(frame, text="ctr", variable=mode_var, value="centered", bg=DIALOG_BG).pack(side="left")
            
            tk.Entry(frame, textvariable=val_var, width=5).pack(side="left", padx=5)
            tk.Label(frame, text="mm", bg=DIALOG_BG).pack(side="left")

        export_frame = tk.LabelFrame(main_frame, text="Export Settings", bg=DIALOG_BG, padx=5, pady=5)
        export_frame.pack(fill="x", pady=5)
        tk.Checkbutton(export_frame, text="Enable Inkscape/Compatibility Mode (unitless SVG)", variable=self.compatibility_mode_var, bg=DIALOG_BG).pack(anchor='w')

        # Max Fill Style
        max_fill_frame = tk.LabelFrame(main_frame, text="Max Fill Style (Polygon Shapes)", bg=DIALOG_BG, padx=5, pady=5)
        max_fill_frame.pack(fill="x", pady=5)
        tk.Radiobutton(max_fill_frame, text="Center Out (fill from center outward)",
                       variable=self.max_fill_style_var, value="center_out", bg=DIALOG_BG).pack(anchor='w')
        tk.Radiobutton(max_fill_frame, text="Longest Edge (fill from longest edge inward)",
                       variable=self.max_fill_style_var, value="longest_edge", bg=DIALOG_BG).pack(anchor='w')


    def save_options(self):
        # Sizing
        self.settings["units"] = self.unit_var.get()
        self.settings["felt_offset"] = self.felt_offset_var.get()
        self.settings["card_to_felt_offset"] = self.card_offset_var.get()
        self.settings["leather_wrap_multiplier"] = self.leather_mult_var.get()
        self.settings["min_hole_size"] = self.min_hole_size_var.get()
        self.settings["felt_thickness"] = self.felt_thickness_var.get()
        self.settings["felt_thickness_unit"] = self.felt_thickness_unit_var.get()
        
        # DART SAVE LOGIC
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
            
        # STAR ENGRAVING SAVE
        self.settings["dart_engraving_on"] = self.dart_engraving_on_var.get()
        self.settings["dart_engraving_loc"] = {
            "mode": self.dart_engraving_mode_var.get(),
            "value": self.dart_engraving_val_var.get()
        }
            
        # Export
        self.settings["compatibility_mode"] = self.compatibility_mode_var.get()

        # Max Fill
        self.settings["max_fill_style"] = self.max_fill_style_var.get()

        self.save_callback()
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
            
            # REVERT DART LOGIC
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

            # Max Fill
            self.max_fill_style_var.set(DEFAULT_SETTINGS.get("max_fill_style", "center_out"))

class LayerColorWindow:
    def __init__(self, parent, settings, save_callback):
        self.settings = settings
        self.save_callback = save_callback
        
        self.top = tk.Toplevel(parent)
        self.top.title("Layer Color Mapping")
        self.top.geometry("450x420")
        self.top.configure(bg=DIALOG_BG)
        self.top.transient(parent)
        self.top.grab_set()

        self.color_map = {name: hex_val for name, hex_val in LIGHTBURN_COLORS}
        color_names = list(self.color_map.keys())
        self.hex_to_name_map = {hex_val: name for name, hex_val in LIGHTBURN_COLORS}

        self.color_vars = {}

        main_frame = tk.Frame(self.top, bg=DIALOG_BG, padx=10, pady=10)
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
            tk.Label(main_frame, text=label_text, bg=DIALOG_BG).grid(row=i, column=0, sticky='w', pady=3)
            
            var = tk.StringVar()
            current_hex = self.settings["layer_colors"].get(key, "#000000")
            current_name = self.hex_to_name_map.get(current_hex, color_names[0])
            var.set(current_name)
            
            combo = ttk.Combobox(main_frame, textvariable=var, values=color_names, state="readonly")
            combo.grid(row=i, column=1, sticky='ew', padx=5)
            self.color_vars[key] = var

        button_frame = tk.Frame(main_frame, bg=DIALOG_BG)
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
        self.update_callback = update_callback 
        self.save_callback = save_callback
        
        self.top = tk.Toplevel(parent)
        self.top.title("Key Height Layout Options")
        self.top.configure(bg=DIALOG_BG)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.geometry("350x450") 

        # --- Main Layout Frames ---
        bottom_button_frame = tk.Frame(self.top, bg=DIALOG_BG)
        bottom_button_frame.pack(side="bottom", fill="x", pady=10, padx=10)
        
        tk.Button(bottom_button_frame, text="Save", command=self.save_options).pack(side="left", padx=5)
        tk.Button(bottom_button_frame, text="Cancel", command=self.top.destroy).pack(side="left", padx=5)

        main_canvas_frame = tk.Frame(self.top)
        main_canvas_frame.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(main_canvas_frame, bg=DIALOG_BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(main_canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=DIALOG_BG, padx=10, pady=10)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        bind_mousewheel(self.top, self.canvas)

        self.key_layout_vars = {}
        # Get a copy to modify
        self.key_layout_settings = DEFAULT_SETTINGS["key_layout"].copy()
        self.key_layout_settings.update(self.settings.get("key_layout", {}))

        # --- Create Widgets ---
        info_frame = tk.LabelFrame(self.scrollable_frame, text="Horn Info Layout", bg=DIALOG_BG, padx=5, pady=5)
        info_frame.pack(fill="x", pady=5)

        # Serial Number Checkbox
        var = tk.BooleanVar(value=self.key_layout_settings.get("show_serial", False))
        self.key_layout_vars["show_serial"] = var
        tk.Checkbutton(info_frame, text="Show 'Serial' field", variable=var, bg=DIALOG_BG).pack(anchor='w')

        # Large Notes Checkbox
        var = tk.BooleanVar(value=self.key_layout_settings.get("large_notes", False))
        self.key_layout_vars["large_notes"] = var
        tk.Checkbutton(info_frame, text="Use large 'Notes' field", variable=var, bg=DIALOG_BG).pack(anchor='w')

        keys_frame = tk.LabelFrame(self.scrollable_frame, text="Visible Key Heights", bg=DIALOG_BG, padx=5, pady=5)
        keys_frame.pack(fill="x", pady=5, expand=True)

        for key_name in ALL_KEY_HEIGHT_FIELDS:
            setting_key = f"show_{key_name.replace(' ', '_')}"
            var = tk.BooleanVar(value=self.key_layout_settings.get(setting_key, True)) 
            self.key_layout_vars[setting_key] = var
            tk.Checkbutton(keys_frame, text=f"Show '{key_name}' field", variable=var, bg=DIALOG_BG).pack(anchor='w')

    def save_options(self):
        for key, var in self.key_layout_vars.items():
            self.key_layout_settings[key] = var.get()
        
        self.settings["key_layout"] = self.key_layout_settings
        
        self.save_callback()
        self.update_callback() 
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
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        main_frame = tk.Frame(self, bg=DIALOG_BG)
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
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        
        tk.Label(self, text="Applying resonance...", bg=DIALOG_BG).pack(pady=10)
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
            self.destroy() 
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
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set() 

        tk.Label(self, text="Uninstalling resonance...", bg=DIALOG_BG).pack(pady=10)
        self.progress = ttk.Progressbar(self, orient="horizontal", length=250, mode="determinate")
        self.progress.pack(pady=5)
        self.update_progress(0)

    def update_progress(self, val):
        self.progress['value'] = val
        if val < 100:
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
        self.presets = presets 
        self.title(title)
        self.default_filename = default_filename
        self.ask_provenance = ask_provenance
        self.geometry("400x500")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.vars = {}

        tk.Label(self, text="Select sets to export:", bg=DIALOG_BG, font=("Helvetica", 12)).pack(pady=10)

        button_frame = tk.Frame(self, bg=DIALOG_BG)
        button_frame.pack(pady=5)
        tk.Button(button_frame, text="Select All", command=self.select_all).pack(side="left", padx=5)
        tk.Button(button_frame, text="Select None", command=self.select_none).pack(side="left", padx=5)

        list_frame = tk.Frame(self, bg=DIALOG_BG)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.canvas = tk.Canvas(list_frame, bg=DIALOG_BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=DIALOG_BG)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        if not presets:
             tk.Label(self.scrollable_frame, text="No local sets found.", bg=DIALOG_BG).pack(pady=10)
        else:
            # Check if this is a nested dictionary (Key Libraries)
            if any(isinstance(v, dict) for v in presets.values()):
                for lib_name in sorted(self.presets.keys()):
                    tk.Label(self.scrollable_frame, text=f"[{lib_name}]", bg=DIALOG_BG, font=("Helvetica", 10, "bold")).pack(anchor='w', pady=(5,0))
                    for preset_name in sorted(self.presets[lib_name].keys()):
                        var = tk.BooleanVar()
                        full_name = f"{lib_name}::{preset_name}" # Internal delimiter
                        cb = tk.Checkbutton(self.scrollable_frame, text=f"  {preset_name}", variable=var, bg=DIALOG_BG)
                        cb.pack(anchor='w')
                        self.vars[full_name] = var
            else: # Flat dictionary (Pad Presets)
                for name in sorted(self.presets.keys()):
                    var = tk.BooleanVar()
                    cb = tk.Checkbutton(self.scrollable_frame, text=name, variable=var, bg=DIALOG_BG)
                    cb.pack(anchor='w')
                    self.vars[name] = var

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        bind_mousewheel(self, self.canvas)

        export_button = tk.Button(self, text="Export Selected", command=self.export_selected, font=("Helvetica", 10, "bold"))
        export_button.pack(pady=10)

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
        self.local_presets_lib = local_presets_lib 
        self.imported_presets = imported_presets 
        self.file_path = file_path
        self.menu_widget = menu_widget
        self.preset_type_name = preset_type_name
        self.save_data = save_data if save_data is not None else local_presets_lib
        
        self.title(f"Import {preset_type_name}s")
        self.geometry("450x500")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.vars = {}

        tk.Label(self, text=f"Select {preset_type_name}s to import:", bg=DIALOG_BG, font=("Helvetica", 12)).pack(pady=10)

        button_frame = tk.Frame(self, bg=DIALOG_BG)
        button_frame.pack(pady=5)
        tk.Button(button_frame, text="Select All", command=self.select_all).pack(side="left", padx=5)
        tk.Button(button_frame, text="Select None", command=self.select_none).pack(side="left", padx=5)

        list_frame = tk.Frame(self, bg=DIALOG_BG)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.canvas = tk.Canvas(list_frame, bg=DIALOG_BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=DIALOG_BG)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        if not imported_presets:
             tk.Label(self.scrollable_frame, text="No presets found in file.", bg=DIALOG_BG).pack(pady=10)
        else:
            for name in sorted(self.imported_presets.keys()):
                var = tk.BooleanVar(value=True) # Default to selected
                cb = tk.Checkbutton(self.scrollable_frame, text=name, variable=var, bg=DIALOG_BG)
                cb.pack(anchor='w')
                self.vars[name] = var

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        bind_mousewheel(self, self.canvas)

        import_button = tk.Button(self, text="Import Selected", command=self.import_selected, font=("Helvetica", 10, "bold"))
        import_button.pack(pady=10)

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
                        pass 
                        
                while new_name in self.local_presets_lib:
                    new_name += "*"
                
                if new_name != name:
                    renamed_count += 1
                
                self.local_presets_lib[new_name] = preset_data
                added_count += 1
        
        if added_count > 0:
            if save_presets(self.save_data, self.file_path):
                # Special refresh
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
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.mode = tk.StringVar(value="existing")
        
        tk.Label(self, text="Where do you want to add these sets?", bg=DIALOG_BG).pack(pady=10)

        existing_frame = tk.Frame(self, bg=DIALOG_BG)
        existing_frame.pack(fill='x', padx=10)
        tk.Radiobutton(existing_frame, text="Add to existing library:", variable=self.mode, value="existing", bg=DIALOG_BG, command=self.toggle_widgets).pack(side="left")
        self.library_dropdown = ttk.Combobox(existing_frame, values=self.existing_libraries, state="readonly", width=15)
        self.library_dropdown.pack(side="left", padx=5)
        if self.existing_libraries:
            self.library_dropdown.set(self.existing_libraries[0])

        new_frame = tk.Frame(self, bg=DIALOG_BG)
        new_frame.pack(fill='x', padx=10, pady=5)
        tk.Radiobutton(new_frame, text="Create new library:", variable=self.mode, value="new", bg=DIALOG_BG, command=self.toggle_widgets).pack(side="left")
        self.new_lib_entry = tk.Entry(new_frame, width=18)
        self.new_lib_entry.pack(side="left", padx=5)
        
        button_frame = tk.Frame(self, bg=DIALOG_BG)
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


class PolygonDrawWindow(tk.Toplevel):
    """
    A dialog for drawing a polygon shape on a grid.
    Used for defining irregular leather skin shapes.
    """
    MAX_POINTS = 8
    CANVAS_PX = 450  # Canvas size in pixels
    POINT_RADIUS = 6  # Radius of drawn points in pixels
    CLOSE_THRESHOLD = 15  # Pixels - how close to first point to auto-close

    def __init__(self, parent, unit="in"):
        super().__init__(parent)
        self.unit = unit
        self.polygon_closed = False
        self.result = None  # Will hold the final polygon or None if cancelled

        # Grid size depends on unit: 15x15 inches or 40x40 cm
        # Each grid square = 1 unit (1 inch or 1 cm)
        if self.unit == "in":
            self.grid_size = 15  # 15x15 inches, 15 squares
        else:
            self.grid_size = 40  # 40x40 cm, 40 squares

        self.points = []  # List of (x, y) in grid units (0-15 for inches, 0-40 for cm)

        self.title("Draw Custom Shape")
        self.geometry("520x620")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        # Calculate pixels per grid unit (1 inch or 1 cm per square)
        self.px_per_unit = self.CANVAS_PX / self.grid_size

        self._create_widgets()
        self._draw_grid()

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.wait_window(self)

    def _create_widgets(self):
        # Instructions
        unit_label = "inches" if self.unit == "in" else "cm"
        instr_text = f"Click grid points to draw shape (max {self.MAX_POINTS} points).\nClick near first point to close. Click a point to remove it."
        tk.Label(self, text=instr_text, bg=DIALOG_BG, justify="center").pack(pady=(10, 5))

        # Grid info - each square = 1 unit (inch or cm)
        tk.Label(self, text=f"Grid: {self.grid_size}x{self.grid_size} {unit_label} (1 square = 1 {self.unit})",
                 bg=DIALOG_BG, font=("Helvetica", 9)).pack(pady=(0, 5))

        # Canvas frame
        canvas_frame = tk.Frame(self, bg=DIALOG_BG)
        canvas_frame.pack(padx=10, pady=5)

        self.canvas = tk.Canvas(canvas_frame, width=self.CANVAS_PX, height=self.CANVAS_PX,
                                 bg="white", highlightthickness=1, highlightbackground="gray")
        self.canvas.pack()
        self.canvas.bind("<Button-1>", self.on_canvas_click)

        # Status label
        self.status_var = tk.StringVar(value="Click to add points...")
        tk.Label(self, textvariable=self.status_var, bg=DIALOG_BG, font=("Helvetica", 10)).pack(pady=5)

        # Buttons
        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Clear", command=self.on_clear, width=10).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Submit", command=self.on_submit, width=10,
                  font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Cancel", command=self.on_cancel, width=10).pack(side="left", padx=5)

    def _draw_grid(self):
        """Draw the grid lines on the canvas."""
        self.canvas.delete("grid")

        # Draw grid lines (1 line per unit)
        # Highlight every 5th line for easier orientation (especially useful for cm grid)
        for i in range(self.grid_size + 1):
            px = i * self.px_per_unit
            # Use reddish color for every 5th line (0, 5, 10, 15, ...)
            if i % 5 == 0:
                line_color = "#CC9999"
            else:
                line_color = "#CCCCCC"
            # Vertical lines
            self.canvas.create_line(px, 0, px, self.CANVAS_PX, fill=line_color, tags="grid")
            # Horizontal lines
            self.canvas.create_line(0, px, self.CANVAS_PX, px, fill=line_color, tags="grid")

        # Draw axis labels - every 5 for inches, every 10 for cm (to avoid crowding)
        label_step = 5 if self.unit == "in" else 10
        for i in range(0, self.grid_size + 1, label_step):
            px = i * self.px_per_unit
            # X-axis labels (bottom) - label value = grid coordinate (1 square = 1 unit)
            self.canvas.create_text(px, self.CANVAS_PX - 5, text=str(i),
                                     font=("Helvetica", 8), anchor="s", tags="grid")
            # Y-axis labels (left) - invert Y so 0 is at bottom
            self.canvas.create_text(5, self.CANVAS_PX - px, text=str(i),
                                     font=("Helvetica", 8), anchor="w", tags="grid")

    def _grid_to_canvas(self, gx, gy):
        """Convert grid coordinates to canvas pixels. Y is inverted (0 at bottom)."""
        cx = gx * self.px_per_unit
        cy = self.CANVAS_PX - (gy * self.px_per_unit)
        return cx, cy

    def _canvas_to_grid(self, cx, cy):
        """Convert canvas pixels to grid coordinates, snapped to nearest grid point."""
        gx = round(cx / self.px_per_unit)
        gy = round((self.CANVAS_PX - cy) / self.px_per_unit)
        # Clamp to grid bounds
        gx = max(0, min(self.grid_size, gx))
        gy = max(0, min(self.grid_size, gy))
        return gx, gy

    def _redraw_polygon(self):
        """Redraw all points and lines."""
        self.canvas.delete("polygon")

        if not self.points:
            return

        # Draw lines between points
        if len(self.points) > 1:
            canvas_points = [self._grid_to_canvas(p[0], p[1]) for p in self.points]
            for i in range(len(canvas_points) - 1):
                x1, y1 = canvas_points[i]
                x2, y2 = canvas_points[i + 1]
                self.canvas.create_line(x1, y1, x2, y2, fill="blue", width=2, tags="polygon")

            # If closed, draw line from last to first
            if self.polygon_closed:
                x1, y1 = canvas_points[-1]
                x2, y2 = canvas_points[0]
                self.canvas.create_line(x1, y1, x2, y2, fill="blue", width=2, tags="polygon")

        # Draw points
        for i, (gx, gy) in enumerate(self.points):
            cx, cy = self._grid_to_canvas(gx, gy)
            color = "green" if i == 0 else "red"
            self.canvas.create_oval(
                cx - self.POINT_RADIUS, cy - self.POINT_RADIUS,
                cx + self.POINT_RADIUS, cy + self.POINT_RADIUS,
                fill=color, outline="black", tags="polygon"
            )

    def _update_status(self):
        """Update the status label."""
        if self.polygon_closed:
            self.status_var.set(f"Shape closed ({len(self.points)} points). Click Submit or adjust points.")
        elif len(self.points) >= self.MAX_POINTS:
            self.status_var.set(f"Max points reached ({self.MAX_POINTS}). Click near first point to close.")
        else:
            remaining = self.MAX_POINTS - len(self.points)
            self.status_var.set(f"{len(self.points)} points. {remaining} remaining. Click near first point to close.")

    def on_canvas_click(self, event):
        """Handle click on canvas."""
        if self.polygon_closed:
            # If already closed, check if clicking on a point to remove it
            clicked_idx = self._get_clicked_point_index(event.x, event.y)
            if clicked_idx is not None:
                self.points.pop(clicked_idx)
                self.polygon_closed = False
                self._redraw_polygon()
                self._update_status()
            return

        # Check if clicking on an existing point (to remove it)
        clicked_idx = self._get_clicked_point_index(event.x, event.y)
        if clicked_idx is not None:
            self.points.pop(clicked_idx)
            self._redraw_polygon()
            self._update_status()
            return

        # Get grid coordinates
        gx, gy = self._canvas_to_grid(event.x, event.y)

        # Check if this would close the polygon (clicking near first point)
        if len(self.points) >= 3:
            first_cx, first_cy = self._grid_to_canvas(self.points[0][0], self.points[0][1])
            dist = ((event.x - first_cx) ** 2 + (event.y - first_cy) ** 2) ** 0.5
            if dist < self.CLOSE_THRESHOLD:
                self.polygon_closed = True
                self._redraw_polygon()
                self._update_status()
                return

        # Check if we can add more points
        if len(self.points) >= self.MAX_POINTS:
            return

        # Check if point already exists at this location
        if (gx, gy) in self.points:
            return

        # Add the point
        self.points.append((gx, gy))
        self._redraw_polygon()
        self._update_status()

    def _get_clicked_point_index(self, cx, cy):
        """Return index of point near click, or None."""
        for i, (gx, gy) in enumerate(self.points):
            pcx, pcy = self._grid_to_canvas(gx, gy)
            dist = ((cx - pcx) ** 2 + (cy - pcy) ** 2) ** 0.5
            if dist <= self.POINT_RADIUS + 4:  # Small tolerance
                return i
        return None

    def on_clear(self):
        """Clear all points."""
        self.points = []
        self.polygon_closed = False
        self._redraw_polygon()
        self._update_status()

    def on_submit(self):
        """Submit the polygon."""
        if len(self.points) < 3:
            from tkinter import messagebox
            messagebox.showwarning("Not Enough Points",
                                   "Please draw at least 3 points to create a shape.",
                                   parent=self)
            return

        if not self.polygon_closed:
            from tkinter import messagebox
            messagebox.showwarning("Shape Not Closed",
                                   "Please close the shape by clicking near the first point.",
                                   parent=self)
            return

        # Return points as list of (x, y) tuples in grid units
        self.result = list(self.points)
        self.destroy()

    def on_cancel(self):
        """Cancel and close."""
        self.result = None
        self.destroy()

    def get_polygon(self):
        """Return the polygon points or None if cancelled."""
        return self.result


# ==========================================
# G-CODE SETTINGS DIALOG
# ==========================================

class GcodeSettingsWindow:
    """Dialog for configuring G-code laser settings per material."""

    MATERIALS = [
        ("felt", "Felt"),
        ("card", "Card"),
        ("leather", "Leather"),
    ]

    OPERATIONS = [
        ("engraving", "Engraving"),
        ("hole", "Center Hole"),
        ("cut", "Outer Cut"),
    ]

    def __init__(self, parent, settings, save_callback):
        self.settings = settings
        self.save_callback = save_callback

        self.top = tk.Toplevel(parent)
        self.top.title("G-code Laser Settings")
        self.top.geometry("550x500")
        self.top.configure(bg=DIALOG_BG)
        self.top.transient(parent)
        self.top.grab_set()

        # Get current gcode settings or defaults
        self.gcode_settings = settings.get("gcode_settings", {})

        # Create variable storage
        self.vars = {}  # vars[material][operation]['speed'|'power']

        self._create_widgets()

    def _create_widgets(self):
        # Header
        header_frame = tk.Frame(self.top, bg=DIALOG_BG)
        header_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(header_frame, text="Configure laser speed and power for each material and operation.",
                 bg=DIALOG_BG, wraplength=500, justify="left").pack(anchor="w")

        tk.Label(header_frame, text="Order: Engraving → Center Hole → Outer Cut",
                 bg=DIALOG_BG, font=("Helvetica", 9, "italic")).pack(anchor="w", pady=(5, 0))

        tk.Label(header_frame, text="Note: Power uses Grbl's S0-S1000 scale. If power seems wrong, check that "
                 "your machine's $30 setting is 1000 (run \"$30=1000\" in your console).",
                 bg=DIALOG_BG, font=("Helvetica", 8), fg="#666666", wraplength=500, justify="left").pack(anchor="w", pady=(5, 0))

        # Main content with scrollable frame
        main_canvas_frame = tk.Frame(self.top)
        main_canvas_frame.pack(fill="both", expand=True, padx=10)

        canvas = tk.Canvas(main_canvas_frame, bg=DIALOG_BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(main_canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=DIALOG_BG)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        bind_mousewheel(self.top, canvas)

        # Create material sections
        for mat_key, mat_label in self.MATERIALS:
            self._create_material_section(scrollable_frame, mat_key, mat_label)

        # Buttons
        button_frame = tk.Frame(self.top, bg=DIALOG_BG)
        button_frame.pack(fill="x", padx=10, pady=10)

        tk.Button(button_frame, text="Save", command=self._on_save).pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self.top.destroy).pack(side="left", padx=5)
        tk.Button(button_frame, text="Reset to Defaults", command=self._reset_defaults).pack(side="right", padx=5)

    def _create_material_section(self, parent, mat_key, mat_label):
        """Create a settings section for one material."""
        frame = tk.LabelFrame(parent, text=mat_label, bg=DIALOG_BG, padx=10, pady=10)
        frame.pack(fill="x", pady=5, padx=5)

        self.vars[mat_key] = {}

        # Get current settings for this material
        mat_settings = self.gcode_settings.get(mat_key, {})

        # Header row
        tk.Label(frame, text="Operation", bg=DIALOG_BG, font=("Helvetica", 9, "bold")).grid(row=0, column=0, sticky="w", padx=5)
        tk.Label(frame, text="Speed (mm/min)", bg=DIALOG_BG, font=("Helvetica", 9, "bold")).grid(row=0, column=1, padx=5)
        tk.Label(frame, text="Power (%)", bg=DIALOG_BG, font=("Helvetica", 9, "bold")).grid(row=0, column=2, padx=5)

        for i, (op_key, op_label) in enumerate(self.OPERATIONS, start=1):
            self.vars[mat_key][op_key] = {}

            # Get defaults from DEFAULT_SETTINGS
            default_speed = self._get_default(mat_key, f"{op_key}_speed")
            default_power = self._get_default(mat_key, f"{op_key}_power")

            # Current values
            current_speed = mat_settings.get(f"{op_key}_speed", default_speed)
            current_power = mat_settings.get(f"{op_key}_power", default_power)

            speed_var = tk.IntVar(value=int(current_speed))
            power_var = tk.DoubleVar(value=current_power)

            self.vars[mat_key][op_key]['speed'] = speed_var
            self.vars[mat_key][op_key]['power'] = power_var

            tk.Label(frame, text=op_label, bg=DIALOG_BG).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            tk.Entry(frame, textvariable=speed_var, width=10).grid(row=i, column=1, padx=5, pady=2)
            tk.Entry(frame, textvariable=power_var, width=10).grid(row=i, column=2, padx=5, pady=2)

        # Kerf width row (after operations)
        kerf_row = len(self.OPERATIONS) + 1
        tk.Label(frame, text="Kerf width:", bg=DIALOG_BG).grid(row=kerf_row, column=0, sticky="w", padx=5, pady=(8, 2))

        default_kerf = self._get_default(mat_key, "kerf_width")
        current_kerf = mat_settings.get("kerf_width", default_kerf)
        kerf_var = tk.DoubleVar(value=current_kerf if current_kerf else 0.0)
        self.vars[mat_key]['kerf_width'] = kerf_var

        kerf_entry = tk.Spinbox(frame, textvariable=kerf_var, from_=0.0, to=1.0,
                                increment=0.05, width=8, format="%.2f")
        kerf_entry.grid(row=kerf_row, column=1, sticky="w", padx=5, pady=(8, 2))
        tk.Label(frame, text="mm", bg=DIALOG_BG).grid(row=kerf_row, column=2, sticky="w", padx=5, pady=(8, 2))

    def _get_default(self, material, setting_key):
        """Get default value from DEFAULT_SETTINGS."""
        defaults = DEFAULT_SETTINGS.get("gcode_settings", {}).get(material, {})
        return defaults.get(setting_key, 100)  # Fallback to 100 if not found

    def _on_save(self):
        """Save settings and close."""
        # Build gcode_settings dict from variables
        new_gcode_settings = {}

        for mat_key, _ in self.MATERIALS:
            new_gcode_settings[mat_key] = {}
            for op_key, _ in self.OPERATIONS:
                try:
                    speed = self.vars[mat_key][op_key]['speed'].get()
                    power = self.vars[mat_key][op_key]['power'].get()
                    new_gcode_settings[mat_key][f"{op_key}_speed"] = speed
                    new_gcode_settings[mat_key][f"{op_key}_power"] = power
                except tk.TclError:
                    messagebox.showerror("Invalid Input",
                                         f"Invalid value for {mat_key} {op_key}. Please enter valid numbers.",
                                         parent=self.top)
                    return

            # Save kerf width per material
            try:
                new_gcode_settings[mat_key]["kerf_width"] = self.vars[mat_key]['kerf_width'].get()
            except tk.TclError:
                messagebox.showerror("Invalid Input",
                                     f"Invalid kerf width for {mat_key}. Please enter a valid number.",
                                     parent=self.top)
                return

        # Also copy leather settings to leather_topgrain
        new_gcode_settings["leather_topgrain"] = dict(new_gcode_settings["leather"])

        # Update settings
        self.settings["gcode_settings"] = new_gcode_settings
        self.save_callback(self.settings)

        self.top.destroy()

    def _reset_defaults(self):
        """Reset all values to defaults."""
        default_gcode = DEFAULT_SETTINGS.get("gcode_settings", {})

        for mat_key, _ in self.MATERIALS:
            mat_defaults = default_gcode.get(mat_key, {})
            for op_key, _ in self.OPERATIONS:
                default_speed = mat_defaults.get(f"{op_key}_speed", 100)
                default_power = mat_defaults.get(f"{op_key}_power", 10)
                self.vars[mat_key][op_key]['speed'].set(default_speed)
                self.vars[mat_key][op_key]['power'].set(default_power)

            # Reset kerf width per material
            default_kerf = mat_defaults.get("kerf_width", 0.0)
            self.vars[mat_key]['kerf_width'].set(default_kerf)
