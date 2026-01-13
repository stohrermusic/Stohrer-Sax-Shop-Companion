import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import json
import os

from config import (
    ALL_KEY_HEIGHT_FIELDS, SCREW_SPECS_FILE, KEY_PRESET_FILE,
    save_presets, DEFAULT_SETTINGS
)
from ui_dialogs import (
    ImportPresetsWindow, ExportPresetsWindow, ImportTargetWindow, 
    get_unique_name
)

# --- Try to load Serial Data ---
try:
    import serials
    SERIAL_DATA = serials.SERIAL_DATA
except ImportError:
    SERIAL_DATA = {}

# ==========================================
# HELPER LOGIC
# ==========================================

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
# THE MIXIN CLASS
# ==========================================

class LibraryFeaturesMixin:
    """
    This class contains all the logic for the 'Companion' features:
    - Key Height Library
    - Serial Number Lookup
    - Screw Specs
    
    The main App class inherits from this to gain these features.
    """

    # ------------------------------------------------------------------
    # TAB 2: KEY HEIGHT LIBRARY
    # ------------------------------------------------------------------

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


    # ------------------------------------------------------------------
    # TAB 3: SERIAL LOOKUP
    # ------------------------------------------------------------------

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


    # ------------------------------------------------------------------
    # TAB 4: SCREW SPECS
    # ------------------------------------------------------------------

    def create_screw_specs_tab(self, parent):
        # Note: self.screw_data is already loaded in __init__ of main

        # --- Main Layout ---
        main_frame = tk.Frame(parent, bg=self.root.cget('bg'), padx=20, pady=20)
        main_frame.pack(expand=True, fill='both')

        title_label = tk.Label(main_frame, text="Screw & Rod Specifications", font=("Helvetica", 16, "bold"), bg=self.root.cget('bg'))
        title_label.pack(pady=(0, 15))

        # --- Controls Frame ---
        controls_frame = tk.Frame(main_frame, bg=self.root.cget('bg'))
        controls_frame.pack(fill='x', pady=5)
        controls_frame.columnconfigure(1, weight=1)

        # Maker Dropdown
        tk.Label(controls_frame, text="Manufacturer:", font=("Helvetica", 12), bg=self.root.cget('bg')).grid(row=0, column=0, sticky='e', padx=10, pady=5)
        
        self.screw_maker_var = tk.StringVar()
        self.screw_maker_dropdown = ttk.Combobox(controls_frame, textvariable=self.screw_maker_var, state="normal", width=25, font=("Helvetica", 12))
        self.screw_maker_dropdown.grid(row=0, column=1, sticky='w', padx=10, pady=5)
        self.screw_maker_dropdown.bind("<<ComboboxSelected>>", self.on_screw_maker_change)

        # Model Dropdown
        tk.Label(controls_frame, text="Model:", font=("Helvetica", 12), bg=self.root.cget('bg')).grid(row=1, column=0, sticky='e', padx=10, pady=5)
        
        self.screw_model_var = tk.StringVar()
        self.screw_model_dropdown = ttk.Combobox(controls_frame, textvariable=self.screw_model_var, state="normal", width=25, font=("Helvetica", 12))
        self.screw_model_dropdown.grid(row=1, column=1, sticky='w', padx=10, pady=5)
        self.screw_model_dropdown.bind("<<ComboboxSelected>>", self.on_screw_model_change)

        # --- Specs Frame (New Grid Layout) ---
        specs_frame = tk.LabelFrame(main_frame, text="OEM Specifications", bg=self.root.cget('bg'), font=("Helvetica", 10, "bold"), padx=10, pady=10)
        specs_frame.pack(fill='both', expand=True, pady=10)
        
        # Headers
        tk.Label(specs_frame, text="Threads / Pitch", font=("Helvetica", 9, "bold"), bg=self.root.cget('bg')).grid(row=0, column=1, padx=5, pady=(0,5))
        tk.Label(specs_frame, text="Dia / Desc", font=("Helvetica", 9, "bold"), bg=self.root.cget('bg')).grid(row=0, column=2, padx=5, pady=(0,5))

        # We will store all the Entry variables in a dictionary for easy saving/loading
        self.screw_vars = {}

        # Define the rows. Structure: (Label Text, Key_Prefix)
        rows = [
            ("Neck Receiver Screw", "neck_screw"), 
            ("Hinge Rod Tiny",   "hinge_tiny"),
            ("Hinge Rod Small",  "hinge_small"),
            ("Hinge Rod Medium", "hinge_med"),
            ("Hinge Rod Large",  "hinge_lrg"),
            ("Pivot Screw Small","pivot_small"),
            ("Pivot Screw Large","pivot_lrg"),
        ]

        row_idx = 1
        for label_text, prefix in rows:
            # Label
            tk.Label(specs_frame, text=f"{label_text}:", font=("Helvetica", 10), bg=self.root.cget('bg')).grid(row=row_idx, column=0, sticky='e', padx=5, pady=2)
            
            # Thread Entry
            th_var = tk.StringVar()
            self.screw_vars[f"{prefix}_th"] = th_var
            tk.Entry(specs_frame, textvariable=th_var, width=15).grid(row=row_idx, column=1, padx=5, pady=2)
            
            # Dia/Desc Entry
            dia_var = tk.StringVar()
            self.screw_vars[f"{prefix}_dia"] = dia_var
            tk.Entry(specs_frame, textvariable=dia_var, width=25).grid(row=row_idx, column=2, padx=5, pady=2)
            
            row_idx += 1

        # Misc Fields (Full width)
        misc_rows = [("Misc 1", "misc1"), ("Misc 2", "misc2")]
        for label_text, key in misc_rows:
            tk.Label(specs_frame, text=f"{label_text}:", font=("Helvetica", 10), bg=self.root.cget('bg')).grid(row=row_idx, column=0, sticky='e', padx=5, pady=2)
            m_var = tk.StringVar()
            self.screw_vars[key] = m_var
            tk.Entry(specs_frame, textvariable=m_var).grid(row=row_idx, column=1, columnspan=2, sticky='ew', padx=5, pady=2)
            row_idx += 1

        # Notes (Text Area)
        tk.Label(specs_frame, text="Notes:", font=("Helvetica", 10), bg=self.root.cget('bg')).grid(row=row_idx, column=0, sticky='ne', padx=5, pady=5)
        self.screw_notes_text = tk.Text(specs_frame, height=4, font=("Helvetica", 10), width=40)
        self.screw_notes_text.grid(row=row_idx, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        
        # Configure grid resizing
        specs_frame.columnconfigure(2, weight=1)

        # --- Buttons ---
        btn_frame = tk.Frame(main_frame, bg=self.root.cget('bg'))
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="Save / Update Spec", command=self.save_screw_spec, font=("Helvetica", 11, "bold")).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Delete Spec", command=self.delete_screw_spec, font=("Helvetica", 11), fg="red").pack(side="right", padx=10)

        # Initialize Dropdowns
        self.update_screw_maker_list()

    def update_screw_maker_list(self):
        makers = sorted(list(self.screw_data.keys()))
        # Add the "(add new)" option at the very top
        self.screw_maker_dropdown['values'] = ["(add new)"] + makers
        
    def on_screw_maker_change(self, event=None):
        maker = self.screw_maker_var.get()
        
        # If user selects the placeholder, clear the field for typing
        if maker == "(add new)":
            self.screw_maker_var.set("")
            self.screw_model_dropdown['values'] = []
            self.screw_model_var.set("")
            
            # Clear all entry fields
            for var in self.screw_vars.values():
                var.set("")
            self.screw_notes_text.delete("1.0", tk.END)
            return

        if maker in self.screw_data:
            models = sorted(list(self.screw_data[maker].keys()))
            self.screw_model_dropdown['values'] = ["(add new)"] + models
            self.screw_model_dropdown.set("(add new)")
            
            # Clear data fields until a valid model is picked
            for var in self.screw_vars.values():
                var.set("")
            self.screw_notes_text.delete("1.0", tk.END)
    
    def on_screw_model_change(self, event=None):
        maker = self.screw_maker_var.get()
        model = self.screw_model_var.get()
        
        # Handle the placeholder
        if model == "(add new)":
            self.screw_model_var.set("")
            for var in self.screw_vars.values():
                var.set("")
            self.screw_notes_text.delete("1.0", tk.END)
            return

        if maker in self.screw_data and model in self.screw_data[maker]:
            data = self.screw_data[maker][model]
            
            # Load each variable from the data dictionary
            for key, var in self.screw_vars.items():
                var.set(data.get(key, ""))
                
            self.screw_notes_text.delete("1.0", tk.END)
            self.screw_notes_text.insert("1.0", data.get("notes", ""))

    def save_screw_spec(self):
        maker = self.screw_maker_var.get().strip()
        model = self.screw_model_var.get().strip()
        
        if maker == "(add new)" or model == "(add new)":
            messagebox.showwarning("Invalid Name", "Please type a real name for the Manufacturer and Model.")
            return
        
        if not maker or not model:
            messagebox.showwarning("Missing Info", "Please enter both a Manufacturer and a Model.")
            return
            
        if maker not in self.screw_data:
            self.screw_data[maker] = {}

        # Build the data dictionary dynamically from our variable map
        spec_data = {}
        for key, var in self.screw_vars.items():
            spec_data[key] = var.get()
            
        spec_data["notes"] = self.screw_notes_text.get("1.0", tk.END).strip()

        self.screw_data[maker][model] = spec_data
        
        if save_presets(self.screw_data, SCREW_SPECS_FILE):
            messagebox.showinfo("Saved", f"Specs for {maker} {model} saved.")
            self.update_screw_maker_list()
            self.screw_maker_var.set(maker)
            self.on_screw_maker_change()
            self.screw_model_var.set(model)
            self.on_screw_model_change()

    def delete_screw_spec(self):
        maker = self.screw_maker_var.get()
        model = self.screw_model_var.get()
        
        if maker in self.screw_data and model in self.screw_data[maker]:
            if messagebox.askyesno("Confirm Delete", f"Delete specs for {maker} {model}?"):
                del self.screw_data[maker][model]
                if not self.screw_data[maker]:
                    del self.screw_data[maker]
                
                save_presets(self.screw_data, SCREW_SPECS_FILE)
                self.update_screw_maker_list()
                self.screw_maker_var.set("")
                self.screw_model_var.set("")
                for var in self.screw_vars.values():
                    var.set("")
                self.screw_notes_text.delete("1.0", tk.END)

    def on_export_screw_specs(self):
        ExportPresetsWindow(self.root, self.screw_data, "Screw Specs", "screw_specs_export.json", False)

    def on_import_screw_specs(self):
        filepath = filedialog.askopenfilename(
            title="Import Screw Specs",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
            initialdir=self.settings.get("last_output_dir", "")
        )
        if not filepath:
            return
            
        try:
            with open(filepath, 'r') as f:
                imported_data = json.load(f)
            
            if not isinstance(imported_data, dict):
                raise TypeError("Invalid JSON structure")

            flat_options = {}
            for maker, models in imported_data.items():
                if isinstance(models, dict):
                    for model, data in models.items():
                        flat_options[f"{maker}::{model}"] = data
            
            if not flat_options:
                messagebox.showinfo("Import", "No specs found in file.")
                return

            top = tk.Toplevel(self.root)
            top.title("Import Screw Specs")
            top.geometry("400x500")
            top.transient(self.root)
            top.grab_set()
            
            vars_dict = {}
            
            tk.Label(top, text="Select specs to import:", pady=10).pack()
            
            canvas = tk.Canvas(top)
            scrollbar = tk.Scrollbar(top, orient="vertical", command=canvas.yview)
            scroll_frame = tk.Frame(canvas)
            
            scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True, padx=10)
            scrollbar.pack(side="right", fill="y")
            
            for key in sorted(flat_options.keys()):
                maker, model = key.split("::", 1)
                var = tk.BooleanVar(value=True)
                vars_dict[key] = var
                tk.Checkbutton(scroll_frame, text=f"[{maker}] {model}", variable=var).pack(anchor='w')
                
            def do_import():
                count = 0
                for key, var in vars_dict.items():
                    if var.get():
                        maker, model = key.split("::", 1)
                        spec_data = flat_options[key]
                        
                        if maker not in self.screw_data:
                            self.screw_data[maker] = {}
                        
                        # Use get_unique_name to prevent overwrite
                        final_model_name = get_unique_name(model, self.screw_data[maker])
                        
                        self.screw_data[maker][final_model_name] = spec_data
                        count += 1
                
                save_presets(self.screw_data, SCREW_SPECS_FILE)
                self.update_screw_maker_list() 
                messagebox.showinfo("Success", f"Imported {count} specs.")
                top.destroy()
                
            tk.Button(top, text="Import Selected", command=do_import, font=("Helvetica", 10, "bold")).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import:\n{e}")
