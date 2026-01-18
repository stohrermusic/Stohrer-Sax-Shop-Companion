import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import json

# --- Local Imports ---
from config import (
    load_settings, save_settings, load_presets, save_presets,
    PAD_PRESET_FILE, KEY_PRESET_FILE, SCREW_SPECS_FILE,
    DEFAULT_SETTINGS,
    find_config_files_in_directory, import_config_files
)
from svg_engine import generate_svg, can_all_pads_fit, check_for_oversized_engravings
from ui_dialogs import (
    OptionsWindow, LayerColorWindow, KeyLayoutWindow,
    ResonanceWindow, ConfirmationDialog,
    ImportPresetsWindow, ExportPresetsWindow, ImportTargetWindow,
    PolygonDrawWindow
)
from library_features import LibraryFeaturesMixin

# ==========================================
# MAIN APP CLASS
# ==========================================

class PadSVGGeneratorApp(LibraryFeaturesMixin):
    def __init__(self, root):
        self.root = root
        self.root.title("Stohrer Sax Shop Companion")
        self.root.geometry("640x720")
        self.default_bg = "#FFFDD0"
        self.root.configure(bg=self.default_bg)

        self.settings = load_settings()
        self.pad_presets = load_presets(PAD_PRESET_FILE, preset_type_name="Pad Preset")
        self.key_presets = load_presets(KEY_PRESET_FILE, preset_type_name="Key Height")
        self.custom_polygon = None  # For custom shape nesting
        
        # --- Screw Specs Init ---
        if not os.path.exists(SCREW_SPECS_FILE):
            save_presets({}, SCREW_SPECS_FILE)
        self.screw_data = load_presets(SCREW_SPECS_FILE, preset_type_name="Screw Specs")
        
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
        pad_file_menu.add_command(label="Import Settings from Folder...", command=self.on_import_settings_folder)
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

        # --- Screw Specs Menu ---
        self.screw_menu = tk.Menu(self.root)
        
        screw_file_menu = tk.Menu(self.screw_menu, tearoff=0)
        self.screw_menu.add_cascade(label="File", menu=screw_file_menu)
        screw_file_menu.add_command(label="Import Screw Specs...", command=self.on_import_screw_specs)
        screw_file_menu.add_command(label="Export Screw Specs...", command=self.on_export_screw_specs)
        screw_file_menu.add_separator()
        screw_file_menu.add_command(label="Exit", command=self.on_exit)

    def on_tab_changed(self, event):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0:
            self.root.config(menu=self.pad_menu)
        elif current_tab == 1:
            self.root.config(menu=self.key_menu)
        elif current_tab == 2:
            self.root.config(menu=tk.Menu(self.root)) # Empty menu for serials
        elif current_tab == 3:
            self.root.config(menu=self.screw_menu)

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

        # --- Create Tab 4: Screw Specs ---
        self.screw_tab = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.screw_tab, text='Screw Specs')
        self.create_screw_specs_tab(self.screw_tab)
        
        self.notebook.pack(expand=True, fill="both", padx=5, pady=5)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # Apply theme colors to the new notebook tabs
        style = ttk.Style()
        style.configure('App.TFrame', background=self.root.cget('bg'))
        style.map('TNotebook.Tab', background=[('selected', self.default_bg), ('!selected', self.default_bg)], foreground=[('selected', 'black')])
        style.configure('TNotebook', background=self.root.cget('bg'))
        
        self.apply_resonance_theme() 

    # ------------------------------------------------------------------
    # TAB 1: PAD GENERATOR LOGIC
    # ------------------------------------------------------------------

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

        # Custom shape controls
        shape_btn_frame = tk.Frame(sheet_frame, bg=self.root.cget('bg'))
        shape_btn_frame.grid(row=2, column=0, columnspan=2, pady=(8, 0))

        tk.Button(shape_btn_frame, text="Draw Custom Shape...", command=self.on_draw_custom_shape).pack(side="left")
        self.shape_status_var = tk.StringVar(value="")
        self.shape_status_label = tk.Label(shape_btn_frame, textvariable=self.shape_status_var,
                                           bg=self.root.cget('bg'), fg="gray", font=("Helvetica", 9))
        self.shape_status_label.pack(side="left", padx=5)

        self.unload_shape_btn = tk.Button(shape_btn_frame, text="Unload", command=self.on_unload_custom_shape)
        # Initially hidden, shown when shape is loaded
        self._update_shape_status()

        tk.Label(parent, text="Output filename base (no extension):", bg=self.root.cget('bg')).pack(pady=5)
        self.filename_entry = tk.Entry(parent)
        self.filename_entry.insert(0, "my_pad_job")
        self.filename_entry.pack(padx=10) 

        tk.Button(parent, text="Generate SVGs", command=self.on_generate, font=('Helvetica', 10, 'bold')).pack(pady=15)

    def toggle_custom_hole_entry(self):
        if self.hole_var.get() == "Custom":
            self.custom_hole_entry.config(state='normal')
        else:
            self.custom_hole_entry.config(state='disabled')

    def _update_shape_status(self):
        """Update the custom shape status indicator."""
        if self.custom_polygon:
            self.shape_status_var.set(f"Drawn shape loaded ({len(self.custom_polygon)} pts)")
            self.shape_status_label.config(fg="green")
            self.unload_shape_btn.pack(side="left", padx=2)
        else:
            self.shape_status_var.set("Using rectangle dimensions")
            self.shape_status_label.config(fg="gray")
            self.unload_shape_btn.pack_forget()

    def _show_draw_shape_tutorial(self):
        """Show first-time tutorial for the polygon drawing tool."""
        unit = self.settings.get("units", "in")
        if unit == "mm":
            unit = "cm"
        unit_label = "inches" if unit == "in" else "centimeters"

        msg = (
            "Draw Custom Shape - How to Use\n\n"
            f"• The grid is 15×15 {unit_label} (1 square = 1 {unit})\n"
            "• Click on grid intersections to add points (max 8)\n"
            "• Click near the first (green) point to close the shape\n"
            "• Click on any point to remove it\n"
            "• Use 'Clear' to start over\n"
            "• Click 'Submit' when your shape is complete\n\n"
            "This is useful for irregular leather skins and scrap pieces.\n\n"
            "Note: Generation can take 5-10x longer for complex shapes."
        )
        messagebox.showinfo("Draw Custom Shape", msg)

    def on_draw_custom_shape(self):
        """Open the polygon drawing window."""
        # Show tutorial on first use
        if not self.settings.get("seen_polygon_tutorial", False):
            self._show_draw_shape_tutorial()
            self.settings["seen_polygon_tutorial"] = True
            save_settings(self.settings)

        unit = self.settings.get("units", "in")
        # Convert mm to appropriate display unit
        if unit == "mm":
            unit = "cm"  # Use cm for the grid when mm is selected

        dialog = PolygonDrawWindow(self.root, unit=unit)
        polygon = dialog.get_polygon()

        if polygon:
            # Convert to mm for internal use
            # Flip Y axis: drawing uses Y=0 at bottom, SVG uses Y=0 at top
            # Grid is always 15x15 units (PolygonDrawWindow.GRID_SIZE)
            grid_size = 15
            if unit == "in":
                # polygon is in inches, convert to mm
                self.custom_polygon = [(x * 25.4, (grid_size - y) * 25.4) for (x, y) in polygon]
            else:
                # polygon is in cm, convert to mm
                self.custom_polygon = [(x * 10, (grid_size - y) * 10) for (x, y) in polygon]
            self._update_shape_status()

    def on_unload_custom_shape(self):
        """Unload the custom shape and return to rectangle mode."""
        self.custom_polygon = None
        self._update_shape_status()

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

            # Validate max pads - only one allowed
            max_pads = [p for p in pads if p['qty'] == 'max']
            if len(max_pads) > 1:
                messagebox.showerror("Error", "Only one pad size can use 'max' quantity at a time.")
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
                if var.get() and not can_all_pads_fit(pads, material, width_mm, height_mm, self.settings, polygon=self.custom_polygon):
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
                    generate_svg(pads, material, width_mm, height_mm, filename, hole_dia, self.settings, polygon=self.custom_polygon)
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
        """
        Parse pad input. Supports:
        - Regular: "18.0 x 5" (size x quantity)
        - Max fill: "18.0 x max" (fill remaining space with this size)
        Only one pad size can use "max" at a time.
        """
        pad_list = []
        for line in pad_input.strip().splitlines():
            line = line.strip().lower()
            if not line:
                continue
            try:
                parts = line.split('x', 1)  # Split only on first 'x' (so 'max' doesn't get split)
                if len(parts) != 2:
                    continue
                size = float(parts[0].strip())
                qty_str = parts[1].strip()
                if qty_str == 'max':
                    pad_list.append({'size': size, 'qty': 'max'})
                else:
                    pad_list.append({'size': size, 'qty': int(float(qty_str))})
            except ValueError:
                continue
        return pad_list

    # --- Pad Presets Wrappers ---

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

    def on_save_pad_preset(self):
        active_library = self.pad_library_var.get()
        if not active_library or active_library == "All Libraries":
            messagebox.showwarning("Save Error", "Please select a specific library to save to.")
            return

        name = simpledialog.askstring("Save Pad Preset", "Enter a name for this preset:")
        if name:
            text_data = self.pad_entry.get("1.0", tk.END)
            if not text_data.strip():
                messagebox.showwarning("Save Pad Preset", "Cannot save an empty list.")
                return
            
            if active_library not in self.pad_presets:
                self.pad_presets[active_library] = {}

            if name in self.pad_presets[active_library]:
                if not messagebox.askyesno("Overwrite", f"A set named '{name}' already exists in this library. Overwrite it?"):
                    return
            
            self.pad_presets[active_library][name] = text_data
            
            if save_presets(self.pad_presets, PAD_PRESET_FILE):
                self.on_pad_library_selected()
                messagebox.showinfo("Preset Saved", f"Preset '{name}' saved successfully.")

    def on_load_pad_preset(self, selected_name):
        if not selected_name or selected_name == "Load Pad Preset":
            return
            
        lib_name = self.pad_library_var.get()
        data = None
        
        if lib_name == "All Libraries":
            try:
                lib_name, preset_name = selected_name.split("] ", 1)
                lib_name = lib_name[1:] 
                if lib_name in self.pad_presets and preset_name in self.pad_presets[lib_name]:
                    data = self.pad_presets[lib_name][preset_name]
            except ValueError:
                return 
        else:
            if lib_name in self.pad_presets and selected_name in self.pad_presets[lib_name]:
                data = self.pad_presets[lib_name][selected_name]

        if data:
            self.pad_entry.delete("1.0", tk.END)
            self.pad_entry.insert(tk.END, data)

    def on_delete_pad_preset(self):
        selected_lib = self.pad_library_var.get()
        selected_preset = self.pad_preset_var.get()

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
        
        if messagebox.askyesno("Delete Pad Preset", f"Are you sure you want to delete the preset '{selected_preset}' from the '{selected_lib}' library?"):
            if selected_lib in self.pad_presets and selected_preset in self.pad_presets[selected_lib]:
                del self.pad_presets[selected_lib][selected_preset]
                if save_presets(self.pad_presets, PAD_PRESET_FILE):
                    self.on_pad_library_selected() 
                    self.pad_entry.delete("1.0", tk.END)
                    messagebox.showinfo("Preset Deleted", f"Preset '{selected_preset}' deleted.")
            else:
                messagebox.showerror("Delete Error", "Could not find the preset to delete.")

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

    def on_import_settings_folder(self):
        """Import all config files from a user-selected folder."""
        folder = filedialog.askdirectory(
            title="Select Folder with Settings",
            initialdir=self.settings.get("last_output_dir", "")
        )
        if not folder:
            return

        found_files = find_config_files_in_directory(folder)
        if not found_files:
            messagebox.showinfo("No Settings Found",
                "No config files found in the selected folder.\n\n"
                "Looking for: app_settings.json, pad_presets.json, key_height_library.json, screw_specs.json")
            return

        file_list = "\n".join(f"  • {f}" for f in found_files)
        msg = (f"The following files will be imported and will REPLACE your current settings:\n\n"
               f"{file_list}\n\n"
               f"Are you sure you want to continue?")

        if not messagebox.askyesno("Confirm Import", msg):
            return

        try:
            import_config_files(folder, found_files)

            # Reload all data from the newly imported files
            self.settings = load_settings()
            self.pad_presets = load_presets(PAD_PRESET_FILE, preset_type_name="Pad Preset")
            self.key_presets = load_presets(KEY_PRESET_FILE, preset_type_name="Key Height")
            self.screw_data = load_presets(SCREW_SPECS_FILE, preset_type_name="Screw Specs")

            # Refresh UI dropdowns
            self.update_pad_library_dropdown()
            self.update_key_library_dropdown()
            self.update_screw_maker_list()

            # Update UI elements that depend on settings
            self.update_ui_from_settings()
            self.apply_resonance_theme()

            messagebox.showinfo("Import Complete",
                f"Successfully imported {len(found_files)} file(s).\n\n"
                f"Imported: {', '.join(found_files)}")
        except Exception as e:
            messagebox.showerror("Import Error", f"Could not import settings:\n{e}")

    # --- Misc Windows ---
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

if __name__ == '__main__':
    root = tk.Tk()
    app = PadSVGGeneratorApp(root)
    root.mainloop()
