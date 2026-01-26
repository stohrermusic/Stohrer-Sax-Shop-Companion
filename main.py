import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import json
import shutil
import sys
import glob
import subprocess

# --- Local Imports ---
from config import (
    load_settings, save_settings, load_presets, save_presets,
    PAD_PRESET_FILE, KEY_PRESET_FILE, SCREW_SPECS_FILE,
    DEFAULT_SETTINGS,
    find_config_files_in_directory, import_config_files
)
from svg_engine import generate_svg, can_all_pads_fit, check_for_oversized_engravings, try_nest_partial, generate_svg_from_placed
from gcode_engine import generate_gcode, generate_gcode_from_placed
from ui_dialogs import (
    OptionsWindow, LayerColorWindow, KeyLayoutWindow,
    ResonanceWindow, ConfirmationDialog,
    ImportPresetsWindow, ExportPresetsWindow, ImportTargetWindow,
    PolygonDrawWindow, GcodeSettingsWindow
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

        # --- Scrap Mode Session State ---
        self.scrap_session = {
            'active': False,
            'original_pads': [],       # Original pads for progress tracking
            'remaining_pads': [],      # [{'size': 18.0, 'qty': 3}, ...]
            'scrap_count': 0,          # Files generated in this session
            'material': None,          # Single material locked for session
            'save_dir': '',            # Output directory
            'hole_dia': 0,             # Locked hole diameter
        }
        self.scrap_remaining_window = None  # Popup showing progress

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
        # Warn if scrap session is active with remaining pads
        if self.scrap_session.get('active', False):
            remaining = self._count_remaining_pads()
            if remaining > 0:
                if not messagebox.askyesno("Scrap Session Active",
                    f"You have {remaining} pads remaining in your scrap session.\n\n"
                    "Exit anyway? (Session will be lost)"):
                    return

        # Save settings from pad generator tab
        self.settings["sheet_width"] = self.width_entry.get()
        self.settings["sheet_height"] = self.height_entry.get()
        self.settings["hole_option"] = self.hole_var.get()
        if self.hole_var.get() == "Custom":
            self.settings["custom_hole_size"] = self.custom_hole_entry.get()
        # Save card paper size settings
        self.settings["card_use_paper_size"] = self.card_paper_var.get()
        dropdown_val = self.card_paper_dropdown.get().lower()
        self.settings["card_paper_size"] = "a4" if dropdown_val.startswith("a4") else "letter"

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
        pad_file_menu.add_command(label="Send G-code to SD Card...", command=self.on_send_to_sd_card)
        pad_file_menu.add_separator()
        pad_file_menu.add_command(label="Exit", command=self.on_exit)

        pad_options_menu = tk.Menu(self.pad_menu, tearoff=0)
        self.pad_menu.add_cascade(label="Options", menu=pad_options_menu)
        pad_options_menu.add_command(label="Sizing Rules...", command=self.open_options_window)
        pad_options_menu.add_command(label="Layer Colors...", command=self.open_color_window)
        pad_options_menu.add_separator()
        pad_options_menu.add_command(label="G-code Settings...", command=self.open_gcode_settings_window)

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
            self.root.config(menu=tk.Menu(self.root))  # Empty menu for serials
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
        self.material_checkboxes = {}  # Store references for enable/disable
        # Two-column layout for materials
        materials_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        materials_frame.pack(anchor='w', padx=20)
        material_list = list(self.material_vars.items())
        for i, (m, var) in enumerate(material_list):
            row, col = i // 2, i % 2
            cb = tk.Checkbutton(materials_frame, text=m.replace('_', ' ').capitalize(),
                               variable=var, bg=self.root.cget('bg'))
            cb.grid(row=row, column=col, sticky='w', padx=(0, 20))
            self.material_checkboxes[m] = cb

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

        # Scrap Mode checkbox and status (right side of sheet frame, centered)
        scrap_inner_frame = tk.Frame(sheet_frame, bg=self.root.cget('bg'))
        scrap_inner_frame.grid(row=0, column=2, rowspan=2, sticky='n', padx=(20, 5))

        self.scrap_mode_var = tk.BooleanVar(value=False)
        tk.Checkbutton(scrap_inner_frame, text="Scrap Mode",
                       variable=self.scrap_mode_var, bg=self.root.cget('bg'),
                       command=self._toggle_scrap_mode).pack()

        # Status label (shown when session active)
        self.scrap_status_var = tk.StringVar(value="")
        self.scrap_status_label = tk.Label(scrap_inner_frame,
                                           textvariable=self.scrap_status_var,
                                           bg=self.root.cget('bg'), font=("Helvetica", 8), fg="blue")

        # Clear button (shown when scrap mode checked)
        self.clear_scrap_btn = tk.Button(scrap_inner_frame, text="Clear", font=("Helvetica", 8),
                                         command=self._on_clear_scrap_clicked)

        # Card paper size option
        card_paper_frame = tk.Frame(sheet_frame, bg=self.root.cget('bg'))
        card_paper_frame.grid(row=2, column=0, columnspan=2, sticky='w', pady=(8, 0))

        self.card_paper_var = tk.BooleanVar(value=self.settings.get("card_use_paper_size", False))
        self.card_paper_checkbox = tk.Checkbutton(
            card_paper_frame, text="Fit card to paper:", variable=self.card_paper_var,
            bg=self.root.cget('bg'), command=self._toggle_card_paper_dropdown
        )
        self.card_paper_checkbox.pack(side="left")

        self.card_paper_size_var = tk.StringVar(value=self.settings.get("card_paper_size", "letter"))
        self.card_paper_dropdown = ttk.Combobox(
            card_paper_frame, textvariable=self.card_paper_size_var, state="readonly", width=18,
            values=["letter (8.5×11 in)", "a4 (210×297 mm)"]
        )
        # Set display value based on stored setting
        if self.card_paper_size_var.get() == "a4":
            self.card_paper_dropdown.set("a4 (210×297 mm)")
        else:
            self.card_paper_dropdown.set("letter (8.5×11 in)")
        self.card_paper_dropdown.pack(side="left", padx=(5, 0))
        self._toggle_card_paper_dropdown()

        # Custom shape controls
        shape_btn_frame = tk.Frame(sheet_frame, bg=self.root.cget('bg'))
        shape_btn_frame.grid(row=3, column=0, columnspan=2, pady=(8, 0))

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

        # Two generate buttons side by side
        generate_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        generate_frame.pack(pady=15)
        tk.Button(generate_frame, text="Generate SVG", command=self.on_generate_svg, font=('Helvetica', 10, 'bold')).pack(side="left", padx=5)
        tk.Button(generate_frame, text="Generate G-code", command=self.on_generate_gcode, font=('Helvetica', 10, 'bold')).pack(side="left", padx=5)

    def toggle_custom_hole_entry(self):
        if self.hole_var.get() == "Custom":
            self.custom_hole_entry.config(state='normal')
        else:
            self.custom_hole_entry.config(state='disabled')

    def _toggle_card_paper_dropdown(self):
        """Enable/disable the card paper size dropdown based on checkbox state."""
        if self.card_paper_var.get():
            self.card_paper_dropdown.config(state="readonly")
        else:
            self.card_paper_dropdown.config(state="disabled")

    def _get_card_paper_dimensions_mm(self):
        """Return (width_mm, height_mm) for the selected paper size, or None if not using paper size."""
        if not self.card_paper_var.get():
            return None
        # Parse the dropdown value to determine paper size
        dropdown_val = self.card_paper_dropdown.get().lower()
        if dropdown_val.startswith("a4"):
            return (210.0, 297.0)  # A4 in mm
        else:
            return (8.5 * 25.4, 11.0 * 25.4)  # Letter: 215.9 x 279.4 mm

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
            # Grid size depends on unit: 15x15 inches or 40x40 cm
            if unit == "in":
                grid_size = 15  # 15x15 inches
                # Convert inches to mm
                self.custom_polygon = [(x * 25.4, (grid_size - y) * 25.4) for (x, y) in polygon]
            else:
                grid_size = 40  # 40x40 cm
                # Convert cm to mm
                self.custom_polygon = [(x * 10, (grid_size - y) * 10) for (x, y) in polygon]
            self._update_shape_status()

    def on_unload_custom_shape(self):
        """Unload the custom shape and return to rectangle mode."""
        self.custom_polygon = None
        self._update_shape_status()

    # --- Scrap Mode Methods ---

    def _toggle_scrap_mode(self):
        """Called when scrap mode checkbox changes."""
        if self.scrap_mode_var.get():
            # Turning on - validate only one material is selected
            selected = [m for m, v in self.material_vars.items() if v.get()]
            if len(selected) != 1:
                messagebox.showwarning("Scrap Mode",
                    "Please select exactly one material to use Scrap Mode.")
                self.scrap_mode_var.set(False)
                return
        else:
            # Turning off - warn if session active
            if self.scrap_session['active']:
                remaining = self._count_remaining_pads()
                if remaining > 0:
                    if not messagebox.askyesno("Clear Session?",
                        f"You have {remaining} pads remaining.\n"
                        "Disabling scrap mode will clear the session.\n\nContinue?"):
                        self.scrap_mode_var.set(True)
                        return
                self._clear_scrap_session()
        self._update_scrap_status_display()

    def _clear_scrap_session(self):
        """Reset scrap session state."""
        self.scrap_session = {
            'active': False,
            'original_pads': [],
            'remaining_pads': [],
            'scrap_count': 0,
            'material': None,
            'save_dir': '',
            'hole_dia': 0,
        }
        self._unlock_material_selection()
        self._close_remaining_pads_window()
        self._update_scrap_status_display()

    def _update_scrap_status_display(self):
        """Update the scrap mode status UI."""
        if self.scrap_mode_var.get():
            # Show Clear button when scrap mode is checked
            self.clear_scrap_btn.pack(pady=(2, 0))

            if self.scrap_session['active']:
                # Show status when session is active
                remaining = self._count_remaining_pads()
                if remaining == 0:
                    self.scrap_status_var.set("Done!")
                    self.scrap_status_label.config(fg="green")
                else:
                    count = self.scrap_session['scrap_count']
                    self.scrap_status_var.set(f"{remaining} left ({count} scraps)")
                    self.scrap_status_label.config(fg="blue")
                self.scrap_status_label.pack()
            else:
                # No active session - hide status label
                self.scrap_status_label.pack_forget()
        else:
            # Scrap mode unchecked - hide everything
            self.scrap_status_label.pack_forget()
            self.clear_scrap_btn.pack_forget()

    def _on_clear_scrap_clicked(self):
        """Handle Clear button click - clears session and unchecks scrap mode."""
        if self.scrap_session['active']:
            remaining = self._count_remaining_pads()
            if remaining > 0:
                if not messagebox.askyesno("Clear Session?",
                    f"You have {remaining} pads remaining.\n\nClear session?"):
                    return
        self._clear_scrap_session()
        self.scrap_mode_var.set(False)
        self._update_scrap_status_display()

    def _count_remaining_pads(self):
        """Count total remaining pads in session."""
        return sum(p['qty'] for p in self.scrap_session['remaining_pads'])

    def _start_scrap_session(self, pads, material, save_dir, hole_dia):
        """Initialize a new scrap session."""
        # Only include fixed-qty pads, not 'max' pads
        # Store both original and remaining for progress tracking
        fixed_pads = [{'size': p['size'], 'qty': p['qty']}
                      for p in pads if p['qty'] != 'max']
        self.scrap_session = {
            'active': True,
            'original_pads': [p.copy() for p in fixed_pads],  # For "done" calculation
            'remaining_pads': fixed_pads,
            'scrap_count': 0,
            'material': material,
            'save_dir': save_dir,
            'hole_dia': hole_dia,
        }
        self._lock_material_selection(material)
        self._open_remaining_pads_window()

    def _open_remaining_pads_window(self):
        """Open or update the floating window showing remaining and done pads."""
        if self.scrap_remaining_window is not None:
            try:
                self.scrap_remaining_window.destroy()
            except tk.TclError:
                pass

        self.scrap_remaining_window = tk.Toplevel(self.root)
        self.scrap_remaining_window.title("Scrap Mode Progress")
        self.scrap_remaining_window.geometry("320x300")
        self.scrap_remaining_window.configure(bg=self.default_bg)
        self.scrap_remaining_window.resizable(True, True)

        # Position to the right of main window
        self.root.update_idletasks()
        x = self.root.winfo_x() + self.root.winfo_width() + 10
        y = self.root.winfo_y()
        self.scrap_remaining_window.geometry(f"+{x}+{y}")

        # Header with progress
        header_frame = tk.Frame(self.scrap_remaining_window, bg=self.default_bg)
        header_frame.pack(fill="x", padx=10, pady=(10, 5))
        self.scrap_window_header = tk.Label(header_frame, text="", bg=self.default_bg,
                                            font=("Helvetica", 10, "bold"))
        self.scrap_window_header.pack(anchor="w")

        # Two-column layout for Remaining and Done
        columns_frame = tk.Frame(self.scrap_remaining_window, bg=self.default_bg)
        columns_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Remaining column
        remaining_frame = tk.Frame(columns_frame, bg=self.default_bg)
        remaining_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        tk.Label(remaining_frame, text="Remaining", bg=self.default_bg,
                 font=("Helvetica", 9, "bold"), fg="blue").pack(anchor="w")
        self.scrap_remaining_listbox = tk.Listbox(remaining_frame, font=("Courier", 10), height=10, width=12)
        self.scrap_remaining_listbox.pack(fill="both", expand=True)

        # Done column
        done_frame = tk.Frame(columns_frame, bg=self.default_bg)
        done_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))
        tk.Label(done_frame, text="Done", bg=self.default_bg,
                 font=("Helvetica", 9, "bold"), fg="green").pack(anchor="w")
        self.scrap_done_listbox = tk.Listbox(done_frame, font=("Courier", 10), height=10, width=12)
        self.scrap_done_listbox.pack(fill="both", expand=True)

        # Footer with scrap count
        self.scrap_window_footer = tk.Label(self.scrap_remaining_window, text="", bg=self.default_bg,
                                            font=("Helvetica", 9))
        self.scrap_window_footer.pack(pady=(0, 10))

        # Handle window close
        self.scrap_remaining_window.protocol("WM_DELETE_WINDOW", self._on_remaining_window_close)

        self._update_remaining_pads_window()

    def _update_remaining_pads_window(self):
        """Update the contents of the remaining pads window with two columns."""
        if self.scrap_remaining_window is None:
            return
        try:
            remaining_pads = self.scrap_session.get('remaining_pads', [])
            original_pads = self.scrap_session.get('original_pads', [])
            scraps = self.scrap_session.get('scrap_count', 0)
            material = self.scrap_session.get('material', '')

            # Calculate done pads (original - remaining)
            remaining_by_size = {p['size']: p['qty'] for p in remaining_pads}
            done_pads = []
            for orig in original_pads:
                size = orig['size']
                orig_qty = orig['qty']
                remaining_qty = remaining_by_size.get(size, 0)
                done_qty = orig_qty - remaining_qty
                if done_qty > 0:
                    done_pads.append({'size': size, 'qty': done_qty})

            total_remaining = sum(p['qty'] for p in remaining_pads)
            total_done = sum(p['qty'] for p in done_pads)
            total_original = sum(p['qty'] for p in original_pads)

            # Update header
            if total_remaining == 0:
                self.scrap_window_header.config(text="All done!", fg="green")
            else:
                self.scrap_window_header.config(
                    text=f"{total_done} / {total_original} pads complete", fg="blue")

            # Update Remaining listbox
            self.scrap_remaining_listbox.delete(0, tk.END)
            if remaining_pads:
                sorted_remaining = sorted(remaining_pads, key=lambda p: -p['size'])
                for pad in sorted_remaining:
                    size_str = f"{pad['size']:.1f}".rstrip('0').rstrip('.')
                    self.scrap_remaining_listbox.insert(tk.END, f" {pad['qty']} x {size_str}")
            else:
                self.scrap_remaining_listbox.insert(tk.END, " (none)")

            # Update Done listbox
            self.scrap_done_listbox.delete(0, tk.END)
            if done_pads:
                sorted_done = sorted(done_pads, key=lambda p: -p['size'])
                for pad in sorted_done:
                    size_str = f"{pad['size']:.1f}".rstrip('0').rstrip('.')
                    self.scrap_done_listbox.insert(tk.END, f" {pad['qty']} x {size_str}")
            else:
                self.scrap_done_listbox.insert(tk.END, " (none)")

            # Update footer
            self.scrap_window_footer.config(
                text=f"{material} | {scraps} scrap(s) used")

        except tk.TclError:
            # Window was closed
            self.scrap_remaining_window = None

    def _close_remaining_pads_window(self):
        """Close the remaining pads window."""
        if self.scrap_remaining_window is not None:
            try:
                self.scrap_remaining_window.destroy()
            except tk.TclError:
                pass
            self.scrap_remaining_window = None

    def _lock_material_selection(self, locked_material):
        """Disable material checkboxes during scrap session, keeping only the selected one checked."""
        for mat, cb in self.material_checkboxes.items():
            if mat == locked_material:
                self.material_vars[mat].set(True)
            else:
                self.material_vars[mat].set(False)
            cb.config(state='disabled')

    def _unlock_material_selection(self):
        """Re-enable material checkboxes after scrap session ends."""
        for cb in self.material_checkboxes.values():
            cb.config(state='normal')

    def _on_remaining_window_close(self):
        """Handle user closing the remaining pads window."""
        self.scrap_remaining_window = None

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

    def _prepare_generation(self):
        """Common validation and setup for SVG/G-code generation. Returns None on error, or a dict with generation params."""
        hole_dia = self.get_hole_dia()
        if hole_dia is None:
            return None

        pads = self.parse_pad_list(self.pad_entry.get("1.0", tk.END))
        if not pads:
            messagebox.showerror("Error", "No valid pad sizes entered.")
            return None

        max_pads = [p for p in pads if p['qty'] == 'max']
        if len(max_pads) > 1:
            messagebox.showerror("Error", "Only one pad size can use 'max' quantity at a time.")
            return None

        if self.settings.get("engraving_on", True):
            oversized_engravings = check_for_oversized_engravings(pads, self.material_vars, self.settings)
            if oversized_engravings and self.settings.get("show_engraving_warning", True):
                message = "Warning: The current font size is too large for some pads and the engraving will be skipped:\n\n"
                for mat, sizes in oversized_engravings.items():
                    message += f"- {mat.replace('_', ' ').capitalize()}: {', '.join(map(str, sorted(sizes)))}\n"
                message += "\nDo you want to proceed?"
                dialog = ConfirmationDialog(self.root, "Engraving Size Warning", message)
                if not dialog.result:
                    return None
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
            return None

        base = self.filename_entry.get().strip()
        if not base:
            messagebox.showerror("Error", "Please enter a base filename.")
            return None

        card_paper_dims = self._get_card_paper_dimensions_mm()

        return {
            'pads': pads, 'hole_dia': hole_dia, 'base': base,
            'width_mm': width_mm, 'height_mm': height_mm,
            'card_paper_dims': card_paper_dims
        }

    def _get_material_dimensions(self, material, width_mm, height_mm, card_paper_dims):
        """Get dimensions and polygon for a material, handling card paper size option."""
        if material == "card" and card_paper_dims:
            return card_paper_dims[0], card_paper_dims[1], None
        return width_mm, height_mm, self.custom_polygon

    def on_generate_svg(self):
        """Generate SVG files."""
        # --- Scrap Mode ---
        if self.scrap_mode_var.get():
            self._generate_svg_scrap_mode()
            return

        # --- Standard Mode ---
        try:
            params = self._prepare_generation()
            if not params:
                return

            pads, hole_dia, base = params['pads'], params['hole_dia'], params['base']
            width_mm, height_mm = params['width_mm'], params['height_mm']
            card_paper_dims = params['card_paper_dims']

            # Validate all materials fit
            for material, var in self.material_vars.items():
                if var.get():
                    mat_w, mat_h, mat_polygon = self._get_material_dimensions(material, width_mm, height_mm, card_paper_dims)
                    if not can_all_pads_fit(pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon):
                        size_desc = "paper" if (material == "card" and card_paper_dims) else "sheet"
                        messagebox.showerror("Nesting Error", f"Could not fit all '{material.replace('_',' ')}' pieces on the specified {size_desc} size.")
                        return

            save_dir = filedialog.askdirectory(title="Select Folder to Save SVGs", initialdir=self.settings.get("last_output_dir", ""))
            if not save_dir:
                return
            self.settings["last_output_dir"] = save_dir

            files_generated = False
            for material, var in self.material_vars.items():
                if var.get():
                    mat_w, mat_h, mat_polygon = self._get_material_dimensions(material, width_mm, height_mm, card_paper_dims)
                    filename = os.path.join(save_dir, f"{base}_{material}.svg")
                    generate_svg(pads, material, mat_w, mat_h, filename, hole_dia, self.settings, polygon=mat_polygon)
                    files_generated = True

            if files_generated:
                save_settings(self.settings)
                messagebox.showinfo("Done", "SVG files generated successfully.")
            else:
                messagebox.showwarning("No Materials Selected", "Please select at least one material.")

        except Exception as e:
            print(f"An error occurred during SVG generation: {e}")
            messagebox.showerror("An Error Occurred", f"Something went wrong during generation:\n\n{e}")

    def _generate_svg_scrap_mode(self):
        """Handle SVG generation in scrap mode."""
        try:
            params = self._prepare_generation()
            if not params:
                return

            pads = params['pads']
            hole_dia = params['hole_dia']
            base = params['base']
            width_mm, height_mm = params['width_mm'], params['height_mm']
            card_paper_dims = params['card_paper_dims']

            # Check material selection - must be exactly one
            selected_materials = [m for m, v in self.material_vars.items() if v.get()]
            if len(selected_materials) != 1:
                messagebox.showerror("Scrap Mode Error",
                    "Please select exactly one material for scrap mode.")
                return
            material = selected_materials[0]

            # Get scrap dimensions (polygon or rectangle)
            mat_w, mat_h, mat_polygon = self._get_material_dimensions(
                material, width_mm, height_mm, card_paper_dims)

            # Initialize or validate session
            if not self.scrap_session['active']:
                # Starting new session - ask for save directory
                save_dir = filedialog.askdirectory(
                    title="Select Folder to Save SVGs",
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return
                self.settings["last_output_dir"] = save_dir

                self._start_scrap_session(pads, material, save_dir, hole_dia)
            else:
                # Continuing existing session - validate material matches
                if self.scrap_session['material'] != material:
                    messagebox.showerror("Material Mismatch",
                        f"Current session is for {self.scrap_session['material']}.\n"
                        f"Clear session to switch materials.")
                    return
                # Use remaining pads from session
                pads = self.scrap_session['remaining_pads']
                hole_dia = self.scrap_session['hole_dia']

            if not pads:
                messagebox.showinfo("Session Complete", "All pads have been placed!")
                return

            # Attempt partial placement
            placed, remaining, any_placed = try_nest_partial(
                pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon)

            if not any_placed:
                min_pad_size = min(p['size'] for p in pads)
                messagebox.showwarning("No Pads Fit",
                    f"No pads could be placed on this scrap.\n\n"
                    f"Smallest remaining pad: {min_pad_size}mm\n"
                    f"Try a larger scrap piece.")
                return

            # Generate filename with scrap number
            self.scrap_session['scrap_count'] += 1
            scrap_num = self.scrap_session['scrap_count']
            save_dir = self.scrap_session['save_dir']
            filename = os.path.join(save_dir, f"{base}_{material}_scrap{scrap_num}.svg")

            # Generate SVG from placed discs
            generate_svg_from_placed(placed, material, mat_w, mat_h, filename,
                                     hole_dia, self.settings, polygon=mat_polygon)

            # Update session with remaining pads
            self.scrap_session['remaining_pads'] = remaining
            self._update_scrap_status_display()
            self._update_remaining_pads_window()

            # Report results
            placed_count = len(placed)
            remaining_count = self._count_remaining_pads()

            if remaining_count == 0:
                save_settings(self.settings)
                messagebox.showinfo("Session Complete!",
                    f"Placed {placed_count} pads on scrap #{scrap_num}.\n\n"
                    f"All pads placed! Session complete.\n"
                    f"Files saved to: {save_dir}")
            else:
                messagebox.showinfo("Scrap Generated",
                    f"Placed {placed_count} pads on scrap #{scrap_num}.\n\n"
                    f"{remaining_count} pads remaining.\n"
                    f"Adjust dimensions and click Generate again.")

        except Exception as e:
            print(f"An error occurred during scrap mode SVG generation: {e}")
            messagebox.showerror("An Error Occurred", f"Something went wrong:\n\n{e}")

    def on_generate_gcode(self):
        """Generate G-code files."""
        # --- Scrap Mode ---
        if self.scrap_mode_var.get():
            self._generate_gcode_scrap_mode()
            return

        # --- Standard Mode ---
        try:
            params = self._prepare_generation()
            if not params:
                return

            pads, hole_dia, base = params['pads'], params['hole_dia'], params['base']
            width_mm, height_mm = params['width_mm'], params['height_mm']
            card_paper_dims = params['card_paper_dims']

            # Check if any supported materials selected (not exact_size)
            supported_materials = [m for m, var in self.material_vars.items() if var.get() and m != "exact_size"]
            if not supported_materials:
                messagebox.showwarning("No Materials Selected", "Please select at least one material (G-code not supported for Exact Size).")
                return

            # Validate all materials fit
            for material in supported_materials:
                mat_w, mat_h, mat_polygon = self._get_material_dimensions(material, width_mm, height_mm, card_paper_dims)
                if not can_all_pads_fit(pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon):
                    size_desc = "paper" if (material == "card" and card_paper_dims) else "sheet"
                    messagebox.showerror("Nesting Error", f"Could not fit all '{material.replace('_',' ')}' pieces on the specified {size_desc} size.")
                    return

            save_dir = filedialog.askdirectory(title="Select Folder to Save G-code", initialdir=self.settings.get("last_output_dir", ""))
            if not save_dir:
                return
            self.settings["last_output_dir"] = save_dir

            # Show working indicator
            working_popup = tk.Toplevel(self.root)
            working_popup.title("Working")
            working_popup.geometry("250x80")
            working_popup.configure(bg="#F0EAD6")
            working_popup.transient(self.root)
            working_popup.resizable(False, False)
            working_popup.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 125
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 40
            working_popup.geometry(f"+{x}+{y}")
            tk.Label(working_popup, text="Generating G-code...", bg="#F0EAD6", font=("Helvetica", 12)).pack(expand=True)
            working_popup.update()

            try:
                for material in supported_materials:
                    mat_w, mat_h, mat_polygon = self._get_material_dimensions(material, width_mm, height_mm, card_paper_dims)
                    filename = os.path.join(save_dir, f"{base}_{material}.gcode")
                    generate_gcode(pads, material, mat_w, mat_h, filename, hole_dia, self.settings, polygon=mat_polygon)
            finally:
                working_popup.destroy()

            save_settings(self.settings)
            messagebox.showinfo("Done", "G-code files generated successfully.")

        except Exception as e:
            print(f"An error occurred during G-code generation: {e}")
            messagebox.showerror("An Error Occurred", f"Something went wrong during G-code generation:\n\n{e}")

    def _generate_gcode_scrap_mode(self):
        """Handle G-code generation in scrap mode."""
        try:
            params = self._prepare_generation()
            if not params:
                return

            pads = params['pads']
            hole_dia = params['hole_dia']
            base = params['base']
            width_mm, height_mm = params['width_mm'], params['height_mm']
            card_paper_dims = params['card_paper_dims']

            # Check material selection - must be exactly one, and not exact_size
            selected_materials = [m for m, v in self.material_vars.items() if v.get() and m != "exact_size"]
            if len(selected_materials) != 1:
                messagebox.showerror("Scrap Mode Error",
                    "Please select exactly one material for scrap mode.\n"
                    "(G-code not supported for Exact Size)")
                return
            material = selected_materials[0]

            # Get scrap dimensions (polygon or rectangle)
            mat_w, mat_h, mat_polygon = self._get_material_dimensions(
                material, width_mm, height_mm, card_paper_dims)

            # Initialize or validate session
            if not self.scrap_session['active']:
                # Starting new session - ask for save directory
                save_dir = filedialog.askdirectory(
                    title="Select Folder to Save G-code",
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return
                self.settings["last_output_dir"] = save_dir

                self._start_scrap_session(pads, material, save_dir, hole_dia)
            else:
                # Continuing existing session - validate material matches
                if self.scrap_session['material'] != material:
                    messagebox.showerror("Material Mismatch",
                        f"Current session is for {self.scrap_session['material']}.\n"
                        f"Clear session to switch materials.")
                    return
                # Use remaining pads from session
                pads = self.scrap_session['remaining_pads']
                hole_dia = self.scrap_session['hole_dia']

            if not pads:
                messagebox.showinfo("Session Complete", "All pads have been placed!")
                return

            # Attempt partial placement
            placed, remaining, any_placed = try_nest_partial(
                pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon)

            if not any_placed:
                min_pad_size = min(p['size'] for p in pads)
                messagebox.showwarning("No Pads Fit",
                    f"No pads could be placed on this scrap.\n\n"
                    f"Smallest remaining pad: {min_pad_size}mm\n"
                    f"Try a larger scrap piece.")
                return

            # Generate filename with scrap number
            self.scrap_session['scrap_count'] += 1
            scrap_num = self.scrap_session['scrap_count']
            save_dir = self.scrap_session['save_dir']
            filename = os.path.join(save_dir, f"{base}_{material}_scrap{scrap_num}.gcode")

            # Generate G-code from placed discs
            generate_gcode_from_placed(placed, material, mat_w, mat_h, filename,
                                       hole_dia, self.settings, polygon=mat_polygon)

            # Update session with remaining pads
            self.scrap_session['remaining_pads'] = remaining
            self._update_scrap_status_display()
            self._update_remaining_pads_window()

            # Report results
            placed_count = len(placed)
            remaining_count = self._count_remaining_pads()

            if remaining_count == 0:
                save_settings(self.settings)
                messagebox.showinfo("Session Complete!",
                    f"Placed {placed_count} pads on scrap #{scrap_num}.\n\n"
                    f"All pads placed! Session complete.\n"
                    f"Files saved to: {save_dir}")
            else:
                messagebox.showinfo("Scrap Generated",
                    f"Placed {placed_count} pads on scrap #{scrap_num}.\n\n"
                    f"{remaining_count} pads remaining.\n"
                    f"Adjust dimensions and click Generate again.")

        except Exception as e:
            print(f"An error occurred during scrap mode G-code generation: {e}")
            messagebox.showerror("An Error Occurred", f"Something went wrong:\n\n{e}")

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

    def open_gcode_settings_window(self):
        GcodeSettingsWindow(self.root, self.settings, lambda s: save_settings(s))

    def open_resonance_window(self):
        ResonanceWindow(self.root, self.settings, lambda: save_settings(self.settings), self.apply_resonance_theme)

    def on_send_to_sd_card(self):
        """Send a G-code file to the SD card, clearing old files and ejecting."""
        # Step 1: Ask user to select SD card folder (remembers last used)
        last_sd = self.settings.get("sd_card_path", "")

        sd_folder = filedialog.askdirectory(
            title="Select Destination (SD Card)",
            initialdir=last_sd if last_sd and os.path.exists(last_sd) else None
        )
        if not sd_folder:
            return

        # Remember this SD card location
        self.settings["sd_card_path"] = sd_folder
        save_settings(self.settings)

        # Step 2: Confirm erasing existing files
        existing_files = glob.glob(os.path.join(sd_folder, "*.gcode"))
        existing_files += glob.glob(os.path.join(sd_folder, "*.nc"))
        existing_files += glob.glob(os.path.join(sd_folder, "*.ngc"))

        if existing_files:
            file_list = "\n".join(f"  • {os.path.basename(f)}" for f in existing_files)
            if not messagebox.askyesno(
                "Erase and Send",
                f"These files will be erased from the SD card:\n\n{file_list}\n\n"
                "Erase and send new file?"
            ):
                return
        else:
            if not messagebox.askyesno(
                "Send to SD Card",
                "Send a G-code file to this SD card?"
            ):
                return

        # Step 3: Ask user to select G-code file (remembers last output dir)
        filepath = filedialog.askopenfilename(
            title="Select G-code File to Send",
            filetypes=[("G-code files", "*.gcode *.nc *.ngc"), ("All files", "*.*")],
            initialdir=self.settings.get("last_output_dir", "")
        )
        if not filepath:
            return

        try:
            # Clear existing G-code files
            for old_file in existing_files:
                os.remove(old_file)

            # Copy the selected file to SD card
            filename = os.path.basename(filepath)
            dest_path = os.path.join(sd_folder, filename)
            shutil.copy2(filepath, dest_path)

            # Safely eject the SD card (Windows only)
            drive_letter = os.path.splitdrive(sd_folder)[0]
            eject_success = self._eject_drive(drive_letter) if drive_letter else False

            if eject_success:
                messagebox.showinfo(
                    "Ready!",
                    f"'{filename}' copied to SD card.\n\n"
                    f"SD card safely ejected - you can remove it now!"
                )
            else:
                messagebox.showinfo(
                    "File Copied",
                    f"'{filename}' copied to SD card.\n\n"
                    f"Please safely eject the SD card before removing."
                )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to send file to SD card:\n\n{e}")

    def _eject_drive(self, drive_letter):
        """Safely eject a drive on Windows. Returns True on success."""
        if sys.platform != 'win32':
            return False

        try:
            # Use PowerShell to eject the drive
            ps_script = f'''
$driveEject = New-Object -comObject Shell.Application
$driveEject.Namespace(17).ParseName("{drive_letter}").InvokeVerb("Eject")
'''
            subprocess.run(
                ["powershell", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=10
            )
            # Give it a moment to complete
            import time
            time.sleep(1)
            # Check if drive is still accessible (if not, eject worked)
            return not os.path.exists(drive_letter + "\\")
        except Exception:
            return False

    def update_ui_from_settings(self):
        self.unit_label.config(text=f"Width ({self.settings['units']}):")
        self.height_label.config(text=f"Height ({self.settings['units']}):")

if __name__ == '__main__':
    root = tk.Tk()
    app = PadSVGGeneratorApp(root)
    root.mainloop()
