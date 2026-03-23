import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import json
import sys
import subprocess
import time

# --- Local Imports ---
from config import (
    load_settings, save_settings, load_presets, save_presets,
    PAD_PRESET_FILE, KEY_PRESET_FILE, SCREW_SPECS_FILE,
    DEFAULT_SETTINGS,
    find_config_files_in_directory, import_config_files,
    get_ssl_context, get_input_devices
)
from svg_engine import generate_svg, can_all_pads_fit, check_for_oversized_engravings, try_nest_partial, generate_svg_from_placed, nest_pads
from gcode_engine import generate_gcode, generate_gcode_from_placed
from ui_dialogs import (
    OptionsWindow, LayerColorWindow, KeyLayoutWindow,
    ResonanceWindow, ConfirmationDialog,
    ImportPresetsWindow, ExportPresetsWindow, WebImportPresetsWindow, ImportTargetWindow,
    PolygonDrawWindow, GcodeSettingsWindow,
    UserGuideWindow, AboutDialog, PadNotesWindow, NestingPreviewWindow
)
from library_features import LibraryFeaturesMixin
from tooling_tab import ToolingTabMixin
from tuner_tab import TunerTabMixin
from toner_tab import TonerTabMixin

# ==========================================
# MAIN APP CLASS
# ==========================================

IS_MACOS = sys.platform == 'darwin'

class PadSVGGeneratorApp(LibraryFeaturesMixin, ToolingTabMixin, TunerTabMixin, TonerTabMixin):
    def __init__(self, root):
        self.root = root
        self.root.title("Stohrer Sax Shop Companion")
        self.root.geometry("640x720")

        # App icon (taskbar + title bar)
        try:
            import os, sys
            if getattr(sys, 'frozen', False):
                base = sys._MEIPASS
            else:
                base = os.path.dirname(__file__)
            ico = os.path.join(base, 'icon.ico')
            if os.path.exists(ico):
                self.root.iconbitmap(ico)
        except Exception:
            pass

        # On macOS, use native system colors (supports dark/light mode).
        # On Windows/Linux, use our custom cream theme.
        if IS_MACOS:
            self.default_bg = self.root.cget('bg')
        else:
            self.default_bg = "#FFFDD0"
            self.root.configure(bg=self.default_bg)

        self.settings = load_settings()

        # Ensure Toner is hidden unless explicitly unlocked
        if not self.settings.get("toner_unlocked"):
            visible = self.settings.get("visible_tabs", {})
            if visible.get("Toner"):
                visible["Toner"] = False
                save_settings(self.settings)

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

        # --- Tooling Tab Init ---
        self._init_tooling_state()

        # --- Tuner Tab Init ---
        self._init_tuner_state()

        # --- Toner Tab Init ---
        self._init_toner_state()

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

        # Save tooling tab settings
        self._update_tooling_settings()

        # Stop tuner/toner and save settings
        self._tuner_stop()
        self._tuner_save_settings()
        self._toner_stop()
        self._toner_save_settings()

        save_settings(self.settings)
        self.root.destroy()

    def _get_theme_color(self):
        """Get the current theme background color based on resonance clicks."""
        if IS_MACOS:
            return self.default_bg
        clicks = self.settings.get("resonance_clicks", 0)
        if 10 <= clicks < 50:
            return "#E0F7FA"  # COOL_BLUE
        elif 50 <= clicks < 100:
            return "#E8F5E9"  # COOL_GREEN
        return self.default_bg

    def apply_resonance_theme(self):
        if IS_MACOS:
            return
        color = self._get_theme_color()
        self.set_background_color(self.root, color)
        clicks = self.settings.get("resonance_clicks", 0)
        if clicks < 100:
            self.root.attributes('-alpha', 1.0)

    def set_background_color(self, parent, color):
        if IS_MACOS:
            return
        try:
            parent.configure(bg=color)
        except tk.TclError:
            pass

        style = ttk.Style()
        style.configure('App.TFrame', background=color)
        style.map('TNotebook.Tab', background=[('selected', color), ('!selected', color)], foreground=[('selected', 'black')])
        style.configure('TNotebook', background=color)

        for widget in parent.winfo_children():
            # Skip entire subtree for dark-themed widgets (e.g. strobe tuner)
            if getattr(widget, '_skip_theme', False):
                continue

            widget_class = widget.winfo_class()

            if widget_class in ('Frame', 'Label', 'Radiobutton', 'Checkbutton', 'LabelFrame', 'Canvas'):
                # Skip dark canvases (e.g. strobe tuner display)
                if widget_class == 'Canvas' and getattr(widget, '_dark_canvas', False):
                    continue
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

            if isinstance(widget, (tk.Frame, tk.LabelFrame, tk.Canvas, ttk.Frame, ttk.LabelFrame, ttk.Notebook)):
                self.set_background_color(widget, color)

    def create_menus(self):
        # --- Pad Generator Menu ---
        self.pad_menu = tk.Menu(self.root)
        
        pad_file_menu = tk.Menu(self.pad_menu, tearoff=0)
        self.pad_menu.add_cascade(label="File", menu=pad_file_menu)
        pad_file_menu.add_command(label="Import Pad Presets...", command=self.on_import_pad_presets)
        pad_file_menu.add_command(label="Export Pad Presets...", command=self.on_export_pad_presets)
        pad_file_menu.add_separator()
        pad_file_menu.add_command(label="Import Matt's Pad Sets", command=self.on_import_matts_pad_sets)
        pad_file_menu.add_separator()
        pad_file_menu.add_command(label="Import Settings from Folder...", command=self.on_import_settings_folder)
        pad_file_menu.add_separator()
        pad_file_menu.add_command(label="Feature Set...", command=self._open_feature_set)
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
        key_file_menu.add_command(label="Import Matt's Key Heights", command=self.on_import_matts_key_heights)
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
        screw_file_menu.add_command(label="Import Matt's Specs", command=self.on_import_matts_specs)
        screw_file_menu.add_command(label="Export Screw Specs...", command=self.on_export_screw_specs)
        screw_file_menu.add_separator()
        screw_file_menu.add_command(label="Exit", command=self.on_exit)

        # --- Serial Lookup Menu (was empty) ---
        self.serial_menu = tk.Menu(self.root)

        # --- Tooling Menu ---
        self.tooling_menu = tk.Menu(self.root)

        tooling_file_menu = tk.Menu(self.tooling_menu, tearoff=0)
        self.tooling_menu.add_cascade(label="File", menu=tooling_file_menu)
        tooling_file_menu.add_command(label="Exit", command=self.on_exit)

        tooling_options_menu = tk.Menu(self.tooling_menu, tearoff=0)
        self.tooling_menu.add_cascade(label="Options", menu=tooling_options_menu)
        tooling_options_menu.add_command(label="Settings...", command=self._open_tooling_gcode_settings)

        # --- Tuner Menu ---
        self.tuner_menu = tk.Menu(self.root)

        tuner_options_menu = tk.Menu(self.tuner_menu, tearoff=0)
        self.tuner_menu.add_cascade(label="Options", menu=tuner_options_menu)
        tuner_options_menu.add_command(label="Settings...", command=self._tuner_open_settings)
        tuner_options_menu.add_command(label="Input Device...", command=self._open_input_device_dialog)

        # --- Toner Menu ---
        self.toner_menu = tk.Menu(self.root)

        toner_file_menu = tk.Menu(self.toner_menu, tearoff=0)
        self.toner_menu.add_cascade(label="File", menu=toner_file_menu)
        toner_file_menu.add_command(label="Profiles...", command=self._toner_open_profile_dialog)
        toner_file_menu.add_separator()
        toner_file_menu.add_command(label="Export Profiles...", command=self._toner_export_profiles)
        toner_file_menu.add_command(label="Import Profiles...", command=self._toner_import_profiles)

        toner_options_menu = tk.Menu(self.toner_menu, tearoff=0)
        self.toner_menu.add_cascade(label="Options", menu=toner_options_menu)
        toner_options_menu.add_command(label="Input Device...", command=self._open_input_device_dialog)
        toner_options_menu.add_command(label="Capture Threshold...", command=self._open_capture_threshold)
        toner_options_menu.add_separator()
        toner_options_menu.add_command(label="Reference Pitch (A=)...", command=self._toner_open_pitch_dialog)
        toner_options_menu.add_command(label="Display Pitch...", command=self._toner_open_pitch_display_dialog)

        # --- Add Help menu to all tab menus ---
        for menu in (self.pad_menu, self.key_menu, self.screw_menu, self.serial_menu, self.tooling_menu, self.tuner_menu, self.toner_menu):
            help_menu = tk.Menu(menu, tearoff=0)
            menu.add_cascade(label="Help", menu=help_menu)
            help_menu.add_command(label="User Guide...", command=self.open_user_guide)
            help_menu.add_separator()
            help_menu.add_command(label="About", command=self.open_about)

    def on_tab_changed(self, event):
        selected = self.notebook.select()

        # Look up menu for the selected tab
        menu = self._tab_menus.get(selected)
        if menu:
            self.root.config(menu=menu)
            # Force menu bar refresh — works around tkinter bug where
            # the menu bar disappears in fullscreen/maximized mode
            self.root.update_idletasks()

        # Start/stop audio tabs
        if selected == str(self.tuner_tab_frame):
            self._toner_stop()
            self._tuner_start()
        elif selected == str(self.toner_tab_frame):
            self._tuner_stop()
            self._toner_start()
        else:
            self._tuner_stop()
            self._toner_stop()

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

        self.tooling_tab_frame = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.tooling_tab_frame, text='Tooling')
        self.create_tooling_tab(self.tooling_tab_frame)

        # --- Create Tab 6: Tuner ---
        self.tuner_tab_frame = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.tuner_tab_frame, text='Tuner')
        self.create_tuner_tab(self.tuner_tab_frame)

        # --- Create Tab 7: Toner ---
        self.toner_tab_frame = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.toner_tab_frame, text='Toner')
        self.create_toner_tab(self.toner_tab_frame)

        self.notebook.pack(expand=True, fill="both", padx=5, pady=5)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # Build tab → menu mapping (uses widget string IDs)
        self._tab_menus = {
            str(self.pad_tab): self.pad_menu,
            str(self.key_tab): self.key_menu,
            str(self.serial_tab): self.serial_menu,
            str(self.screw_tab): self.screw_menu,
            str(self.tooling_tab_frame): self.tooling_menu,
            str(self.tuner_tab_frame): self.tuner_menu,
            str(self.toner_tab_frame): self.toner_menu,
        }

        # Hide tabs based on Feature Set settings
        visible = self.settings.get("visible_tabs", {})
        tab_visibility = {
            self.key_tab: visible.get("Key Height Library", True),
            self.serial_tab: visible.get("Serial Lookup", True),
            self.screw_tab: visible.get("Screw Specs", True),
            self.tooling_tab_frame: visible.get("Tooling", True),
            self.tuner_tab_frame: visible.get("Tuner", True),
            self.toner_tab_frame: visible.get("Toner", True),
        }
        for tab_frame, is_visible in tab_visibility.items():
            if not is_visible:
                self.notebook.hide(tab_frame)

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
        self.pad_entry = tk.Text(parent, height=10, undo=True, maxundo=-1)
        self.pad_entry.pack(fill="x", padx=10)

        # Row 1: Library and preset dropdowns
        preset_select_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        preset_select_frame.pack(pady=(10, 2), fill='x', padx=10)

        tk.Label(preset_select_frame, text="Library:", bg=self.root.cget('bg')).pack(side="left", padx=(0, 2))
        self.pad_library_var = tk.StringVar()
        self.pad_library_dropdown = ttk.Combobox(preset_select_frame, textvariable=self.pad_library_var, state="readonly", width=15)
        self.pad_library_dropdown.pack(side="left")
        self.pad_library_dropdown.bind("<<ComboboxSelected>>", self.on_pad_library_selected)

        preset_names = []
        self.pad_preset_var = tk.StringVar()
        self.pad_preset_menu = ttk.Combobox(preset_select_frame, textvariable=self.pad_preset_var, values=preset_names, state="readonly", width=40)
        self.pad_preset_menu.set("Load Pad Preset")
        self.pad_preset_menu.pack(side="left", padx=5)
        self.pad_preset_menu.bind("<<ComboboxSelected>>", lambda e: self.on_load_pad_preset(self.pad_preset_var.get()))

        # Row 2: Save/Notes on left, Delete on right
        preset_btn_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        preset_btn_frame.pack(pady=(2, 10), fill='x', padx=10)

        left_btns = tk.Frame(preset_btn_frame, bg=self.root.cget('bg'))
        left_btns.pack(side="left")
        tk.Button(left_btns, text="Save as Preset", command=self.on_save_pad_preset).pack(side="left", padx=(0, 5))
        self.pad_notes_btn = tk.Button(left_btns, text="View Notes", command=self.on_pad_notes, state="disabled")
        self.pad_notes_btn.pack(side="left", padx=5)

        tk.Button(preset_btn_frame, text="Delete Preset", command=self.on_delete_pad_preset).pack(side="right")

        self.pad_preset_loaded_library = None
        self.pad_preset_loaded_name = None

        self.update_pad_library_dropdown()

        self.material_vars = {
            'felt': tk.BooleanVar(value=True),
            'card': tk.BooleanVar(value=True),
            'leather': tk.BooleanVar(value=True),
            'exact_size': tk.BooleanVar(value=False)
        }
        self.material_checkboxes = {}  # Store references for enable/disable

        options_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        options_frame.pack(pady=10, fill='x', padx=10)

        # Materials and Center Hole side by side
        mat_hole_row = tk.Frame(options_frame, bg=self.root.cget('bg'))
        mat_hole_row.pack(fill="x")

        mat_frame = tk.LabelFrame(mat_hole_row, text="Materials", bg=self.root.cget('bg'), padx=5, pady=5)
        mat_frame.pack(side="left", fill="y")
        material_list = list(self.material_vars.items())
        for i, (m, var) in enumerate(material_list):
            row, col = i // 2, i % 2
            cb = tk.Checkbutton(mat_frame, text=m.replace('_', ' ').capitalize(),
                               variable=var, bg=self.root.cget('bg'))
            cb.grid(row=row, column=col, sticky='w', padx=(0, 10))
            self.material_checkboxes[m] = cb

        hole_frame = tk.LabelFrame(mat_hole_row, text="Center Hole", bg=self.root.cget('bg'), padx=5, pady=5)
        hole_frame.pack(side="left", fill="both", expand=True, padx=(10, 0))
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
        sheet_frame.columnconfigure(2, weight=1)  # Scrap Mode column absorbs extra space

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
        scrap_inner_frame.grid(row=0, column=2, rowspan=4, sticky='n', padx=(20, 5))

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

        # Edge Bias d-pad (right side of sheet frame)
        bias_frame = tk.Frame(sheet_frame, bg=self.root.cget('bg'))
        bias_frame.grid(row=0, column=3, rowspan=4, sticky='n', padx=(15, 20))

        tk.Label(bias_frame, text="Edge Bias", font=("Helvetica", 8),
                 bg=self.root.cget('bg')).grid(row=0, column=0, columnspan=3)

        self.edge_bias_var = tk.StringVar(value=self.settings.get("edge_bias", "center"))
        self._edge_bias_buttons = {}
        _tmp_btn = tk.Button(bias_frame)
        self._edge_bias_default_bg = _tmp_btn.cget('bg')
        _tmp_btn.destroy()
        dpad_positions = {
            "nw": (1, 0), "n": (1, 1), "ne": (1, 2),
            "w":  (2, 0), "center": (2, 1), "e":  (2, 2),
            "sw": (3, 0), "s": (3, 1), "se": (3, 2),
        }
        dpad_labels = {
            "nw": "\u2196", "n": "\u2191", "ne": "\u2197",
            "w":  "\u2190", "center": "\u00b7", "e":  "\u2192",
            "sw": "\u2199", "s": "\u2193", "se": "\u2198",
        }
        for direction, (row, col) in dpad_positions.items():
            if direction == "center":
                # Center button toggles between "center" and "off"
                btn = tk.Button(bias_frame, text="ctr", width=2, height=1,
                               font=("Helvetica", 8), relief="raised",
                               command=self._toggle_center_bias)
            else:
                btn = tk.Button(bias_frame, text=dpad_labels[direction], width=2, height=1,
                               font=("Helvetica", 8), relief="raised",
                               command=lambda d=direction: self._set_edge_bias(d))
            btn.grid(row=row, column=col, padx=1, pady=1)
            self._edge_bias_buttons[direction] = btn
        self._update_edge_bias_display()

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
        generate_frame.pack(pady=(15, 5))
        tk.Button(generate_frame, text="Generate SVG", command=self.on_generate_svg, font=('Helvetica', 10, 'bold')).pack(side="left", padx=5)
        tk.Button(generate_frame, text="Generate G-code", command=self.on_generate_gcode, font=('Helvetica', 10, 'bold')).pack(side="left", padx=5)

        # Options below generate buttons
        options_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        options_frame.pack(pady=(0, 10))

        self.preview_var = tk.BooleanVar(value=self.settings.get("show_preview", False))
        tk.Checkbutton(options_frame, text="Preview before saving",
                       variable=self.preview_var, bg=self.root.cget('bg'),
                       command=lambda: self._save_checkbox("show_preview", self.preview_var)
                       ).pack(side="left", padx=(0, 15))

        self.eject_sd_var = tk.BooleanVar(value=self.settings.get("eject_sd_after_gcode", False))
        if sys.platform == 'win32':
            tk.Checkbutton(options_frame, text="Eject SD card after G-code export",
                           variable=self.eject_sd_var, bg=self.root.cget('bg'),
                           command=self._on_eject_sd_changed).pack(side="left")

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
                raw = [(x * 25.4, (grid_size - y) * 25.4) for (x, y) in polygon]
            else:
                grid_size = 40  # 40x40 cm
                # Convert cm to mm
                raw = [(x * 10, (grid_size - y) * 10) for (x, y) in polygon]
            # Normalize so bounding box starts at (0, 0)
            min_x = min(p[0] for p in raw)
            min_y = min(p[1] for p in raw)
            self.custom_polygon = [(x - min_x, y - min_y) for (x, y) in raw]
            self._update_shape_status()

    def on_unload_custom_shape(self):
        """Unload the custom shape and return to rectangle mode."""
        self.custom_polygon = None
        self._update_shape_status()

    def _show_scrap_continue_dialog(self, placed_count, scrap_num, remaining_count):
        """Show scrap continue dialog. If a polygon is loaded, offer to unload it."""
        msg = (f"Placed {placed_count} pads on scrap #{scrap_num}.\n\n"
               f"{remaining_count} pads remaining.\n"
               f"Adjust dimensions and click Generate again.")

        if not self.custom_polygon:
            messagebox.showinfo("Scrap Generated", msg)
            return

        # Custom dialog with shape options
        dlg = tk.Toplevel(self.root)
        dlg.title("Scrap Generated")
        dlg.configure(bg=self._get_theme_color())
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        bg = self._get_theme_color()
        tk.Label(dlg, text=msg, wraplength=380, bg=bg, justify="left",
                 font=("Helvetica", 10)).pack(padx=15, pady=(15, 10))
        tk.Label(dlg, text="A custom shape is loaded for the next scrap:",
                 bg=bg, font=("Helvetica", 9, "italic")).pack(padx=15)

        btn_frame = tk.Frame(dlg, bg=bg)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="Unload Shape", width=16,
                  command=lambda: self._scrap_dialog_close(dlg, unload=True)
                  ).pack(side="left", padx=8)
        tk.Button(btn_frame, text="Keep Shape", width=16,
                  command=lambda: self._scrap_dialog_close(dlg, unload=False)
                  ).pack(side="left", padx=8)

        dlg.protocol("WM_DELETE_WINDOW", lambda: self._scrap_dialog_close(dlg, unload=False))
        dlg.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dlg.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dlg.winfo_height() // 2)
        dlg.geometry(f"+{x}+{y}")
        dlg.wait_window()

    def _scrap_dialog_close(self, dlg, unload):
        """Handle scrap continue dialog button press."""
        if unload:
            self.custom_polygon = None
            self._update_shape_status()
        dlg.destroy()

    # --- Edge Bias Methods ---

    def _toggle_center_bias(self):
        """Handle center button clicks.

        If already on off or center: toggle between them.
        If on a directional bias: switch to whatever center was last
        (off or ctr), defaulting to off.
        """
        current = self.edge_bias_var.get()
        if current == "off":
            self._set_edge_bias("center")
        elif current == "center":
            self._set_edge_bias("off")
        else:
            # Coming from a direction — activate center at its last state
            last_center = getattr(self, '_last_center_mode', 'off')
            self._set_edge_bias(last_center)

    def _set_edge_bias(self, direction):
        """Set the edge bias direction and update button display."""
        # Remember center mode state so it persists when switching away
        if direction in ("off", "center"):
            self._last_center_mode = direction
        self.edge_bias_var.set(direction)
        self.settings["edge_bias"] = direction
        self._update_edge_bias_display()

    def _update_edge_bias_display(self):
        """Highlight the active edge bias button."""
        active = self.edge_bias_var.get()
        default_bg = self._edge_bias_default_bg

        # Center button label shows its internal state (off/ctr)
        center_btn = self._edge_bias_buttons.get("center")
        if center_btn:
            last_center = getattr(self, '_last_center_mode', 'off')
            center_btn.configure(text="off" if last_center == "off" else "ctr")

        for direction, btn in self._edge_bias_buttons.items():
            # "off" highlights the center button
            is_active = (direction == active) or (direction == "center" and active == "off")
            if is_active:
                if IS_MACOS:
                    btn.configure(relief="sunken")
                else:
                    btn.configure(relief="sunken", bg="#4a90d9", fg="white")
            else:
                if IS_MACOS:
                    btn.configure(relief="raised")
                else:
                    btn.configure(relief="raised", bg=default_bg, fg="black")

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
        theme_bg = self._get_theme_color()
        self.scrap_remaining_window.configure(bg=theme_bg)
        self.scrap_remaining_window.resizable(True, True)

        # Position to the right of main window
        self.root.update_idletasks()
        x = self.root.winfo_x() + self.root.winfo_width() + 10
        y = self.root.winfo_y()
        self.scrap_remaining_window.geometry(f"+{x}+{y}")

        # Header with progress
        header_frame = tk.Frame(self.scrap_remaining_window, bg=theme_bg)
        header_frame.pack(fill="x", padx=10, pady=(10, 5))
        self.scrap_window_header = tk.Label(header_frame, text="", bg=theme_bg,
                                            font=("Helvetica", 10, "bold"))
        self.scrap_window_header.pack(anchor="w")

        # Two-column layout for Remaining and Done
        columns_frame = tk.Frame(self.scrap_remaining_window, bg=theme_bg)
        columns_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Remaining column
        remaining_frame = tk.Frame(columns_frame, bg=theme_bg)
        remaining_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        tk.Label(remaining_frame, text="Remaining", bg=theme_bg,
                 font=("Helvetica", 9, "bold"), fg="blue").pack(anchor="w")
        self.scrap_remaining_listbox = tk.Listbox(remaining_frame, font=("Courier", 10), height=10, width=12)
        self.scrap_remaining_listbox.pack(fill="both", expand=True)

        # Done column
        done_frame = tk.Frame(columns_frame, bg=theme_bg)
        done_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))
        tk.Label(done_frame, text="Done", bg=theme_bg,
                 font=("Helvetica", 9, "bold"), fg="green").pack(anchor="w")
        self.scrap_done_listbox = tk.Listbox(done_frame, font=("Courier", 10), height=10, width=12)
        self.scrap_done_listbox.pack(fill="both", expand=True)

        # Footer with scrap count
        self.scrap_window_footer = tk.Label(self.scrap_remaining_window, text="", bg=theme_bg,
                                            font=("Helvetica", 9))
        self.scrap_window_footer.pack(pady=(0, 5))

        # Close button
        tk.Button(self.scrap_remaining_window, text="Close",
                  command=self._close_remaining_pads_window).pack(pady=(0, 10))

        # Handle window close
        self.scrap_remaining_window.protocol("WM_DELETE_WINDOW", self._on_remaining_window_close)

        self._update_remaining_pads_window()

    def _update_remaining_pads_window(self):
        """Update the contents of the remaining pads window with two columns."""
        if self.scrap_remaining_window is None:
            # Reopen window if session is active but window was closed
            if self.scrap_session.get('active', False):
                self._open_remaining_pads_window()
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
        """Handle user closing the remaining pads window via X button."""
        self._close_remaining_pads_window()

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

        try:
            width_val = float(self.width_entry.get())
            height_val = float(self.height_entry.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Sheet width and height must be valid numbers.")
            return None

        if width_val <= 0 or height_val <= 0:
            messagebox.showerror("Invalid Input", "Sheet width and height must be greater than zero.")
            return None

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

            selected_materials = [m for m, var in self.material_vars.items() if var.get()]
            if not selected_materials:
                messagebox.showwarning("No Materials Selected", "Please select at least one material.")
                return

            use_preview = self.preview_var.get()

            # Preview requires single material
            if use_preview and len(selected_materials) > 1:
                messagebox.showinfo("Preview",
                    "Preview works with one material at a time.\n"
                    "Please select a single material to preview its layout.")
                return

            save_dir = None

            # Process each material (with optional per-material preview)
            for material in selected_materials:
                mat_w, mat_h, mat_polygon = self._get_material_dimensions(material, width_mm, height_mm, card_paper_dims)

                # Nest (may re-run if user adjusts and retries)
                placed = nest_pads(pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon)

                # Validate fit
                if not can_all_pads_fit(pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon):
                    size_desc = "paper" if (material == "card" and card_paper_dims) else "sheet"
                    messagebox.showerror("Nesting Error", f"Could not fit all '{material.replace('_',' ')}' pieces on the specified {size_desc} size.")
                    return

                # Preview this material
                if use_preview:
                    preview = NestingPreviewWindow(
                        self.root, {material: placed}, mat_w, mat_h,
                        polygon=mat_polygon)
                    if preview.result != "save":
                        return  # User clicked Adjust — go back for this material

                # Ask for save directory once (on first material)
                if save_dir is None:
                    save_dir = filedialog.askdirectory(title="Select Folder to Save SVGs", initialdir=self.settings.get("last_output_dir", ""))
                    if not save_dir:
                        return
                    self.settings["last_output_dir"] = save_dir

                filename = os.path.join(save_dir, f"{base}_{material}.svg")
                generate_svg_from_placed(placed, material, mat_w, mat_h, filename, hole_dia, self.settings, polygon=mat_polygon)

            save_settings(self.settings)
            messagebox.showinfo("Done", "SVG files generated successfully.")

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

            # Preview before saving
            if self.preview_var.get():
                preview = NestingPreviewWindow(
                    self.root, {material: placed}, mat_w, mat_h,
                    polygon=mat_polygon)
                if preview.result != "save":
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
                self._show_scrap_continue_dialog(placed_count, scrap_num, remaining_count)

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

            use_preview = self.preview_var.get()

            if use_preview and len(supported_materials) > 1:
                messagebox.showinfo("Preview",
                    "Preview works with one material at a time.\n"
                    "Please select a single material to preview its layout.")
                return
            save_dir = None

            # Process each material with optional per-material preview
            all_placed = {}
            for material in supported_materials:
                mat_w, mat_h, mat_polygon = self._get_material_dimensions(material, width_mm, height_mm, card_paper_dims)

                placed = nest_pads(pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon)

                if not can_all_pads_fit(pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon):
                    size_desc = "paper" if (material == "card" and card_paper_dims) else "sheet"
                    messagebox.showerror("Nesting Error", f"Could not fit all '{material.replace('_',' ')}' pieces on the specified {size_desc} size.")
                    return

                if use_preview:
                    preview = NestingPreviewWindow(
                        self.root, {material: placed}, mat_w, mat_h,
                        polygon=mat_polygon)
                    if preview.result != "save":
                        return

                all_placed[material] = (placed, mat_w, mat_h, mat_polygon)

            if save_dir is None:
                save_dir = filedialog.askdirectory(title="Select Folder to Save G-code", initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return
                self.settings["last_output_dir"] = save_dir

            # Show working indicator
            working_popup = tk.Toplevel(self.root)
            working_popup.title("Working")
            working_popup.geometry("250x80")
            popup_bg = self._get_theme_color()
            working_popup.configure(bg=popup_bg)
            working_popup.transient(self.root)
            working_popup.resizable(False, False)
            working_popup.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 125
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 40
            working_popup.geometry(f"+{x}+{y}")
            tk.Label(working_popup, text="Generating G-code...", bg=popup_bg, font=("Helvetica", 12)).pack(expand=True)
            working_popup.update()

            try:
                for material, (placed, mat_w, mat_h, mat_polygon) in all_placed.items():
                    filename = os.path.join(save_dir, f"{base}_{material}.gcode")
                    generate_gcode_from_placed(placed, material, mat_w, mat_h, filename, hole_dia, self.settings, polygon=mat_polygon)
            finally:
                working_popup.destroy()

            save_settings(self.settings)

            # Auto-eject SD card if checkbox is checked and destination is removable
            if self.eject_sd_var.get() and self._is_removable_drive(save_dir):
                drive_letter = os.path.splitdrive(os.path.abspath(save_dir))[0]
                eject_success = self._eject_drive(drive_letter) if drive_letter else False
                if eject_success:
                    messagebox.showinfo("Done", "G-code files generated successfully.\n\nSD card safely ejected — you can remove it now!")
                else:
                    messagebox.showinfo("Done", "G-code files generated successfully.\n\nPlease safely eject the SD card before removing.")
            else:
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

            # Preview before saving
            if self.preview_var.get():
                preview = NestingPreviewWindow(
                    self.root, {material: placed}, mat_w, mat_h,
                    polygon=mat_polygon)
                if preview.result != "save":
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
                self._show_scrap_continue_dialog(placed_count, scrap_num, remaining_count)

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
                if size <= 0:
                    continue
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

        # Remember last used library
        self.settings["last_pad_library"] = lib_name

    def update_pad_library_dropdown(self):
        lib_names = ["All Libraries"] + sorted(self.pad_presets.keys())
        self.pad_library_dropdown['values'] = lib_names
        # Default to My Presets if it exists, otherwise All Libraries
        if "My Presets" in lib_names:
            self.pad_library_var.set("My Presets")
        else:
            self.pad_library_var.set("All Libraries")
        self.on_pad_library_selected()

    def _get_pad_preset_data(self, raw):
        """Extract pads text and notes from a preset entry (handles both old string and new dict formats)."""
        if isinstance(raw, dict):
            return raw.get("pads", ""), raw.get("notes", "")
        return raw, ""

    def on_save_pad_preset(self):
        active_library = self.pad_library_var.get()
        if not active_library or active_library == "All Libraries":
            # No library selected — ask for one or create "My Presets"
            lib_name = simpledialog.askstring("Library Name",
                "Enter a library name to save to:",
                initialvalue="My Presets")
            if not lib_name:
                return
            active_library = lib_name.strip()
            if not active_library:
                return
            if active_library not in self.pad_presets:
                self.pad_presets[active_library] = {}
            self.update_pad_library_dropdown()
            self.pad_library_var.set(active_library)
            self.on_pad_library_selected()

        name = simpledialog.askstring("Save Pad Preset", "Enter a name for this preset:")
        if name:
            text_data = self.pad_entry.get("1.0", tk.END)
            if not text_data.strip():
                messagebox.showwarning("Save Pad Preset", "Cannot save an empty list.")
                return

            if active_library not in self.pad_presets:
                self.pad_presets[active_library] = {}

            # Preserve existing notes if overwriting
            existing_notes = ""
            if name in self.pad_presets[active_library]:
                if not messagebox.askyesno("Overwrite", f"A set named '{name}' already exists in this library. Overwrite it?"):
                    return
                _, existing_notes = self._get_pad_preset_data(self.pad_presets[active_library][name])

            # Check for duplicate pad lists across all libraries
            new_lines = sorted(line.strip() for line in text_data.strip().splitlines() if line.strip())
            for lib, presets in self.pad_presets.items():
                for pname, pdata in presets.items():
                    if lib == active_library and pname == name:
                        continue  # Skip self when overwriting
                    existing_pads, _ = self._get_pad_preset_data(pdata)
                    existing_lines = sorted(line.strip() for line in existing_pads.strip().splitlines() if line.strip())
                    if new_lines == existing_lines:
                        if not messagebox.askyesno("Duplicate Detected",
                                f"This pad list is identical to '{pname}' "
                                f"in '{lib}'.\n\nSave anyway?"):
                            return

            self.pad_presets[active_library][name] = {"pads": text_data, "notes": existing_notes}

            if save_presets(self.pad_presets, PAD_PRESET_FILE):
                self.pad_preset_loaded_library = active_library
                self.pad_preset_loaded_name = name
                self.pad_notes_btn.config(state="normal")
                self.on_pad_library_selected()
                messagebox.showinfo("Preset Saved", f"Preset '{name}' saved successfully.")

    def on_load_pad_preset(self, selected_name):
        if not selected_name or selected_name == "Load Pad Preset":
            return

        lib_name = self.pad_library_var.get()
        preset_name = selected_name
        raw = None

        if lib_name == "All Libraries":
            try:
                lib_name, preset_name = selected_name.split("] ", 1)
                lib_name = lib_name[1:]
                if lib_name in self.pad_presets and preset_name in self.pad_presets[lib_name]:
                    raw = self.pad_presets[lib_name][preset_name]
            except ValueError:
                return
        else:
            if lib_name in self.pad_presets and preset_name in self.pad_presets[lib_name]:
                raw = self.pad_presets[lib_name][preset_name]

        if raw is not None:
            pads_text, _ = self._get_pad_preset_data(raw)
            self.pad_entry.delete("1.0", tk.END)
            self.pad_entry.insert(tk.END, pads_text)
            self.pad_preset_loaded_library = lib_name
            self.pad_preset_loaded_name = preset_name
            self.pad_notes_btn.config(state="normal")

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
                    self.pad_preset_loaded_library = None
                    self.pad_preset_loaded_name = None
                    self.pad_notes_btn.config(state="disabled")
                    messagebox.showinfo("Preset Deleted", f"Preset '{selected_preset}' deleted.")
            else:
                messagebox.showerror("Delete Error", "Could not find the preset to delete.")

    def on_pad_notes(self):
        lib = self.pad_preset_loaded_library
        name = self.pad_preset_loaded_name
        if not lib or not name or lib not in self.pad_presets or name not in self.pad_presets[lib]:
            messagebox.showwarning("Notes", "No preset loaded.")
            return

        _, current_notes = self._get_pad_preset_data(self.pad_presets[lib][name])
        dlg = PadNotesWindow(self.root, name, current_notes)
        if dlg.result is not None:
            # Update notes in the preset data
            pads_text, _ = self._get_pad_preset_data(self.pad_presets[lib][name])
            self.pad_presets[lib][name] = {"pads": pads_text, "notes": dlg.result}
            save_presets(self.pad_presets, PAD_PRESET_FILE)

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

    def on_import_matts_pad_sets(self):
        """Fetch pad set presets from stohrermusic.com and let user pick which to import."""
        import urllib.request
        import urllib.error

        MATTS_PADS_URL = "https://www.stohrermusic.com/data/pad_presets.json"

        try:
            req = urllib.request.Request(MATTS_PADS_URL, headers={"User-Agent": "StohrerSaxShopCompanion"})
            with urllib.request.urlopen(req, timeout=10, context=get_ssl_context()) as resp:
                web_data = json.loads(resp.read().decode("utf-8"))
        except (urllib.error.URLError, urllib.error.HTTPError, OSError) as e:
            messagebox.showerror("Connection Error",
                f"Could not fetch pad sets from stohrermusic.com:\n\n{e}")
            return
        except (json.JSONDecodeError, ValueError) as e:
            messagebox.showerror("Data Error",
                f"Invalid data received from server:\n\n{e}")
            return

        if not isinstance(web_data, dict) or not web_data:
            messagebox.showinfo("No Data", "No pad set data found on the server.")
            return

        WebImportPresetsWindow(self.root, web_data, self.pad_presets, PAD_PRESET_FILE, self)

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
        # Show only pad materials (felt/card/leather) from the pad generator tab
        pad_materials = [("felt", "Felt"), ("card", "Card"), ("leather", "Leather")]
        GcodeSettingsWindow(self.root, self.settings, lambda s: save_settings(s),
                            materials=pad_materials)

    def open_resonance_window(self):
        ResonanceWindow(self.root, self.settings, lambda: save_settings(self.settings), self.apply_resonance_theme)

    def _open_input_device_dialog(self):
        """Open a dialog to select the audio input device."""
        devices = get_input_devices()
        if not devices:
            messagebox.showinfo("No Devices", "No audio input devices found.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Input Device")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = self.root.cget('bg') if not IS_MACOS else "systemWindowBackgroundColor"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Select audio input device:", bg=bg,
                 font=("Helvetica", 10)).pack(pady=(0, 4))
        tk.Label(frame, text="Showing devices with 44.1 kHz+ sample rate.\n"
                 "Bluetooth headsets are excluded (sample rate too low).",
                 bg=bg, fg="#666666", font=("Helvetica", 8)).pack(pady=(0, 8))

        current_dev = self.settings.get("audio_input_device")
        dev_names = ["System Default"] + [name for _, name in devices]
        dev_indices = [None] + [idx for idx, _ in devices]

        mic_var = tk.StringVar(value="System Default")
        if current_dev is not None:
            for idx, name in devices:
                if idx == current_dev:
                    mic_var.set(name)
                    break

        listbox = tk.Listbox(frame, height=min(10, len(dev_names)),
                              width=45, font=("Helvetica", 10))
        listbox.pack(pady=(0, 10))
        for name in dev_names:
            listbox.insert(tk.END, name)

        # Select current device
        current_idx = 0
        if current_dev is not None:
            for i, (idx, _) in enumerate(devices):
                if idx == current_dev:
                    current_idx = i + 1
                    break
        listbox.selection_set(current_idx)
        listbox.see(current_idx)

        def apply():
            sel = listbox.curselection()
            if not sel:
                return
            dev_idx = dev_indices[sel[0]]
            self.settings["audio_input_device"] = dev_idx
            save_settings(self.settings)

            # Restart active audio engine with new device
            if hasattr(self, '_tuner_engine') and self._tuner_engine and self._tuner_engine.is_running:
                self._tuner_stop()
                self._tuner_start()
            if hasattr(self, '_toner_engine') and self._toner_engine and self._toner_engine.is_running:
                self._toner_stop()
                self._toner_start()

            dlg.destroy()
            dev_name = dev_names[sel[0]]
            messagebox.showinfo("Input Device",
                f"Audio input set to: {dev_name}")

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="Apply", command=apply).pack(
            side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(
            side="left")

    def _open_capture_threshold(self):
        """Open capture threshold dialog with live level meter."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Capture Threshold")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = self.root.cget('bg') if not IS_MACOS else "systemWindowBackgroundColor"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Set the minimum signal level to trigger capture.\n"
                 "Raise the threshold if chair noises or breathing\n"
                 "trigger false captures.",
                 bg=bg, font=("Helvetica", 9), justify="left").pack(pady=(0, 10))

        # Live level meter
        meter_frame = tk.Frame(frame, bg=bg)
        meter_frame.pack(fill="x", pady=(0, 10))

        tk.Label(meter_frame, text="Current level:", bg=bg,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 5))

        level_cv = tk.Canvas(meter_frame, bg="#333333", highlightthickness=1,
                              highlightbackground="#888888",
                              width=200, height=16)
        level_cv.pack(side="left")
        level_bar = level_cv.create_rectangle(0, 0, 0, 16, fill="#44AA44", outline="")
        level_text = level_cv.create_text(100, 8, text="", fill="white",
                                           font=("Helvetica", 7))

        # Threshold slider
        slider_frame = tk.Frame(frame, bg=bg)
        slider_frame.pack(fill="x", pady=(0, 10))

        tk.Label(slider_frame, text="Threshold:", bg=bg,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 5))

        threshold_var = tk.IntVar(value=self.settings.get("capture_threshold", 50))
        threshold_slider = tk.Scale(slider_frame, variable=threshold_var,
                                     from_=0, to=100, orient="horizontal",
                                     length=200, width=12, showvalue=True,
                                     bg=bg, highlightthickness=0)
        threshold_slider.pack(side="left")

        # Threshold line on the meter
        threshold_line = level_cv.create_line(0, 0, 0, 16, fill="#FF4444", width=2)

        # Animation loop for live level
        running = [True]

        def update_meter():
            if not running[0]:
                return
            # Read signal level from toner engine
            level = 0.0
            if hasattr(self, '_toner_engine') and self._toner_engine and self._toner_engine.is_running:
                result = self._toner_engine.analyze()
                level = result.signal_level

            # Update bar
            bar_w = int(level * 200)
            color = "#44AA44" if level < threshold_var.get() / 100.0 else "#FF8800"
            level_cv.coords(level_bar, 0, 0, bar_w, 16)
            level_cv.itemconfigure(level_bar, fill=color)
            level_cv.itemconfigure(level_text, text=f"{level:.0%}")

            # Update threshold line position
            thresh_x = int(threshold_var.get() / 100.0 * 200)
            level_cv.coords(threshold_line, thresh_x, 0, thresh_x, 16)

            dlg.after(50, update_meter)

        update_meter()

        def apply():
            running[0] = False
            self.settings["capture_threshold"] = threshold_var.get()
            if hasattr(self, '_toner_engine') and self._toner_engine:
                self._toner_engine.set_sensitivity(100 - threshold_var.get())
            save_settings(self.settings)
            dlg.destroy()

        def on_close():
            running[0] = False
            dlg.destroy()

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="Apply", command=apply).pack(
            side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Cancel", command=on_close).pack(side="left")

        dlg.protocol("WM_DELETE_WINDOW", on_close)

    def _open_feature_set(self):
        """Open the Feature Set dialog to show/hide tabs."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Feature Set")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = self.root.cget('bg') if not IS_MACOS else "systemWindowBackgroundColor"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Feature Set", bg=bg,
                 font=("Helvetica", 14, "bold")).pack(pady=(0, 5))
        tk.Label(frame, text="Choose which tabs to show.\n"
                 "The SVG Generator is always on.",
                 bg=bg, font=("Helvetica", 9)).pack(pady=(0, 10))

        visible = self.settings.get("visible_tabs", {})

        main_tabs = [
            "Key Height Library",
            "Serial Lookup",
            "Screw Specs",
            "Tooling",
        ]
        experimental_tabs = [
            "Tuner",
            "Toner",
        ]

        check_vars = {}
        for name in main_tabs:
            var = tk.BooleanVar(value=visible.get(name, True))
            tk.Checkbutton(frame, text=name, variable=var, bg=bg,
                           font=("Helvetica", 10)).pack(anchor="w")
            check_vars[name] = var

        tk.Label(frame, text="\nExperimental / In Progress", bg=bg,
                 font=("Helvetica", 11, "bold")).pack(anchor="w")
        for name in experimental_tabs:
            var = tk.BooleanVar(value=visible.get(name, True))
            tk.Checkbutton(frame, text=name, variable=var, bg=bg,
                           font=("Helvetica", 10)).pack(anchor="w")
            check_vars[name] = var

        def apply():
            new_visible = {name: var.get() for name, var in check_vars.items()}

            # Toner requires password if not previously unlocked
            if new_visible.get("Toner") and not self.settings.get("toner_unlocked"):
                pw = simpledialog.askstring("Toner Access",
                    "The Toner is in early development.\n"
                    "Enter the access code to enable it:",
                    show="*", parent=dlg)
                if pw != "iunderstand":
                    if pw is not None:  # None = cancelled
                        messagebox.showinfo("Access Denied",
                            "That's not the right code.", parent=dlg)
                    new_visible["Toner"] = False
                    check_vars["Toner"].set(False)
                    return
                self.settings["toner_unlocked"] = True

            self.settings["visible_tabs"] = new_visible
            save_settings(self.settings)
            dlg.destroy()
            messagebox.showinfo("Feature Set",
                "Changes will take effect next time you open the app.")

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x", pady=(10, 0))
        tk.Button(btn_frame, text="Apply", command=apply).pack(
            side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(
            side="left")

    def open_user_guide(self):
        selected = self.notebook.select()
        tab_sections = {
            str(self.pad_tab): "pad_generator",
            str(self.key_tab): "key_heights",
            str(self.serial_tab): "serial_lookup",
            str(self.screw_tab): "screw_specs",
            str(self.tooling_tab_frame): "tooling",
            str(self.tuner_tab_frame): "tuner",
            str(self.toner_tab_frame): "toner",
        }
        section = tab_sections.get(selected, None)
        UserGuideWindow(self.root, section=section)

    def open_about(self):
        AboutDialog(self.root)

    def _save_checkbox(self, key, var):
        """Save a checkbox setting."""
        self.settings[key] = var.get()
        save_settings(self.settings)

    def _on_eject_sd_changed(self):
        """Persist the eject SD card checkbox state."""
        self.settings["eject_sd_after_gcode"] = self.eject_sd_var.get()
        save_settings(self.settings)

    def _is_removable_drive(self, path):
        """Check if path is on a removable drive (USB/SD). Windows only."""
        if sys.platform != 'win32':
            return False
        try:
            import ctypes
            drive = os.path.splitdrive(os.path.abspath(path))[0]
            if not drive:
                return False
            # GetDriveTypeW: 2 = DRIVE_REMOVABLE
            return ctypes.windll.kernel32.GetDriveTypeW(drive + "\\") == 2
        except Exception:
            return False

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
