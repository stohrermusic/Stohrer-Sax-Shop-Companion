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
    PAD_PRESET_FILE, KEY_PRESET_FILE, SCREW_SPECS_FILE, SIZING_PRESET_FILE,
    GCODE_PRESET_FILE, GCODE_PRESET_MATERIALS,
    find_config_files_in_directory, import_config_files,
    get_ssl_context, get_input_devices,
    setup_logging, get_log_file,
    settings_to_sizing_preset, settings_to_gcode_presets,
)

# Initialize translations BEFORE importing UI modules. Any module-level
# strings in ui_dialogs/main.py/toner_tab/etc. that use _() need the
# translation catalog already loaded when their module body runs.
from i18n import init_translation
init_translation(load_settings().get("language", "en"))

from svg_engine import can_all_pads_fit, check_for_oversized_engravings, try_nest_partial, generate_svg_from_placed, nest_pads  # noqa: E402
from gcode_engine import generate_gcode_from_placed  # noqa: E402
from ui_dialogs import (  # noqa: E402
    OptionsWindow, LayerColorWindow, KeyLayoutWindow,
    ResonanceWindow, ConfirmationDialog,
    ImportPresetsWindow, ExportPresetsWindow, WebImportPresetsWindow, ImportTargetWindow,
    PolygonDrawWindow, GcodeSettingsWindow,
    UserGuideWindow, AboutDialog, PadNotesWindow, NestingPreviewWindow
)
from library_features import LibraryFeaturesMixin  # noqa: E402
from tooling_tab import ToolingTabMixin  # noqa: E402
from tuner_tab import TunerTabMixin  # noqa: E402
from toner_tab import TonerTabMixin  # noqa: E402

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
            import os
            import sys
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

        # Apply tooltip on/off setting at startup so settings dialogs
        # opened later honor the user's choice from the Feature Set menu.
        from ui_dialogs import set_tooltips_enabled
        set_tooltips_enabled(self.settings.get("tooltips_enabled", True))

        # Ensure Toner is hidden unless explicitly unlocked
        if not self.settings.get("toner_unlocked"):
            visible = self.settings.get("visible_tabs", {})
            if visible.get("Toner"):
                visible["Toner"] = False
                save_settings(self.settings)

        self.pad_presets = load_presets(PAD_PRESET_FILE, preset_type_name="Pad Preset")
        self.key_presets = load_presets(KEY_PRESET_FILE, preset_type_name="Key Height")
        self.sizing_presets = load_presets(SIZING_PRESET_FILE, preset_type_name="Sizing Preset")

        # The app requires at least one sizing preset; if the library is
        # empty (first run, or a wiped config) bootstrap a "Default" preset
        # from the current settings so users always have something to load.
        if not self.sizing_presets:
            self.sizing_presets["Default"] = settings_to_sizing_preset(self.settings)
            save_presets(self.sizing_presets, SIZING_PRESET_FILE)

        # G-code presets: shape is {material: {preset_name: data}}. Bootstrap
        # a Default per material on first run so every material always has
        # at least one preset to load. Backfill any new materials added later.
        self.gcode_presets = load_presets(GCODE_PRESET_FILE, preset_type_name="G-code Preset")
        bootstrap = settings_to_gcode_presets(self.settings)
        gcode_presets_dirty = False
        for mat in GCODE_PRESET_MATERIALS:
            if mat not in self.gcode_presets or not self.gcode_presets[mat]:
                self.gcode_presets[mat] = bootstrap.get(mat, {})
                gcode_presets_dirty = True
        if gcode_presets_dirty:
            save_presets(self.gcode_presets, GCODE_PRESET_FILE)

        # --- Custom polygon state ---
        # custom_polygon holds the active scrap outline used for
        # nesting + framing. ALWAYS stored in SVG-Y-DOWN convention
        # (origin (0, 0) at top-left of polygon bbox). Both drawn and
        # camera-captured polygons go through the same Y-flip-then-
        # normalize path (see _adopt_camera_polygon and the drawn
        # path in on_draw_custom_shape) so downstream code never has
        # to care which source produced it.
        #
        # ---- Y-axis conventions across the codebase ----
        # The flip sites below cross EXACTLY one boundary each.
        # Adding a new flip without crossing a boundary is a bug.
        #
        #   Layer                                  | Y direction
        #   ---------------------------------------|------------
        #   Tk canvas (PolygonDrawWindow display)  | DOWN (Tk native)
        #   Grid storage (PolygonDrawWindow.points)| UP   (graph-paper)
        #   OpenCV pixel coords                    | DOWN (image native)
        #   OpenCV ChArUco board frame             | DOWN (matches image)
        #   Machine coords (Grbl/Falcon)           | UP   (Y=0 at front)
        #   custom_polygon (this attribute)        | DOWN (SVG-style)
        #   SVG output                             | DOWN (SVG native)
        #   G-code emission                        | UP   (matches Grbl)
        #
        # Boundary flip sites (one per boundary, no redundancy):
        #   board (DOWN) → machine (UP)
        #     camera_capture._board_corner_to_machine_mm
        #   legacy schema-1 pixels_to_mm output → corrected machine
        #     camera_capture.pixels_to_mm (compat shim)
        #   grid (UP) ↔ canvas (DOWN)
        #     ui_dialogs.PolygonDrawWindow._{grid,canvas}_to_{canvas,grid}
        #   grid storage (UP) → custom_polygon (DOWN)
        #     on_draw_custom_shape (drawn path)
        #   machine-mm (UP) → custom_polygon (DOWN)
        #     _adopt_camera_polygon (camera path) — only the SHAPE is
        #     kept; the machine-coord offset isn't tracked since the
        #     auto-frame feature was removed. The polygon gets nested
        #     in local coords and cut at the user's manually-jogged
        #     work origin via G92.
        #   custom_polygon (DOWN) → G-code (UP)
        #     gcode_engine.generate_polygon_framing_gcode
        #     gcode_engine.generate_gcode_from_placed (and die variant)
        #   calibration-card strokes (DOWN) → G-code (UP)
        #     gcode_engine.generate_calibration_card_gcode
        #   G-code (UP) → SVG (DOWN) — arc text only
        #     svg_engine die/holder renderers
        self.custom_polygon = None
        # The polygon's FULL OUTLINE (un-insetted), normalized to share
        # the same coordinate origin as custom_polygon. Used by Frame &
        # Cut to trace the actual scrap edge for visual placement
        # verification, while custom_polygon (the insetted shape) still
        # governs pad nesting placement. None when no inset was applied
        # (plain hand-drawn polygons) — Frame & Cut falls back to
        # custom_polygon in that case.
        self.custom_polygon_outline = None
        # The polygon's leftmost-lowest vertex in absolute machine coords
        # (Y-up), populated whenever the polygon has a machine reference
        # (camera-captured or hand-traced over the live camera overlay).
        # Drives Frame & Cut's "Try Auto Locate" button. None for plain
        # hand-drawn polygons (no machine reference exists, so the button
        # hides).
        self._custom_polygon_lb_machine = None

        # --- Falcon (direct serial) state ---
        # Detected on startup; the "Frame & Cut" button only appears when
        # a Grbl controller answered the handshake on some serial port.
        self.falcon_port = None
        self._falcon_detect_attempted = False
        # Tracks whether the laser has been homed in this SSC session.
        # Frame & Cut and the calibration dialog both check this to
        # decide whether to auto-home or skip. Reset on Machine > Reset
        # (soft-reset loses position).
        self._falcon_homed_this_session = False

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
        # WM_DELETE_WINDOW only covers the window close button. On macOS,
        # Cmd-Q and the app menu's Quit go through Tk's ::tk::mac::Quit
        # handler, whose default exits the process without running on_exit
        # — settings would silently never save on the standard mac quit
        # path. Route it through the same close handler.
        if IS_MACOS:
            self.root.createcommand("::tk::mac::Quit", self.on_exit)

    def on_exit(self):
        # Warn if scrap session is active with remaining pads
        if self.scrap_session.get('active', False):
            remaining = self._count_remaining_pads()
            if remaining > 0:
                if not messagebox.askyesno(_("Scrap Session Active"),
                    _("You have {remaining} pads remaining in your scrap session.\n\n"
                    "Exit anyway? (Session will be lost)").format(remaining=remaining)):
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
        self.pad_menu.add_cascade(label=_("File"), menu=pad_file_menu)
        pad_file_menu.add_command(label=_("Import Pad Presets..."), command=self.on_import_pad_presets)
        pad_file_menu.add_command(label=_("Export Pad Presets..."), command=self.on_export_pad_presets)
        pad_file_menu.add_separator()
        pad_file_menu.add_command(label=_("Import Matt's Pad Sets"), command=self.on_import_matts_pad_sets)
        pad_file_menu.add_separator()
        pad_file_menu.add_command(label=_("Import Settings from Folder..."), command=self.on_import_settings_folder)
        pad_file_menu.add_separator()
        pad_file_menu.add_command(label=_("Feature Set..."), command=self._open_feature_set)
        pad_file_menu.add_separator()
        pad_file_menu.add_command(label=_("Exit"), command=self.on_exit)

        pad_options_menu = tk.Menu(self.pad_menu, tearoff=0)
        self.pad_menu.add_cascade(label=_("Options"), menu=pad_options_menu)
        pad_options_menu.add_command(label=_("Sizing Rules..."), command=self.open_options_window)
        pad_options_menu.add_command(label=_("Layer Colors..."), command=self.open_color_window)
        pad_options_menu.add_separator()
        pad_options_menu.add_command(label=_("G-code Settings..."), command=self.open_gcode_settings_window)
        pad_options_menu.add_separator()

        # Machine submenu — direct Falcon serial control + camera
        # setup. Gated behind the Feature Set experimental toggle;
        # off by default. Within the cascade, Camera Calibration is
        # the only item enabled until calibration has been done
        # (after which the rest un-greys via _refresh_machine_ui_state).
        # Camera Calibration is the "initiation" entry point for
        # camera-based scrap capture and other Falcon features.
        self._machine_menu_indices = {}
        if self._machine_enabled():
            machine_menu = tk.Menu(pad_options_menu, tearoff=0)
            pad_options_menu.add_cascade(label=_("Machine"), menu=machine_menu)
            self._machine_menu = machine_menu
            # Order: Camera Calibration first (always enabled when
            # toggle is on — it's the gateway). Then the Falcon-direct
            # items below, which require a working calibration.
            machine_menu.add_command(label=_("Camera Calibration..."),
                                       command=self._on_machine_recalibrate)
            self._machine_menu_indices['camera_cal'] = (
                machine_menu.index('end'))
            machine_menu.add_command(
                label=_("Camera-Polygon Inset Margin..."),
                command=self._on_machine_inset_settings)
            self._machine_menu_indices['inset'] = (
                machine_menu.index('end'))
            machine_menu.add_separator()
            machine_menu.add_command(label=_("Home Laser"),
                                       command=self._on_machine_home_falcon)
            self._machine_menu_indices['home'] = (
                machine_menu.index('end'))
            machine_menu.add_command(label=_("Test Connection"),
                                       command=self._on_machine_test_connection)
            self._machine_menu_indices['test'] = (
                machine_menu.index('end'))
            machine_menu.add_command(label=_("Clear Errors ($X)"),
                                       command=self._on_machine_clear_errors)
            self._machine_menu_indices['clear'] = (
                machine_menu.index('end'))
            machine_menu.add_command(label=_("Reset Falcon (soft-reset)"),
                                       command=self._on_machine_reset_falcon)
            self._machine_menu_indices['reset'] = (
                machine_menu.index('end'))

        # --- Key Height Library Menu ---
        self.key_menu = tk.Menu(self.root)

        key_file_menu = tk.Menu(self.key_menu, tearoff=0)
        self.key_menu.add_cascade(label=_("File"), menu=key_file_menu)
        key_file_menu.add_command(label=_("Import Key Sets..."), command=self.on_import_key_sets)
        key_file_menu.add_command(label=_("Import Matt's Key Heights"), command=self.on_import_matts_key_heights)
        key_file_menu.add_command(label=_("Export Key Sets..."), command=self.on_export_key_sets)
        key_file_menu.add_separator()
        key_file_menu.add_command(label=_("Exit"), command=self.on_exit)

        key_options_menu = tk.Menu(self.key_menu, tearoff=0)
        self.key_menu.add_cascade(label=_("Options"), menu=key_options_menu)
        key_options_menu.add_command(label=_("Layout Options..."), command=self.open_key_layout_window)

        # --- Screw Specs Menu ---
        self.screw_menu = tk.Menu(self.root)

        screw_file_menu = tk.Menu(self.screw_menu, tearoff=0)
        self.screw_menu.add_cascade(label=_("File"), menu=screw_file_menu)
        screw_file_menu.add_command(label=_("Import Screw Specs..."), command=self.on_import_screw_specs)
        screw_file_menu.add_command(label=_("Import Matt's Specs"), command=self.on_import_matts_specs)
        screw_file_menu.add_command(label=_("Export Screw Specs..."), command=self.on_export_screw_specs)
        screw_file_menu.add_separator()
        screw_file_menu.add_command(label=_("Exit"), command=self.on_exit)

        # --- Serial Lookup Menu (was empty) ---
        self.serial_menu = tk.Menu(self.root)

        # --- Tooling Menu ---
        self.tooling_menu = tk.Menu(self.root)

        tooling_file_menu = tk.Menu(self.tooling_menu, tearoff=0)
        self.tooling_menu.add_cascade(label=_("File"), menu=tooling_file_menu)
        tooling_file_menu.add_command(label=_("Exit"), command=self.on_exit)

        tooling_options_menu = tk.Menu(self.tooling_menu, tearoff=0)
        self.tooling_menu.add_cascade(label=_("Options"), menu=tooling_options_menu)
        tooling_options_menu.add_command(label=_("Settings..."), command=self._open_tooling_gcode_settings)
        # Camera Calibration lives in Pad Maker > Options > Machine —
        # it's a calibration workflow that produces a card as a
        # byproduct of running it, not a standalone tool, so it doesn't
        # belong in the tools selector.

        # --- Tuner Menu ---
        self.tuner_menu = tk.Menu(self.root)

        tuner_options_menu = tk.Menu(self.tuner_menu, tearoff=0)
        self.tuner_menu.add_cascade(label=_("Options"), menu=tuner_options_menu)
        tuner_options_menu.add_command(label=_("Settings..."), command=self._tuner_open_settings)
        tuner_options_menu.add_command(label=_("Input Device..."), command=self._open_input_device_dialog)

        # --- Toner Menu ---
        self.toner_menu = tk.Menu(self.root)

        toner_file_menu = tk.Menu(self.toner_menu, tearoff=0)
        self.toner_menu.add_cascade(label=_("File"), menu=toner_file_menu)
        toner_file_menu.add_command(label=_("Presets..."), command=self._toner_open_preset_dialog)
        toner_file_menu.add_command(label=_("Analyze..."), command=self._toner_open_analyze_dialog)
        toner_file_menu.add_separator()
        toner_transfer_menu = tk.Menu(toner_file_menu, tearoff=0)
        toner_file_menu.add_cascade(label=_("Transfer Data"), menu=toner_transfer_menu)
        toner_transfer_menu.add_command(label=_("Export Preset Library..."), command=self._toner_export_presets)
        toner_transfer_menu.add_command(label=_("Import Preset Library..."), command=self._toner_import_presets)

        toner_options_menu = tk.Menu(self.toner_menu, tearoff=0)
        self.toner_menu.add_cascade(label=_("Options"), menu=toner_options_menu)
        toner_options_menu.add_command(label=_("Settings..."), command=self._toner_open_settings)
        toner_options_menu.add_command(label=_("Capture Threshold..."), command=self._open_capture_threshold)

        # --- Add Help menu to all tab menus ---
        for menu in (self.pad_menu, self.key_menu, self.screw_menu, self.serial_menu, self.tooling_menu, self.tuner_menu, self.toner_menu):
            help_menu = tk.Menu(menu, tearoff=0)
            menu.add_cascade(label=_("Help"), menu=help_menu)
            help_menu.add_command(label=_("User Guide..."), command=self.open_user_guide)
            help_menu.add_separator()
            help_menu.add_command(label=_("Open Log File"), command=self._open_log_file)
            help_menu.add_separator()
            help_menu.add_command(label=_("About"), command=self.open_about)

    def on_tab_changed(self, event):
        selected = self.notebook.select()

        # Look up menu for the selected tab
        menu = self._tab_menus.get(selected)
        if menu:
            self.root.config(menu=menu)
            # Force menu bar refresh — works around tkinter/X11 bug where
            # the menu bar disappears in maximized mode on Linux.
            # A geometry toggle forces the window manager to redraw.
            self.root.update_idletasks()
            if sys.platform == 'linux':
                geo = self.root.geometry()
                self.root.geometry(geo)

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
        
        # --- Create Tab 1: Pad Maker ---
        self.pad_tab = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.pad_tab, text=_('Pad Maker'))
        self.create_pad_generator_tab(self.pad_tab)

        # --- Create Tab 2: Key Height Library ---
        self.key_tab = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.key_tab, text=_('Key Height Library'))
        self.create_key_library_tab(self.key_tab)

        # --- Create Tab 3: Serial Lookup ---
        self.serial_tab = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.serial_tab, text=_('Serial Lookup'))
        self.create_serial_lookup_tab(self.serial_tab)

        # --- Create Tab 4: Screw Specs ---
        self.screw_tab = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.screw_tab, text=_('Screw Specs'))
        self.create_screw_specs_tab(self.screw_tab)

        self.tooling_tab_frame = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.tooling_tab_frame, text=_('Tooling'))
        self.create_tooling_tab(self.tooling_tab_frame)

        # --- Create Tab 6: Tuner ---
        self.tuner_tab_frame = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.tuner_tab_frame, text=_('Tuner'))
        self.create_tuner_tab(self.tuner_tab_frame)

        # --- Create Tab 7: Toner ---
        self.toner_tab_frame = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.toner_tab_frame, text=_('Toner (beta)'))
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
        # Scrollable wrapper so the tab scrolls when content exceeds window
        self._pad_tab_canvas = tk.Canvas(parent, highlightthickness=0,
                                         bg=self.root.cget('bg'))
        self._pad_tab_scrollbar = tk.Scrollbar(parent, orient="vertical",
                                                command=self._pad_tab_canvas.yview)
        self._pad_tab_inner = tk.Frame(self._pad_tab_canvas, bg=self.root.cget('bg'))

        self._pad_tab_inner.bind("<Configure>",
            lambda e: self._pad_tab_canvas.configure(
                scrollregion=self._pad_tab_canvas.bbox("all")))
        self._pad_tab_canvas_window = self._pad_tab_canvas.create_window(
            (0, 0), window=self._pad_tab_inner, anchor="nw")
        self._pad_tab_canvas.configure(yscrollcommand=self._pad_tab_scrollbar.set)

        # Keep inner frame width in sync with canvas
        def _sync_inner_width(event):
            self._pad_tab_canvas.itemconfig(self._pad_tab_canvas_window,
                                            width=event.width)
        self._pad_tab_canvas.bind("<Configure>", _sync_inner_width)

        self._pad_tab_canvas.pack(side="left", fill="both", expand=True)
        self._pad_tab_scrollbar.pack(side="right", fill="y")

        # Enable mousewheel scrolling on the canvas and all children
        from ui_dialogs import bind_mousewheel
        bind_mousewheel(self._pad_tab_canvas, self._pad_tab_canvas)
        bind_mousewheel(self._pad_tab_inner, self._pad_tab_canvas)

        # Redirect mousewheel from child widgets to the outer canvas
        def _bind_children_mousewheel(frame):
            for child in frame.winfo_children():
                bind_mousewheel(child, self._pad_tab_canvas)
                if isinstance(child, (tk.Frame, tk.LabelFrame, ttk.Frame, ttk.LabelFrame)):
                    _bind_children_mousewheel(child)
        # Defer until all children exist
        self._pad_tab_bind_children_mousewheel = _bind_children_mousewheel

        # All content goes into _pad_tab_inner instead of parent
        parent = self._pad_tab_inner

        tk.Label(parent, text=_("Enter pad sizes (e.g. 42.0x3):"), bg=self.root.cget('bg')).pack(pady=5)
        self.pad_entry = tk.Text(parent, height=10, undo=True, maxundo=-1)
        self.pad_entry.pack(fill="x", padx=10)

        # Auto-resize text widget to fit content
        def _auto_resize_pad_entry(event=None):
            line_count = int(self.pad_entry.index("end-1c").split(".")[0])
            new_height = max(10, min(line_count + 1, 30))
            self.pad_entry.configure(height=new_height)
        self.pad_entry.bind("<KeyRelease>", _auto_resize_pad_entry)
        self.pad_entry.bind("<<Paste>>", lambda e: self.pad_entry.after(10, _auto_resize_pad_entry))
        self._auto_resize_pad_entry = _auto_resize_pad_entry

        # Row 1: Library and preset dropdowns
        preset_select_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        preset_select_frame.pack(pady=(10, 2), fill='x', padx=10)

        tk.Label(preset_select_frame, text=_("Library:"), bg=self.root.cget('bg')).pack(side="left", padx=(0, 2))
        self.pad_library_var = tk.StringVar()
        self.pad_library_dropdown = ttk.Combobox(preset_select_frame, textvariable=self.pad_library_var, state="readonly", width=15)
        self.pad_library_dropdown.pack(side="left")
        self.pad_library_dropdown.bind("<<ComboboxSelected>>", self.on_pad_library_selected)

        preset_names = []
        self.pad_preset_var = tk.StringVar()
        self.pad_preset_menu = ttk.Combobox(preset_select_frame, textvariable=self.pad_preset_var, values=preset_names, state="readonly", width=40)
        self.pad_preset_menu.set(_("Load Pad Preset"))
        self.pad_preset_menu.pack(side="left", padx=5)
        self.pad_preset_menu.bind("<<ComboboxSelected>>", lambda e: self.on_load_pad_preset(self.pad_preset_var.get()))

        # Row 2: Save/Notes on left, Delete on right
        preset_btn_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        preset_btn_frame.pack(pady=(2, 10), fill='x', padx=10)

        left_btns = tk.Frame(preset_btn_frame, bg=self.root.cget('bg'))
        left_btns.pack(side="left")
        tk.Button(left_btns, text=_("Save as Preset"), command=self.on_save_pad_preset).pack(side="left", padx=(0, 5))
        self.pad_notes_btn = tk.Button(left_btns, text=_("View Notes"), command=self.on_pad_notes, state="disabled")
        self.pad_notes_btn.pack(side="left", padx=5)

        self.pad_delete_btn = tk.Button(preset_btn_frame, text=_("Delete Preset"), command=self.on_delete_pad_preset)
        self.pad_delete_btn.pack(side="right")

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

        mat_frame = tk.LabelFrame(mat_hole_row, text=_("Materials"), bg=self.root.cget('bg'), padx=5, pady=5)
        mat_frame.pack(side="left", fill="y")
        material_list = list(self.material_vars.items())
        for i, (m, var) in enumerate(material_list):
            row, col = i // 2, i % 2
            cb = tk.Checkbutton(mat_frame, text=_(m.replace('_', ' ').capitalize()),
                               variable=var, bg=self.root.cget('bg'))
            cb.grid(row=row, column=col, sticky='w', padx=(0, 10))
            self.material_checkboxes[m] = cb

        hole_frame = tk.LabelFrame(mat_hole_row, text=_("Center Hole"), bg=self.root.cget('bg'), padx=5, pady=5)
        hole_frame.pack(side="left", fill="both", expand=True, padx=(10, 0))
        self.hole_var = tk.StringVar(value=self.settings["hole_option"])

        tk.Radiobutton(hole_frame, text=_("None"), variable=self.hole_var, value="No center holes", bg=self.root.cget('bg'), command=self.toggle_custom_hole_entry).pack(side="left")
        tk.Radiobutton(hole_frame, text="3.0mm", variable=self.hole_var, value="3.0mm", bg=self.root.cget('bg'), command=self.toggle_custom_hole_entry).pack(side="left")
        tk.Radiobutton(hole_frame, text="3.5mm", variable=self.hole_var, value="3.5mm", bg=self.root.cget('bg'), command=self.toggle_custom_hole_entry).pack(side="left")
        tk.Radiobutton(hole_frame, text=_("Custom:"), variable=self.hole_var, value="Custom", bg=self.root.cget('bg'), command=self.toggle_custom_hole_entry).pack(side="left")

        self.custom_hole_entry = tk.Entry(hole_frame, width=6)
        self.custom_hole_entry.insert(0, self.settings.get("custom_hole_size", "4.0"))
        self.custom_hole_entry.pack(side="left", padx=2)
        tk.Label(hole_frame, text="mm", bg=self.root.cget('bg')).pack(side="left")
        self.toggle_custom_hole_entry()

        sheet_frame = tk.LabelFrame(options_frame, text=_("Sheet Size"), bg=self.root.cget('bg'), padx=5, pady=5)
        sheet_frame.pack(fill="x", pady=(10,0))
        sheet_frame.columnconfigure(2, weight=1)  # Scrap Mode column absorbs extra space

        self.unit_label = tk.Label(sheet_frame, text=_("Width ({units}):").format(units=self.settings['units']), bg=self.root.cget('bg'))
        self.unit_label.grid(row=0, column=0, sticky='w', padx=5)
        self.width_entry = tk.Entry(sheet_frame)
        self.width_entry.insert(0, self.settings["sheet_width"])
        self.width_entry.grid(row=0, column=1, sticky='w')

        self.height_label = tk.Label(sheet_frame, text=_("Height ({units}):").format(units=self.settings['units']), bg=self.root.cget('bg'))
        self.height_label.grid(row=1, column=0, sticky='w', padx=5)
        self.height_entry = tk.Entry(sheet_frame)
        self.height_entry.insert(0, self.settings["sheet_height"])
        self.height_entry.grid(row=1, column=1, sticky='w')

        # Scrap Mode row — own row beneath the sheet inputs so the label
        # always fits, even with longer translated labels (e.g. Spanish
        # "Modo retales") and on narrower window widths. Previously this
        # was crammed into column 2 between the entries and the d-pad,
        # which caused the checkbox label to clip and sometimes disappear
        # entirely on resize.
        scrap_inner_frame = tk.Frame(sheet_frame, bg=self.root.cget('bg'))
        scrap_inner_frame.grid(row=4, column=0, columnspan=3, sticky='w',
                               padx=(5, 0), pady=(8, 0))

        self.scrap_mode_var = tk.BooleanVar(value=False)
        tk.Checkbutton(scrap_inner_frame, text=_("Scrap Mode"),
                       variable=self.scrap_mode_var, bg=self.root.cget('bg'),
                       command=self._toggle_scrap_mode).pack(side='left')

        # Status label (shown when session active)
        self.scrap_status_var = tk.StringVar(value="")
        self.scrap_status_label = tk.Label(scrap_inner_frame,
                                           textvariable=self.scrap_status_var,
                                           bg=self.root.cget('bg'), font=("Helvetica", 8), fg="blue")

        # Clear button (shown when scrap mode checked)
        self.clear_scrap_btn = tk.Button(scrap_inner_frame, text=_("Clear"), font=("Helvetica", 8),
                                         command=self._on_clear_scrap_clicked)

        # Edge Bias d-pad (right side of sheet frame)
        bias_frame = tk.Frame(sheet_frame, bg=self.root.cget('bg'))
        bias_frame.grid(row=0, column=3, rowspan=4, sticky='n', padx=(15, 20))

        tk.Label(bias_frame, text=_("Edge Bias"), font=("Helvetica", 8),
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
                btn = tk.Button(bias_frame, text=_("ctr"), width=2, height=1,
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
            card_paper_frame, text=_("Fit card to paper:"), variable=self.card_paper_var,
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

        # Single entry point for drawn-by-hand shapes (always) and
        # camera-captured shapes (only when machine integration is
        # enabled). Label reflects what the dialog can actually do:
        # "Draw / Capture Shape" when camera path is available,
        # "Draw Shape" when it isn't.
        shape_btn_label = (_("Draw / Capture Shape...")
                           if self._machine_enabled()
                           else _("Draw Shape..."))
        tk.Button(shape_btn_frame, text=shape_btn_label,
                  command=self.on_draw_custom_shape).pack(side="left")
        self.shape_status_var = tk.StringVar(value="")
        self.shape_status_label = tk.Label(shape_btn_frame, textvariable=self.shape_status_var,
                                           bg=self.root.cget('bg'), fg="gray", font=("Helvetica", 9))
        self.shape_status_label.pack(side="left", padx=5)

        self.unload_shape_btn = tk.Button(shape_btn_frame, text=_("Unload"), command=self.on_unload_custom_shape)
        # Initially hidden, shown when shape is loaded
        self._update_shape_status()

        tk.Label(parent, text=_("Output filename base (no extension):"), bg=self.root.cget('bg')).pack(pady=5)
        self.filename_entry = tk.Entry(parent)
        self.filename_entry.insert(0, "my_pad_job")
        self.filename_entry.pack(padx=10)

        # Generate buttons. The "Frame & Cut" button is created here but
        # only packed when a Falcon (Grbl) controller is detected on USB.
        # The check happens after the UI is built — see _detect_falcon_async.
        generate_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        generate_frame.pack(pady=(15, 5))
        tk.Button(generate_frame, text=_("Generate SVG"), command=self.on_generate_svg, font=('Helvetica', 10, 'bold')).pack(side="left", padx=5)
        tk.Button(generate_frame, text=_("Generate G-code"), command=self.on_generate_gcode, font=('Helvetica', 10, 'bold')).pack(side="left", padx=5)
        self._frame_cut_btn = tk.Button(
            generate_frame, text=_("Frame & Cut"),
            command=self.on_frame_and_cut,
            font=('Helvetica', 10, 'bold'))
        # Not packed yet — _detect_falcon_async packs it when a
        # Grbl controller answers the handshake. Frame & Cut is
        # manual-mode only: user jogs the head to their material's
        # bottom-left corner, then framing + cutting use G92 to
        # zero work coords at that position. (An earlier camera-
        # offset AUTO mode was removed — calibration accuracy was
        # never tight enough to make it worth the surprise factor.)

        # Options below generate buttons
        options_frame = tk.Frame(parent, bg=self.root.cget('bg'))
        options_frame.pack(pady=(0, 10))

        self.preview_var = tk.BooleanVar(value=self.settings.get("show_preview", False))
        tk.Checkbutton(options_frame, text=_("Preview before saving"),
                       variable=self.preview_var, bg=self.root.cget('bg'),
                       command=lambda: self._save_checkbox("show_preview", self.preview_var)
                       ).pack(side="left", padx=(0, 15))

        # Live camera preview during Frame & Cut. Only relevant when
        # machine integration is enabled AND camera calibration is on
        # disk — created here but pack/unpack happens in
        # _refresh_machine_ui_state so the checkbox appears/disappears
        # alongside Frame & Cut as those gates resolve.
        self.live_camera_var = tk.BooleanVar(
            value=self.settings.get("show_live_camera", False))
        self._live_camera_chk = tk.Checkbutton(
            options_frame, text=_("Live camera preview"),
            variable=self.live_camera_var, bg=self.root.cget('bg'),
            command=lambda: self._save_checkbox("show_live_camera",
                                                  self.live_camera_var))

        self.eject_sd_var = tk.BooleanVar(value=self.settings.get("eject_sd_after_gcode", False))
        if sys.platform == 'win32':
            tk.Checkbutton(options_frame, text=_("Eject SD card after G-code export"),
                           variable=self.eject_sd_var, bg=self.root.cget('bg'),
                           command=self._on_eject_sd_changed).pack(side="left")

        # Now that all children exist, bind mousewheel on them
        self._pad_tab_bind_children_mousewheel(self._pad_tab_inner)

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
        """Return (width_mm, height_mm) for the selected paper size with a
        1/4" safety inset on each dimension, or None if not using paper size.

        The inset shrinks the cut area away from the physical paper edge so
        the user has room to align the paper in the laser cutter without
        running the bed clear of room for tape, fixtures, or eyeballed
        registration. Without it the nest can pack discs right up against
        the edge, which is awkward to align cleanly.
        """
        if not self.card_paper_var.get():
            return None
        inset_mm = 0.25 * 25.4  # 1/4 inch margin
        dropdown_val = self.card_paper_dropdown.get().lower()
        if dropdown_val.startswith("a4"):
            return (210.0 - inset_mm, 297.0 - inset_mm)  # A4 minus margin
        else:
            return (8.5 * 25.4 - inset_mm, 11.0 * 25.4 - inset_mm)  # Letter minus margin

    def _update_shape_status(self):
        """Update the custom shape status indicator."""
        if self.custom_polygon:
            n = len(self.custom_polygon)
            self.shape_status_var.set(
                _("Shape loaded ({n} pts)").format(n=n))
            self.shape_status_label.config(fg="green")
            self.unload_shape_btn.pack(side="left", padx=2)
        else:
            self.shape_status_var.set(_("Using rectangle dimensions"))
            self.shape_status_label.config(fg="gray")
            self.unload_shape_btn.pack_forget()

    def _show_draw_shape_tutorial(self):
        """Show first-time tutorial for the polygon drawing tool."""
        unit = self.settings.get("units", "in")
        if unit == "mm":
            unit = "cm"
        unit_label = _("inches") if unit == "in" else _("centimeters")

        msg = _(
            "Draw Custom Shape - How to Use\n\n"
            "• The grid is 15×15 {unit_label} (1 square = 1 {unit})\n"
            "• Click on grid intersections to add points (max 8)\n"
            "• Click near the first (green) point to close the shape\n"
            "• Click on any point to remove it\n"
            "• Use 'Clear' to start over\n"
            "• Click 'Submit' when your shape is complete\n\n"
            "This is useful for irregular leather skins and scrap pieces.\n\n"
            "Note: Generation can take 5-10x longer for complex shapes."
        ).format(unit_label=unit_label, unit=unit)
        messagebox.showinfo(_("Draw Custom Shape"), msg)

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

        dialog = PolygonDrawWindow(self.root, unit=unit,
                                     settings=self.settings)
        # Camera-captured / overlay-traced polygons come back in
        # absolute machine-mm; adopt those (inset + LB-machine
        # tracking applied inside _adopt_camera_polygon).
        machine_polygon = dialog.get_machine_polygon_mm()
        if machine_polygon:
            self._adopt_camera_polygon(machine_polygon)
            return

        polygon = dialog.get_polygon()
        if polygon:
            # Plain hand-drawn (no overlay, no capture) — no machine
            # reference. Clear any LB-machine tracking from a previous
            # polygon so the stale value doesn't leak into Frame & Cut.
            self._custom_polygon_lb_machine = None
            # Grid stores Y-UP (graph-paper); custom_polygon stores
            # Y-DOWN (SVG). Scale to mm in Y-UP, then hand off to
            # the shared adopter which applies the boundary flip.
            mm_per_unit = 25.4 if unit == "in" else 10.0
            y_up_mm = [(x * mm_per_unit, y * mm_per_unit) for (x, y) in polygon]
            self._set_custom_polygon_from_y_up(y_up_mm)

    def on_unload_custom_shape(self):
        """Unload the custom shape and return to rectangle mode."""
        self.custom_polygon = None
        self.custom_polygon_outline = None
        self._custom_polygon_lb_machine = None
        self._update_shape_status()

    def _camera_capture_ready(self):
        """True iff the experimental toggle is on AND OpenCV is installed
        AND a camera calibration file exists. Toggle-off short-circuits
        so the 'Re-capture from camera' button in the scrap-continue
        dialog (and any other caller using this helper) hides when the
        user has opted out of machine integration or is on a platform
        we don't support."""
        if not self._machine_enabled():
            return False
        try:
            import camera_capture
        except ImportError:
            return False
        if not camera_capture.HAS_OPENCV:
            return False
        return camera_capture.load_calibration(
            camera_capture.default_calibration_path()) is not None

    def _adopt_camera_polygon(self, machine_polygon):
        """Adopt an ABSOLUTE machine-Y-up polygon (from the camera's
        pixel→machine homography) as the active scrap shape.

        Stores BOTH:
          - self.custom_polygon: the INSET shape (safety-shrunk),
            used for pad nesting so cuts land safely inside the
            actual scrap edges.
          - self.custom_polygon_outline: the ORIGINAL un-insetted
            shape, used for Frame & Cut's framing trace so the trace
            visually matches the scrap edge.

        Both are normalized to the OUTLINE's bbox-BL so they share a
        coordinate origin — pad placements (in custom_polygon coords)
        land at the right machine positions during cuts even though
        the framing trace uses the outline.

        Also captures the polygon's LB-vertex MACHINE COORDINATE on
        self._custom_polygon_lb_machine for Frame & Cut's "Try Auto
        Locate" button.

        Used by:
          - File menu / standalone "Get from camera" flow
          - The polygon dialog's "Capture from camera" sub-action
          - The polygon dialog's overlay-traced path (via
            get_machine_polygon_mm)
        """
        if not machine_polygon:
            return
        # Capture the LB-vertex BEFORE inset shrinks the polygon — the
        # user is going to jog to the visible bottom-left corner of
        # their material, not the inset corner.
        self._custom_polygon_lb_machine = min(
            machine_polygon, key=lambda p: (p[0], p[1]))
        outline = list(machine_polygon)
        inset_polygon = outline
        try:
            import camera_capture
            inset_mm = float(
                self.settings.get("camera_polygon_inset_mm", 3.0))
            if inset_mm > 0:
                try:
                    inset_polygon = camera_capture.inset_polygon_mm(
                        outline, inset_mm)
                except Exception:
                    pass  # fall back to uninsetted on failure
        except ImportError:
            pass
        # Pass BOTH to the shared adopter so the inset polygon (for
        # nesting) and the outline (for framing) share an origin.
        self._set_custom_polygon_from_y_up(
            inset_polygon, outline_y_up_mm=outline)

    def _set_custom_polygon_from_y_up(self, points_y_up_mm,
                                        outline_y_up_mm=None):
        """Single boundary crossing into custom_polygon storage.

        Input:
          - points_y_up_mm: the polygon used for NESTING (the inset
            shape when there is one).
          - outline_y_up_mm (optional): the polygon used for FRAMING
            (the un-inset original). When None, frame and nest use
            the same shape — appropriate for plain hand-drawn
            polygons where no inset was applied.

        Both polygons are normalized to the OUTLINE's bbox-BL so they
        share a coordinate origin. The framing trace then matches the
        scrap edge while pad placements still land inside the inset
        boundary.

        Output:
          - self.custom_polygon: Y-DOWN, nesting boundary
          - self.custom_polygon_outline: Y-DOWN, framing trace
            (= self.custom_polygon when no separate outline)
        """
        if not points_y_up_mm:
            return
        outline = (outline_y_up_mm if outline_y_up_mm is not None
                    else points_y_up_mm)
        min_x = min(p[0] for p in outline)
        max_y = max(p[1] for p in outline)
        # Both polygons get the same translation + Y-flip so they
        # share an origin.
        self.custom_polygon = [(x - min_x, max_y - y)
                               for (x, y) in points_y_up_mm]
        self.custom_polygon_outline = [(x - min_x, max_y - y)
                                        for (x, y) in outline]
        self._update_shape_status()

    def _show_scrap_continue_dialog(self, placed_count, scrap_num, remaining_count):
        """Show scrap continue dialog. If a polygon is loaded, offer to unload it."""
        msg = _("Placed {placed_count} pads on scrap #{scrap_num}.\n\n"
                "{remaining_count} pads remaining.\n"
                "Adjust dimensions and click Generate again.").format(
                    placed_count=placed_count, scrap_num=scrap_num,
                    remaining_count=remaining_count)

        if not self.custom_polygon:
            messagebox.showinfo(_("Scrap Generated"), msg)
            return

        # Custom dialog with shape options
        dlg = tk.Toplevel(self.root)
        dlg.title(_("Scrap Generated"))
        dlg.configure(bg=self._get_theme_color())
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        bg = self._get_theme_color()
        tk.Label(dlg, text=msg, wraplength=380, bg=bg, justify="left",
                 font=("Helvetica", 10)).pack(padx=15, pady=(15, 10))
        tk.Label(dlg, text=_("A custom shape is loaded for the next scrap:"),
                 bg=bg, font=("Helvetica", 9, "italic")).pack(padx=15)

        btn_frame = tk.Frame(dlg, bg=bg)
        btn_frame.pack(pady=15)
        # Re-capture from camera: shortcut for the per-scrap workflow.
        # Most users with camera calibration will swap scraps between
        # cuts, so they want a new polygon each time without unloading
        # + opening the polygon dialog manually.
        if self._camera_capture_ready():
            tk.Button(btn_frame, text=_("Re-capture from camera"),
                       width=20,
                       command=lambda: self._scrap_dialog_recapture(dlg)
                       ).pack(side="left", padx=4)
        tk.Button(btn_frame, text=_("Unload Shape"), width=14,
                  command=lambda: self._scrap_dialog_close(dlg, unload=True)
                  ).pack(side="left", padx=4)
        tk.Button(btn_frame, text=_("Keep Shape"), width=14,
                  command=lambda: self._scrap_dialog_close(dlg, unload=False)
                  ).pack(side="left", padx=4)

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
            self.custom_polygon_outline = None
            self._custom_polygon_lb_machine = None
            self._update_shape_status()
        dlg.destroy()

    def _scrap_dialog_recapture(self, dlg):
        """Scrap-continue shortcut: clear the current polygon, close
        this dialog, and immediately open the polygon dialog to
        capture a fresh shape for the next scrap."""
        self.custom_polygon = None
        self.custom_polygon_outline = None
        self._custom_polygon_lb_machine = None
        self._update_shape_status()
        dlg.destroy()
        self.on_draw_custom_shape()

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
            center_btn.configure(text=_("off") if last_center == "off" else _("ctr"))

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
                messagebox.showwarning(_("Scrap Mode"),
                    _("Please select exactly one material to use Scrap Mode."))
                self.scrap_mode_var.set(False)
                return
        else:
            # Turning off - warn if session active
            if self.scrap_session['active']:
                remaining = self._count_remaining_pads()
                if remaining > 0:
                    if not messagebox.askyesno(_("Clear Session?"),
                        _("You have {remaining} pads remaining.\n"
                        "Disabling scrap mode will clear the session.\n\nContinue?").format(remaining=remaining)):
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
            'optimize': None,
        }
        self._unlock_material_selection()
        self._close_remaining_pads_window()
        self._update_scrap_status_display()

    def _update_scrap_status_display(self):
        """Update the scrap mode status UI."""
        if self.scrap_mode_var.get():
            # Show Clear button when scrap mode is checked. Lay out
            # horizontally beside the checkbox.
            self.clear_scrap_btn.pack(side='left', padx=(8, 0))

            if self.scrap_session['active']:
                # Show status when session is active
                remaining = self._count_remaining_pads()
                if remaining == 0:
                    self.scrap_status_var.set(_("Done!"))
                    self.scrap_status_label.config(fg="green")
                else:
                    count = self.scrap_session['scrap_count']
                    self.scrap_status_var.set(_("{remaining} left ({count} scraps)").format(remaining=remaining, count=count))
                    self.scrap_status_label.config(fg="blue")
                self.scrap_status_label.pack(side='left', padx=(8, 0))
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
                if not messagebox.askyesno(_("Clear Session?"),
                    _("You have {remaining} pads remaining.\n\nClear session?").format(remaining=remaining)):
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
            # Large-batch optimization opt-in. None = not yet asked
            # (prompt on first scrap with ≥ LARGE_BATCH_THRESHOLD pads
            # remaining), True/False = user's session-level answer.
            'optimize': None,
        }
        self._lock_material_selection(material)
        self._open_remaining_pads_window()

    LARGE_BATCH_THRESHOLD = 75  # remaining-pad count above which the
                                # optimization opt-in popup appears

    def _maybe_prompt_large_batch_optimization(self, pads):
        """If this scrap qualifies (≥ threshold pads remaining) AND the
        user hasn't been asked yet this session, ask once. Mutates
        scrap_session['optimize'] with the answer. Subsequent scraps
        in the same session use the same answer with no further prompt.
        """
        if self.scrap_session.get('optimize') is not None:
            return  # already asked + answered
        total = sum(p.get('qty', 0) for p in pads
                    if isinstance(p.get('qty'), int))
        if total < self.LARGE_BATCH_THRESHOLD:
            return  # not enough pads to bother
        answer = messagebox.askyesno(
            _("Large Batch Optimization"),
            _("You have {n} pads remaining in this scrap session.\n\n"
              "Use large-batch optimization for this session?\n\n"
              "The nester will try several pad orderings per scrap and "
              "keep the best result. Typically fits 5-15% more pads per "
              "scrap on large batches, but adds ~5-30 seconds of compute "
              "per scrap.\n\nApplies to every remaining scrap in this "
              "session.").format(n=total))
        self.scrap_session['optimize'] = bool(answer)

    def _scrap_begin_partial(self, pads, hole_dia, material, mat_w, mat_h,
                             mat_polygon, ask_save_dir):
        """Shared scrap-mode front half: start or continue the session,
        then run the partial nest for the current scrap.

        Used by both the file-export path (_generate_gcode_scrap_mode,
        ask_save_dir=True) and the Frame & Cut path (ask_save_dir=False)
        so the session bookkeeping can't drift between them.

        ask_save_dir: True prompts for an output folder when STARTING a
        new session (file export writes .gcode there). False starts the
        session with an empty save_dir — Frame & Cut streams to the
        laser and never writes files.

        Returns (placed, remaining) on success, or None if the caller
        should abort: folder prompt cancelled (silent), material
        mismatch, session already complete, or nothing fit (each shows
        its own message). On a successful return self.scrap_session is
        active and its 'hole_dia' is authoritative for generation.
        """
        if not self.scrap_session['active']:
            save_dir = ''
            if ask_save_dir:
                save_dir = filedialog.askdirectory(
                    title=_("Select Folder to Save G-code"),
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return None
                self.settings["last_output_dir"] = save_dir
            self._start_scrap_session(pads, material, save_dir, hole_dia)
        else:
            if self.scrap_session['material'] != material:
                messagebox.showerror(_("Material Mismatch"),
                    _("Current session is for {material}.\n"
                    "Clear session to switch materials.").format(material=self.scrap_session['material']))
                return None
            # A session started by Frame & Cut carries no save_dir. If
            # we're now exporting files into it, ask for a folder once
            # and remember it on the session for later file exports.
            if ask_save_dir and not self.scrap_session.get('save_dir'):
                save_dir = filedialog.askdirectory(
                    title=_("Select Folder to Save G-code"),
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return None
                self.settings["last_output_dir"] = save_dir
                self.scrap_session['save_dir'] = save_dir
            pads = self.scrap_session['remaining_pads']

        if not pads:
            messagebox.showinfo(_("Session Complete"), _("All pads have been placed!"))
            return None

        # Large-batch optimization opt-in (prompted once per session,
        # only on scraps with >= LARGE_BATCH_THRESHOLD pads remaining).
        self._maybe_prompt_large_batch_optimization(pads)
        _optimize = bool(self.scrap_session.get('optimize'))

        placed, remaining, any_placed = try_nest_partial(
            pads, material, mat_w, mat_h, self.settings,
            polygon=mat_polygon, optimize=_optimize)

        if not any_placed:
            min_pad_size = min(p['size'] for p in pads)
            messagebox.showwarning(_("No Pads Fit"),
                _("No pads could be placed on this scrap.\n\n"
                "Smallest remaining pad: {min_pad_size}mm\n"
                "Try a larger scrap piece.").format(min_pad_size=min_pad_size))
            return None

        return placed, remaining

    def _frame_cut_scrap_advance(self, placed_count, remaining):
        """Post-cut scrap bookkeeping for Frame & Cut, mirroring the tail
        of _generate_gcode_scrap_mode: bump the scrap count, record what's
        left for the next scrap, refresh the status UI, and either
        announce session completion or show the continue/recapture dialog.

        Called only after a cut has streamed to completion — if the user
        stopped the cut or it errored, the session is left untouched so
        the scrap can be re-cut.
        """
        self.scrap_session['scrap_count'] += 1
        scrap_num = self.scrap_session['scrap_count']
        self.scrap_session['remaining_pads'] = remaining
        self._update_scrap_status_display()
        self._update_remaining_pads_window()

        remaining_count = self._count_remaining_pads()
        if remaining_count == 0:
            messagebox.showinfo(_("Session Complete!"),
                                _("All pads have been placed!"))
        else:
            self._show_scrap_continue_dialog(
                placed_count, scrap_num, remaining_count)

    def _open_remaining_pads_window(self):
        """Open or update the floating window showing remaining and done pads."""
        if self.scrap_remaining_window is not None:
            try:
                self.scrap_remaining_window.destroy()
            except tk.TclError:
                pass

        self.scrap_remaining_window = tk.Toplevel(self.root)
        self.scrap_remaining_window.title(_("Scrap Mode Progress"))
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
        tk.Label(remaining_frame, text=_("Remaining"), bg=theme_bg,
                 font=("Helvetica", 9, "bold"), fg="blue").pack(anchor="w")
        self.scrap_remaining_listbox = tk.Listbox(remaining_frame, font=("Courier", 10), height=10, width=12)
        self.scrap_remaining_listbox.pack(fill="both", expand=True)

        # Done column
        done_frame = tk.Frame(columns_frame, bg=theme_bg)
        done_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))
        tk.Label(done_frame, text=_("Done"), bg=theme_bg,
                 font=("Helvetica", 9, "bold"), fg="green").pack(anchor="w")
        self.scrap_done_listbox = tk.Listbox(done_frame, font=("Courier", 10), height=10, width=12)
        self.scrap_done_listbox.pack(fill="both", expand=True)

        # Footer with scrap count
        self.scrap_window_footer = tk.Label(self.scrap_remaining_window, text="", bg=theme_bg,
                                            font=("Helvetica", 9))
        self.scrap_window_footer.pack(pady=(0, 5))

        # Close button
        tk.Button(self.scrap_remaining_window, text=_("Close"),
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
                self.scrap_window_header.config(text=_("All done!"), fg="green")
            else:
                self.scrap_window_header.config(
                    text=_("{done} / {total} pads complete").format(done=total_done, total=total_original), fg="blue")

            # Update Remaining listbox
            self.scrap_remaining_listbox.delete(0, tk.END)
            if remaining_pads:
                sorted_remaining = sorted(remaining_pads, key=lambda p: -p['size'])
                for pad in sorted_remaining:
                    size_str = f"{pad['size']:.1f}".rstrip('0').rstrip('.')
                    self.scrap_remaining_listbox.insert(tk.END, f" {pad['qty']} x {size_str}")
            else:
                self.scrap_remaining_listbox.insert(tk.END, _(" (none)"))

            # Update Done listbox
            self.scrap_done_listbox.delete(0, tk.END)
            if done_pads:
                sorted_done = sorted(done_pads, key=lambda p: -p['size'])
                for pad in sorted_done:
                    size_str = f"{pad['size']:.1f}".rstrip('0').rstrip('.')
                    self.scrap_done_listbox.insert(tk.END, f" {pad['qty']} x {size_str}")
            else:
                self.scrap_done_listbox.insert(tk.END, _(" (none)"))

            # Update footer
            self.scrap_window_footer.config(
                text=_("{material} | {scraps} scrap(s) used").format(material=material, scraps=scraps))

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
                val = float(self.custom_hole_entry.get())
            except (ValueError, TypeError):
                messagebox.showerror(_("Invalid Input"), _("Custom hole size must be a valid number."))
                return None
            if val <= 0:
                messagebox.showerror(_("Invalid Input"), _("Custom hole size must be greater than zero."))
                return None
            return val
        return 0

    def _prepare_generation(self):
        """Common validation and setup for SVG/G-code generation. Returns None on error, or a dict with generation params."""
        hole_dia = self.get_hole_dia()
        if hole_dia is None:
            return None

        rejected = []
        pads = self.parse_pad_list(self.pad_entry.get("1.0", tk.END), rejected)
        if not pads:
            messagebox.showerror(_("Error"), _("No valid pad sizes entered."))
            return None

        if rejected:
            shown = "\n".join(f"  • {ln}" for ln in rejected[:10])
            if len(rejected) > 10:
                shown += "\n" + _("  … and {n} more").format(n=len(rejected) - 10)
            if not messagebox.askyesno(
                _("Some lines skipped"),
                _("These lines couldn't be read as \"size x quantity\" and "
                  "won't be cut:\n\n"
                  "{lines}\n\n"
                  "Continue without them?").format(lines=shown)):
                return None

        max_pads = [p for p in pads if p['qty'] == 'max']
        if len(max_pads) > 1:
            messagebox.showerror(_("Error"), _("Only one pad size can use 'max' quantity at a time."))
            return None

        if self.settings.get("engraving_on", True):
            oversized_engravings = check_for_oversized_engravings(pads, self.material_vars, self.settings)
            if oversized_engravings and self.settings.get("show_engraving_warning", True):
                message = _("Warning: The current font size is too large for some pads and the engraving will be skipped:\n\n")
                for mat, sizes in oversized_engravings.items():
                    message += f"- {_(mat.replace('_', ' ').capitalize())}: {', '.join(map(str, sorted(sizes)))}\n"
                message += _("\nDo you want to proceed?")
                dialog = ConfirmationDialog(self.root, _("Engraving Size Warning"), message)
                if not dialog.result:
                    return None
                if dialog.dont_show_again.get():
                    self.settings["show_engraving_warning"] = False

        try:
            width_val = float(self.width_entry.get())
            height_val = float(self.height_entry.get())
        except ValueError:
            messagebox.showerror(_("Invalid Input"), _("Sheet width and height must be valid numbers."))
            return None

        if width_val <= 0 or height_val <= 0:
            messagebox.showerror(_("Invalid Input"), _("Sheet width and height must be greater than zero."))
            return None

        if self.settings['units'] == 'in':
            width_mm, height_mm = width_val * 25.4, height_val * 25.4
        elif self.settings['units'] == 'cm':
            width_mm, height_mm = width_val * 10, height_val * 10
        elif self.settings['units'] == 'mm':
            width_mm, height_mm = width_val, height_val
        else:
            messagebox.showerror(_("Error"), _("Unknown unit '{units}' in settings.").format(units=self.settings['units']))
            return None

        base = self.filename_entry.get().strip()
        if not base:
            messagebox.showerror(_("Error"), _("Please enter a base filename."))
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
                messagebox.showwarning(_("No Materials Selected"), _("Please select at least one material."))
                return

            use_preview = self.preview_var.get()

            # Preview requires single material
            if use_preview and len(selected_materials) > 1:
                messagebox.showinfo(_("Preview"),
                    _("Preview works with one material at a time.\n"
                    "Please select a single material to preview its layout."))
                return

            save_dir = None

            # Process each material (with optional per-material preview)
            for material in selected_materials:
                mat_w, mat_h, mat_polygon = self._get_material_dimensions(material, width_mm, height_mm, card_paper_dims)

                # Nest (may re-run if user adjusts and retries)
                placed = nest_pads(pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon)

                # Validate fit
                if not can_all_pads_fit(pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon):
                    size_desc = _("paper") if (material == "card" and card_paper_dims) else _("sheet")
                    messagebox.showerror(_("Nesting Error"), _("Could not fit all '{material}' pieces on the specified {size_desc} size.").format(material=material.replace('_', ' '), size_desc=size_desc))
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
                    save_dir = filedialog.askdirectory(title=_("Select Folder to Save SVGs"), initialdir=self.settings.get("last_output_dir", ""))
                    if not save_dir:
                        return
                    self.settings["last_output_dir"] = save_dir

                filename = os.path.join(save_dir, f"{base}_{material}.svg")
                generate_svg_from_placed(placed, material, mat_w, mat_h, filename, hole_dia, self.settings, polygon=mat_polygon)

            save_settings(self.settings)
            messagebox.showinfo(_("Done"), _("SVG files generated successfully."))

        except Exception as e:
            print(f"An error occurred during SVG generation: {e}")
            messagebox.showerror(_("An Error Occurred"), _("Something went wrong during generation:\n\n{error}").format(error=e))

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
                messagebox.showerror(_("Scrap Mode Error"),
                    _("Please select exactly one material for scrap mode."))
                return
            material = selected_materials[0]

            # Get scrap dimensions (polygon or rectangle)
            mat_w, mat_h, mat_polygon = self._get_material_dimensions(
                material, width_mm, height_mm, card_paper_dims)

            # Initialize or validate session
            if not self.scrap_session['active']:
                # Starting new session - ask for save directory
                save_dir = filedialog.askdirectory(
                    title=_("Select Folder to Save SVGs"),
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return
                self.settings["last_output_dir"] = save_dir

                self._start_scrap_session(pads, material, save_dir, hole_dia)
            else:
                # Continuing existing session - validate material matches
                if self.scrap_session['material'] != material:
                    messagebox.showerror(_("Material Mismatch"),
                        _("Current session is for {material}.\n"
                        "Clear session to switch materials.").format(material=self.scrap_session['material']))
                    return
                # Use remaining pads from session
                pads = self.scrap_session['remaining_pads']
                hole_dia = self.scrap_session['hole_dia']

            if not pads:
                messagebox.showinfo(_("Session Complete"), _("All pads have been placed!"))
                return

            # Large-batch optimization opt-in (prompted once per session,
            # only on scraps with ≥ LARGE_BATCH_THRESHOLD pads remaining).
            self._maybe_prompt_large_batch_optimization(pads)
            _optimize = bool(self.scrap_session.get('optimize'))

            # Attempt partial placement
            placed, remaining, any_placed = try_nest_partial(
                pads, material, mat_w, mat_h, self.settings,
                polygon=mat_polygon, optimize=_optimize)

            if not any_placed:
                min_pad_size = min(p['size'] for p in pads)
                messagebox.showwarning(_("No Pads Fit"),
                    _("No pads could be placed on this scrap.\n\n"
                    "Smallest remaining pad: {min_pad_size}mm\n"
                    "Try a larger scrap piece.").format(min_pad_size=min_pad_size))
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
                messagebox.showinfo(_("Session Complete!"),
                    _("Placed {placed_count} pads on scrap #{scrap_num}.\n\n"
                    "All pads placed! Session complete.\n"
                    "Files saved to: {save_dir}").format(placed_count=placed_count, scrap_num=scrap_num, save_dir=save_dir))
            else:
                self._show_scrap_continue_dialog(placed_count, scrap_num, remaining_count)

        except Exception as e:
            print(f"An error occurred during scrap mode SVG generation: {e}")
            messagebox.showerror(_("An Error Occurred"), _("Something went wrong:\n\n{error}").format(error=e))

    def _detect_falcon_async(self):
        """Probe USB serial ports for a Grbl controller; show the button if found.

        Runs in a background thread (Probe takes ~1-2 seconds per port) and
        marshals the result back to the Tk main thread. Safe to call any
        time the user wants to re-scan.
        """
        if self._falcon_detect_attempted:
            return
        self._falcon_detect_attempted = True
        try:
            import falcon_sender
            if not falcon_sender.HAS_PYSERIAL:
                return
        except ImportError:
            return

        import threading

        def _worker():
            try:
                port = falcon_sender.auto_detect_falcon()
            except Exception:
                port = None
            self.root.after(0, self._on_falcon_detected, port)

        threading.Thread(target=_worker, name='falcon-detect', daemon=True).start()

    def _on_falcon_detected(self, port):
        """Called on the Tk main thread once detection finishes."""
        if port:
            self.falcon_port = port
            # Frame & Cut button is gated on calibration as well as
            # Falcon detection — _refresh_machine_ui_state checks both.
            self._refresh_machine_ui_state()
            return
        # Not found this round. Auto-retry in 10s so users who started
        # SSC before powering on the Falcon get the button as soon as
        # the controller comes online — without having to restart the
        # app or click a Machine menu item to manually retrigger
        # detection. Stops retrying once a port is found.
        self._falcon_detect_attempted = False
        self.root.after(10000, self._detect_falcon_async)

    def _machine_enabled(self):
        """True iff the user has opted into machine integration.

        The underlying Falcon serial path (pyserial + Grbl 1.1
        character-counting protocol) is cross-platform — Linux and
        macOS users with a Falcon on USB can use it via `/dev/ttyUSB*`
        or `/dev/tty.usbserial-*`. The Falcon-camera auto-detect
        heuristic is per-platform (PowerShell on Windows,
        system_profiler on macOS, /sys/class/video4linux on Linux);
        wherever it fails, _resolve_camera_index falls through to the
        last-enumerated camera and the user can Switch camera in any
        dialog if that picks wrong.

        Single source of truth used by every gating site. Tested on
        Windows; the macOS/Linux paths are best-effort and the user
        is opting in via the experimental toggle."""
        return bool(self.settings.get("experimental_machine_menu", False))

    def _camera_calibration_present(self):
        """True iff a usable (non-legacy) camera calibration is on disk."""
        try:
            import camera_capture
        except ImportError:
            return False
        if not camera_capture.HAS_OPENCV:
            return False
        return camera_capture.load_calibration(
            camera_capture.default_calibration_path()) is not None

    def _refresh_machine_ui_state(self):
        """Single source of truth for which machine UI is enabled.

        Two levels of gating:
          - experimental_machine_menu OFF → nothing machine-related
            visible (Machine menu wasn't built at all)
          - experimental_machine_menu ON + no calibration → Machine
            menu visible with only Camera Calibration enabled; rest
            disabled, Frame & Cut hidden
          - experimental_machine_menu ON + calibration → all Machine
            items enabled; Frame & Cut shown when Falcon detected

        Safe to call any time; idempotent. Wired into:
          - app startup (after menu construction)
          - _on_falcon_detected (Falcon comes online)
          - after CameraCalibrationDialog returns saved (live un-grey)
        """
        toggle_on = self._machine_enabled()
        has_cal = self._camera_calibration_present()

        # Machine menu item states. Only meaningful if the cascade was
        # actually built (toggle was on at startup).
        if (toggle_on and hasattr(self, '_machine_menu')
                and self._machine_menu_indices):
            # Camera Calibration is ALWAYS enabled when the cascade
            # exists — it's the gateway. Everything else depends on
            # whether a calibration is on disk.
            other_state = "normal" if has_cal else "disabled"
            menu = self._machine_menu
            indices = self._machine_menu_indices
            try:
                menu.entryconfig(indices['camera_cal'], state="normal")
                for key in ('inset', 'home', 'test', 'clear', 'reset'):
                    if key in indices:
                        menu.entryconfig(indices[key], state=other_state)
            except tk.TclError:
                pass

        # Frame & Cut button visibility: needs toggle ON + calibration
        # AND a Falcon detected on USB. The packing happens here so
        # all three gates resolve through one helper.
        btn = getattr(self, '_frame_cut_btn', None)
        if btn is not None:
            should_show = bool(
                toggle_on and has_cal and self.falcon_port)
            try:
                is_packed = bool(btn.winfo_manager())
            except tk.TclError:
                is_packed = False
            if should_show and not is_packed:
                try:
                    btn.pack(side="left", padx=5)
                except tk.TclError:
                    pass
            elif not should_show and is_packed:
                try:
                    btn.pack_forget()
                except tk.TclError:
                    pass

        # Live camera preview checkbox: same gate as the camera-capture
        # features (toggle + calibration). Falcon doesn't need to be
        # detected — the checkbox is a setting consumed by Frame & Cut
        # only when that button is also packed (which DOES need Falcon),
        # but the user can preset it any time post-calibration.
        chk = getattr(self, '_live_camera_chk', None)
        if chk is not None:
            should_show_chk = bool(toggle_on and has_cal)
            try:
                is_packed = bool(chk.winfo_manager())
            except tk.TclError:
                is_packed = False
            if should_show_chk and not is_packed:
                try:
                    chk.pack(side="left", padx=(0, 15))
                except tk.TclError:
                    pass
            elif not should_show_chk and is_packed:
                try:
                    chk.pack_forget()
                except tk.TclError:
                    pass

    # ==================================================================
    # Machine menu handlers (Pad Maker > Machine cascade)
    # ==================================================================

    def _machine_require_falcon(self):
        """Common pre-flight: ensure Falcon is detected. Returns True
        if good, False (with an error popup) if not."""
        if not self.falcon_port:
            messagebox.showerror(
                _("Falcon Not Detected"),
                _("No Grbl controller detected. Plug in the Falcon "
                  "over USB and try again."))
            self._falcon_detect_attempted = False
            self._detect_falcon_async()
            return False
        return True

    def _resolve_camera_index(self):
        """Pick a camera index for the camera dialogs.

        Order: persisted override → Falcon-name heuristic → highest
        enumerated index. Returns None if no camera is available or
        OpenCV isn't installed.

        Persist a working index after a successful calibration save (via
        _persist_camera_index) so subsequent runs skip the guess.
        """
        try:
            import camera_capture
        except ImportError:
            return None
        if not camera_capture.HAS_OPENCV:
            return None
        override = self.settings.get("camera_index_override")
        if override is not None:
            try:
                return int(override)
            except (TypeError, ValueError):
                pass
        idx = camera_capture.find_falcon_camera_index()
        if idx is not None:
            return idx
        cams = camera_capture.enumerate_cameras()
        if cams:
            return cams[-1]['index']
        return None

    def _persist_camera_index(self, index):
        """Save a known-working camera index to settings so future
        camera dialogs open it directly without re-guessing. Called
        from the calibration dialog after a successful save."""
        if index is None:
            return
        try:
            current = self.settings.get("camera_index_override")
            if current == int(index):
                return
            self.settings["camera_index_override"] = int(index)
            from config import save_settings
            save_settings(self.settings)
        except Exception:
            pass  # never block on settings save

    def _run_blocking_home(self, sender):
        """Send $H on a worker thread, show the homing-status modal,
        block the caller until homing completes (or fails). Returns
        True on success, False on failure (with an error popup
        already shown).

        Uses an after()-poll loop on the Tk thread instead of the
        older `while t.is_alive(): root.update()` pattern — root.update()
        pumps the full Tk event loop including WM_DELETE_WINDOW, which
        could destroy root mid-loop during the 60s home and produce a
        TclError on the next iteration.
        """
        import threading
        result = {}

        def _worker():
            try:
                ok, m = sender.home(timeout_s=60.0)
                result['ok'] = ok
                result['msg'] = m
            except Exception as exc:
                result['ok'] = False
                result['msg'] = str(exc)

        self._show_homing_status()
        t = threading.Thread(target=_worker, daemon=True)
        t.start()
        # after()-poll until the worker finishes. wait_variable lets
        # the Tk event loop keep running (status modal stays painted,
        # other dialogs respond) without burning CPU in a tight loop.
        done_var = tk.BooleanVar(value=False)

        def _check():
            if t.is_alive():
                self.root.after(50, _check)
            else:
                done_var.set(True)

        self.root.after(50, _check)
        self.root.wait_variable(done_var)
        self._hide_homing_status()
        if result.get('ok'):
            self._falcon_homed_this_session = True
            return True
        messagebox.showerror(
            _("Homing Failed"),
            _("Falcon returned: {m}").format(
                m=result.get('msg', 'no response')))
        return False

    def _on_machine_home_falcon(self):
        """Standalone $H. Same flow as Frame & Cut's home prompt."""
        if not self._machine_require_falcon():
            return
        try:
            import falcon_sender
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Home needs pyserial: {e}").format(e=e))
            return
        if not messagebox.askyesno(
                _("Home Laser?"),
                _("Send the laser home? The head will travel to the "
                  "home corner at full speed. Make sure nothing is "
                  "in its path.")):
            return
        sender = falcon_sender.FalconSender(port=self.falcon_port)
        try:
            sender.connect()
            sender.unlock()
        except Exception as e:
            messagebox.showerror(_("Connection Failed"),
                                  _("Couldn't connect: {e}").format(e=e))
            try:
                sender.disconnect()
            except Exception:
                pass
            return
        success = self._run_blocking_home(sender)
        try:
            sender.disconnect()
        except Exception:
            pass
        if success:
            messagebox.showinfo(_("Homed"),
                                 _("Laser homed. Head is parked at "
                                   "the home corner."))

    def _on_machine_test_connection(self):
        """Quick Falcon ping: connect, read status, disconnect, report."""
        if not self._machine_require_falcon():
            return
        try:
            import falcon_sender
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Test needs pyserial: {e}").format(e=e))
            return
        sender = falcon_sender.FalconSender(port=self.falcon_port)
        try:
            sender.connect()
            status = sender.get_status(timeout=0.5)
        except Exception as e:
            messagebox.showerror(
                _("Connection Failed"),
                _("Couldn't connect to the Falcon at {p}: {e}").format(
                    p=self.falcon_port, e=e))
            try:
                sender.disconnect()
            except Exception:
                pass
            return
        finally:
            pass
        try:
            sender.disconnect()
        except Exception:
            pass
        if not status:
            messagebox.showwarning(
                _("Connected, no response"),
                _("Port opened at {p}, but Grbl didn't reply to a "
                  "status query. Try the Reset Falcon menu item, "
                  "then test again.").format(p=self.falcon_port))
            return
        mpos = status.get('mpos') or [0, 0, 0]
        messagebox.showinfo(
            _("Connection OK"),
            _("Falcon connected on {p}.\n\n"
              "State: {s}\n"
              "MPos: X={x:.2f}  Y={y:.2f}  Z={z:.2f}").format(
                p=self.falcon_port, s=status.get('state', '?'),
                x=mpos[0], y=mpos[1], z=mpos[2]))

    def _on_machine_clear_errors(self):
        """Send $X to unlock motion after an alarm. Doesn't reset
        position info, just clears the alarm state."""
        if not self._machine_require_falcon():
            return
        try:
            import falcon_sender
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Clear needs pyserial: {e}").format(e=e))
            return
        sender = falcon_sender.FalconSender(port=self.falcon_port)
        try:
            sender.connect()
            sender.unlock()
        except Exception as e:
            messagebox.showerror(_("Connection Failed"),
                                  _("Couldn't connect: {e}").format(e=e))
            try:
                sender.disconnect()
            except Exception:
                pass
            return
        finally:
            pass
        try:
            sender.disconnect()
        except Exception:
            pass
        messagebox.showinfo(
            _("Errors Cleared"),
            _("Sent $X unlock. Motion is allowed again. If the "
              "Falcon was in an alarm state, it should now be Idle "
              "— check with Test Connection."))

    def _on_machine_reset_falcon(self):
        """Send Ctrl-X (RT_SOFT_RESET) to reboot the controller's
        firmware state. Stops any in-progress motion immediately,
        clears alarms, resets the planner buffer."""
        if not self._machine_require_falcon():
            return
        if not messagebox.askyesno(
                _("Reset Falcon?"),
                _("Send a soft-reset (Ctrl-X) to the Falcon. This:\n"
                  "  • Immediately stops any motion in progress\n"
                  "  • Clears alarm + error states\n"
                  "  • Resets the planner buffer\n"
                  "  • Loses the machine's known position (you'll "
                  "need to home again)\n\n"
                  "Only use this when the Falcon is stuck in a bad "
                  "state and Clear Errors didn't help.\n\n"
                  "Continue?")):
            return
        try:
            import falcon_sender
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Reset needs pyserial: {e}").format(e=e))
            return
        sender = falcon_sender.FalconSender(port=self.falcon_port)
        try:
            sender.connect()
            sender.send_realtime(falcon_sender.RT_SOFT_RESET)
            time.sleep(1.0)  # let firmware reboot
        except Exception as e:
            messagebox.showerror(_("Reset Failed"),
                                  _("Couldn't reset: {e}").format(e=e))
        try:
            sender.disconnect()
        except Exception:
            pass
        # Soft-reset loses machine position — clear the homed flag so
        # the next Frame & Cut auto-homes again.
        self._falcon_homed_this_session = False
        messagebox.showinfo(
            _("Falcon Reset"),
            _("Soft-reset sent. The Falcon is now in alarm state — "
              "use Home Laser to recover, then jog as needed."))

    def _on_machine_recalibrate(self):
        """Open Camera Calibration with a strong warning since the
        engrave + capture flow takes the better part of an hour."""
        try:
            import camera_capture
        except ImportError:
            messagebox.showerror(
                _("OpenCV Required"),
                _("Camera calibration needs OpenCV:\n\n"
                  "    pip install opencv-python Pillow"))
            return
        cal_path = camera_capture.default_calibration_path()
        existing = camera_capture.load_calibration(cal_path) is not None
        if existing:
            if not messagebox.askyesno(
                    _("Recalibrate Camera?"),
                    _("⚠ This will REPLACE your existing calibration. "
                      "⚠\n\n"
                      "The full calibration takes ~60 minutes — the "
                      "calibration card has to be engraved on fresh "
                      "basswood, then you take 12+ photos.\n\n"
                      "Only recalibrate if:\n"
                      "  • The camera was bumped / re-mounted\n"
                      "  • Cuts are landing in the wrong spot\n"
                      "  • You're sure the existing one is wrong\n\n"
                      "Proceed?")):
                return
        # Delegate to the existing handler (inherited from
        # ToolingTabMixin) — same dialog, same flow.
        self._open_camera_calibration()

    def _on_machine_inset_settings(self):
        """Small dialog: edit camera_polygon_inset_mm."""
        dlg = tk.Toplevel(self.root)
        dlg.title(_("Camera-Polygon Inset Margin"))
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        bg = self._get_theme_color() if hasattr(
            self, '_get_theme_color') else None
        if bg:
            dlg.configure(bg=bg)

        tk.Label(dlg, justify='left', wraplength=420,
                  text=_("Camera-captured polygons are shrunk by this "
                          "many millimeters on every edge before "
                          "nesting. This safety margin compensates "
                          "for camera measurement error near the "
                          "edges of its view — keeps pads from "
                          "landing where a chunk of leather might "
                          "actually be missing.\n\n"
                          "Typical: 2-5mm. Set to 0 to disable.")
                  ).pack(padx=20, pady=(15, 10))

        row = tk.Frame(dlg)
        row.pack(pady=(0, 15))
        tk.Label(row, text=_("Inset margin (mm):")).pack(side='left',
                                                            padx=5)
        var = tk.StringVar(
            value=str(self.settings.get("camera_polygon_inset_mm", 3.0)))
        entry = tk.Entry(row, textvariable=var, width=8,
                          font=("Helvetica", 11))
        entry.pack(side='left', padx=5)
        entry.focus_set()

        def _ok():
            try:
                v = float(var.get())
                if v < 0 or v > 50:
                    raise ValueError("out of range")
            except (ValueError, TypeError):
                messagebox.showerror(_("Invalid"),
                                      _("Inset must be a number 0-50."),
                                      parent=dlg)
                return
            self.settings["camera_polygon_inset_mm"] = v
            from config import save_settings
            try:
                save_settings(self.settings)
            except Exception as e:
                messagebox.showerror(_("Save Failed"),
                                      str(e), parent=dlg)
                return
            dlg.destroy()

        btn_row = tk.Frame(dlg)
        btn_row.pack(pady=(0, 15))
        tk.Button(btn_row, text=_("Save"), command=_ok,
                   width=10, font=("Helvetica", 10, "bold")
                   ).pack(side='left', padx=5)
        tk.Button(btn_row, text=_("Cancel"), command=dlg.destroy,
                   width=10).pack(side='left', padx=5)

    def _show_homing_status(self):
        """Modal 'Homing...' window held open while $H blocks.

        Created as a Toplevel rather than a messagebox so we can dismiss
        it programmatically when homing returns. Lives in
        self._homing_status_win.
        """
        win = tk.Toplevel(self.root)
        win.title(_("Homing"))
        win.transient(self.root)
        win.resizable(False, False)
        try:
            win.grab_set()
        except Exception:
            pass
        tk.Label(win, text=_("Homing the laser...\n\n"
                              "Head is moving to the home corner."),
                 padx=30, pady=20,
                 font=("Helvetica", 11), justify='left').pack()
        win.update_idletasks()
        # Center on root window
        try:
            x = (self.root.winfo_x()
                 + (self.root.winfo_width() // 2)
                 - (win.winfo_width() // 2))
            y = (self.root.winfo_y()
                 + (self.root.winfo_height() // 2)
                 - (win.winfo_height() // 2))
            win.geometry(f"+{x}+{y}")
        except Exception:
            pass
        self._homing_status_win = win
        win.update()

    def _hide_homing_status(self):
        win = getattr(self, '_homing_status_win', None)
        if win is not None:
            try:
                win.destroy()
            except Exception:
                pass
            self._homing_status_win = None

    def on_frame_and_cut(self):
        """Generate G-code in memory, frame the bbox, then stream to Falcon."""
        try:
            import falcon_sender
            from gcode_engine import (
                extract_gcode_bbox,
                generate_framing_gcode,
                generate_polygon_framing_gcode,
            )
            from ui_dialogs import FalconRunDialog
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Frame & Cut needs pyserial:\n\n{e}").format(e=e))
            return

        # Long cuts can run for tens of minutes; block sleep so the
        # USB serial doesn't get killed mid-job. Released in the
        # finally block below.
        try:
            import sleep_lock
            sleep_lock.prevent_sleep()
        except Exception:
            pass

        if not self.falcon_port:
            messagebox.showerror(_("Falcon Not Detected"),
                                  _("No Grbl controller detected. Plug in "
                                    "the Falcon over USB and try again."))
            self._falcon_detect_attempted = False
            self._detect_falcon_async()
            return

        # No "would you like to recapture from camera?" prompt here —
        # the Draw / Capture Shape dialog already offered both paths
        # before the user picked one. If they drew a polygon, that's
        # their choice; proceeding in manual mode is the right
        # default. (An earlier two-button layout had Draw and Capture
        # as separate top-level buttons, and this prompt was the
        # bridge between them; with the consolidated dialog it just
        # makes users repeat work they already chose against.)

        # --- Input validation up through nesting (mirrors on_generate_gcode) ---
        try:
            params = self._prepare_generation()
            if not params:
                return
            pads, hole_dia = params['pads'], params['hole_dia']
            width_mm, height_mm = params['width_mm'], params['height_mm']
            card_paper_dims = params['card_paper_dims']

            supported_materials = [m for m, v in self.material_vars.items()
                                     if v.get() and m != "exact_size"]
            if len(supported_materials) != 1:
                messagebox.showerror(
                    _("Pick One Material"),
                    _("Frame & Cut works with one material at a time. "
                      "Please select exactly one material (not Exact Size)."))
                return
            material = supported_materials[0]
            mat_w, mat_h, mat_polygon = self._get_material_dimensions(
                material, width_mm, height_mm, card_paper_dims)

            # Scrap mode places a partial batch on this scrap (one scrap
            # per click, mirroring _generate_gcode_scrap_mode); standard
            # mode nests the whole batch and requires it all to fit.
            # scrap_remaining is handed to _frame_cut_scrap_advance once
            # the cut completes; it stays None in standard mode.
            scrap_mode = self.scrap_mode_var.get()
            scrap_remaining = None
            if scrap_mode:
                result = self._scrap_begin_partial(
                    pads, hole_dia, material, mat_w, mat_h, mat_polygon,
                    ask_save_dir=False)
                if result is None:
                    return
                placed, scrap_remaining = result
                hole_dia = self.scrap_session['hole_dia']
            else:
                placed = nest_pads(pads, material, mat_w, mat_h, self.settings,
                                    polygon=mat_polygon)
                if not can_all_pads_fit(pads, material, mat_w, mat_h,
                                        self.settings, polygon=mat_polygon):
                    messagebox.showerror(
                        _("Nesting Error"),
                        _("Could not fit all '{m}' pieces in the available "
                          "area.").format(m=material.replace('_', ' ')))
                    return

            # Generate G-code to a temp file then read it back into memory.
            import tempfile
            import os as _os
            with tempfile.NamedTemporaryFile(suffix='.gcode', delete=False,
                                              mode='w') as tmp:
                tmp_path = tmp.name
            try:
                generate_gcode_from_placed(
                    placed, material, mat_w, mat_h, tmp_path, hole_dia,
                    self.settings, polygon=mat_polygon)
                with open(tmp_path, 'r') as f:
                    gcode_text = f.read()
            finally:
                try:
                    _os.remove(tmp_path)
                except Exception:
                    pass

            # Framing pass. If a custom polygon is loaded (drawn or camera-
            # captured), trace the polygon outline — on irregular scrap a
            # bbox rectangle overhangs every concave edge and tells the user
            # nothing about whether the cuts land on material. Fall back to
            # the bbox rectangle for rectangular jobs (no polygon loaded).
            power_s = self.settings.get("laser_framing_power_s", 10)
            feed = self.settings.get("laser_framing_feed", 2000)
            framing_lines = []
            # Frame the OUTLINE polygon (un-insetted) so the trace
            # matches the actual scrap edge. Cuts still use the inset
            # polygon for placement — both share a coordinate origin
            # so the G92 math below works either way. Fall back to
            # custom_polygon for plain hand-drawn polygons that don't
            # have a separate outline.
            framing_polygon = (self.custom_polygon_outline
                                or self.custom_polygon)
            if framing_polygon and len(framing_polygon) >= 3:
                framing_lines = generate_polygon_framing_gcode(
                    framing_polygon, power_s=power_s, feed=feed,
                )
            else:
                bbox = extract_gcode_bbox(gcode_text)
                if bbox:
                    xmin, xmax, ymin, ymax = bbox
                    framing_lines = generate_framing_gcode(
                        xmin, ymin, xmax, ymax, power_s=power_s, feed=feed,
                    )

            sender = falcon_sender.FalconSender(port=self.falcon_port)
            try:
                sender.connect()
            except Exception as e:
                messagebox.showerror(_("Connection Failed"),
                                      _("Could not connect to the Falcon at "
                                        "{p}: {e}").format(p=self.falcon_port, e=e))
                return

            # Optional live camera window — opens alongside the run dialog
            # so the user can watch the laser head while it works, and stays
            # open after the cut so the user can see the finished result.
            # Stored on self so the reference persists past this function.
            # Only ONE window at a time: if a previous Frame & Cut left
            # a live cam open (it intentionally outlives the cut), focus
            # that one instead of opening a duplicate.
            if self.live_camera_var.get():
                try:
                    import camera_capture
                    from ui_dialogs import LiveCameraWindow
                    existing = getattr(self, '_live_cam_window', None)
                    if existing is not None and existing.winfo_exists():
                        try:
                            existing.lift()
                            existing.focus_force()
                        except Exception:
                            pass
                    elif camera_capture.HAS_OPENCV:
                        cam_idx = self._resolve_camera_index()
                        if cam_idx is not None:
                            self._live_cam_window = LiveCameraWindow(
                                self.root, camera_index=cam_idx,
                                settings=self.settings)
                except Exception:
                    pass  # never block the cut on preview failure

            # After our connect+soft-reset, Grbl is in alarm state
            # (homing required). $X clears that so motion is allowed.
            # We don't $H here — that would fly the head to the homing
            # switches and override whatever position the user manually
            # set as their material origin.
            sender.unlock()

            # Manual positioning step: jog dialog stays open so the
            # user can fine-tune with the arrow buttons. They can ALSO
            # have positioned the head physically (open lid → push head
            # to material's lower-left → close lid → click Start Frame),
            # so the Start Frame button is enabled immediately — no
            # forced jog click. If the head's already where they want,
            # they just click Start Frame straight through.
            # If the polygon has a known machine LB-vertex, offer
            # "Try Auto Locate" — drives the head directly there so
            # the user doesn't have to jog manually. Only enabled once
            # the laser is homed (otherwise MPos drift could put the
            # auto-drive somewhere wrong).
            auto_locate_target = self._custom_polygon_lb_machine
            jog_dlg = FalconRunDialog(
                self.root, sender, [],  # nothing to stream
                title=_("Position the head at your material's "
                         "bottom-left corner (jog here OR open the "
                         "lid and physically move it), then click "
                         "Start Frame"),
                show_pause_resume=False, show_cut_button=False,
                stop_needs_confirm=False,
                done_button_label=_("Start Frame →"),
                auto_locate_target=auto_locate_target,
                is_homed=self._falcon_homed_this_session)
            if jog_dlg._final_reason != "complete":
                return

            # Compute the polygon's leftmost-lowest vertex in framing's
            # Y-UP frame. This is the corner the user intuitively jogs
            # to ("BL of the scrap"). Computed from the OUTLINE (not
            # the inset) so the LB-vertex sits on an actual scrap
            # edge — for axis-aligned scraps this is bbox-(0,0); for
            # tilted scraps it's the visible bottom-left corner.
            #
            # The inset polygon shares the same coordinate origin
            # (see _set_custom_polygon_from_y_up), so cut placements
            # remain correctly anchored even though the G92 offset
            # comes from the outline.
            g92_x, g92_y = 0.0, 0.0
            offset_source = (self.custom_polygon_outline
                              or self.custom_polygon)
            if offset_source:
                _y_max_storage = max(p[1] for p in offset_source)
                _flipped = [(x, _y_max_storage - y)
                             for (x, y) in offset_source]
                _lb_idx = min(
                    range(len(_flipped)),
                    key=lambda i: _flipped[i][0] ** 2 + _flipped[i][1] ** 2)
                g92_x, g92_y = _flipped[_lb_idx]

            # Both modes share the same framing-loop recipe:
            # G92 X{lb_x} Y{lb_y} sets work coords such that the
            # polygon's bbox-BL ends up at work (0, 0) when the head
            # is physically at the polygon's LB vertex — what the
            # user thinks of as "BL of the material."
            prefix = [f'G92 X{g92_x:.3f} Y{g92_y:.3f}']
            if framing_lines:
                framing_lines = prefix + framing_lines

            try:
                if framing_lines:
                    # Framing always runs when there's something to
                    # frame — the user's escape hatch is clicking
                    # "Looks Good — Cut!" immediately in the loop
                    # dialog to skip to the cut. Loop mode lets them
                    # take their time verifying alignment and nudge
                    # the head with the jog buttons between/during
                    # passes.
                    fdlg = FalconRunDialog(
                        self.root, sender, framing_lines,
                        title=_("Framing — verify cut area, jog if "
                                 "needed, click Cut when ready"),
                        loop=True)
                    if fdlg._final_reason != "complete":
                        return  # user stopped or error during framing
                    # No "Frame Looks Good?" prompt — the user already
                    # confirmed by clicking Cut Now.

                # Cut uses the same G92 offset as framing so the cut
                # placements align with the LB-vertex jog convention.
                cut_lines = [f'G92 X{g92_x:.3f} Y{g92_y:.3f}'] + gcode_text.splitlines()
                cut_dlg = FalconRunDialog(
                    self.root, sender, cut_lines,
                    title=_("Cutting — {m}").format(m=material))

                # Scrap mode: commit this scrap (decrement remaining,
                # offer continue/recapture) only once the cut has
                # streamed to completion. A stopped or errored cut
                # leaves the session untouched so it can be re-cut.
                if scrap_mode and cut_dlg._final_reason == "complete":
                    self._frame_cut_scrap_advance(len(placed), scrap_remaining)
            finally:
                # Intentionally don't auto-close live_cam — users want
                # to see the finished result. They close it manually,
                # or it dies with the app when SSC exits.
                try:
                    sender.disconnect()
                except Exception:
                    pass

        except Exception as e:
            messagebox.showerror(_("An Error Occurred"),
                                  _("Something went wrong during Frame & Cut:\n\n"
                                    "{error}").format(error=e))
        finally:
            try:
                import sleep_lock
                sleep_lock.allow_sleep()
            except Exception:
                pass

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
                messagebox.showwarning(_("No Materials Selected"), _("Please select at least one material (G-code not supported for Exact Size)."))
                return

            use_preview = self.preview_var.get()

            if use_preview and len(supported_materials) > 1:
                messagebox.showinfo(_("Preview"),
                    _("Preview works with one material at a time.\n"
                    "Please select a single material to preview its layout."))
                return
            save_dir = None

            # Process each material with optional per-material preview
            all_placed = {}
            for material in supported_materials:
                mat_w, mat_h, mat_polygon = self._get_material_dimensions(material, width_mm, height_mm, card_paper_dims)

                placed = nest_pads(pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon)

                if not can_all_pads_fit(pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon):
                    size_desc = _("paper") if (material == "card" and card_paper_dims) else _("sheet")
                    messagebox.showerror(_("Nesting Error"), _("Could not fit all '{material}' pieces on the specified {size_desc} size.").format(material=material.replace('_', ' '), size_desc=size_desc))
                    return

                if use_preview:
                    preview = NestingPreviewWindow(
                        self.root, {material: placed}, mat_w, mat_h,
                        polygon=mat_polygon)
                    if preview.result != "save":
                        return

                all_placed[material] = (placed, mat_w, mat_h, mat_polygon)

            if save_dir is None:
                save_dir = filedialog.askdirectory(title=_("Select Folder to Save G-code"), initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return
                self.settings["last_output_dir"] = save_dir

            # Show working indicator
            working_popup = tk.Toplevel(self.root)
            working_popup.title(_("Working"))
            working_popup.geometry("250x80")
            popup_bg = self._get_theme_color()
            working_popup.configure(bg=popup_bg)
            working_popup.transient(self.root)
            working_popup.resizable(False, False)
            working_popup.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 125
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 40
            working_popup.geometry(f"+{x}+{y}")
            tk.Label(working_popup, text=_("Generating G-code..."), bg=popup_bg, font=("Helvetica", 12)).pack(expand=True)
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
                    messagebox.showinfo(_("Done"), _("G-code files generated successfully.\n\nSD card safely ejected — you can remove it now!"))
                else:
                    messagebox.showinfo(_("Done"), _("G-code files generated successfully.\n\nPlease safely eject the SD card before removing."))
            else:
                messagebox.showinfo(_("Done"), _("G-code files generated successfully."))

        except Exception as e:
            print(f"An error occurred during G-code generation: {e}")
            messagebox.showerror(_("An Error Occurred"), _("Something went wrong during G-code generation:\n\n{error}").format(error=e))

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
                messagebox.showerror(_("Scrap Mode Error"),
                    _("Please select exactly one material for scrap mode.\n"
                    "(G-code not supported for Exact Size)"))
                return
            material = selected_materials[0]

            # Get scrap dimensions (polygon or rectangle)
            mat_w, mat_h, mat_polygon = self._get_material_dimensions(
                material, width_mm, height_mm, card_paper_dims)

            # Start or continue the session and place this scrap's pads.
            # ask_save_dir=True: file export needs an output folder.
            result = self._scrap_begin_partial(
                pads, hole_dia, material, mat_w, mat_h, mat_polygon,
                ask_save_dir=True)
            if result is None:
                return
            placed, remaining = result
            hole_dia = self.scrap_session['hole_dia']

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
                messagebox.showinfo(_("Session Complete!"),
                    _("Placed {placed_count} pads on scrap #{scrap_num}.\n\n"
                    "All pads placed! Session complete.\n"
                    "Files saved to: {save_dir}").format(placed_count=placed_count, scrap_num=scrap_num, save_dir=save_dir))
            else:
                self._show_scrap_continue_dialog(placed_count, scrap_num, remaining_count)

        except Exception as e:
            print(f"An error occurred during scrap mode G-code generation: {e}")
            messagebox.showerror(_("An Error Occurred"), _("Something went wrong:\n\n{error}").format(error=e))

    def parse_pad_list(self, pad_input, rejected=None):
        """
        Parse pad input. Supports:
        - Regular: "18.0 x 5" (size x quantity)
        - Max fill: "18.0 x max" (fill remaining space with this size)
        Only one pad size can use "max" at a time.

        If `rejected` is a list, lines that can't be parsed are appended
        to it so the caller can warn — a typo'd line silently vanishing
        from a cut job means missing pads discovered after the sheet is cut.
        """
        pad_list = []
        for line in pad_input.strip().splitlines():
            line = line.strip().lower()
            if not line:
                continue
            ok = False
            try:
                parts = line.split('x', 1)  # Split only on first 'x' (so 'max' doesn't get split)
                if len(parts) == 2:
                    size = float(parts[0].strip())
                    if size > 0:
                        qty_str = parts[1].strip()
                        if qty_str == 'max':
                            pad_list.append({'size': size, 'qty': 'max'})
                            ok = True
                        else:
                            qty = int(float(qty_str))
                            if qty > 0:
                                pad_list.append({'size': size, 'qty': qty})
                                ok = True
            except ValueError:
                pass
            if not ok and rejected is not None:
                rejected.append(line)
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
        self.pad_preset_menu.set(_("Load Pad Preset"))

        # Toggle delete button: "Delete Library" when library is empty
        if (lib_name != "All Libraries"
                and lib_name in self.pad_presets
                and not self.pad_presets[lib_name]):
            self.pad_delete_btn.config(text=_("Delete Library"))
        else:
            self.pad_delete_btn.config(text=_("Delete Preset"))

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
            lib_name = simpledialog.askstring(_("Library Name"),
                _("Enter a library name to save to:"),
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

        name = simpledialog.askstring(_("Save Pad Preset"), _("Enter a name for this preset:"))
        if name:
            text_data = self.pad_entry.get("1.0", tk.END)
            if not text_data.strip():
                messagebox.showwarning(_("Save Pad Preset"), _("Cannot save an empty list."))
                return

            if active_library not in self.pad_presets:
                self.pad_presets[active_library] = {}

            # Preserve existing notes if overwriting
            existing_notes = ""
            if name in self.pad_presets[active_library]:
                if not messagebox.askyesno(_("Overwrite"), _("A set named '{name}' already exists in this library. Overwrite it?").format(name=name)):
                    return
                _unused, existing_notes = self._get_pad_preset_data(self.pad_presets[active_library][name])

            # Check for duplicate pad lists across all libraries
            new_lines = sorted(line.strip() for line in text_data.strip().splitlines() if line.strip())
            for lib, presets in self.pad_presets.items():
                for pname, pdata in presets.items():
                    if lib == active_library and pname == name:
                        continue  # Skip self when overwriting
                    existing_pads, _unused = self._get_pad_preset_data(pdata)
                    existing_lines = sorted(line.strip() for line in existing_pads.strip().splitlines() if line.strip())
                    if new_lines == existing_lines:
                        if not messagebox.askyesno(_("Duplicate Detected"),
                                _("This pad list is identical to '{pname}' "
                                "in '{lib}'.\n\nSave anyway?").format(pname=pname, lib=lib)):
                            return

            self.pad_presets[active_library][name] = {"pads": text_data, "notes": existing_notes}

            if save_presets(self.pad_presets, PAD_PRESET_FILE):
                self.pad_preset_loaded_library = active_library
                self.pad_preset_loaded_name = name
                self.pad_notes_btn.config(state="normal")
                self.on_pad_library_selected()
                messagebox.showinfo(_("Preset Saved"), _("Preset '{name}' saved successfully.").format(name=name))

    def on_load_pad_preset(self, selected_name):
        if not selected_name or selected_name == _("Load Pad Preset"):
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
            pads_text, _unused = self._get_pad_preset_data(raw)
            self.pad_entry.delete("1.0", tk.END)
            self.pad_entry.insert(tk.END, pads_text)
            self._auto_resize_pad_entry()
            self.pad_preset_loaded_library = lib_name
            self.pad_preset_loaded_name = preset_name
            self.pad_notes_btn.config(state="normal")

    def on_delete_pad_preset(self):
        selected_lib = self.pad_library_var.get()

        # Delete empty library
        if (selected_lib != "All Libraries"
                and selected_lib in self.pad_presets
                and not self.pad_presets[selected_lib]):
            if messagebox.askyesno(_("Delete Library"),
                    _("Are you sure you want to delete the empty library '{lib}'?").format(lib=selected_lib)):
                del self.pad_presets[selected_lib]
                save_presets(self.pad_presets, PAD_PRESET_FILE)
                self.update_pad_library_dropdown()
                messagebox.showinfo(_("Library Deleted"), _("Library '{lib}' deleted.").format(lib=selected_lib))
            return

        # Delete individual preset
        selected_preset = self.pad_preset_var.get()

        if not selected_preset or selected_preset.startswith("Load"):
            messagebox.showwarning(_("Delete Error"), _("Please load a set to delete."))
            return

        if selected_lib == "All Libraries":
            try:
                selected_lib, selected_preset = selected_preset.split("] ", 1)
                selected_lib = selected_lib[1:]
            except ValueError:
                messagebox.showerror(_("Delete Error"), _("Cannot delete from 'All Libraries' view. Please select the specific library first."))
                return

        if messagebox.askyesno(_("Delete Pad Preset"), _("Are you sure you want to delete the preset '{preset}' from the '{lib}' library?").format(preset=selected_preset, lib=selected_lib)):
            if selected_lib in self.pad_presets and selected_preset in self.pad_presets[selected_lib]:
                del self.pad_presets[selected_lib][selected_preset]
                if save_presets(self.pad_presets, PAD_PRESET_FILE):
                    self.on_pad_library_selected()
                    self.pad_entry.delete("1.0", tk.END)
                    self._auto_resize_pad_entry()
                    self.pad_preset_loaded_library = None
                    self.pad_preset_loaded_name = None
                    self.pad_notes_btn.config(state="disabled")
                    messagebox.showinfo(_("Preset Deleted"), _("Preset '{preset}' deleted.").format(preset=selected_preset))
            else:
                messagebox.showerror(_("Delete Error"), _("Could not find the preset to delete."))

    def on_pad_notes(self):
        lib = self.pad_preset_loaded_library
        name = self.pad_preset_loaded_name
        if not lib or not name or lib not in self.pad_presets or name not in self.pad_presets[lib]:
            messagebox.showwarning(_("Notes"), _("No preset loaded."))
            return

        _unused, current_notes = self._get_pad_preset_data(self.pad_presets[lib][name])
        dlg = PadNotesWindow(self.root, name, current_notes)
        if dlg.result is not None:
            # Update notes in the preset data
            pads_text, _unused = self._get_pad_preset_data(self.pad_presets[lib][name])
            self.pad_presets[lib][name] = {"pads": pads_text, "notes": dlg.result}
            save_presets(self.pad_presets, PAD_PRESET_FILE)

    def on_import_pad_presets(self):
        filepath = filedialog.askopenfilename(
            title=_("Import Pad Presets"),
            filetypes=((_("JSON files"), "*.json"), (_("All files"), "*.*")),
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
            messagebox.showerror(_("Import Error"), _("Could not import pad presets:\n{error}").format(error=e))

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
            messagebox.showerror(_("Connection Error"),
                _("Could not fetch pad sets from stohrermusic.com:\n\n{error}").format(error=e))
            return
        except (json.JSONDecodeError, ValueError) as e:
            messagebox.showerror(_("Data Error"),
                _("Invalid data received from server:\n\n{error}").format(error=e))
            return

        if not isinstance(web_data, dict) or not web_data:
            messagebox.showinfo(_("No Data"), _("No pad set data found on the server."))
            return

        WebImportPresetsWindow(self.root, web_data, self.pad_presets, PAD_PRESET_FILE, self)

    def on_import_settings_folder(self):
        """Import all config files from a user-selected folder."""
        folder = filedialog.askdirectory(
            title=_("Select Folder with Settings"),
            initialdir=self.settings.get("last_output_dir", "")
        )
        if not folder:
            return

        found_files = find_config_files_in_directory(folder)
        if not found_files:
            messagebox.showinfo(_("No Settings Found"),
                _("No config files found in the selected folder.\n\n"
                "Looking for: app_settings.json, pad_presets.json, sizing_presets.json, "
                "key_height_library.json, screw_specs.json, toner_data.json"))
            return

        file_list = "\n".join(f"  • {f}" for f in found_files)
        msg = _("The following files will be imported and will REPLACE your current settings:\n\n"
                "{file_list}\n\n"
                "Are you sure you want to continue?").format(file_list=file_list)

        if not messagebox.askyesno(_("Confirm Import"), msg):
            return

        try:
            import_config_files(folder, found_files)

            # Reload all data from the newly imported files
            self.settings = load_settings()
            self.pad_presets = load_presets(PAD_PRESET_FILE, preset_type_name="Pad Preset")
            self.key_presets = load_presets(KEY_PRESET_FILE, preset_type_name="Key Height")
            self.screw_data = load_presets(SCREW_SPECS_FILE, preset_type_name="Screw Specs")
            self.sizing_presets = load_presets(SIZING_PRESET_FILE, preset_type_name="Sizing Preset")
            # The app guarantees at least one sizing preset exists
            if not self.sizing_presets:
                self.sizing_presets["Default"] = settings_to_sizing_preset(self.settings)
                save_presets(self.sizing_presets, SIZING_PRESET_FILE)

            # Refresh UI dropdowns
            self.update_pad_library_dropdown()
            self.update_key_library_dropdown()
            self.update_screw_maker_list()

            # Update UI elements that depend on settings
            self.refresh_widgets_from_settings()
            self.apply_resonance_theme()

            messagebox.showinfo(_("Import Complete"),
                _("Successfully imported {count} file(s).\n\n"
                "Imported: {files}").format(count=len(found_files), files=', '.join(found_files)))
        except Exception as e:
            messagebox.showerror(_("Import Error"), _("Could not import settings:\n{error}").format(error=e))

    # --- Misc Windows ---
    def open_options_window(self):
        OptionsWindow(
            self.root, self, self.settings,
            self.update_ui_from_settings,
            lambda: save_settings(self.settings),
            sizing_presets=self.sizing_presets,
            sizing_presets_save_callback=lambda: save_presets(self.sizing_presets, SIZING_PRESET_FILE),
        )


    def open_key_layout_window(self):
        KeyLayoutWindow(self.root, self.settings, self.rebuild_key_tab, lambda: save_settings(self.settings))

    def open_color_window(self):
        LayerColorWindow(self.root, self.settings, lambda: save_settings(self.settings))

    def open_gcode_settings_window(self):
        # Show only pad materials (felt/card/leather) from the pad generator tab
        pad_materials = [("felt", _("Felt")), ("card", _("Card")), ("leather", _("Leather"))]
        GcodeSettingsWindow(self.root, self.settings, lambda s: save_settings(s),
                            materials=pad_materials,
                            gcode_presets=self.gcode_presets,
                            gcode_presets_save_callback=lambda: save_presets(self.gcode_presets, GCODE_PRESET_FILE))

    def open_resonance_window(self):
        ResonanceWindow(self.root, self.settings, lambda: save_settings(self.settings), self.apply_resonance_theme)

    def _open_input_device_dialog(self):
        """Open a dialog to select the audio input device."""
        if sys.platform == 'linux':
            messagebox.showinfo(_("Input Device"),
                _("On Linux, audio input is locked to the system default.\n\n"
                "Set your preferred device in your system audio settings\n"
                "(PulseAudio or PipeWire)."))
            return

        devices = get_input_devices()
        if not devices:
            messagebox.showinfo(_("No Devices"), _("No audio input devices found."))
            return

        dlg = tk.Toplevel(self.root)
        dlg.title(_("Input Device"))
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = self.root.cget('bg') if not IS_MACOS else "systemWindowBackgroundColor"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=_("Select audio input device:"), bg=bg,
                 font=("Helvetica", 10)).pack(pady=(0, 4))
        tk.Label(frame, text=_("Showing devices with 44.1 kHz+ sample rate.\n"
                 "Bluetooth headsets are excluded (sample rate too low)."),
                 bg=bg, fg="#666666", font=("Helvetica", 8)).pack(pady=(0, 8))

        current_dev = self.settings.get("audio_input_device")
        dev_names = [_("System Default")] + [name for _idx, name in devices]
        dev_indices = [None] + [idx for idx, _name in devices]

        mic_var = tk.StringVar(value=_("System Default"))
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
            for i, (idx, _name) in enumerate(devices):
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
            messagebox.showinfo(_("Input Device"),
                _("Audio input set to: {dev_name}").format(dev_name=dev_name))

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text=_("Apply"), command=apply).pack(
            side="left", padx=(0, 5))
        tk.Button(btn_frame, text=_("Cancel"), command=dlg.destroy).pack(
            side="left")

    def _open_capture_threshold(self):
        """Open capture threshold dialog with live level meter."""
        dlg = tk.Toplevel(self.root)
        dlg.title(_("Capture Threshold"))
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = self.root.cget('bg') if not IS_MACOS else "systemWindowBackgroundColor"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=_("Set the minimum signal level to trigger capture.\n"
                 "Raise the threshold if chair noises or breathing\n"
                 "trigger false captures."),
                 bg=bg, font=("Helvetica", 9), justify="left").pack(pady=(0, 10))

        # Live level meter
        meter_frame = tk.Frame(frame, bg=bg)
        meter_frame.pack(fill="x", pady=(0, 10))

        tk.Label(meter_frame, text=_("Current level:"), bg=bg,
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

        tk.Label(slider_frame, text=_("Threshold:"), bg=bg,
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
        tk.Button(btn_frame, text=_("Apply"), command=apply).pack(
            side="left", padx=(0, 5))
        tk.Button(btn_frame, text=_("Cancel"), command=on_close).pack(side="left")

        dlg.protocol("WM_DELETE_WINDOW", on_close)

    def _open_feature_set(self):
        """Open the Feature Set dialog to show/hide tabs."""
        from i18n import available_languages, current_language
        dlg = tk.Toplevel(self.root)
        dlg.title(_("Feature Set"))
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = self.root.cget('bg') if not IS_MACOS else "systemWindowBackgroundColor"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=_("Feature Set"), bg=bg,
                 font=("Helvetica", 14, "bold")).pack(pady=(0, 5))
        tk.Label(frame, text=_("Choose which tabs to show.\n"
                 "The Pad Maker is always on."),
                 bg=bg, font=("Helvetica", 9)).pack(pady=(0, 10))

        # --- Language picker (requires restart) ---
        tk.Label(frame, text=_("Language"), bg=bg,
                 font=("Helvetica", 11, "bold")).pack(anchor="w")
        lang_row = tk.Frame(frame, bg=bg)
        lang_row.pack(anchor="w", pady=(2, 8))
        langs = available_languages()
        lang_codes = [code for code, _name in langs]
        lang_displays = [name for _code, name in langs]
        current = current_language()
        try:
            current_idx = lang_codes.index(current)
        except ValueError:
            current_idx = 0
        lang_var = tk.StringVar(value=lang_displays[current_idx])
        lang_combo = ttk.Combobox(lang_row, textvariable=lang_var,
                                  values=lang_displays, state="readonly",
                                  width=18)
        lang_combo.pack(side="left")
        tk.Label(lang_row, text=_(" (restart to apply)"), bg=bg,
                 font=("Helvetica", 9, "italic")).pack(side="left")

        # --- Tooltips toggle (applies immediately, no restart) ---
        tk.Label(frame, text=_("Tool Tips"), bg=bg,
                 font=("Helvetica", 11, "bold")).pack(anchor="w")
        tooltips_var = tk.BooleanVar(value=self.settings.get("tooltips_enabled", True))
        tk.Checkbutton(frame,
                       text=_("Show tool tips when hovering over settings"),
                       variable=tooltips_var, bg=bg,
                       font=("Helvetica", 10)).pack(anchor="w")
        tk.Label(frame, text=_("\nTabs"), bg=bg,
                 font=("Helvetica", 11, "bold")).pack(anchor="w")

        visible = self.settings.get("visible_tabs", {})

        # (key, display) — key is the storage identifier in visible_tabs and
        # MUST stay English. The display string is wrapped here as a literal
        # so pybabel can extract it; falls through to the live translator.
        main_tabs = [
            ("Key Height Library", _("Key Height Library")),
            ("Serial Lookup", _("Serial Lookup")),
            ("Screw Specs", _("Screw Specs")),
            ("Tooling", _("Tooling")),
            ("Tuner", _("Tuner")),
        ]
        experimental_tabs = [
            ("Toner", _("Toner")),
        ]

        check_vars = {}
        for key, display in main_tabs:
            var = tk.BooleanVar(value=visible.get(key, True))
            tk.Checkbutton(frame, text=display, variable=var, bg=bg,
                           font=("Helvetica", 10)).pack(anchor="w")
            check_vars[key] = var

        tk.Label(frame, text=_("\nExperimental / In Progress"), bg=bg,
                 font=("Helvetica", 11, "bold")).pack(anchor="w")
        for key, display in experimental_tabs:
            var = tk.BooleanVar(value=visible.get(key, True))
            tk.Checkbutton(frame, text=display, variable=var, bg=bg,
                           font=("Helvetica", 10)).pack(anchor="w")
            check_vars[key] = var

        # Experimental MENU toggles (not tabs — stored as top-level
        # settings keys, not under visible_tabs). Restart-required.
        # Available on all platforms — the serial path is cross-
        # platform (pyserial), and camera detection degrades gracefully
        # to manual Switch-camera on platforms where we don't have a
        # name-based auto-detect. Windows is the primary tested
        # platform; the macOS/Linux paths are best-effort.
        machine_menu_var = tk.BooleanVar(
            value=self.settings.get("experimental_machine_menu", False))
        tk.Checkbutton(
            frame,
            text=_("Machine menu (direct Falcon serial control)"),
            variable=machine_menu_var, bg=bg,
            font=("Helvetica", 10)).pack(anchor="w")

        def _show_toner_terms(parent_dlg):
            """Show toner beta terms acceptance dialog. Returns True if accepted."""
            terms = tk.Toplevel(parent_dlg)
            terms.title(_("Toner — Beta Notice"))
            terms.resizable(True, True)
            terms.transient(parent_dlg)
            terms.grab_set()
            terms.minsize(460, 400)

            tbg = parent_dlg.cget('bg') if not IS_MACOS else "systemWindowBackgroundColor"
            frm = tk.Frame(terms, bg=tbg, padx=20, pady=15)
            frm.pack(fill="both", expand=True)

            tk.Label(frm, text=_("Harmonic Tone Analyzer"), bg=tbg,
                     font=("Helvetica", 14, "bold")).pack(pady=(0, 10))

            txt_frame = tk.Frame(frm, bg=tbg)
            txt_frame.pack(fill="both", expand=True)
            txt = tk.Text(txt_frame, wrap="word", font=("Helvetica", 10),
                          bg=tbg, relief="flat", highlightthickness=0,
                          padx=8, pady=8, spacing3=4)
            txt.pack(side="left", fill="both", expand=True)
            sb = tk.Scrollbar(txt_frame, command=txt.yview)
            sb.pack(side="right", fill="y")
            txt.configure(yscrollcommand=sb.set)

            # Tag for the data-sharing link
            txt.tag_configure("link",
                              font=("Helvetica", 10, "underline"),
                              foreground="#0066CC")

            txt.insert("end",
                _("This feature is in beta. Run it in full screen.\n\n"

                "The Toner is a real-time harmonic analyzer that captures the "
                "frequency content of your recorded sound, extracting and "
                "analyzing the saxophone-specific content: the fundamental "
                "and its overtones. A microphone picks up your sound, an FFT "
                "extracts the harmonic series, and the display shows you "
                "what's happening in real time. You can capture sessions, "
                "save presets for different setups (horn + mouthpiece + "
                "player + reed + mic), and compare them side by side.\n\n"

                "The formulas and scaling may still shift as we gather more data. "
                "The raw harmonic measurements are always saved, so any future "
                "analytical improvements apply retroactively to all your "
                "historical captures.\n\n"

                "THREE WAYS TO USE IT\n\n"

                "1. Instant feedback \u2014 see how you affect the sound with "
                "your embouchure, support, voicing, mic placement, etc.\n\n"

                "2. Controlled comparison \u2014 change a variable and compare "
                "readings to see what changes in your sound. More readings, "
                "better data.\n\n"

                "3. Tracking your sound over time \u2014 same horn, same setup, "
                "just you and your practice. Sessions are dated, so you "
                "can see how your tone evolves week to week or month "
                "to month.\n\n"

                "BEFORE YOUR FIRST CAPTURE\n\n"

                "Check Options > Settings:\n\n"

                "1. Input Device \u2014 select your audio input. It should "
                "automatically choose your default. Best practice is to have "
                "your mic plugged in before running the program.\n\n"

                "2. Recording Folder \u2014 this program does its best work by "
                "analyzing WAV files it records during your sessions. Choose "
                "where these WAV files are stored, and whether to keep them "
                "after analysis or automatically delete.\n\n"

                "3. Preset Fields \u2014 choose which variable fields to show when "
                "creating presets.\n\n"

                "Then create a preset in File > Presets. Six fields are always "
                "required: Make, Model, Player, Mouthpiece, Mic Type, and "
                "Mic Model. If you want accuracy, a high quality condenser "
                "mic will give the most accurate readings, but you can also "
                "use the toner to compare and learn about mic and mic "
                "placement effects.\n\n"

                "HELP IMPROVE THE TONER\n\n"

                "I need more data! Please help by contributing recordings "
                "and preset files. Every recording helps calibrate the "
                "descriptors and uncover what they actually measure across "
                "different horns, mouthpieces, mics, and players. If you'd "
                "like to contribute, export your preset library via File > "
                "Transfer Data > Export Preset Library... and make yourself "
                "a folder and drop the resulting JSON file (and any WAV "
                "recordings you'd like to share) in the shared folder below:\n\n")
            )

            # Clickable Google Drive link
            import webbrowser
            drive_url = ("https://drive.google.com/drive/folders/"
                         "1fZndlhDv57vdr6lApOQcvvVFP6pK0sj_?usp=drive_link")
            txt.insert("end", drive_url + "\n", "link")
            txt.tag_bind("link", "<Button-1>",
                         lambda e: webbrowser.open(drive_url))
            txt.tag_bind("link", "<Enter>",
                         lambda e: txt.config(cursor="hand2"))
            txt.tag_bind("link", "<Leave>",
                         lambda e: txt.config(cursor=""))

            txt.insert("end",
                _("\nThis feature is in active development. Please contact me "
                "with suggestions, questions, problems. If you experience a "
                "bug or problem, screenshots are super helpful.\n")
            )

            txt.configure(state="disabled")

            result = {"accepted": False}

            def accept():
                result["accepted"] = True
                terms.destroy()

            btn_frame = tk.Frame(frm, bg=tbg)
            btn_frame.pack(fill="x", pady=(10, 0))
            tk.Button(btn_frame, text=_("I Understand — Enable Toner"),
                      command=accept).pack(side="left", padx=(0, 5))
            tk.Button(btn_frame, text=_("Cancel"),
                      command=terms.destroy).pack(side="left")

            terms.wait_window()
            return result["accepted"]

        def apply():
            new_visible = {name: var.get() for name, var in check_vars.items()}

            # Toner requires terms acceptance if not previously unlocked
            if new_visible.get("Toner") and not self.settings.get("toner_unlocked"):
                if not _show_toner_terms(dlg):
                    new_visible["Toner"] = False
                    check_vars["Toner"].set(False)
                    return
                self.settings["toner_unlocked"] = True

            tabs_changed = (new_visible != self.settings.get("visible_tabs", {}))

            # Experimental menu toggle (Machine cascade). Restart-required
            # — the menu bar is built once at startup; rebuilding it
            # live would require tearing down + recreating all menus.
            new_machine = bool(machine_menu_var.get())
            machine_changed = new_machine != self.settings.get(
                "experimental_machine_menu", False)
            self.settings["experimental_machine_menu"] = new_machine

            # Language change requires restart to take effect.
            selected_display = lang_var.get()
            try:
                selected_code = lang_codes[lang_displays.index(selected_display)]
            except ValueError:
                selected_code = self.settings.get("language", "en")
            lang_changed = selected_code != self.settings.get("language", "en")
            self.settings["language"] = selected_code

            # Tooltip toggle takes effect immediately — no restart needed.
            from ui_dialogs import set_tooltips_enabled
            self.settings["tooltips_enabled"] = tooltips_var.get()
            set_tooltips_enabled(tooltips_var.get())

            self.settings["visible_tabs"] = new_visible
            save_settings(self.settings)
            dlg.destroy()
            if lang_changed:
                messagebox.showinfo(_("Feature Set"),
                    _("Restart the app for the language change to take effect."))
            elif tabs_changed or machine_changed:
                messagebox.showinfo(_("Feature Set"),
                    _("Changes will take effect next time you open the app."))

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x", pady=(10, 0))
        tk.Button(btn_frame, text=_("Apply"), command=apply).pack(
            side="left", padx=(0, 5))
        tk.Button(btn_frame, text=_("Cancel"), command=dlg.destroy).pack(
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

    def _open_log_file(self):
        """Open the log file in the system's default text editor."""
        log_path = get_log_file()
        if log_path and os.path.exists(log_path):
            if sys.platform == 'win32':
                os.startfile(log_path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', log_path])
            else:
                subprocess.Popen(['xdg-open', log_path])
        else:
            messagebox.showinfo(_("Log File"), _("No log file found.\nExpected at: {log_path}").format(log_path=log_path))

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
        self.unit_label.config(text=_("Width ({units}):").format(units=self.settings['units']))
        self.height_label.config(text=_("Height ({units}):").format(units=self.settings['units']))

    def refresh_widgets_from_settings(self):
        """Repopulate main-tab widgets from self.settings after an import.

        on_exit() writes these widgets' state BACK into settings, so if
        they aren't refreshed after Import Settings from Folder, the stale
        pre-import values silently clobber the imported config at exit.
        Kept separate from update_ui_from_settings(), which also runs on
        Options-Apply where resetting mid-session sheet edits would be rude.
        """
        self.update_ui_from_settings()
        self.width_entry.delete(0, tk.END)
        self.width_entry.insert(0, self.settings["sheet_width"])
        self.height_entry.delete(0, tk.END)
        self.height_entry.insert(0, self.settings["sheet_height"])
        self.hole_var.set(self.settings["hole_option"])
        self.custom_hole_entry.config(state='normal')
        self.custom_hole_entry.delete(0, tk.END)
        self.custom_hole_entry.insert(0, self.settings.get("custom_hole_size", "4.0"))
        self.toggle_custom_hole_entry()
        self.card_paper_var.set(self.settings.get("card_use_paper_size", False))
        if self.settings.get("card_paper_size", "letter") == "a4":
            self.card_paper_dropdown.set("a4 (210×297 mm)")
        else:
            self.card_paper_dropdown.set("letter (8.5×11 in)")
        self._toggle_card_paper_dropdown()

if __name__ == '__main__':
    setup_logging()

    def _handle_exception(exc_type, exc_value, exc_tb):
        """Log unhandled exceptions to the log file and show a dialog."""
        import traceback
        import logging
        tb_text = ''.join(traceback.format_exception(exc_type, exc_value, exc_tb))
        logging.error("Unhandled exception:\n%s", tb_text)
        try:
            messagebox.showerror(
                _("Unexpected Error"),
                _("{exc_type}: {exc_value}\n\n"
                "Details saved to log file (Help > Open Log File).").format(
                    exc_type=exc_type.__name__, exc_value=exc_value))
        except Exception:
            pass  # GUI may not be available

    sys.excepthook = _handle_exception

    root = tk.Tk()

    def _handle_tk_exception(exc_type, exc_value, exc_tb):
        """Handle exceptions in tkinter callbacks."""
        _handle_exception(exc_type, exc_value, exc_tb)

    root.report_callback_exception = _handle_tk_exception

    app = PadSVGGeneratorApp(root)
    # Run the initial machine-UI refresh so menu items + Frame & Cut
    # button start in the correct state (disabled / hidden when no
    # calibration is on disk yet, even though the toggle is on).
    try:
        app._refresh_machine_ui_state()
    except Exception:
        pass
    # Detect the Falcon (Grbl) controller in the background only if
    # the user has opted into the experimental Machine menu — there's
    # no point probing USB serial ports for a user who never plans to
    # use direct Falcon control.
    if app._machine_enabled():
        root.after(500, app._detect_falcon_async)
    root.mainloop()
