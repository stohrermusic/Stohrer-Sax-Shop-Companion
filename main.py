import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import json
import sys
import subprocess
import time
import datetime
import logging
import math

# --- Local Imports ---
from config import (
    load_settings, save_settings, load_presets, save_presets,
    PAD_PRESET_FILE, KEY_PRESET_FILE, SCREW_SPECS_FILE, SIZING_PRESET_FILE,
    GCODE_PRESET_FILE, GCODE_PRESET_MATERIALS,
    append_job_history, load_job_history, save_job_history,
    find_config_files_in_directory, import_config_files,
    get_ssl_context, get_input_devices, migrate_analyzer_tab,
    setup_logging, get_log_file,
    settings_to_sizing_preset, settings_to_gcode_presets,
)

# Initialize translations BEFORE importing UI modules. Any module-level
# strings in ui_dialogs/main.py/toner_tab/etc. that use _() need the
# translation catalog already loaded when their module body runs.
from i18n import init_translation
init_translation(load_settings().get("language", "en"))

from svg_engine import check_for_oversized_engravings, try_nest_partial, generate_svg_from_placed, nest_with_zones  # noqa: E402
from gcode_engine import generate_gcode_from_placed  # noqa: E402
import tkinter.font as tkfont  # noqa: E402
from ui_dialogs import (  # noqa: E402
    OptionsWindow, LayerColorWindow, KeyLayoutWindow,
    ResonanceWindow, ConfirmationDialog,
    ImportPresetsWindow, ExportPresetsWindow, WebImportPresetsWindow, ImportTargetWindow,
    PolygonDrawWindow, GcodeSettingsWindow,
    UserGuideWindow, AboutDialog, PadNotesWindow, NestingPreviewWindow,
    JobHistoryWindow, gcode_file_note, WelcomeWindow, QuickStartWindow
)
from library_features import LibraryFeaturesMixin  # noqa: E402
from tooling_tab import ToolingTabMixin  # noqa: E402
from tuner_tab import TunerTabMixin  # noqa: E402
from toner_tab import TonerTabMixin  # noqa: E402
# The three constants moved to machine_control.py with the machine methods; the suites read them from main.
from machine_control import MachineControlMixin, MACHINE_MENU_LASER_ITEMS, MACHINE_MENU_CAMERA_ITEMS, LASER_WATCH_PERIOD_MS  # noqa: E402, F401

# ==========================================
# MAIN APP CLASS
# ==========================================

IS_MACOS = sys.platform == 'darwin'


def audio_supported(platform=None, audio_available=None):
    """Whether this build can run the Tuner and Toner at all.

    False only on a Mac without the audio stack -- the Intel build, which
    ships without sounddevice by design. There the tabs used to sit in the
    notebook saying "not available on this Mac"; since 2026-10-15 they
    are hidden and left out of Feature Set instead (Matt: "let's just not
    support the features on an intel mac that it cannot use. they are
    getting rare"). Windows and Linux keep their panel, because there a
    missing audio library is something the user can install."""
    platform = platform or sys.platform
    if audio_available is None:
        import tuner_tab
        audio_available = tuner_tab.AUDIO_AVAILABLE
    return not (platform == 'darwin' and not audio_available)


def machine_supported():
    """Whether this build can talk to a laser at all (pyserial bundled).
    False on the Intel Mac build, which has neither pyserial nor OpenCV:
    the Machine menu toggle isn't offered there, and an old setting that
    turned it on is ignored (same decision as audio_supported)."""
    try:
        import falcon_sender
        return bool(falcon_sender.HAS_PYSERIAL)
    except Exception:
        return False


def reveal_in_file_manager_command(path, platform=None):
    """The command that opens the system file manager with ``path``
    selected (Explorer, Finder), or its folder where selecting isn't
    available (Linux has no common 'select' flag across file managers).

    Windows gets a string, not a list: explorer wants ``/select,"<path>"``
    as one argument with the quotes inside it, which list-to-command-line
    quoting can't produce for a path with spaces in it."""
    platform = platform or sys.platform
    if platform == 'win32':
        return f'explorer /select,"{os.path.normpath(path)}"'
    if platform == 'darwin':
        return ['open', '-R', path]
    return ['xdg-open', os.path.dirname(path) or '.']


class PadSVGGeneratorApp(LibraryFeaturesMixin, ToolingTabMixin, TunerTabMixin, TonerTabMixin, MachineControlMixin):
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

        # The Analyzer tab is on by default (2026-11-10); the one-time migration
        # turns it on for a user whose toner was forced hidden, once. The beta
        # terms go on gating the tab's Harmonics switch, not the tab. In memory
        # only: building the app writes nothing (--selftest, --tour, the suites);
        # the flag goes to disk with the next regular save.
        migrate_analyzer_tab(self.settings)

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
            'pending': {},             # {size: qty} nested on the current scrap, not cut yet
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
        pad_file_menu.add_command(label=_("Job History..."), command=self.open_job_history)
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
            machine_menu.add_command(
                label=_("Framing Power..."),
                command=self._on_machine_framing_power)
            self._machine_menu_indices['framing_power'] = (
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
            help_menu.add_command(label=_("Welcome"), command=self.open_welcome)
            help_menu.add_command(label=_("Quick Start Guide"), command=self.open_quick_start)
            help_menu.add_command(label=_("User Guide..."), command=self.open_user_guide)
            help_menu.add_separator()
            help_menu.add_command(label=_("Open Log File"), command=self._open_log_file)
            help_menu.add_separator()
            help_menu.add_command(label=_("About"), command=self.open_about)

    def _fit_window_to_tabs(self):
        """Widen the window so every visible tab label fits, and so the
        Tooling tab's rows of section buttons are never squeezed.

        Aqua's ttk notebook clips tab labels instead of growing the tab
        row: at the 640 px default, seven tabs read "Pad Mak / Toner
        (bet" on a Mac (CI screenshot, 2026-10-06). Windows and X11 grow
        the row, so there this is a no-op in practice. Measures the labels
        with the notebook's own font rather than trusting a requested
        width that Aqua may not report.

        The Tooling tab's selector rows cannot shrink: three buttons to a
        row, 22 characters each, is about 500 px on Windows but about 720
        px in Aqua, so with fewer tabs shown a Mac or Linux window stayed
        at the 640 px it starts at and cut the buttons down (CI, 2026-11-09;
        with all seven tabs the tab row happened to widen the window
        enough). Only that tab is counted, not the notebook's whole
        request: the library and audio tabs ask for 850 px and more on
        every platform and are built to live in less.

        The Analyzer's control strip (about 910 px in English, 1040 in
        German, which pack clips from the right when the page is narrower)
        is NOT counted here: the window widens for it when that tab is
        selected and goes back when the user leaves it
        (_widen_for_analyzer / _unwiden_after_analyzer; Matt, 2026-11-10:
        "have it resize itself" rather than a wider start-up).

        The room between the window's width and a notebook page's width
        (the pack padx round the notebook, its own border and inset) is
        MEASURED on the page that is mapped, not assumed: it is the same
        for every page, and on Aqua it is more than the 24 px it is on
        Windows. With the constant, the Apple Silicon picture of c671b33
        had the window 15 px wider than the released build's and the
        third column of section buttons still running past the page's
        right edge, clipped there (CI run 42719069476). The constant is
        only the fallback for a window that has no size yet.
        """
        try:
            self.root.update_idletasks()
            style = ttk.Style(self.root)
            spec = style.lookup('TNotebook.Tab', 'font') or 'TkDefaultFont'
            font = tkfont.nametofont(spec) if spec in tkfont.names(self.root) else tkfont.Font(self.root, font=spec)
            visible = [t for t in self.notebook.tabs() if self.notebook.tab(t, 'state') != 'hidden']
            need = sum(font.measure(self.notebook.tab(t, 'text')) for t in visible)
            need += TAB_LABEL_PADDING_PX * len(visible) + 40
            tooling = getattr(self, 'tooling_tab_frame', None)
            if tooling is not None and str(tooling) in visible:
                need = max(need, tooling.winfo_reqwidth() + self._notebook_page_chrome_px())
            cur_w = self.root.winfo_width()
            if cur_w <= 1:
                cur_w = 640
            if need > cur_w:
                h = self.root.winfo_height()
                self.root.geometry(f"{int(need)}x{h if h > 1 else 720}")
        except tk.TclError:
            pass

    def _analyzer_strip_need_px(self):
        """The window width the Analyzer's control strip needs: its request,
        the strip's own padx (5 a side inside the tab's main frame) and the
        notebook's chrome; 0 while the tab has no strip (the fallback UI)."""
        strip = getattr(self, '_toner_ctrl_frame', None)
        if strip is None:
            return 0
        return int(strip.winfo_reqwidth() + 10 + self._notebook_page_chrome_px())

    def _widen_for_analyzer(self):
        """On selecting the Analyzer tab: a window narrower than its control
        strip (pack would drop the Capture button off the page) is widened
        to fit, never past the screen, height and position unchanged. When
        even the screen cannot take one row of the strip, the strip goes to
        two rows first (the right column under the rest) and the window is
        widened to what two rows need. A maximized window gets its rows
        decided the same way (it is as wide as the screen) but is never
        resized, and a window already wide enough is left as it is. What was
        widened is remembered for _unwiden_after_analyzer. A restore still
        pending from the last leave is dropped first: it must never run
        after this widening and take it for the old one (it did: CI, Windows
        and Linux, 2026-11-11)."""
        try:
            dropped = self._cancel_width_restore("superseded: the Analyzer was selected again")
            if getattr(self, '_analyzer_widened', None):
                return
            screen_w = self.root.winfo_screenwidth()
            # the decision is by what ONE row needs, whatever the strip stands at now
            self._toner_strip_two_rows(False)
            self.root.update_idletasks()
            one_row_need = self._analyzer_strip_need_px()
            if one_row_need > screen_w:
                self._toner_strip_two_rows(True)
                self.root.update_idletasks()
            if self.root.state() == 'zoomed':
                return
            need = self._analyzer_strip_need_px()
            cur_w, h = self.root.winfo_width(), self.root.winfo_height()
            if dropped:
                cur_w = dropped["before"]   # the restore to it is on its way, whatever the window reads now
            if cur_w <= 1 or h <= 1 or need <= cur_w:
                return
            new_w = min(need, screen_w)
            if new_w <= cur_w:
                return
            self.root.geometry(f"{new_w}x{h}")
            self._analyzer_widened = (cur_w, new_w)
        except tk.TclError:
            pass

    def _unwiden_after_analyzer(self):
        """On leaving the Analyzer tab: the width goes back to what it was
        before _widen_for_analyzer, but only if the app did the widening and
        the window is still exactly that wide (a size the user chose since,
        or a maximize, is theirs and stays). Height and position unchanged.
        The geometry is asked for, checked after idle and asked again while
        the window still reads the widened width (Aqua applies a geometry
        late, or drops one); what happened is kept in `_analyzer_restore`."""
        record = getattr(self, '_analyzer_widened', None)
        self._analyzer_widened = None
        if not record:
            return
        before, widened_to = record
        self._cancel_width_restore("superseded: left again")
        try:
            state = self.root.state()
            self.root.update_idletasks()
            cur = self.root.winfo_width()
            self._analyzer_restore = {"before": before, "widened_to": widened_to, "found": cur, "state": state,
                                      "screen": (self.root.winfo_screenwidth(), self.root.winfo_screenheight()),
                                      "attempts": 0, "outcome": "pending"}
            if state == 'zoomed':
                self._analyzer_restore["outcome"] = "left alone: maximized"
                return
            if cur != widened_to:
                self._analyzer_restore["outcome"] = "left alone: resized by the user"
                return
            self._restore_width(before, widened_to, tries=8)
        except tk.TclError:
            pass

    def _restore_width(self, width, from_width, tries):
        """geometry(width) now, then after idle: done if the window reads
        `width`; left alone if it reads something else (the user's); asked
        again, up to `tries` times 100 ms apart, while it still reads
        `from_width` (the geometry not yet applied, or dropped)."""
        info = self._analyzer_restore
        try:
            info["attempts"] += 1
            self.root.geometry(f"{width}x{self.root.winfo_height()}")
        except tk.TclError:
            return

        def check():
            self._analyzer_restore_after = None
            if self._analyzer_restore is not info or info["outcome"] != "pending":
                return          # a newer restore owns the window now, or this one was closed
            try:
                self.root.update_idletasks()
                w = self.root.winfo_width()
            except tk.TclError:
                return
            if w == width:
                info["outcome"] = "restored"
            elif w != from_width:
                info["outcome"] = f"left at {w}: not the widened width any more"
            elif tries > 1:
                self._restore_width(width, from_width, tries - 1)
            else:
                info["outcome"] = f"gave up at {w} after {info['attempts']} attempts"
        try:
            self._analyzer_restore_after = self.root.after(100, check)
        except tk.TclError:
            pass

    def _cancel_width_restore(self, reason):
        """Drop the restore check still pending from the last leave, if any,
        and close its record with `reason`; returns the record it closed, or
        None. The check is scheduled 100 ms after the leave; a select within
        that time (a test does it, a user can) would otherwise have it read
        the new widening as the old one not yet undone and take it back (CI,
        Windows and Linux, 2026-11-11: the window went back to its start
        width with the tab selected). The record's `before` is the width the
        window is on its way to, which is what a new widening starts from
        (on Aqua the window may still read the old width)."""
        after_id = getattr(self, '_analyzer_restore_after', None)
        self._analyzer_restore_after = None
        if after_id is not None:
            try:
                self.root.after_cancel(after_id)
            except (tk.TclError, ValueError):
                pass
        info = getattr(self, '_analyzer_restore', None)
        if info is not None and info.get("outcome") == "pending":
            info["outcome"] = reason
            return info                  # the dropped restore: its `before` is where the window is going
        return None

    def _notebook_page_chrome_px(self):
        """How much wider the window is than a notebook page: the window's
        width less the mapped page's, measured once on whichever page is
        selected (the inset is the same for every page) and kept, since
        the notebook's padding and border never change while the app runs.
        Measured once, not every time: right after a geometry() call the
        window's width is already the new one while the page is still laid
        out for the old, and the difference is meaningless (found on
        Windows with Aqua's chrome planted). TOOLING_CONTENT_MARGIN_PX
        while there is no consistent pair to measure (the window not
        mapped: both read 1) and whenever the reading is not credible."""
        cached = getattr(self, '_page_chrome_px', None)
        if cached is not None:
            return cached
        try:
            page = self.notebook.nametowidget(self.notebook.select())
            win_w, page_w = self.root.winfo_width(), page.winfo_width()
        except (tk.TclError, KeyError):
            return TOOLING_CONTENT_MARGIN_PX
        chrome = win_w - page_w
        if win_w > 1 and page_w > 1 and 0 < chrome <= PAGE_CHROME_MAX_PX:
            self._page_chrome_px = chrome
            return chrome
        return TOOLING_CONTENT_MARGIN_PX

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
            self._unwiden_after_analyzer()
            self._toner_stop()
            self._tuner_start()
        elif selected == str(self.toner_tab_frame):
            self._tuner_stop()
            self._widen_for_analyzer()
            self._toner_start()
        else:
            self._unwiden_after_analyzer()
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

        # --- Create Tab 7: Analyzer (the Toner until 2026-11-10; the attribute and keys keep the name) ---
        self.toner_tab_frame = ttk.Frame(self.notebook, style='App.TFrame')
        self.notebook.add(self.toner_tab_frame, text=_('Analyzer'))
        self.create_toner_tab(self.toner_tab_frame)

        self.notebook.pack(expand=True, fill="both", padx=5, pady=5)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        # After the Feature Set hides have been applied (end of __init__).
        self.root.after_idle(self._fit_window_to_tabs)

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
            self.tuner_tab_frame: visible.get("Tuner", True) and audio_supported(),
            self.toner_tab_frame: visible.get("Toner", True) and audio_supported(),
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
        self._mat_frame = mat_frame
        material_list = list(self.material_vars.items())
        for i, (m, var) in enumerate(material_list):
            row, col = i // 2, i % 2
            cb = tk.Checkbutton(mat_frame, text=_(m.replace('_', ' ').capitalize()),
                               variable=var, bg=self.root.cget('bg'))
            cb.grid(row=row, column=col, sticky='w', padx=(0, 10))
            self.material_checkboxes[m] = cb
        # Adhesive rings join the list only while the feature is on.
        self._refresh_adhesive_ring_ui()

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

        # The shape dialog's overlay and its "Get from camera" open the
        # camera; the app's own live preview must not be holding it.
        self._release_live_camera()
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

    def _show_scrap_continue_dialog(self, placed_count, scrap_num, remaining_count,
                                    gcode_file=False):
        """Show scrap continue dialog. If a polygon is loaded, offer to unload it.

        gcode_file: this scrap was just written as a G-code file, so the
        message ends with the shared "send this file to the laser as it
        is" line (ui_dialogs.gcode_file_note). Not for an SVG scrap, and
        not for Frame & Cut, which writes no file."""
        msg = _("Placed {placed_count} pads on scrap #{scrap_num}.\n\n"
                "{remaining_count} pads remaining.\n"
                "Adjust dimensions and click Generate again.").format(
                    placed_count=placed_count, scrap_num=scrap_num,
                    remaining_count=remaining_count)
        if gcode_file:
            msg += "\n\n" + gcode_file_note()

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
            'done': {},
            'pending': {},
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
                    pending = self._count_pending_pads()
                    if pending:
                        self.scrap_status_var.set(
                            _("{remaining} left, {pending} on this scrap ({count} scraps)").format(
                                remaining=remaining - pending, pending=pending, count=count))
                    else:
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
            # Per-size tally of pads already cut. The pad list itself is
            # live (re-read every Generate); this is the only thing the
            # session has to remember about it.
            'done': {},
            'pending': {},
        }
        self._lock_material_selection(material)
        self._open_remaining_pads_window()

    def _scrap_pads_for_this_scrap(self, pads):
        """Reconcile the pad list in the box with what this session has
        already cut, and return the list to nest on this scrap.

        The pad list stays live for the whole session. The session only
        remembers, per size, how many pads are already cut ('done'), so
        remaining is "what the box says now" minus that tally -- a size
        added, raised, lowered or removed mid-session simply shows up on
        the next scrap, and lowering below the cut count clamps to zero.
        Until 2026-09-20 the session copied the list at its start and
        ignored the box until it ended; a user who added sizes to fill a
        half-empty last scrap got a scrap without them and no message.

        Duplicate lines for one size are merged. A 'max' entry keeps its
        existing behaviour: nested on the first scrap only.
        """
        session = self.scrap_session
        done = session.setdefault('done', {})
        wanted = {}
        for p in pads:
            if p['qty'] != 'max':
                wanted[p['size']] = wanted.get(p['size'], 0) + p['qty']
        session['original_pads'] = [{'size': s, 'qty': q} for s, q in wanted.items()]
        remaining = [{'size': s, 'qty': q - done.get(s, 0)}
                     for s, q in wanted.items() if q - done.get(s, 0) > 0]
        session['remaining_pads'] = remaining
        if session['scrap_count'] == 0:
            return remaining + [p for p in pads if p['qty'] == 'max']
        return remaining

    def _scrap_record_result(self, remaining):
        """Credit a finished scrap to the session: add what it placed to
        the per-size 'done' tally and store the new remaining list.

        The credit is the drop from the pre-scrap list (still sitting in
        remaining_pads) to `remaining`, so this has to run before anything
        else overwrites remaining_pads. All three commit sites (SVG,
        G-code, Frame & Cut) go through here.
        """
        session = self.scrap_session
        done = session.setdefault('done', {})
        after = {p['size']: p['qty'] for p in remaining}
        for p in session['remaining_pads']:
            cut = p['qty'] - after.get(p['size'], 0)
            if cut > 0:
                done[p['size']] = done.get(p['size'], 0) + cut
        session['remaining_pads'] = remaining
        session['pending'] = {}

    @staticmethod
    def _count_by_size(placed):
        """{pad_size: count} for a list of placed (pad_size, cx, cy, r)."""
        counts = {}
        for pad_size, _cx, _cy, _r in placed:
            counts[pad_size] = counts.get(pad_size, 0) + 1
        return counts

    def _count_pending_pads(self):
        """Pads nested on the current scrap and not yet cut."""
        return sum((self.scrap_session.get('pending') or {}).values())

    def _scrap_set_pending(self, placed):
        """Record what the current scrap's nest placed, before anything is
        cut or written. The progress window shows it in its "On this scrap"
        column and takes it off Remaining from this moment, so the preview
        and the running cut no longer read as "0 done, everything remaining"
        -- which looked exactly like pads being dropped (Matt, 2026-10-26).
        Cleared by _scrap_record_result (commit) or _scrap_abandon_pending.
        """
        self.scrap_session['pending'] = self._count_by_size(placed)
        self._update_scrap_status_display()
        self._update_remaining_pads_window()

    def _scrap_abandon_pending(self):
        """The nested scrap never became a cut: preview backed out, jog or
        framing dialog cancelled, cut stopped or errored, file not written.
        Put the pads back in Remaining. No-op when nothing is pending, so
        the file-export and Frame & Cut paths call it from a finally.
        """
        if self.scrap_session.get('pending'):
            self.scrap_session['pending'] = {}
            self._update_scrap_status_display()
            self._update_remaining_pads_window()

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

        # The box is live: nest what it says now, minus what's already cut.
        pads = self._scrap_pads_for_this_scrap(pads)

        if not pads:
            messagebox.showinfo(_("Session Complete"), _("All pads have been placed!"))
            return None

        placed, remaining, any_placed = try_nest_partial(
            pads, material, mat_w, mat_h, self.settings,
            polygon=mat_polygon)

        if not any_placed:
            min_pad_size = min(p['size'] for p in pads)
            messagebox.showwarning(_("No Pads Fit"),
                _("No pads could be placed on this scrap.\n\n"
                "Smallest remaining pad: {min_pad_size}mm\n"
                "Try a larger scrap piece.").format(min_pad_size=min_pad_size))
            return None

        self._scrap_set_pending(placed)
        return placed, remaining

    def _open_remaining_pads_window(self):
        """Open or update the floating window showing remaining and done pads."""
        if self.scrap_remaining_window is not None:
            try:
                self.scrap_remaining_window.destroy()
            except tk.TclError:
                pass

        self.scrap_remaining_window = tk.Toplevel(self.root)
        self.scrap_remaining_window.title(_("Scrap Mode Progress"))
        self.scrap_remaining_window.geometry("480x300")
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

        # Three columns: Remaining (not yet placed), On this scrap (nested,
        # waiting for the cut), Done (cut).
        columns_frame = tk.Frame(self.scrap_remaining_window, bg=theme_bg)
        columns_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Remaining column
        remaining_frame = tk.Frame(columns_frame, bg=theme_bg)
        remaining_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        tk.Label(remaining_frame, text=_("Remaining"), bg=theme_bg,
                 font=("Helvetica", 9, "bold"), fg="blue").pack(anchor="w")
        self.scrap_remaining_listbox = tk.Listbox(remaining_frame, font=("Courier", 10), height=10, width=12)
        self.scrap_remaining_listbox.pack(fill="both", expand=True)

        # On-this-scrap column
        pending_frame = tk.Frame(columns_frame, bg=theme_bg)
        pending_frame.pack(side="left", fill="both", expand=True, padx=5)
        tk.Label(pending_frame, text=_("On this scrap"), bg=theme_bg,
                 font=("Helvetica", 9, "bold"), fg="#B35900").pack(anchor="w")
        self.scrap_pending_listbox = tk.Listbox(pending_frame, font=("Courier", 10), height=10, width=12)
        self.scrap_pending_listbox.pack(fill="both", expand=True)

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

            # Done comes from the session's own tally, not original minus
            # remaining: the list is live, so a size lowered below what's
            # already cut would otherwise under-report.
            done_pads = [{'size': s, 'qty': q}
                         for s, q in self.scrap_session.get('done', {}).items() if q > 0]

            # Pads nested on the current scrap but not cut yet sit in the
            # middle column and come off Remaining; a backed-out preview
            # or a stopped cut puts them straight back (see
            # _scrap_set_pending / _scrap_abandon_pending).
            pending = self.scrap_session.get('pending') or {}
            pending_pads = [{'size': s, 'qty': q} for s, q in pending.items() if q > 0]
            if pending:
                remaining_pads = [{'size': p['size'], 'qty': p['qty'] - pending.get(p['size'], 0)}
                                  for p in remaining_pads]
                remaining_pads = [p for p in remaining_pads if p['qty'] > 0]

            total_remaining = sum(p['qty'] for p in remaining_pads)
            total_pending = sum(p['qty'] for p in pending_pads)
            total_done = sum(p['qty'] for p in done_pads)
            total_original = sum(p['qty'] for p in original_pads)

            # Update header
            if total_remaining == 0 and total_pending == 0:
                self.scrap_window_header.config(text=_("All done!"), fg="green")
            elif total_pending:
                self.scrap_window_header.config(
                    text=_("{placed} on scrap #{num}, {done} / {total} cut").format(
                        placed=total_pending, num=scraps + 1, done=total_done, total=total_original),
                    fg="#B35900")
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

            # Update On-this-scrap listbox
            self.scrap_pending_listbox.delete(0, tk.END)
            if pending_pads:
                for pad in sorted(pending_pads, key=lambda p: -p['size']):
                    size_str = f"{pad['size']:.1f}".rstrip('0').rstrip('.')
                    self.scrap_pending_listbox.insert(tk.END, f" {pad['qty']} x {size_str}")
            else:
                self.scrap_pending_listbox.insert(tk.END, _(" (none)"))

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

    def _adhesive_ring_enabled(self):
        """True when the opt-in adhesive ring material is switched on
        (Sizing Rules > Lesser-used settings)."""
        return bool(self.settings.get("adhesive_ring_enabled", False))

    def _refresh_adhesive_ring_ui(self):
        """Add or remove the Adhesive rings material to match the setting.

        The material is not merely hidden while the feature is off: its
        entry leaves self.material_vars altogether, so every flow that
        walks the checked materials (Generate, Frame & Cut, scrap mode,
        job reload) can't pick it up from a checkbox nobody can see.
        Called when the tab is built and whenever Sizing Rules is applied.
        """
        enabled = self._adhesive_ring_enabled()
        present = 'adhesive_ring' in self.material_vars
        session = getattr(self, 'scrap_session', None) or {}
        if enabled and not present:
            var = tk.BooleanVar(value=False)
            i = len(self.material_checkboxes)
            cb = tk.Checkbutton(self._mat_frame, text=_("Adhesive rings"),
                                variable=var, bg=self._mat_frame.cget('bg'))
            cb.grid(row=i // 2, column=i % 2, sticky='w', padx=(0, 10))
            if session.get('active'):
                cb.config(state='disabled')   # materials are locked for the session
            self.material_vars['adhesive_ring'] = var
            self.material_checkboxes['adhesive_ring'] = cb
        elif present and not enabled:
            # A scrap session cutting rings can't outlive the material.
            if session.get('active') and session.get('material') == 'adhesive_ring':
                self._clear_scrap_session()
                self.scrap_mode_var.set(False)
                self._update_scrap_status_display()
            self.material_checkboxes.pop('adhesive_ring').destroy()
            del self.material_vars['adhesive_ring']

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

        # A labeled group is a fixed grid, so it needs a fixed count —
        # "max" has none. Refuse rather than silently cutting that size
        # loose and unlabeled, which is the exact confusion zones exist
        # to prevent. Filling the leftover space with an ungrouped size
        # still works; it's only a max on a GROUPED size that can't.
        # Adhesive rings never group (nest_with_zones), so a job that is
        # only rings has nothing to clash with.
        zoned_materials = [m for m, v in self.material_vars.items()
                           if v.get() and m != 'adhesive_ring']
        if self.settings.get("zone_labels_enabled", False) and max_pads and zoned_materials:
            lo = float(self.settings.get("zone_label_min_size", 7.0))
            hi = float(self.settings.get("zone_label_max_size", 12.5))
            clash = [p['size'] for p in max_pads if lo <= p['size'] <= hi]
            if clash:
                size = f"{clash[0]:g}"
                messagebox.showerror(
                    _("Can't group a 'max' size"),
                    _("{size} mm is set to 'max', but Labeled Zones is "
                      "grouping {lo}–{hi} mm. A group needs a fixed "
                      "count to lay out.\n\n"
                      "Give {size} a quantity, or use 'max' on a size "
                      "outside that range to fill the rest.").format(
                          size=size, lo=f"{lo:g}", hi=f"{hi:g}"))
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

    def _record_job(self, output, materials, pads, params, placed_count=None,
                    scrap_num=None, save_dir=None, status="complete"):
        """Log a job that reached an output stage to the history file.

        `output` is "svg", "gcode", or "laser". Call this only once real
        output exists — files on disk, or G-code streamed to the machine.

        Swallows every error on purpose: the history is a convenience
        log, and nothing downstream reads it back, so a bad write must
        never turn a successful cut into an error dialog.
        """
        try:
            requested = sum(p['qty'] for p in pads if isinstance(p.get('qty'), int))
            has_max = any(p.get('qty') == 'max' for p in pads)

            entry = {
                "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
                "output": output,
                "status": status,
                "materials": list(materials),
                # Raw text as typed — this is what gets restored, so the
                # user gets their list back exactly (including "x max").
                "pad_text": self.pad_entry.get("1.0", tk.END).strip(),
                "pads": [{"size": p.get('size'), "qty": p.get('qty')} for p in pads],
                "requested_count": requested,
                "has_max": has_max,
                # Discs actually placed on the sheet. Differs from
                # requested when a "max" line filled the leftover space,
                # and is the per-scrap subset in scrap mode.
                "placed_count": placed_count,
                "base": params.get('base'),
                "units": self.settings.get('units'),
                "sheet_w": self.width_entry.get().strip(),
                "sheet_h": self.height_entry.get().strip(),
                "sheet_w_mm": params.get('width_mm'),
                "sheet_h_mm": params.get('height_mm'),
                "card_paper": bool(params.get('card_paper_dims')),
                "hole_option": self.hole_var.get(),
                "custom_hole": self.custom_hole_entry.get().strip(),
                "scrap_num": scrap_num,
                "save_dir": save_dir,
                "preset_library": self.pad_preset_loaded_library,
                "preset_name": self.pad_preset_loaded_name,
                # Recorded for display only. The shape itself isn't
                # restorable in any meaningful way — a camera-captured
                # polygon is anchored to a scrap that's now been cut up.
                "polygon_vertices": len(self.custom_polygon) if self.custom_polygon else 0,
            }
            append_job_history(entry)
        except Exception:
            logging.getLogger(__name__).exception("Could not record job history")

    def _record_tooling_job(self, job, output, status="complete", save_dir=None):
        """Log a Tooling job that reached an output stage (Matt,
        2026-11-09: "tooling runs go into job history if easy").

        `job` carries the fields ToolingTabMixin._tooling_history builds
        (tool, description, material, count, sheet in mm, base, scrap
        number); `output` is "svg", "gcode", "stl" or "laser". The entry
        has the same shape as _record_job's so the storage, the list and
        delete work untouched, with `tool` added and the description where
        the pad list goes. The dialog shows the tool's name on these rows
        and does not load them into the Pad Maker's form. Swallows every
        error, for the same reason _record_job does.
        """
        try:
            count = job.get('count')
            entry = {
                "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
                "output": output,
                "status": status,
                "materials": [job['material']] if job.get('material') else [],
                "tool": job['tool'],
                "pad_text": job.get('description') or "",
                "pads": [],
                "requested_count": count,
                "has_max": False,
                "placed_count": count,
                "base": job.get('base'),
                "units": None,
                "sheet_w": None,
                "sheet_h": None,
                "sheet_w_mm": job.get('sheet_w_mm'),
                "sheet_h_mm": job.get('sheet_h_mm'),
                "card_paper": False,
                "hole_option": None,
                "custom_hole": "",
                "scrap_num": job.get('scrap_num'),
                "save_dir": save_dir,
                "preset_library": None,
                "preset_name": None,
                "polygon_vertices": 0,
            }
            append_job_history(entry)
        except Exception:
            logging.getLogger(__name__).exception("Could not record job history")

    def open_job_history(self):
        """File > Job History — browse past jobs, optionally reload one."""
        dlg = JobHistoryWindow(self.root, load_job_history(), save_job_history)
        if dlg.result:
            self._load_job_into_form(dlg.result)

    def _load_job_into_form(self, job):
        """Restore a history entry's inputs into the Pad Maker form.

        Restores what the user typed: pad list, materials, sheet size,
        center hole, base filename. Deliberately does NOT touch sizing
        rules, G-code settings, or the custom polygon — those are
        current-state settings the user may have tuned since, and a
        camera-captured polygon is anchored to a scrap that no longer
        exists.
        """
        # A Tooling entry has no pad list to put back; the dialog refuses
        # it before it gets here, this is the backstop.
        if job.get("tool"):
            return
        # A live scrap session owns its own pad list and locks the
        # material checkboxes; loading over the top would desync them.
        if self.scrap_session.get('active'):
            messagebox.showwarning(
                _("Scrap Session Active"),
                _("Finish or clear the current scrap session before "
                  "loading a job from history."))
            return

        self.pad_entry.delete("1.0", tk.END)
        self.pad_entry.insert(tk.END, job.get("pad_text", ""))
        self._auto_resize_pad_entry()

        # Loading a job is not loading a preset — drop the preset link
        # so the Notes button doesn't point at an unrelated preset.
        self.pad_preset_loaded_library = None
        self.pad_preset_loaded_name = None
        self.pad_notes_btn.config(state="disabled")

        mats = job.get("materials") or []
        if mats:
            for m, var in self.material_vars.items():
                var.set(m in mats)

        # Sheet size is stored in mm as well as as-typed. Restore the
        # typed text when the units still match, otherwise convert —
        # units are a Sizing Rules setting and loading a job must not
        # silently change them.
        units = self.settings.get('units', 'in')
        w_txt, h_txt = None, None
        if job.get("units") == units and job.get("sheet_w") and job.get("sheet_h"):
            w_txt, h_txt = job["sheet_w"], job["sheet_h"]
        else:
            w_mm, h_mm = job.get("sheet_w_mm"), job.get("sheet_h_mm")
            if isinstance(w_mm, (int, float)) and isinstance(h_mm, (int, float)):
                per_unit = {'in': 25.4, 'cm': 10.0, 'mm': 1.0}.get(units, 25.4)
                w_txt = f"{w_mm / per_unit:g}"
                h_txt = f"{h_mm / per_unit:g}"
        if w_txt is not None:
            self.width_entry.delete(0, tk.END)
            self.width_entry.insert(0, w_txt)
            self.height_entry.delete(0, tk.END)
            self.height_entry.insert(0, h_txt)

        if job.get("hole_option"):
            self.hole_var.set(job["hole_option"])
        if job.get("custom_hole"):
            self.custom_hole_entry.config(state='normal')
            self.custom_hole_entry.delete(0, tk.END)
            self.custom_hole_entry.insert(0, job["custom_hole"])
        self.toggle_custom_hole_entry()

        if job.get("base"):
            self.filename_entry.delete(0, tk.END)
            self.filename_entry.insert(0, job["base"])

        # The one thing that can't come back — say so rather than let
        # the user assume the scrap shape came along with the numbers.
        if job.get("polygon_vertices"):
            messagebox.showinfo(
                _("Job Loaded"),
                _("Pad list, materials and sheet size are back.\n\n"
                  "The custom shape from that job wasn't restored — "
                  "draw or capture the piece you're cutting now."))

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
            generated = []      # materials actually written, for the history log
            last_placed = 0

            # Process each material (with optional per-material preview)
            for material in selected_materials:
                mat_w, mat_h, mat_polygon = self._get_material_dimensions(material, width_mm, height_mm, card_paper_dims)

                # Nest once (may re-run if user adjusts and retries). The same
                # pass reports whether every fixed-quantity pad landed, so
                # there is no second nest just to validate the fit.
                placed, zones, fixed_placed, fixed_total = nest_with_zones(
                    pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon)

                # Validate fit
                if fixed_placed != fixed_total:
                    size_desc = _("paper") if (material == "card" and card_paper_dims) else _("sheet")
                    messagebox.showerror(_("Nesting Error"), _("Could not fit all '{material}' pieces on the specified {size_desc} size.").format(material=material.replace('_', ' '), size_desc=size_desc))
                    return

                # Preview this material
                if use_preview:
                    preview = NestingPreviewWindow(
                        self.root, {material: placed}, mat_w, mat_h,
                        polygon=mat_polygon, zones=zones, settings=self.settings)
                    if preview.result != "save":
                        return  # User clicked Adjust — go back for this material

                # Ask for save directory once (on first material)
                if save_dir is None:
                    save_dir = filedialog.askdirectory(title=_("Select Folder to Save SVGs"), initialdir=self.settings.get("last_output_dir", ""))
                    if not save_dir:
                        return
                    self.settings["last_output_dir"] = save_dir

                filename = os.path.join(save_dir, f"{base}_{material}.svg")
                generate_svg_from_placed(placed, material, mat_w, mat_h, filename, hole_dia, self.settings,
                                         polygon=mat_polygon, zones=zones)
                generated.append(material)
                last_placed = len(placed)

            self._record_job("svg", generated, pads, params,
                             placed_count=last_placed, save_dir=save_dir)
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
                hole_dia = self.scrap_session['hole_dia']

            # The box is live: nest what it says now, minus what's already cut.
            pads = self._scrap_pads_for_this_scrap(pads)

            if not pads:
                messagebox.showinfo(_("Session Complete"), _("All pads have been placed!"))
                return

            # Attempt partial placement
            placed, remaining, any_placed = try_nest_partial(
                pads, material, mat_w, mat_h, self.settings,
                polygon=mat_polygon)

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
                    polygon=mat_polygon, settings=self.settings)
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

            self._record_job("svg", [material], pads, params,
                             placed_count=len(placed), scrap_num=scrap_num,
                             save_dir=save_dir)

            # Credit this scrap to the session's done tally
            self._scrap_record_result(remaining)
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
        finally:
            # A nest that never became a file is not "on this scrap".
            self._scrap_abandon_pending()

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

                # One nest; its fixed counts are the fit check.
                placed, zones, fixed_placed, fixed_total = nest_with_zones(
                    pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon)

                if fixed_placed != fixed_total:
                    size_desc = _("paper") if (material == "card" and card_paper_dims) else _("sheet")
                    messagebox.showerror(_("Nesting Error"), _("Could not fit all '{material}' pieces on the specified {size_desc} size.").format(material=material.replace('_', ' '), size_desc=size_desc))
                    return

                if use_preview:
                    preview = NestingPreviewWindow(
                        self.root, {material: placed}, mat_w, mat_h,
                        polygon=mat_polygon, zones=zones, settings=self.settings)
                    if preview.result != "save":
                        return

                all_placed[material] = (placed, mat_w, mat_h, mat_polygon, zones)

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

            last_placed = 0
            try:
                for material, (placed, mat_w, mat_h, mat_polygon, zones) in all_placed.items():
                    filename = os.path.join(save_dir, f"{base}_{material}.gcode")
                    generate_gcode_from_placed(placed, material, mat_w, mat_h, filename, hole_dia, self.settings,
                                               polygon=mat_polygon, zones=zones)
                    last_placed = len(placed)
            finally:
                working_popup.destroy()

            self._record_job("gcode", list(all_placed.keys()), pads, params,
                             placed_count=last_placed, save_dir=save_dir)
            save_settings(self.settings)

            # Auto-eject SD card if checkbox is checked and destination is removable
            if self.eject_sd_var.get() and self._is_removable_drive(save_dir):
                drive_letter = os.path.splitdrive(os.path.abspath(save_dir))[0]
                eject_success = self._eject_drive(drive_letter) if drive_letter else False
                if eject_success:
                    messagebox.showinfo(_("Done"), _("G-code files generated successfully.\n\nSD card safely ejected — you can remove it now!")
                                        + "\n\n" + gcode_file_note())
                else:
                    messagebox.showinfo(_("Done"), _("G-code files generated successfully.\n\nPlease safely eject the SD card before removing.")
                                        + "\n\n" + gcode_file_note())
            else:
                messagebox.showinfo(_("Done"), _("G-code files generated successfully.")
                                    + "\n\n" + gcode_file_note())

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
                    polygon=mat_polygon, settings=self.settings)
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

            self._record_job("gcode", [material], pads, params,
                             placed_count=len(placed), scrap_num=scrap_num,
                             save_dir=save_dir)

            # Credit this scrap to the session's done tally
            self._scrap_record_result(remaining)
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
                    "Files saved to: {save_dir}").format(placed_count=placed_count, scrap_num=scrap_num, save_dir=save_dir)
                    + "\n\n" + gcode_file_note())
            else:
                # The flag goes by position: several suites stand in for this
                # dialog with a `lambda *args`, which takes no keywords.
                self._show_scrap_continue_dialog(placed_count, scrap_num, remaining_count, True)

        except Exception as e:
            print(f"An error occurred during scrap mode G-code generation: {e}")
            messagebox.showerror(_("An Error Occurred"), _("Something went wrong:\n\n{error}").format(error=e))
        finally:
            # A nest that never became a file is not "on this scrap".
            self._scrap_abandon_pending()

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
                    # inf, 1e400 and nan parse as floats but are no size and
                    # no quantity (2026-11-10): int(inf) raised OverflowError
                    # past this handler and reached the Unexpected Error box.
                    if size > 0 and math.isfinite(size):
                        qty_str = parts[1].strip()
                        if qty_str == 'max':
                            pad_list.append({'size': size, 'qty': 'max'})
                            ok = True
                        else:
                            qty_f = float(qty_str)
                            if not math.isfinite(qty_f):
                                raise ValueError(qty_str)
                            qty = int(qty_f)
                            if qty > 0:
                                pad_list.append({'size': size, 'qty': qty})
                                ok = True
            except (ValueError, OverflowError):
                pass
            if not ok and rejected is not None:
                rejected.append(line)
        return pad_list

    # --- Pad Presets Wrappers ---

    def on_pad_library_selected(self, event=None):
        lib_name = self.pad_library_var.get()
        preset_list = []
        if lib_name == _("All Libraries"):
            for library, presets in sorted(self.pad_presets.items()):
                for name in sorted(presets.keys()):
                    preset_list.append(f"[{library}] {name}")
        else:
            preset_list = sorted(self.pad_presets.get(lib_name, {}).keys())

        self.pad_preset_menu['values'] = preset_list
        self.pad_preset_menu.set(_("Load Pad Preset"))

        # Toggle delete button: "Delete Library" when library is empty
        if (lib_name != _("All Libraries")
                and lib_name in self.pad_presets
                and not self.pad_presets[lib_name]):
            self.pad_delete_btn.config(text=_("Delete Library"))
        else:
            self.pad_delete_btn.config(text=_("Delete Preset"))

        # Remember last used library
        self.settings["last_pad_library"] = lib_name

    def update_pad_library_dropdown(self):
        lib_names = [_("All Libraries")] + sorted(self.pad_presets.keys())
        self.pad_library_dropdown['values'] = lib_names
        # Default to My Presets if it exists, otherwise All Libraries
        if "My Presets" in lib_names:
            self.pad_library_var.set("My Presets")
        else:
            self.pad_library_var.set(_("All Libraries"))
        self.on_pad_library_selected()

    def _get_pad_preset_data(self, raw):
        """Extract pads text and notes from a preset entry (handles both old string and new dict formats)."""
        if isinstance(raw, dict):
            return raw.get("pads", ""), raw.get("notes", "")
        return raw, ""

    def on_save_pad_preset(self):
        active_library = self.pad_library_var.get()
        if not active_library or active_library == _("All Libraries"):
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

        if lib_name == _("All Libraries"):
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
        if (selected_lib != _("All Libraries")
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

        if selected_lib == _("All Libraries"):
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
        # Adhesive rings get their section only while the feature is on.
        if self._adhesive_ring_enabled():
            pad_materials.append(("adhesive_ring", _("Adhesive rings")))
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
            ("Toner", _("Analyzer")),    # the storage key stays "Toner" (2026-11-10)
        ]
        if not audio_supported():
            # A build that can't run them doesn't offer them (Intel Mac).
            main_tabs = [t for t in main_tabs if t[0] not in ("Tuner", "Toner")]

        check_vars = {}
        for key, display in main_tabs:
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
        if machine_supported():  # not offered on a build without pyserial
            tk.Label(frame, text=_("\nExperimental / In Progress"), bg=bg,
                     font=("Helvetica", 11, "bold")).pack(anchor="w")
            tk.Checkbutton(
                frame,
                text=_("Direct machine control (GRBL)"),
                variable=machine_menu_var, bg=bg,
                font=("Helvetica", 10)).pack(anchor="w")

        def apply():
            # Merge, don't replace: tabs this build doesn't offer (Tuner and
            # Toner on the Intel Mac) keep whatever was saved for them.
            new_visible = dict(self.settings.get("visible_tabs", {}))
            new_visible.update({name: var.get() for name, var in check_vars.items()})

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

    def open_user_guide_search(self):
        """The whole User Guide with the cursor in its search box: where
        the Welcome window's "Search the User Guide" link goes."""
        guide = UserGuideWindow(self.root)
        guide.focus_search()
        return guide

    def open_quick_start(self):
        """Help > Quick Start Guide, and the Welcome window's link. Its last
        line opens the whole User Guide."""
        return QuickStartWindow(self.root, show_frame_cut=machine_supported(),
                                on_user_guide=lambda: UserGuideWindow(self.root))

    def open_welcome(self):
        """Help > Welcome, and the once-per-version showing below. One at
        a time: a second request brings the open one forward."""
        from config import APP_VERSION
        existing = getattr(self, '_welcome_window', None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.lift()
                    return existing
            except tk.TclError:
                pass
        self._welcome_window = WelcomeWindow(
            self.root, APP_VERSION,
            on_search_guide=self.open_user_guide_search,
            on_quick_start=self.open_quick_start)
        return self._welcome_window

    def maybe_show_welcome(self):
        """Show the Welcome window if this version has not been welcomed
        on this install yet: the first launch after a new install or an
        update. Returns the window, or None when there was nothing to show.

        Called from the real launch path only (the __main__ block, after
        the main window is up, on an after() so it costs startup nothing).
        Building the app does not call it, so --selftest, --tour and every
        test suite that constructs PadSVGGeneratorApp never see a Welcome
        window they did not ask for. The version is recorded when the
        window is shown, not when it is closed: a crash or a force-quit
        must not bring it back every launch.
        """
        from config import APP_VERSION
        if self.settings.get("welcomed_version") == APP_VERSION:
            return None
        self.settings["welcomed_version"] = APP_VERSION
        save_settings(self.settings)
        return self.open_welcome()

    def _open_log_file(self):
        """Show the log file selected in the system file manager.

        Opening it in a text editor (before v2.86) showed the contents but
        left the user with no idea where the file was, and what we need
        from a user is the file itself, attached to an email, often with
        last_calibration_frame.png and last_frame_cut.gcode beside it."""
        log_path = get_log_file()
        if log_path and os.path.exists(log_path):
            try:
                subprocess.Popen(reveal_in_file_manager_command(log_path))
            except Exception:
                logging.getLogger(__name__).exception(
                    "could not show the log file in the file manager")
                if sys.platform == 'win32':
                    os.startfile(log_path)
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
        # Sizing Rules can switch the adhesive ring material on or off.
        self._refresh_adhesive_ring_ui()

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

# Per-tab horizontal padding Aqua adds round a notebook label, plus slack.
TAB_LABEL_PADDING_PX = 30
# Room between the window and a notebook page (the notebook's padding and
# border), used by _fit_window_to_tabs ONLY as the fallback when the window
# has no size to measure yet. 14 px on Windows, more on Aqua, which is why
# it is measured (_notebook_page_chrome_px) and not assumed; the 24 keeps a
# little slack for the moment before the first measurement.
TOOLING_CONTENT_MARGIN_PX = 24
# A measured chrome above this is a window and a page read mid-resize, not
# a notebook's padding: it is not kept.
PAGE_CHROME_MAX_PX = 200


def run_tour(root, app, shots_dir=None, step_ms=1500, on_done=None, appearance=None):
    """Walk every tab and dialog, optionally screenshotting each.

    This is how the Mac gets looked at: nobody on the project owns one, so
    the macOS build job runs the frozen .app with `--tour all --shots DIR`
    and uploads the pictures. With shots_dir=None it is a gate instead:
    every dialog is constructed and torn down on the live app, on every
    platform, and `tour_errors` holds anything that raised.

    Each step is (name, open, close). `open` may block in a modal
    wait_window — Tk keeps firing after() callbacks meanwhile, so the
    step's finish is scheduled BEFORE open is called, the way the zone
    preview tests do it. Modal or not, the same code works.
    """
    tour_log = []
    tour_errors = []
    pre_toplevels = set()
    brightness = []

    def apply_appearance(win):
        # macOS only. `defaults write -g AppleInterfaceStyle Dark` does not
        # change a running login session, so the CI "dark" tour came back
        # light (2026-10-06). Tk can force a window's appearance directly.
        if appearance and sys.platform == 'darwin':
            try:
                win.tk.call('::tk::unsupported::MacWindowStyle', 'appearance', win,
                            'darkaqua' if appearance == 'dark' else 'aqua')
                win.update_idletasks()
            except tk.TclError as e:
                tour_errors.append(f"appearance {appearance}: {e}")
    apply_appearance(root)

    def new_toplevels():
        return [w for w in root.winfo_children()
                if isinstance(w, tk.Toplevel) and w not in pre_toplevels]

    def destroy_new():
        for w in new_toplevels():
            try:
                w.destroy()
            except tk.TclError:
                pass

    def shot(name):
        if not shots_dir:
            return
        os.makedirs(shots_dir, exist_ok=True)
        path = os.path.join(shots_dir, f"{len(tour_log):02d}-{name}.png")
        try:
            from PIL import ImageGrab
            # Crop to the app's own windows (root + any open Toplevel) so the
            # picture is the app, not a desktop — and so a multi-monitor
            # machine whose primary screen isn't the one the app opened on
            # still gets the app. all_screens only matters on Windows.
            wins = [root] + [w for w in root.winfo_children()
                             if isinstance(w, tk.Toplevel) and w.winfo_viewable()]
            x0 = min(w.winfo_rootx() for w in wins) - 12
            y0 = min(w.winfo_rooty() for w in wins) - 40
            x1 = max(w.winfo_rootx() + w.winfo_width() for w in wins) + 12
            y1 = max(w.winfo_rooty() + w.winfo_height() for w in wins) + 12
            try:
                img = ImageGrab.grab(bbox=(x0, y0, x1, y1), all_screens=True)
            except TypeError:
                img = ImageGrab.grab(bbox=(x0, y0, x1, y1))
            img.save(path)
            # Mean brightness of the picture: says which appearance was
            # really captured, so the artifact can't quietly be the wrong one.
            try:
                g = img.convert("L")
                brightness.append(sum(g.getdata()) / (g.width * g.height))
            except Exception:
                pass
        except Exception as e:  # noqa: BLE001 — fall back to the OS tool
            if sys.platform == 'darwin':
                subprocess.run(["screencapture", "-x", path], check=False)
            else:
                tour_errors.append(f"{name}: screenshot failed: {e}")

    held = {}

    def open_sizing_rules():
        opts = OptionsWindow(root, app, app.settings, app.update_ui_from_settings, lambda: None,
                             sizing_presets=app.sizing_presets, sizing_presets_save_callback=lambda: None)
        if not opts.lesser_open:
            opts._toggle_lesser_used()
        opts.show_preview_var.set(True)
        opts._toggle_preview_window()
        for w in opts.top.winfo_children():        # the PadPreviewWindow, if it opened
            if hasattr(w, 'preview_size_var'):
                try:
                    w.preview_size_var.set(10.0)
                except Exception:
                    pass
        held['opts'] = opts

    def close_sizing_rules():
        o = held.pop('opts', None)
        if o is not None:
            o.top.destroy()

    def open_nesting_preview():
        from svg_engine import nest_pads_with_zones
        pads = [{'size': 7.0, 'qty': 6}, {'size': 12.5, 'qty': 4}, {'size': 20.0, 'qty': 3}]
        s = dict(app.settings)
        s['zone_labels_enabled'] = True
        placed, zones = nest_pads_with_zones(pads, 'felt', 200.0, 150.0, s)
        NestingPreviewWindow(root, {'felt': placed}, 200.0, 150.0, zones=zones)

    def open_toner():
        engine = getattr(app, '_toner_engine', None)
        if engine is not None:
            engine.synthetic_hz = 233.08
        app._toner_mic_checked = True
        if app.settings.get("toner_unlocked") and hasattr(app, '_toner_harmonics_var'):
            app._toner_harmonics_var.set("on")      # the spectrum on the tone; the chart otherwise
        app.notebook.add(app.toner_tab_frame)
        app._fit_window_to_tabs()
        app.notebook.select(app.toner_tab_frame)

    def open_tuner_no_mic():
        # What a user with no (or a denied) microphone sees: the real
        # engine start, which fails on a runner and draws the audio-error
        # overlay. On a machine with a mic this is just the live tuner.
        app.notebook.select(app.tuner_tab_frame)

    def open_tuner():
        engine = getattr(app, '_tuner_engine', None)
        if engine is not None:
            engine.synthetic_hz = 440.0
        app.notebook.select(app.tuner_tab_frame)
        # The tab may already be selected (previous stop), so restart the
        # engine explicitly to pick up the synthetic source.
        try:
            app._tuner_stop()
            app._tuner_start()
        except Exception:
            pass

    steps = [
        ("pad-maker", lambda: app.notebook.select(app.pad_tab), None),
        ("sizing-rules-with-preview", open_sizing_rules, close_sizing_rules),
        ("layer-colors", lambda: LayerColorWindow(root, app.settings, lambda: None), destroy_new),
        ("gcode-settings", lambda: GcodeSettingsWindow(
            root, app.settings, lambda s: None,
            materials=([("felt", _("Felt")), ("card", _("Card")), ("leather", _("Leather"))]
                       # as Options > G-code Settings: the ring section while that feature is on
                       + ([("adhesive_ring", _("Adhesive rings"))] if app._adhesive_ring_enabled() else [])),
            gcode_presets=app.gcode_presets, gcode_presets_save_callback=lambda: None), destroy_new),
        ("nesting-preview", open_nesting_preview, destroy_new),
        ("polygon-draw", lambda: PolygonDrawWindow(root, unit="in", settings=app.settings), destroy_new),
        ("job-history", lambda: JobHistoryWindow(root, load_job_history(), lambda jobs: None), destroy_new),
        ("feature-set", app._open_feature_set, destroy_new),
        ("user-guide", lambda: UserGuideWindow(root, section="pad_generator"), destroy_new),
        ("about", lambda: AboutDialog(root), destroy_new),
        ("welcome", app.open_welcome, destroy_new),
        ("quick-start", app.open_quick_start, destroy_new),
        ("key-heights", lambda: app.notebook.select(app.key_tab), None),
        ("key-layout", lambda: KeyLayoutWindow(root, app.settings, lambda: None, lambda: None), destroy_new),
        ("serial-lookup", lambda: app.notebook.select(app.serial_tab), None),
        ("screw-specs", lambda: app.notebook.select(app.screw_tab), None),
        ("tooling", lambda: app.notebook.select(app.tooling_tab_frame), None),
        ("tuner-no-mic", open_tuner_no_mic, None),
        ("tuner", open_tuner, None),
        ("toner", open_toner, None),
    ]
    if not audio_supported():
        # Hidden on this build; selecting a hidden ttk tab would show it.
        steps = [s for s in steps if s[0] not in ("tuner-no-mic", "tuner", "toner")]

    def run_step(i):
        if i >= len(steps):
            try:
                app._tuner_stop()
                app._toner_stop()
            except Exception:
                pass
            if shots_dir:
                with open(os.path.join(shots_dir, "tour-done.txt"), "w", encoding="utf-8") as f:
                    f.write("\n".join(tour_log) + "\n")
                    if brightness:
                        mean = sum(brightness) / len(brightness)
                        f.write(f"appearance requested: {appearance or 'system'}; measured: "
                                f"{'dark' if mean < 110 else 'light'} (mean {mean:.0f}/255)\n")
                    if tour_errors:
                        f.write("ERRORS:\n" + "\n".join(tour_errors) + "\n")
            if on_done is not None:
                on_done()
            return
        name, open_fn, close_fn = steps[i]
        pre_toplevels.clear()
        pre_toplevels.update(w for w in root.winfo_children() if isinstance(w, tk.Toplevel))

        def finish():
            try:
                for w in new_toplevels():
                    apply_appearance(w)
                root.update_idletasks()
                shot(name)
                tour_log.append(name)
            except Exception as e:  # noqa: BLE001
                tour_errors.append(f"{name}: {e!r}")
            try:
                if close_fn is not None:
                    close_fn()
            except Exception as e:  # noqa: BLE001
                tour_errors.append(f"{name} close: {e!r}")
            root.after(150, lambda: run_step(i + 1))

        # Audio tabs need a couple of seconds of frames before they look alive.
        wait = step_ms * (2 if name in ("tuner-no-mic", "tuner", "toner") else 1)
        root.after(wait, finish)
        try:
            open_fn()
        except Exception as e:  # noqa: BLE001
            tour_errors.append(f"{name} open: {e!r}")

    root.after(200, lambda: run_step(0))
    run_tour.log = tour_log
    run_tour.errors = tour_errors
    return tour_log, tour_errors


def _selftest(root):
    """Frozen-build probe: build the whole app headless and exit 0 or 1.

    `SaxShopCompanion --selftest` constructs every tab and every import in
    a withdrawn root, checks the tabs exist, and exits without saving
    anything. CI runs it against the PyInstaller output on each platform —
    the one check that runs the *shipped* bundle rather than the source
    tree, so a missing hidden import, a bad data path, or a module that
    must not be bundled on a platform (the macOS GPU renderer) fails the
    build instead of a user's first launch. Nothing is shown and no
    dialog can block: the excepthook is replaced for the run.
    """
    import traceback

    def _say(msg):
        # os._exit skips Python's buffer flush, so write the verdict straight
        # to fd 1; a windowed .exe may have no fd 1 at all, hence the guard.
        try:
            os.write(1, (msg + "\n").encode("utf-8", "replace"))
        except OSError:
            pass

    def _plain_hook(exc_type, exc_value, exc_tb):
        _say("selftest FAILED:\n" + "".join(traceback.format_exception(exc_type, exc_value, exc_tb)))
        try:
            root.destroy()
        except Exception:
            pass
        os._exit(1)

    sys.excepthook = _plain_hook
    root.report_callback_exception = _plain_hook
    root.withdraw()
    try:
        app = PadSVGGeneratorApp(root)
        app._refresh_machine_ui_state()
        for attr in ('pad_tab', 'key_tab', 'serial_tab', 'screw_tab', 'tooling_tab_frame'):
            if getattr(app, attr, None) is None:
                raise RuntimeError(f"selftest: tab missing: {attr}")
        root.update_idletasks()
        root.update()
        # The shipped bundle can verify stohrermusic.com: certifi's CA bundle
        # was packed and is found (no network; the Intel Mac's only check).
        from config import assert_tls_ready
        n_ca = assert_tls_ready()
        from config import APP_VERSION as _ver
        _say(f"selftest ok: {_ver} on {sys.platform}, "
             f"{len(app.notebook.tabs())} tabs, frozen={getattr(sys, 'frozen', False)}, "
             f"{n_ca} CA certificates")
    except BaseException:
        _say("selftest FAILED:\n" + traceback.format_exc())
        try:
            root.destroy()
        except Exception:
            pass
        os._exit(1)
    root.destroy()
    os._exit(0)


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
    try:
        from config import log_runtime_versions
        log_runtime_versions(root)
    except Exception:
        logging.getLogger(__name__).exception("could not log runtime versions")

    if '--selftest' in sys.argv[1:]:
        _selftest(root)

    # --tour tuner: open on the Tuner tab with a synthetic 440 Hz tone in
    # place of the microphone, so CI can screenshot the tuner running on a
    # machine with no audio input (the macOS runners). Visible, normal
    # mainloop; only the audio source is faked.
    tour = None
    shots_dir = None
    appearance = None
    for i, arg in enumerate(sys.argv[1:], 1):
        if arg.startswith('--tour='):
            tour = arg.split('=', 1)[1]
        elif arg == '--tour' and i < len(sys.argv) - 1:
            tour = sys.argv[i + 1]
        elif arg.startswith('--shots='):
            shots_dir = arg.split('=', 1)[1]
        elif arg == '--shots' and i < len(sys.argv) - 1:
            shots_dir = sys.argv[i + 1]
        elif arg.startswith('--appearance='):
            appearance = arg.split('=', 1)[1]
        elif arg == '--appearance' and i < len(sys.argv) - 1:
            appearance = sys.argv[i + 1]

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
        logging.getLogger(__name__).exception(
            "could not set up the Machine menu and Frame & Cut buttons at launch")
    # Detect the Falcon (Grbl) controller in the background only if
    # the user has opted into the experimental Machine menu — there's
    # no point probing USB serial ports for a user who never plans to
    # use direct Falcon control.
    if app._machine_enabled():
        root.after(500, app._start_laser_watch)

    # The Welcome window, on the first launch of each version. Only here,
    # in a real launch: --selftest has exited above, a --tour has its own
    # Welcome stop, and nothing that merely builds the app (every test
    # suite) reaches this line. On an after(), so the main window is up
    # first and startup is no slower.
    if not tour:
        root.after(600, app.maybe_show_welcome)

    if tour == 'all':
        # Full walk: every tab and dialog, screenshot each when --shots is
        # given, then exit. See run_tour().
        root.geometry("+40+60")

        def _tour_finished():
            try:
                root.destroy()
            finally:
                os._exit(0)
        run_tour(root, app, shots_dir=shots_dir, on_done=_tour_finished, appearance=appearance)
    elif tour in ('tuner', 'toner'):
        def _start_tour():
            try:
                if tour == 'tuner':
                    engine = getattr(app, '_tuner_engine', None)
                    if engine is not None:
                        engine.synthetic_hz = 440.0
                    app.notebook.select(app.tuner_tab_frame)
                else:
                    # Analyzer: Bb3 with harmonics. With the harmonics unlocked
                    # in the profile (SAXSHOP_CONFIG_DIR lets CI supply one) the
                    # switch goes on and the spectrum shows; else the chart.
                    # Re-adding the frame un-hides a tab a profile hid.
                    engine = getattr(app, '_toner_engine', None)
                    if engine is not None:
                        engine.synthetic_hz = 233.08
                    app._toner_mic_checked = True   # no modal mic notice
                    if app.settings.get("toner_unlocked") and hasattr(app, '_toner_harmonics_var'):
                        app._toner_harmonics_var.set("on")
                    app.notebook.add(app.toner_tab_frame)
                    app._fit_window_to_tabs()
                    app.notebook.select(app.toner_tab_frame)
            except Exception as e:
                logging.getLogger(__name__).warning("tour %s: could not open the tab: %s", tour, e)
        root.after(1500, _start_tour)
    root.mainloop()
