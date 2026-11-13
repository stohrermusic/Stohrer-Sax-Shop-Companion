"""The dialog windows, as a package (2026-11-10).

Cut from one 11,330-line ui_dialogs.py by what the dialogs share, not by size: common (what every
dialog uses), sizing, gcode_settings, presets_io, resonance, previews, history, guide, machine. Every
body moved byte for byte. This file re-exports every name the app, the tabs and the suites import,
so `from ui_dialogs import X` and `ui_dialogs.X` keep working. A dialog that another dialog constructs
and that suites stand in for on the package (ui_dialogs.FalconRunDialog = Fake, and the other four
named in each module's `_pkg` line) is looked up through this package at call time, so the stand-in
is what gets built. History: docs/decisions/user-interface.md.
"""
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from .common import (
    IS_MACOS,
    DIALOG_BG,
    mlog,
    log,
    material_display_name,
    bind_mousewheel,
    _TOOLTIPS_ENABLED,
    set_tooltips_enabled,
    tooltips_enabled,
    Tooltip,
    gcode_file_note,
    add_tooltip,
    add_tooltips,
    get_unique_name,
    ConfirmationDialog,
    _PREVIEW_COLORS,
)
from .sizing import (
    grey_value_when_mid,
    OptionsWindow,
    PadPreviewWindow,
    SaveSizingPresetDialog,
    LayerColorWindow,
)
from .gcode_settings import (
    GcodeSettingsWindow,
)
from .presets_io import (
    KeyLayoutWindow,
    ExportPresetsWindow,
    ImportPresetsWindow,
    WebImportPresetsWindow,
    ImportTargetWindow,
    PadNotesWindow,
)
from .resonance import (
    ResonanceWindow,
    ResonanceProgressDialog,
    UninstallResonanceDialog,
)
from .previews import (
    NestingPreviewWindow,
    SpeedPowerTestPreview,
)
from .history import (
    tooling_tool_name,
    tooling_job_not_loadable,
    JobHistoryWindow,
)
from .guide import (
    LATEST_RELEASE_URL,
    _FigureCanvasHost,
    UserGuideWindow,
    QuickStartWindow,
    WelcomeWindow,
    AboutDialog,
)
from .machine import (
    _holds_port,
    PolygonDrawWindow,
    CameraCalibrationDialog,
    CameraCaptureDialog,
    _GRBL_ERROR_TEXT,
    _GRBL_ALARM_TEXT,
    _decode_grbl_error,
    LidClosedDialog,
    FalconRunDialog,
    LiveCameraWindow,
)

__all__ = [
    'tk',
    'filedialog',
    'messagebox',
    'simpledialog',
    'ttk',
    'IS_MACOS',
    'DIALOG_BG',
    'mlog',
    'log',
    'material_display_name',
    'bind_mousewheel',
    '_TOOLTIPS_ENABLED',
    'set_tooltips_enabled',
    'tooltips_enabled',
    'Tooltip',
    'gcode_file_note',
    'add_tooltip',
    'add_tooltips',
    'get_unique_name',
    'ConfirmationDialog',
    '_PREVIEW_COLORS',
    'grey_value_when_mid',
    'OptionsWindow',
    'PadPreviewWindow',
    'SaveSizingPresetDialog',
    'LayerColorWindow',
    'GcodeSettingsWindow',
    'KeyLayoutWindow',
    'ExportPresetsWindow',
    'ImportPresetsWindow',
    'WebImportPresetsWindow',
    'ImportTargetWindow',
    'PadNotesWindow',
    'ResonanceWindow',
    'ResonanceProgressDialog',
    'UninstallResonanceDialog',
    'NestingPreviewWindow',
    'SpeedPowerTestPreview',
    'tooling_tool_name',
    'tooling_job_not_loadable',
    'JobHistoryWindow',
    'LATEST_RELEASE_URL',
    '_FigureCanvasHost',
    'UserGuideWindow',
    'QuickStartWindow',
    'WelcomeWindow',
    'AboutDialog',
    '_holds_port',
    'PolygonDrawWindow',
    'CameraCalibrationDialog',
    'CameraCaptureDialog',
    '_GRBL_ERROR_TEXT',
    '_GRBL_ALARM_TEXT',
    '_decode_grbl_error',
    'LidClosedDialog',
    'FalconRunDialog',
    'LiveCameraWindow',
]
