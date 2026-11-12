"""The type of every setting, in one place, and the check that runs at load.

`DEFAULT_SETTINGS` (config.py) holds the values; this module holds what each key may
be: its type, the values it may take (an enumeration, a range, None or not), and for
the nested dicts their own key schemas. `config.load_settings()` runs `validate_settings`
after its merge and migrations, so a value of the wrong type or outside its declared
values is put back to its default and logged once with the key, the bad value and the
file; a typo in a hand-edited file or a value from an old version then stops at the door
instead of travelling until something breaks far from the cause. The dict stays the
runtime object: nothing that reads `settings["key"]` changes.

Rules (2026-11-10, structural step 4; history: docs/decisions/settings.md):

- The schema covers exactly the keys of `DEFAULT_SETTINGS`, both ways, at every level
  of the twelve nested dicts; `tools/test_settings_schema.py` fails on a key in one and
  not the other. A new setting is added to `DEFAULT_SETTINGS` AND here.
- A range or an enumeration is declared only where the code already enforces or
  assumes it (a dialog's radio buttons, a slider's ends, a clamp, an engine's
  requirement). Where nothing bounds a number today, only its type is declared, and the
  comment says so. No limit was invented for this module.
- Numbers: an int is accepted where a float is declared; a bool is never a number
  (`True` is not 1 here); NaN and infinity are never numbers.
- A nested dict's value of the wrong type replaces just that value, never the dict. A
  key a block carries that the schema does not know (one from a newer version, kept by
  the two-level merge) is left alone.
- A list is checked as a whole: its type and its elements' type (a range rule that is
  not a dict), not the fields inside each element.
- `tools/test_settings_schema.py` also scans the app's modules for every string-literal
  key read from or written to a settings dict and fails on one the schema does not know.
"""
import math
import re

# ---- the kinds of field ------------------------------------------------------------------------


class Field:
    """What one setting may be. `kind` is one of bool, int, float, str, list, dict."""

    __slots__ = ("kind", "choices", "min", "max", "gt", "nullable", "pattern", "items", "length",
                 "keys", "extra", "note")

    def __init__(self, kind, *, choices=None, min=None, max=None, gt=None, nullable=False, pattern=None,
                 items=None, length=None, keys=None, extra=None, note=""):
        self.kind = kind
        self.choices = tuple(choices) if choices is not None else None
        self.min, self.max, self.gt = min, max, gt
        self.nullable = nullable
        self.pattern = re.compile(pattern) if pattern else None
        self.items = items            # list: the Field every element must satisfy (None: any element)
        self.length = length          # list: the exact length, when fixed
        self.keys = keys              # dict: {key: Field}; None means any dict
        self.extra = extra            # dict: the Field for values under keys not in `keys` (None: left alone)
        self.note = note

    def accepts(self, value):
        """Whether `value` is acceptable as this field, looking no deeper than one level
        into a dict (its known keys are checked one by one by `validate`)."""
        if value is None:
            return self.nullable
        k = self.kind
        if k == "bool":
            return isinstance(value, bool)
        if k in ("int", "float"):
            if isinstance(value, bool):
                return False
            if k == "int" and not isinstance(value, int):
                return False
            if k == "float" and not isinstance(value, (int, float)):
                return False
            if isinstance(value, float) and not math.isfinite(value):
                return False
            if self.choices is not None and value not in self.choices:
                return False
            if self.min is not None and value < self.min:
                return False
            if self.max is not None and value > self.max:
                return False
            if self.gt is not None and not value > self.gt:
                return False
            return True
        if k == "str":
            if not isinstance(value, str):
                return False
            if self.choices is not None and value not in self.choices:
                return False
            if self.pattern is not None and not self.pattern.fullmatch(value):
                return False
            return True
        if k == "list":
            if not isinstance(value, list):
                return False
            if self.length is not None and len(value) != self.length:
                return False
            if self.items is not None and not all(self.items.accepts(v) for v in value):
                return False
            return True
        if k == "dict":
            return isinstance(value, dict)
        raise ValueError(f"unknown field kind {k!r}")

    def describe(self):
        """One phrase for the log line: 'a number 0 to 100', 'one of in, cm, mm', ..."""
        if self.kind == "dict":
            return "a mapping"
        if self.kind == "list":
            return "a list" + (f" of {self.items.describe()}s" if self.items is not None else "")
        if self.kind == "str":
            if self.choices:
                return "one of " + ", ".join(self.choices)
            return "text" + (" (a colour like #RRGGBB)" if self.pattern else "")
        if self.kind == "bool":
            return "true or false"
        what = "a whole number" if self.kind == "int" else "a number"
        if self.choices:
            return what + " among " + ", ".join(str(c) for c in self.choices)
        if self.min is not None and self.max is not None:
            return f"{what} {self.min} to {self.max}"
        if self.min is not None:
            return f"{what} of at least {self.min}"
        if self.gt is not None:
            return f"{what} above {self.gt}"
        return what


def Bool(**kw):
    return Field("bool", **kw)


def Int(**kw):
    return Field("int", **kw)


def Float(**kw):
    """A number: an int is accepted too (never a bool, never NaN or infinity)."""
    return Field("float", **kw)


def Str(**kw):
    return Field("str", **kw)


def List(items=None, **kw):
    return Field("list", items=items, **kw)


def Dict(keys=None, extra=None, **kw):
    return Field("dict", keys=keys, extra=extra, **kw)


# ---- the enumerations the code already assumes ------------------------------------------------
# Each is checked against its source of truth by tools/test_settings_schema.py.
LANGUAGES = ("en", "es", "de", "fr", "it")                      # i18n.LANGUAGE_NAMES
UNITS = ("in", "cm", "mm")                                      # the Sizing Rules radios; main._get_sheet_mm
THICKNESS_UNITS = ("mm", "in")                                  # the felt thickness radios
HOLE_OPTIONS = ("3.5mm", "3.0mm", "Custom", "No center holes")  # the Pad Maker's hole radios (main.get_hole_dia)
RANGE_MODES = ("universal", "range")                            # the four Universal / Per Size Range switches
PLACEMENT_MODES = ("from_outside", "centered", "from_inside")   # size-label placement radios (out / mid / int)
LOCATOR_STYLES = ("lines", "circle", "both")                    # svg_engine.LOCATOR_STYLES
EDGE_BIASES = ("off", "center", "n", "s", "e", "w", "nw", "ne", "sw", "se")  # the d-pad (main._set_edge_bias)
PAPER_SIZES = ("letter", "a4")                                  # the card paper combobox
CUT_GROUPINGS = ("layer", "pad")                                # G-code Settings radios
ENGRAVING_MODES = ("line", "filled")                            # G-code Settings radios, per material
RING_LABEL_LOCATIONS = ("centered", "from_outside")             # Tooling G-code Settings radios
TRANSPOSITIONS = ("C", "Eb", "Bb", "F")                         # tuner_tab.TRANSPOSITION_SHIFTS
TUNER_FPS = ("60", "90", "120")                                 # the tuner's FPS switch
WAVEFORMS = ("pure", "rich")                                    # tuner_engine.ReferencePlayer.play
TONER_FPS = ("30", "60")                                        # toner_tab.FRAME_RATES
TONER_VIEWS = ("spectrum", "bars")                              # the toner's VIEW switch
TONER_SCALES = ("linear", "db")                                 # the toner's SCALE switch
SAX_TYPES = ("Bass", "Baritone", "Tenor", "C Melody", "Alto", "F Mezzo", "Soprano", "Sopranino")  # toner_engine
HOLDER_LAYERS = (5, 6)                                          # the Die Holders radios
HEX_COLOUR = r"#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})"               # what Tk and SVG both take; the chooser writes 6

ORGANIZER_MATERIALS = ("basswood", "acrylic")                   # tooling_tab._ORGANIZER_MATERIALS
CRESCENT_MATERIALS = ("cork", "acrylic", "basswood", "leather", "felt", "card")  # tooling_tab._CRESCENT_MATERIALS
PAD_MATERIALS = ("felt", "card", "leather", "exact_size")       # the Pad Maker's four, as the dicts key them
GCODE_MATERIALS = ("felt", "card", "leather", "acrylic", "basswood", "adhesive_ring", "cork")  # config.GCODE_PRESET_MATERIALS
CUT_ONLY_MATERIALS = ("adhesive_ring",)                         # config.GCODE_CUT_ONLY_MATERIALS

# ---- the nested blocks -------------------------------------------------------------------------

# A size-label placement: where the number sits and how far (mm).
PLACEMENT = Dict({"mode": Str(choices=PLACEMENT_MODES), "value": Float()})

# One material's laser settings. Speeds and powers: Grbl takes F and S as non-negative
# numbers, and 0 means "not set" (config.engraving_is_set); passes: the dialog writes
# max(1, ...); kerf: the dialog's spinbox runs 0 to 1, but typing is not clamped, so only
# the floor is declared; filled_line_spacing: the scan-line fill steps by it, so it must be
# above 0. Nothing bounds the top of a speed or power today.
_CUT_KEYS = {
    "cut_speed": Float(min=0), "cut_power": Float(min=0), "cut_passes": Int(min=1),
    "kerf_width": Float(min=0), "air_assist_cut": Bool(),
}
_ENGRAVE_AND_HOLE_KEYS = {
    "engraving_mode": Str(choices=ENGRAVING_MODES),
    "engraving_speed": Float(min=0), "engraving_power": Float(min=0), "engraving_passes": Int(min=1),
    "filled_engraving_speed": Float(min=0), "filled_engraving_power": Float(min=0),
    "filled_engraving_passes": Int(min=1), "filled_line_spacing": Float(gt=0),
    "hole_speed": Float(min=0), "hole_power": Float(min=0), "hole_passes": Int(min=1),
    "air_assist_engraving": Bool(), "air_assist_filled_engraving": Bool(), "air_assist_hole": Bool(),
}
GCODE_BLOCK = Dict({**_ENGRAVE_AND_HOLE_KEYS, **_CUT_KEYS})       # config.GCODE_PRESET_KEYS
CUT_ONLY_BLOCK = Dict(dict(_CUT_KEYS))                             # config.GCODE_CUT_ONLY_KEYS

TUNER_SETTINGS = Dict({
    "stripe_color": Str(pattern=HEX_COLOUR),
    "reference_pitch": Float(note="the A= slider runs 420 to 460; a file may hold any pitch"),
    "transposition": Str(choices=TRANSPOSITIONS),
    "sensitivity": Int(min=0, max=100),                      # SENS slider
    "waveform": Str(choices=WAVEFORMS),
    "fps": Str(choices=TUNER_FPS),
    "ring_brightness": Int(min=0, max=100),                  # NOTE slider
    "overall_brightness": Int(min=10, max=150),              # BRIGHT slider
    "octave_boost": Int(min=0, max=100),                     # OCT. slider
    "faceplate_color": Str(pattern=HEX_COLOUR),
    "show_fps": Bool(),
})

TONER_SETTINGS = Dict({
    "reference_pitch": Float(),
    "sensitivity": Int(min=0, max=100),                      # the high/low switch writes 75 or 25
    "fps": Str(choices=TONER_FPS),
    "view_mode": Str(choices=TONER_VIEWS),
    "scale_mode": Str(choices=TONER_SCALES),
    "gauge_bias": Dict({"richness": Float(), "warmth": Float()},
                       note="the live descriptor gauges were removed 2026-04-06; nothing reads this"),
    "sax_type": Str(choices=SAX_TYPES),
    "concert_pitch": Bool(),
    "chart_low_notes_left": Bool(),                          # the intonation chart's direction
    "harmonics_on": Bool(),                                  # the HARMONICS switch's last position
    "analysis_descriptors": Dict({"richness": Bool(), "warmth": Bool(), "even_odd": Bool(), "rolloff_shape": Bool()}),
})

# The Speed & Power Test's remembered form. Its numbers are typed into entries and the tab
# stores every value, start and end as float() (the passes too: Matt's file holds 1.0 and
# 3.0) and every stops count as int(); nothing bounds them in the dialog (stops are clamped
# 2..10 on read).
FEEDS_SPEEDS_TESTER = Dict({
    "material": Str(note="the combobox's display label, translated; not an enumeration"),
    "disc_diameter_mm": Float(), "inner_diameter_mm": Float(),
    "speed_value": Float(), "speed_sweep": Bool(), "speed_start": Float(), "speed_end": Float(), "speed_stops": Int(),
    "power_value": Float(), "power_sweep": Bool(), "power_start": Float(), "power_end": Float(), "power_stops": Int(),
    "passes_value": Float(), "passes_sweep": Bool(), "passes_start": Float(), "passes_end": Float(), "passes_stops": Int(),
    "engraving_on": Bool(), "eng_speed_value": Float(), "eng_power_value": Float(),
    "air_assist": Bool(), "also_test_no_air": Bool(), "show_preview": Bool(),
    "sheet_w": Str(), "sheet_h": Str(), "sheet_unit": Str(choices=("in", "mm")), "filename": Str(),
})

TOOLING_SETTINGS = Dict({
    "sheet_width": Str(), "sheet_height": Str(),             # entry text, parsed at generate
    "engrave_ring": Bool(), "engrave_cutout": Bool(),
    "engraving_mode": Str(choices=ENGRAVING_MODES),
    "step_size": Str(),
    "ring_font_size": Float(), "cutout_font_size": Float(),  # "must be valid numbers", nothing more
    "ring_engraving_location": Str(choices=RING_LABEL_LOCATIONS),
    "ring_engraving_offset": Float(),
    "holder_layer_count": Int(choices=HOLDER_LAYERS),
    "holder_sheet_width": Str(), "holder_sheet_height": Str(),
    "crescent_tonehole_d": Str(), "crescent_percent": Str(),
    "crescent_curve": Float(min=0, max=100),                 # the curve slider
    "crescent_thickness": Str(), "crescent_qty": Str(),
    "crescent_list": List(items=Dict(), note="[{tonehole_d, percent, curve, qty}, ...]"),
    # The Die Organizer's material and sheet, the Crescent's material (declared 2026-11-10,
    # when the gate found the tab writing them to keys that no default held).
    "organizer_material": Str(choices=ORGANIZER_MATERIALS),
    "organizer_sheet_width": Str(), "organizer_sheet_height": Str(),   # entry text; "" = size from the app's units
    "organizer_sheet_units": Str(choices=("", "in", "mm")),            # "" = follow the app's units
    "crescent_material": Str(choices=CRESCENT_MATERIALS),
    "feeds_speeds_tester": FEEDS_SPEEDS_TESTER,
})

# ---- the schema: every key of DEFAULT_SETTINGS -------------------------------------------------
SCHEMA = {
    "language": Str(choices=LANGUAGES),
    "units": Str(choices=UNITS),
    "felt_offset": Float(),                  # nothing bounds the sizing numbers today
    "card_to_felt_offset": Float(),
    "leather_wrap_multiplier": Float(),
    "sheet_width": Str(), "sheet_height": Str(),             # entry text; "must be greater than zero" at generate
    "hole_option": Str(choices=HOLE_OPTIONS),
    "custom_hole_size": Str(),                               # entry text; "must be greater than zero" at generate
    "min_hole_size": Float(),
    "felt_thickness": Float(),
    "felt_thickness_unit": Str(choices=THICKNESS_UNITS),
    "engraving_on": Bool(),
    "show_engraving_warning": Bool(),
    "show_sd_card_warning": Bool(),
    "last_output_dir": Str(),
    "resonance_clicks": Int(),                               # a counter
    "compatibility_mode": Bool(),
    "darts_enabled": Bool(),
    "dart_threshold": Float(),
    "dart_overwrap": Float(), "dart_wrap_bonus": Float(), "dart_frequency_multiplier": Float(),
    "dart_shape_factor": Float(min=0.0, max=1.0),            # the Darts slider; the v2 migration clamps to it too
    "dart_shape_v2": Bool(),
    "dart_engraving_on": Bool(),
    "dart_engraving_loc": PLACEMENT,
    "dart_range_mode": Str(choices=RANGE_MODES),
    "dart_ranges": List(items=Dict(), note="[{min_size, max_size, overwrap, wrap_bonus, frequency_multiplier, shape_factor, engraving_on, engraving_loc}]"),
    "sizing_range_mode": Str(choices=RANGE_MODES),
    "sizing_ranges": List(items=Dict(), note="[{min_size, max_size, felt_offset, card_to_felt_offset, leather_wrap_multiplier, min_hole_size, felt_thickness, felt_thickness_unit}]"),
    "engraving_settings_range_mode": Str(choices=RANGE_MODES),
    "engraving_settings_ranges": List(items=Dict(), note="[{min_size, max_size, engraving_on, engraving_font_size: {material: size}}]"),
    "engraving_placement_range_mode": Str(choices=RANGE_MODES),
    "engraving_placement_ranges": List(items=Dict(), note="[{min_size, max_size, engraving_location: {material: {mode, value}}}]"),
    "darted_out_from_valley": Bool(),
    "zone_labels_enabled": Bool(),
    "zone_label_min_size": Float(), "zone_label_max_size": Float(),
    "zone_gutter_mm": Float(),
    "locator_marks_enabled": Bool(),
    "locator_marks_min_size": Float(), "locator_marks_max_size": Float(),
    "locator_marks_style": Str(choices=LOCATOR_STYLES),
    "locator_marks_dashed": Bool(),
    "adhesive_ring_enabled": Bool(),
    "adhesive_ring_trim_mm": Float(),                        # may be negative on purpose (hides the ring under the wrap)
    "zone_border_mm": Float(), "zone_label_font_mm": Float(), "zone_edge_margin_mm": Float(), "zone_group_gap_mm": Float(),
    "edge_bias": Str(choices=EDGE_BIASES),
    "card_use_paper_size": Bool(),
    "card_paper_size": Str(choices=PAPER_SIZES),
    "gcode_output_enabled": Bool(note="nothing reads this key"),
    "gcode_settings": Dict({m: (CUT_ONLY_BLOCK if m in CUT_ONLY_MATERIALS else GCODE_BLOCK) for m in GCODE_MATERIALS}),
    "tuner_settings": TUNER_SETTINGS,
    "toner_settings": TONER_SETTINGS,
    "visible_preset_fields": Dict({k: Bool() for k in ("serial", "reed", "ligature", "room", "preamp", "notes")}),
    "tooling_settings": TOOLING_SETTINGS,
    "filled_overscan_enabled": Bool(),
    "filled_overscan_mm": Float(),
    "gcode_return_speed": Int(note="the dialog's IntVar; nothing bounds it"),
    "gcode_cut_grouping": Str(choices=CUT_GROUPINGS),
    "sd_card_path": Str(note="nothing reads this key; the SD menu item is gone"),
    "eject_sd_after_gcode": Bool(),
    "falcon_serial_port_override": Str(nullable=True, note="nothing reads this key"),
    "laser_framing_power_s": Int(note="clamped to 0..FRAMING_POWER_S_MAX when read (config.get_framing_power_s), so not bounded here"),
    "laser_framing_power_by_material": Dict({m: Int() for m in GCODE_MATERIALS}, extra=Int(),
                                            note="the same clamp on read; the Framing Power dialog writes 0..FRAMING_POWER_S_MAX"),
    "laser_framing_feed": Float(note="a Grbl F value; read, never written by a dialog"),
    "laser_bed_x_max": Float(), "laser_bed_y_max": Float(),
    "laser_seed_feed": Float(note="nothing reads this key"),
    "camera_polygon_inset_mm": Float(min=0, max=50),         # the Inset dialog: "a number 0-50"
    "camera_detection_threshold_bias": Int(min=-80, max=80),  # the calibration dialog's slider
    "camera_detection_invert": Bool(),
    "camera_calibration_engrave_offset_mm": List(items=Float(), length=2, nullable=True),   # [x, y] mm, or None
    "camera_index_override": Int(nullable=True),              # a camera's index, or None for the resolver
    "show_live_camera": Bool(),                               # the Pad Maker's "Live camera preview" tick
    "polygon_draw_grid_size_in": Int(), "polygon_draw_grid_size_cm": Int(),
    "seen_polygon_tutorial": Bool(),
    "welcomed_version": Str(),
    "seen_kerf_test_tutorial": Bool(), "seen_calibration_tutorial": Bool(),
    "audio_input_device": Int(nullable=True),                 # a sounddevice index, or None for the default
    "audio_input_device_name": Str(nullable=True),            # that device's name; what is looked up (config.input_device_for)
    "capture_threshold": Int(min=0, max=100),                 # the Capture Threshold slider
    "toner_record_wav": Bool(),
    "toner_recording_dir": Str(),
    "toner_wav_reanalyze": Bool(), "toner_wav_auto_delete": Bool(),
    "visible_sax_types": List(items=Str(), nullable=True),    # None = all types
    "toner_unlocked": Bool(), "toner_sandbox_enabled": Bool(),
    "analyzer_tab_shown_v3": Bool(),                          # config.migrate_analyzer_tab has run
    "experimental_machine_menu": Bool(),
    "show_preview": Bool(), "seen_preview_tutorial": Bool(),
    "last_pad_library": Str(), "last_key_library": Str(),
    "tooltips_enabled": Bool(),
    "visible_tabs": Dict({k: Bool() for k in ("Key Height Library", "Serial Lookup", "Screw Specs", "Tooling", "Tuner", "Toner")},
                         extra=Bool()),
    "key_layout": Dict({k: Bool() for k in ("show_serial", "large_notes", "show_B", "show_F", "show_Palm_F", "show_Palm_E",
                                            "show_Palm_Eb", "show_Palm_D", "show_G", "show_D", "show_Low_C", "show_Low_B",
                                            "show_Low_Bb")}),
    "engraving_font_size": Dict({m: Float() for m in PAD_MATERIALS}),
    "engraving_location": Dict({m: PLACEMENT for m in PAD_MATERIALS + ("darted_leather",)}),
    "layer_colors": Dict({k: Str(pattern=HEX_COLOUR) for k in (
        "felt_outline", "felt_center_hole", "felt_engraving",
        "card_outline", "card_center_hole", "card_engraving",
        "leather_outline", "leather_center_hole", "leather_engraving",
        "exact_size_outline", "exact_size_center_hole", "exact_size_engraving",
        "leather_locator", "adhesive_ring_cut",
        "die_outer_cut", "die_inner_cut", "die_engraving", "die_cutout_engraving",
        "die_holder_cut", "die_holder_hole")}),
}


# ---- validation --------------------------------------------------------------------------------

def validate_settings(settings, defaults, schema=SCHEMA, _path=""):
    """Put every value that is not what the schema declares back to its default, in place.

    `defaults` is the dict the replacements come from (DEFAULT_SETTINGS, or the matching
    block of it). Returns the corrections as a list of (path, bad_value, field): the
    caller logs them. Keys of `settings` the schema does not know are left alone; keys
    the schema knows and `settings` lacks are left alone too (load_settings has already
    filled them from the defaults)."""
    import copy

    corrections = []
    for key, field in schema.items():
        if key not in settings:
            continue
        value = settings[key]
        path = f"{_path}{key}"
        if not field.accepts(value):
            settings[key] = copy.deepcopy(defaults[key])
            corrections.append((path, value, field))
            continue
        if field.kind == "dict" and isinstance(value, dict):
            if field.keys:
                corrections += validate_settings(value, defaults[key], field.keys, _path=path + ".")
            if field.extra is not None:
                for sub, sub_value in list(value.items()):
                    if sub in (field.keys or {}):
                        continue
                    if not field.extra.accepts(sub_value):
                        # an extra key has no default of its own: it goes
                        del value[sub]
                        corrections.append((f"{path}.{sub}", sub_value, field.extra))
    return corrections
