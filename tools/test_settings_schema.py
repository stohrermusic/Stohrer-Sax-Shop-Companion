"""The settings schema (settings_schema.py) and the gate on settings keys.

  1. The schema covers exactly the keys of DEFAULT_SETTINGS, both ways, at every level of
     the nested dicts, and every default passes its own field.
  2. The enumerations the schema declares are the ones the code defines (languages, locator
     styles, transpositions, sax types, frame rates, materials, the G-code preset keys, the
     switches' literals).
  3. Validation: the defaults come through untouched; planted faults (a wrong type, a value
     off its enumeration, out of its range, a bool where a number is declared, None where
     it is not allowed, a bad value inside a nested block, a bad element in a list, a bad
     extra key) are each put back to the default and reported, and nothing else moves.
  4. load_settings runs it after the merge and the migrations, logs each correction with the
     key, the value and the file, and a fault in the check itself cannot cost the user the
     file.
  5. The gate: an AST scan of the app's modules for every string-literal key read from or
     written to a settings dict (`settings["k"]`, `settings.get("k"`, `"k" in settings`,
     `self.settings[...]`, `app.settings`, `self._settings`, and the local names assigned
     from those), top level and one level down through the known sub-dicts; a key the
     schema does not know fails with file and line. KNOWN_HITS is the explicit allow-list
     for what the first run found and was judged not a bug; a stale entry fails too.

Headless. Run:
    python tools/test_settings_schema.py
"""
import ast
import copy
import json
import logging
import os
import re
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Never the user's real settings folder: an isolated profile, set before any
# project module is imported (tools/test_config_isolation.py checks this).
PROFILE = tempfile.mkdtemp(prefix="ssc-settings-schema-")
os.environ["SAXSHOP_CONFIG_DIR"] = PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

import config  # noqa: E402
import i18n  # noqa: E402
import settings_schema as ss  # noqa: E402
import svg_engine  # noqa: E402
import toner_engine  # noqa: E402
import toner_tab  # noqa: E402
import tuner_tab  # noqa: E402
from config import DEFAULT_SETTINGS  # noqa: E402

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
results = []


def check(name, fn):
    try:
        fn()
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


print("Settings schema and the key gate")
print("=" * 60)


# ---- 1. coverage ----------------------------------------------------------------------------------
def cover(schema, defaults, path=""):
    bad = []
    for k in sorted(set(schema) - set(defaults)):
        bad.append(f"{path}{k}: in the schema, not in DEFAULT_SETTINGS")
    for k in sorted(set(defaults) - set(schema)):
        bad.append(f"{path}{k}: in DEFAULT_SETTINGS, not in the schema")
    for k in sorted(set(schema) & set(defaults)):
        f = schema[k]
        if not f.accepts(defaults[k]):
            bad.append(f"{path}{k}: the default {defaults[k]!r} fails its own field ({f.describe()})")
        if f.kind == "dict" and isinstance(defaults[k], dict):
            if f.keys is None and defaults[k]:
                bad.append(f"{path}{k}: a dict default with no key schema")
            elif f.keys is not None:
                bad += cover(f.keys, defaults[k], path + k + ".")
    return bad


def schema_covers_the_defaults_both_ways():
    bad = cover(ss.SCHEMA, DEFAULT_SETTINGS)
    assert not bad, "\n    " + "\n    ".join(bad)
    assert len(ss.SCHEMA) == len(DEFAULT_SETTINGS) == 107, len(ss.SCHEMA)
check("the schema covers exactly DEFAULT_SETTINGS' keys, both ways, at every level; every default passes",
      schema_covers_the_defaults_both_ways)


def a_missing_or_extra_key_is_caught():
    schema = dict(ss.SCHEMA)
    del schema["units"]
    assert any("units: in DEFAULT_SETTINGS, not in the schema" in b for b in cover(schema, DEFAULT_SETTINGS))
    schema = dict(ss.SCHEMA, brand_new_key=ss.Bool())
    assert any("brand_new_key: in the schema, not in DEFAULT_SETTINGS" in b for b in cover(schema, DEFAULT_SETTINGS))
    defaults = copy.deepcopy(DEFAULT_SETTINGS)
    defaults["tuner_settings"]["new_sub"] = 1
    assert any("tuner_settings.new_sub: in DEFAULT_SETTINGS" in b for b in cover(ss.SCHEMA, defaults))
    defaults = copy.deepcopy(DEFAULT_SETTINGS)
    defaults["units"] = "furlongs"
    assert any("units: the default" in b for b in cover(ss.SCHEMA, defaults))
check("the coverage check fails on a key in one and not the other, at either level, and on a default off its field",
      a_missing_or_extra_key_is_caught)


# ---- 2. the enumerations are the code's ---------------------------------------------------------
def enumerations_match_their_sources():
    assert ss.LANGUAGES == tuple(i18n.LANGUAGE_NAMES), (ss.LANGUAGES, tuple(i18n.LANGUAGE_NAMES))
    assert ss.LOCATOR_STYLES == tuple(svg_engine.LOCATOR_STYLES)
    assert ss.TRANSPOSITIONS == tuple(tuner_tab.TRANSPOSITION_SHIFTS)
    assert ss.SAX_TYPES == tuple(toner_engine.SAX_TYPES)
    assert set(ss.SAX_TYPES) == set(toner_engine.SAX_NOTE_RANGES)
    assert ss.TONER_FPS == tuple(toner_tab.FRAME_RATES)
    import tooling_tab
    assert ss.ORGANIZER_MATERIALS == tuple(tooling_tab._ORGANIZER_MATERIALS)
    assert ss.CRESCENT_MATERIALS == tuple(tooling_tab._CRESCENT_MATERIALS)
    assert ss.GCODE_MATERIALS == tuple(config.GCODE_PRESET_MATERIALS)
    assert ss.CUT_ONLY_MATERIALS == tuple(config.GCODE_CUT_ONLY_MATERIALS)
    assert set(ss.GCODE_BLOCK.keys) == set(config.GCODE_PRESET_KEYS), set(ss.GCODE_BLOCK.keys) ^ set(config.GCODE_PRESET_KEYS)
    assert set(ss.CUT_ONLY_BLOCK.keys) == set(config.GCODE_CUT_ONLY_KEYS)
    assert set(ss.SCHEMA["gcode_settings"].keys) == set(DEFAULT_SETTINGS["gcode_settings"])
    assert set(ss.SCHEMA["laser_framing_power_by_material"].keys) == set(DEFAULT_SETTINGS["laser_framing_power_by_material"])
    # literals that live inside methods: read from the source
    tuner_src = open(os.path.join(REPO, "tuner_tab.py"), encoding="utf-8").read()
    m = re.search(r"fps_values = \[([^\]]*)\]", tuner_src)
    assert m and tuple(re.findall(r'"(\d+)"', m.group(1))) == ss.TUNER_FPS, m and m.group(1)
    toner_src = open(os.path.join(REPO, "toner_tab.py"), encoding="utf-8").read()
    assert re.search(r'_make_vswitch\(left_col, "VIEW", self._toner_view_var,\s*"spectrum", "bars"', toner_src)
    assert re.search(r'_make_vswitch\(left_col, "SCALE", self._toner_scale_var,\s*"linear", "db"', toner_src)
    main_src = open(os.path.join(REPO, "main.py"), encoding="utf-8").read()
    for opt in ss.HOLE_OPTIONS:
        assert f'"{opt}"' in main_src, f"hole option {opt!r} is not in main.py"
    for bias in ss.EDGE_BIASES:
        assert re.search(rf'["\']{bias}["\']', main_src) or re.search(rf'["\']{bias}["\']', svg_engine.__dict__.get("__file__") and open(svg_engine.__file__, encoding="utf-8").read()), bias
    sizing_src = open(os.path.join(REPO, "ui_dialogs", "sizing.py"), encoding="utf-8").read()
    for mode in ss.PLACEMENT_MODES:
        assert f'value="{mode}"' in sizing_src, mode
    for unit in ss.UNITS:
        assert f'value="{unit}"' in sizing_src, unit
    assert ss.WAVEFORMS == ("pure", "rich")
    assert '"rich"' in open(os.path.join(REPO, "tuner_engine.py"), encoding="utf-8").read()
check("the enumerations are the code's own (languages, styles, transpositions, sax types, fps, materials, preset keys, switches)",
      enumerations_match_their_sources)


# ---- 3. validation ---------------------------------------------------------------------------------
def validate(settings):
    return ss.validate_settings(settings, DEFAULT_SETTINGS)


def the_defaults_come_through_untouched():
    s = copy.deepcopy(DEFAULT_SETTINGS)
    assert validate(s) == [] and s == DEFAULT_SETTINGS
check("validation leaves the defaults exactly as they are", the_defaults_come_through_untouched)


def planted_faults_are_corrected_and_nothing_else():
    s = copy.deepcopy(DEFAULT_SETTINGS)
    s["units"] = "furlongs"                                   # off its enumeration
    s["felt_offset"] = "0.75"                                 # a string where a number is declared
    s["engraving_on"] = 1                                     # an int where a bool is declared
    s["capture_threshold"] = 250                              # out of its range
    s["dart_shape_factor"] = True                             # a bool where a number is declared
    s["welcomed_version"] = None                              # None where it is not allowed
    s["camera_index_override"] = None                         # None where it IS allowed: untouched
    s["gcode_settings"]["felt"]["cut_passes"] = 0             # inside a block, below its floor
    s["gcode_settings"]["felt"]["cut_speed"] = 613            # an int where a float is declared: accepted
    s["gcode_settings"]["cork"]["engraving_speed"] = 0        # cork's "not set": accepted
    s["gcode_settings"]["leather"]["engraving_mode"] = "dotted"
    s["gcode_settings"]["leather"]["future_key"] = 7          # unknown to the schema: left alone
    s["tuner_settings"]["fps"] = "45"
    s["tuner_settings"]["sensitivity"] = 50.0                 # a float where an int is declared
    s["toner_settings"]["analysis_descriptors"]["warmth"] = "yes"
    s["layer_colors"]["felt_outline"] = "red"
    s["layer_colors"]["card_outline"] = "#F00"                # the short form is a colour too: accepted
    s["dart_ranges"] = [{"min_size": 7.0, "max_size": 12.0}, "not a range"]   # a bad element: the list goes
    s["sizing_ranges"] = [{"min_size": 7.0, "max_size": 12.0}]                 # a good list: untouched
    s["laser_framing_power_by_material"]["felt"] = "12"
    s["laser_framing_power_by_material"]["linen"] = "x"      # an extra key with a bad value: removed
    s["laser_framing_power_by_material"]["denim"] = 15        # an extra key with a good value: kept
    s["camera_calibration_engrave_offset_mm"] = [1.0]         # the wrong length
    s["visible_sax_types"] = ["Alto", 3]
    s["felt_thickness"] = float("nan")
    s["tooling_settings"]["holder_layer_count"] = 4
    s["tooling_settings"]["crescent_curve"] = 101
    s["tooling_settings"]["feeds_speeds_tester"]["sheet_unit"] = "cm"
    corrections = validate(s)
    paths = sorted(p for p, _v, _f in corrections)
    expected = sorted([
        "units", "felt_offset", "engraving_on", "capture_threshold", "dart_shape_factor", "welcomed_version",
        "gcode_settings.felt.cut_passes", "gcode_settings.leather.engraving_mode", "tuner_settings.fps",
        "tuner_settings.sensitivity", "toner_settings.analysis_descriptors.warmth", "layer_colors.felt_outline",
        "dart_ranges", "laser_framing_power_by_material.felt", "laser_framing_power_by_material.linen",
        "camera_calibration_engrave_offset_mm", "visible_sax_types", "felt_thickness",
        "tooling_settings.holder_layer_count", "tooling_settings.crescent_curve",
        "tooling_settings.feeds_speeds_tester.sheet_unit",
    ])
    assert paths == expected, f"\n    got      {paths}\n    expected {expected}"
    # Each corrected value is the default at its path; what was accepted stayed.
    assert s["units"] == "in" and s["felt_offset"] == 0.75 and s["engraving_on"] is True
    assert s["capture_threshold"] == 50 and s["dart_shape_factor"] == 0.5 and s["welcomed_version"] == ""
    assert s["camera_index_override"] is None
    assert s["gcode_settings"]["felt"]["cut_passes"] == 1 and s["gcode_settings"]["felt"]["cut_speed"] == 613
    assert s["gcode_settings"]["cork"]["engraving_speed"] == 0
    assert s["gcode_settings"]["leather"]["engraving_mode"] == "line" and s["gcode_settings"]["leather"]["future_key"] == 7
    assert s["tuner_settings"]["fps"] == "60" and s["tuner_settings"]["sensitivity"] == 50
    assert s["toner_settings"]["analysis_descriptors"]["warmth"] is True
    assert s["layer_colors"]["felt_outline"] == "#000000" and s["layer_colors"]["card_outline"] == "#F00"
    assert s["dart_ranges"] == [] and s["sizing_ranges"] == [{"min_size": 7.0, "max_size": 12.0}]
    assert s["laser_framing_power_by_material"]["felt"] == 10
    assert "linen" not in s["laser_framing_power_by_material"] and s["laser_framing_power_by_material"]["denim"] == 15
    assert s["camera_calibration_engrave_offset_mm"] is None and s["visible_sax_types"] is None
    assert s["felt_thickness"] == 3.175
    assert s["tooling_settings"]["holder_layer_count"] == 6 and s["tooling_settings"]["crescent_curve"] == 75.0
    assert s["tooling_settings"]["feeds_speeds_tester"]["sheet_unit"] == "in"
    # the report carries the bad value and a description of what was expected
    by_path = {p: (v, f) for p, v, f in corrections}
    assert by_path["units"][0] == "furlongs" and by_path["units"][1].describe() == "one of in, cm, mm"
    assert by_path["capture_threshold"][1].describe() == "a whole number 0 to 100"
    assert by_path["gcode_settings.felt.cut_passes"][1].describe() == "a whole number of at least 1"
    # everything not planted is still the default
    planted_top = {"units", "felt_offset", "engraving_on", "capture_threshold", "dart_shape_factor", "welcomed_version",
                   "camera_index_override", "gcode_settings", "tuner_settings", "toner_settings", "layer_colors",
                   "dart_ranges", "sizing_ranges", "laser_framing_power_by_material", "camera_calibration_engrave_offset_mm",
                   "visible_sax_types", "felt_thickness", "tooling_settings"}
    for k in DEFAULT_SETTINGS:
        if k not in planted_top:
            assert s[k] == DEFAULT_SETTINGS[k], k
check("planted faults of every kind are put back to their defaults and reported; the rest is untouched",
      planted_faults_are_corrected_and_nothing_else)


def numbers_are_numbers():
    f = ss.Float(min=0)
    assert f.accepts(0) and f.accepts(2) and f.accepts(2.5) and not f.accepts(True) and not f.accepts(False)
    assert not f.accepts(-0.1) and not f.accepts(float("inf")) and not f.accepts("3") and not f.accepts(None)
    i = ss.Int(min=1)
    assert i.accepts(1) and not i.accepts(1.0) and not i.accepts(True) and not i.accepts(0)
    assert ss.Int(nullable=True).accepts(None) and not ss.Int().accepts(None)
    assert ss.Float(gt=0).accepts(0.001) and not ss.Float(gt=0).accepts(0)
    assert ss.Int(choices=(5, 6)).accepts(6) and not ss.Int(choices=(5, 6)).accepts(7)
check("an int is a number, a bool and NaN are not; None only where allowed; floors, ceilings and choices hold",
      numbers_are_numbers)


# ---- 4. load_settings runs it -----------------------------------------------------------------------
class Grab(logging.Handler):
    def __init__(self):
        super().__init__()
        self.lines = []

    def emit(self, record):
        self.lines.append(record.getMessage())


def write_profile(data):
    with open(config.SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f)


def load_settings_validates_and_logs():
    grab = Grab()
    logging.getLogger("ssc.settings").addHandler(grab)
    try:
        bad = copy.deepcopy(DEFAULT_SETTINGS)
        bad["units"] = "furlongs"
        bad["gcode_settings"]["felt"]["cut_passes"] = 0
        bad["tuner_settings"]["fps"] = "45"
        bad["edge_bias"] = "up"
        write_profile(bad)
        s = config.load_settings()
        assert s["units"] == "in" and s["gcode_settings"]["felt"]["cut_passes"] == 1
        assert s["tuner_settings"]["fps"] == "60" and s["edge_bias"] == "center"
        assert len(grab.lines) == 4, grab.lines
        for ln in grab.lines:
            assert config.SETTINGS_FILE in ln and "the default is used" in ln, ln
        assert any("units was 'furlongs', not one of in, cm, mm" in ln for ln in grab.lines), grab.lines
        assert any("gcode_settings.felt.cut_passes was 0, not a whole number of at least 1" in ln for ln in grab.lines)
        # a valid file: nothing logged, nothing changed
        del grab.lines[:]
        good = copy.deepcopy(DEFAULT_SETTINGS)
        good["units"] = "mm"
        good["gcode_settings"]["felt"]["cut_speed"] = 613
        write_profile(good)
        s = config.load_settings()
        assert s == good and grab.lines == [], grab.lines
    finally:
        logging.getLogger("ssc.settings").removeHandler(grab)
        if os.path.exists(config.SETTINGS_FILE):
            os.remove(config.SETTINGS_FILE)
check("load_settings validates after the merge, logs each correction with key, value and file; a valid file is silent",
      load_settings_validates_and_logs)


def the_hardening_and_migrations_run_first():
    # The old hardening still does its work before validation: a null, a non-dict block,
    # a LightBurn colour code and the dart-shape migration.
    bad = {"units": None, "gcode_settings": "nope", "layer_colors": {"felt_outline": "C10"},
           "dart_shape_factor": 0.0, "dart_shape_v2": False}
    write_profile(bad)
    grab = Grab()
    logging.getLogger("ssc.settings").addHandler(grab)
    try:
        s = config.load_settings()
    finally:
        logging.getLogger("ssc.settings").removeHandler(grab)
        os.remove(config.SETTINGS_FILE)
    assert s["units"] == "in" and s["gcode_settings"] == DEFAULT_SETTINGS["gcode_settings"]
    assert s["layer_colors"]["felt_outline"] == "#000000"
    assert s["dart_shape_factor"] == 0.5 and s["dart_shape_v2"] is True        # migrated from the old scale
    assert grab.lines == [], f"validation found something the hardening should have fixed first: {grab.lines}"
check("null rejection, the non-dict skip, the LightBurn colours and the dart-shape migration still run before validation",
      the_hardening_and_migrations_run_first)


def a_fault_in_the_check_cannot_cost_the_file():
    write_profile(dict(DEFAULT_SETTINGS, units="mm"))
    keep = ss.validate_settings

    def broken(*a, **k):
        raise RuntimeError("the schema is on fire")
    ss.validate_settings = broken
    grab = Grab()
    logging.getLogger("ssc.settings").addHandler(grab)
    try:
        s = config.load_settings()
    finally:
        ss.validate_settings = keep
        logging.getLogger("ssc.settings").removeHandler(grab)
        os.remove(config.SETTINGS_FILE)
    assert s["units"] == "mm", "the user's values must survive a failing check"
    assert any("validation failed" in ln for ln in grab.lines), grab.lines
    assert not os.path.exists(config.SETTINGS_FILE + ".corrupt.bak"), "a failing check must not archive the file as corrupt"
check("a fault in the check itself is logged and the merged values are used; the file is not archived as corrupt",
      a_fault_in_the_check_cannot_cost_the_file)


def the_keys_the_first_scan_found_now_survive_a_restart():
    # The first run of the gate (2026-11-10) found six keys the app saved but never
    # declared in DEFAULT_SETTINGS, so load_settings dropped them on every start: the
    # Pad Maker's "Live camera preview" tick, the Die Organizer's material, sheet and
    # units, the Tonehole Crescent's material. A profile holding them must come back
    # with them; before the fix every one came back as if never set.
    saved = copy.deepcopy(DEFAULT_SETTINGS)
    saved["show_live_camera"] = True
    saved["tooling_settings"].update({"organizer_material": "acrylic", "organizer_sheet_width": "10",
                                      "organizer_sheet_height": "16", "organizer_sheet_units": "in",
                                      "crescent_material": "leather"})
    write_profile(saved)
    try:
        s = config.load_settings()
    finally:
        os.remove(config.SETTINGS_FILE)
    assert s["show_live_camera"] is True, "the live-camera tick is forgotten on restart"
    t = s["tooling_settings"]
    assert (t["organizer_material"], t["organizer_sheet_width"], t["organizer_sheet_height"], t["organizer_sheet_units"],
            t["crescent_material"]) == ("acrylic", "10", "16", "in", "leather"), t
    # and the empty sheet defaults still mean "size it from the app's units": the tab reads them with `or`
    src = open(os.path.join(REPO, "tooling_tab.py"), encoding="utf-8").read()
    assert 'tooling.get("organizer_sheet_width") or (' in src and 'tooling.get("organizer_sheet_height") or (' in src
check("the six keys the first scan found (live-camera tick, organizer material/sheet/units, crescent material) "
      "survive a restart", the_keys_the_first_scan_found_now_survive_a_restart)


# ---- 5. the gate: every literal settings key in the code is in the schema -----------------------------
APP_MODULES = sorted(
    [f for f in os.listdir(REPO) if f.endswith(".py") and f not in ("build.py", "settings_schema.py")]
    + ["ui_dialogs/" + f for f in os.listdir(os.path.join(REPO, "ui_dialogs")) if f.endswith(".py")])

# What the first run found and was judged not a bug (2026-11-10), by (file, key): the reason.
# A hit not listed here fails; an entry here that is no longer hit fails too.
KNOWN_HITS = {
}

TOP = ss.Dict(ss.SCHEMA)
SETTINGS_ATTRS = ("settings", "_settings")


def uniform_child(field):
    """The Field a dynamic key reaches under this dict: its extra Field, or the one its known
    keys share (the largest block when they are dicts, as gcode_settings' are)."""
    if field.extra is not None:
        return field.extra
    if not field.keys:
        return None
    children = list(field.keys.values())
    if all(c.kind == "dict" for c in children):
        return max(children, key=lambda c: len(c.keys or {}))
    if all(c is children[0] for c in children) or len({c.kind for c in children}) == 1:
        return children[0]
    return None


class Scan(ast.NodeVisitor):
    """One module: every literal key read from or written to a settings dict."""

    def __init__(self, path):
        self.path = path
        self.hits = []          # (line, key, where)
        self.seen = set()       # (line, column, key) already looked up: a nested access is classified from both ends
        self.unclassified = 0   # literal accesses to a dict we could not tell is the settings, whose key the schema knows
        self.aliases = [{}]     # per function: local name -> Field (a settings dict or a known sub-dict)

    @property
    def checked(self):
        """Literal accesses that were checked against the schema."""
        return len(self.seen)

    # -- classification
    def classify(self, node):
        """The Field `node` evaluates to when it is (part of) the settings dict, else None."""
        if isinstance(node, ast.Name):
            if node.id in ("settings", "DEFAULT_SETTINGS"):
                return TOP
            return self.aliases[-1].get(node.id)
        if isinstance(node, ast.Attribute):
            if node.attr in SETTINGS_ATTRS or node.attr == "DEFAULT_SETTINGS":
                return TOP
            return None
        if isinstance(node, ast.Subscript):
            base = self.classify(node.value)
            if base is None or base.kind != "dict":
                return None
            key = node.slice
            if isinstance(key, ast.Constant) and isinstance(key.value, str):
                return self.lookup(base, key.value, key)
            return uniform_child(base)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and node.func.attr in ("get", "setdefault", "pop"):
            base = self.classify(node.func.value)
            if base is None or base.kind != "dict" or not node.args:
                return None
            key = node.args[0]
            if isinstance(key, ast.Constant) and isinstance(key.value, str):
                return self.lookup(base, key.value, key)
            return uniform_child(base)
        return None

    def lookup(self, base, key, key_node):
        """Check `key` against `base` (a dict Field); the child Field, or None."""
        first = (key_node.lineno, key_node.col_offset, key) not in self.seen
        self.seen.add((key_node.lineno, key_node.col_offset, key))
        if base.keys is not None and key in base.keys:
            return base.keys[key]
        if base.extra is not None:
            return base.extra
        if base.keys is None:
            return None         # an unschema'd dict (a range rule): anything goes
        if first:
            where = "top level" if base is TOP else "a nested block"
            self.hits.append((key_node.lineno, key, where))
        return None

    # -- what to look at
    def visit_FunctionDef(self, node):
        self.aliases.append(dict(self.aliases[-1]))
        self.generic_visit(node)
        self.aliases.pop()

    visit_AsyncFunctionDef = visit_FunctionDef

    def visit_Assign(self, node):
        field = self.classify(node.value)
        for t in node.targets:
            if isinstance(t, ast.Name):
                if field is not None and field.kind == "dict":
                    self.aliases[-1][t.id] = field
                else:
                    self.aliases[-1].pop(t.id, None)
        self.generic_visit(node)

    def visit_Subscript(self, node):
        self.classify(node)          # checks the literal key, if any
        self.note_unclassified(node.value, node.slice)
        self.generic_visit(node)

    def visit_Call(self, node):
        if isinstance(node.func, ast.Attribute) and node.func.attr in ("get", "setdefault", "pop") and node.args:
            self.classify(node)
            self.note_unclassified(node.func.value, node.args[0])
        self.generic_visit(node)

    def visit_Compare(self, node):
        if (isinstance(node.left, ast.Constant) and isinstance(node.left.value, str) and len(node.ops) == 1
                and isinstance(node.ops[0], (ast.In, ast.NotIn))):
            base = self.classify(node.comparators[0])
            if base is not None and base.kind == "dict":
                self.lookup(base, node.left.value, node.left)
        self.generic_visit(node)

    def note_unclassified(self, base_node, key_node):
        """A literal key the schema knows, read off a dict the scan cannot tell is the settings
        (a range rule, a material block passed under its own name, a preset): counted by the
        base's name, so the report can say what the gate does not cover."""
        if isinstance(key_node, ast.Constant) and isinstance(key_node.value, str) and self.classify(base_node) is None:
            if key_node.value in ss.SCHEMA or any(key_node.value in (f.keys or {}) for f in ss.SCHEMA.values() if f.kind == "dict"):
                self.unclassified += 1
                name = base_node.id if isinstance(base_node, ast.Name) else (
                    "." + base_node.attr if isinstance(base_node, ast.Attribute) else type(base_node).__name__)
                UNCLASSIFIED_BASES[name] = UNCLASSIFIED_BASES.get(name, 0) + 1


UNCLASSIFIED_BASES = {}


def run_scan():
    hits, checked, unclassified = [], 0, 0
    for rel in APP_MODULES:
        src = open(os.path.join(REPO, rel), encoding="utf-8").read()
        scan = Scan(rel)
        scan.visit(ast.parse(src))
        hits += [(rel, line, key, where) for line, key, where in scan.hits]
        checked += scan.checked
        unclassified += scan.unclassified
    return hits, checked, unclassified


def the_gate():
    hits, checked, unclassified = run_scan()
    top = sorted(UNCLASSIFIED_BASES.items(), key=lambda kv: -kv[1])[:8]
    print(f"        {checked} literal key accesses checked in {len(APP_MODULES)} modules; "
          f"{unclassified} on dicts the scan cannot tell are the settings "
          f"({', '.join(f'{n} {c}' for n, c in top)})")
    unknown = [(rel, line, key, where) for rel, line, key, where in hits if (rel, key) not in KNOWN_HITS]
    stale = [k for k in KNOWN_HITS if not any((rel, key) == k for rel, _l, key, _w in hits)]
    for rel, line, key, where in unknown:
        print(f"        {rel}:{line}: {key!r} is not in the schema ({where})")
    assert not unknown, f"{len(unknown)} settings key(s) the schema does not know (above)"
    assert not stale, f"allow-list entries no longer hit: {stale}"
check("every string-literal settings key in the app's modules is in the schema (or listed, with its reason)", the_gate)


def the_gate_catches_a_typo():
    src = ("def f(self, settings):\n"
           "    a = settings['feld_offset']\n"
           "    b = self.settings.get('units')\n"
           "    ts = self.settings['tuner_settings']\n"
           "    c = ts.get('fpz')\n"
           "    d = settings['gcode_settings']['felt']['cut_sped']\n"
           "    e = settings['gcode_settings'][m]['kerf_width']\n"
           "    f = 'edge_bais' in settings\n"
           "    settings['brand_new_runtime_key'] = 1\n"
           "    g = other['feld_offset']\n")
    scan = Scan("typo.py")
    scan.visit(ast.parse(src))
    keys = sorted(k for _l, k, _w in scan.hits)
    assert keys == ["brand_new_runtime_key", "cut_sped", "edge_bais", "feld_offset", "fpz"], keys
    # eleven literal keys looked up (the dynamic [m] is not one); the other dict's is ignored
    assert scan.unclassified == 0 and scan.checked == 11, (scan.unclassified, scan.checked)
check("the scan catches a typo at the top level, in a sub-dict (through a local alias), two levels down, in an "
      "`in` test and in a write, and ignores other dicts", the_gate_catches_a_typo)


print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
