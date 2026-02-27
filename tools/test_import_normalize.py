"""Test pad preset import unwrapping and normalization."""
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import PadSVGGeneratorApp

passed = 0
failed = 0

def test(name, actual, expected):
    global passed, failed
    if actual == expected:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        print(f"    Expected: {expected!r}")
        print(f"    Got:      {actual!r}")
        failed += 1

# --- normalize_pad_text ---
print("=== _normalize_pad_text ===")

# Comma-separated SIZExQTY
result = PadSVGGeneratorApp._normalize_pad_text("36x2, 34x2, 30x2, 26x2, 22, 20x2")
test("comma-separated SIZExQTY", result, "36.0 x 2\n34.0 x 2\n30.0 x 2\n26.0 x 2\n22.0 x 1\n20.0 x 2\n")

# Already standard format
result = PadSVGGeneratorApp._normalize_pad_text("36.0 x 2\n34.0 x 2\n")
test("already standard format", result, "36.0 x 2\n34.0 x 2\n")

# SIZExQTY without commas (newline separated)
result = PadSVGGeneratorApp._normalize_pad_text("36x2\n34x2\n30x1")
test("newline SIZExQTY", result, "36.0 x 2\n34.0 x 2\n30.0 x 1\n")

# With trailing whitespace and blank lines
result = PadSVGGeneratorApp._normalize_pad_text("  36.0 x 2  \n\n34.0 x 1\n\n\n")
test("trailing whitespace/blanks", result, "36.0 x 2\n34.0 x 1\n")

# Max fill
result = PadSVGGeneratorApp._normalize_pad_text("18.0 x max\n36x2")
test("max fill preserved", result, "18.0 x max\n36.0 x 2\n")

# Empty/None
result = PadSVGGeneratorApp._normalize_pad_text("")
test("empty string passthrough", result, "")
result = PadSVGGeneratorApp._normalize_pad_text("  \n  \n")
test("whitespace-only passthrough", result, "  \n  \n")

# Decimals in comma format
result = PadSVGGeneratorApp._normalize_pad_text("48.2x2, 43.2, 40.2x3, 36.2")
test("decimals in comma format", result, "48.2 x 2\n43.2 x 1\n40.2 x 3\n36.2 x 1\n")

# --- _unwrap_pad_presets ---
print("\n=== _unwrap_pad_presets ===")

# Library-wrapped format (like Music Center file)
wrapped = {
    "Music Center": {
        "Soprano - Yamaha YSS-62": {
            "pads": "36x2, 34x2, 30x2",
            "notes": "via Oscar"
        },
        "Alto - Jupiter": {
            "pads": "48.2x2, 43.2, 40.2x3",
            "notes": "via Oscar"
        }
    }
}
suggested, flat = PadSVGGeneratorApp._unwrap_pad_presets(wrapped)
test("library name detected", suggested, "Music Center")
test("presets unwrapped count", len(flat), 2)
test("preset names correct", sorted(flat.keys()), ["Alto - Jupiter", "Soprano - Yamaha YSS-62"])
test("pad text normalized", flat["Soprano - Yamaha YSS-62"]["pads"], "36.0 x 2\n34.0 x 2\n30.0 x 2\n")
test("notes preserved", flat["Soprano - Yamaha YSS-62"]["notes"], "via Oscar")

# Already flat format
flat_input = {
    "My Preset": {"pads": "36x2, 34x2", "notes": "test"},
    "Other": "48.0 x 2\n44.0 x 1\n"
}
suggested, flat = PadSVGGeneratorApp._unwrap_pad_presets(flat_input)
test("flat format: no suggested lib", suggested, None)
test("flat format: count preserved", len(flat), 2)
test("flat dict normalized", flat["My Preset"]["pads"], "36.0 x 2\n34.0 x 2\n")
test("flat string normalized", flat["Other"], "48.0 x 2\n44.0 x 1\n")

# Multiple libraries in one file
multi_lib = {
    "Lib A": {"Preset 1": {"pads": "36x2", "notes": ""}},
    "Lib B": {"Preset 2": {"pads": "48x1", "notes": ""}}
}
suggested, flat = PadSVGGeneratorApp._unwrap_pad_presets(multi_lib)
test("multi-lib: no single suggested name", suggested, None)
test("multi-lib: all presets merged", len(flat), 2)

# Empty dict
suggested, flat = PadSVGGeneratorApp._unwrap_pad_presets({})
test("empty: no suggestion", suggested, None)
test("empty: empty result", flat, {})

# --- Test with actual Music Center file ---
print("\n=== Music Center file test ===")
mc_path = r"C:\Users\abadc\Downloads\music_center_pad_presets.json"
if os.path.exists(mc_path):
    with open(mc_path) as f:
        mc_data = json.load(f)
    suggested, flat = PadSVGGeneratorApp._unwrap_pad_presets(mc_data)
    test("MC file: library detected", suggested, "Music Center")
    test("MC file: has presets", len(flat) > 0, True)
    print(f"    ({len(flat)} presets found)")
    # Check first preset is normalized
    first_name = sorted(flat.keys())[0]
    first_data = flat[first_name]
    pads = first_data["pads"] if isinstance(first_data, dict) else first_data
    has_newlines = '\n' in pads
    no_commas = ',' not in pads
    has_spaces = ' x ' in pads
    test(f"MC file: '{first_name}' has newlines", has_newlines, True)
    test(f"MC file: '{first_name}' no commas", no_commas, True)
    test(f"MC file: '{first_name}' has ' x ' spacing", has_spaces, True)
    # Print a sample
    print(f"\n  Sample preset '{first_name}':")
    for line in pads.strip().split('\n')[:5]:
        print(f"    {line}")
    if pads.strip().count('\n') > 4:
        print(f"    ... ({pads.strip().count(chr(10)) + 1} lines total)")
else:
    print(f"  SKIP: {mc_path} not found")

# --- Summary ---
print(f"\n{'='*40}")
print(f"Results: {passed} passed, {failed} failed")
if failed:
    sys.exit(1)
