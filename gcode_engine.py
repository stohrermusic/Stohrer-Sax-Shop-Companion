"""
G-code generation engine for Stohrer Sax Shop Companion.

Generates Grbl-compatible G-code for laser cutting pad materials.
Includes a single-stroke digit font for engraving pad sizes.
"""

import math

# =============================================================================
# SINGLE-STROKE DIGIT FONT
# =============================================================================
# Each digit is defined as a list of strokes (pen-down segments).
# Each stroke is a list of (x, y) points in a unit coordinate system.
# Character cell is approximately 0.6 wide x 1.0 tall.
# Origin (0, 0) is bottom-left of character cell.

STROKE_FONT = {
    '0': [
        [(0.1, 0.2), (0.1, 0.8), (0.2, 0.95), (0.4, 0.95), (0.5, 0.8), (0.5, 0.2), (0.4, 0.05), (0.2, 0.05), (0.1, 0.2)]
    ],
    '1': [
        [(0.15, 0.75), (0.3, 0.95), (0.3, 0.05)],
        [(0.15, 0.05), (0.45, 0.05)]
    ],
    '2': [
        [(0.1, 0.75), (0.15, 0.9), (0.3, 0.95), (0.45, 0.9), (0.5, 0.75), (0.5, 0.6), (0.1, 0.05), (0.5, 0.05)]
    ],
    '3': [
        [(0.1, 0.85), (0.2, 0.95), (0.4, 0.95), (0.5, 0.85), (0.5, 0.6), (0.4, 0.5), (0.25, 0.5)],
        [(0.4, 0.5), (0.5, 0.4), (0.5, 0.15), (0.4, 0.05), (0.2, 0.05), (0.1, 0.15)]
    ],
    '4': [
        [(0.4, 0.05), (0.4, 0.95), (0.1, 0.35), (0.5, 0.35)]
    ],
    '5': [
        [(0.5, 0.95), (0.1, 0.95), (0.1, 0.55), (0.35, 0.6), (0.5, 0.45), (0.5, 0.2), (0.4, 0.05), (0.2, 0.05), (0.1, 0.15)]
    ],
    '6': [
        [(0.45, 0.9), (0.3, 0.95), (0.15, 0.85), (0.1, 0.6), (0.1, 0.2), (0.2, 0.05), (0.4, 0.05), (0.5, 0.2), (0.5, 0.35), (0.4, 0.5), (0.2, 0.5), (0.1, 0.35)]
    ],
    '7': [
        [(0.1, 0.95), (0.5, 0.95), (0.25, 0.05)]
    ],
    '8': [
        [(0.25, 0.5), (0.15, 0.6), (0.15, 0.8), (0.25, 0.95), (0.4, 0.95), (0.5, 0.8), (0.5, 0.6), (0.4, 0.5), (0.25, 0.5), (0.1, 0.4), (0.1, 0.15), (0.2, 0.05), (0.4, 0.05), (0.5, 0.15), (0.5, 0.4), (0.4, 0.5)]
    ],
    '9': [
        [(0.15, 0.1), (0.3, 0.05), (0.45, 0.15), (0.5, 0.4), (0.5, 0.8), (0.4, 0.95), (0.2, 0.95), (0.1, 0.8), (0.1, 0.65), (0.2, 0.5), (0.4, 0.5), (0.5, 0.65)]
    ],
    '.': [
        [(0.2, 0.05), (0.2, 0.15), (0.3, 0.15), (0.3, 0.05), (0.2, 0.05)]
    ],
    '-': [
        [(0.1, 0.5), (0.5, 0.5)]
    ],
}

# Character width for spacing (proportional)
CHAR_WIDTHS = {
    '0': 0.6, '1': 0.5, '2': 0.6, '3': 0.6, '4': 0.6, '5': 0.6,
    '6': 0.6, '7': 0.6, '8': 0.6, '9': 0.6, '.': 0.4, '-': 0.6
}

DEFAULT_CHAR_WIDTH = 0.6
CHAR_HEIGHT = 1.0
CHAR_SPACING = 0.1  # Space between characters as fraction of height


def get_text_strokes(text, font_size_mm, center_x, center_y):
    """
    Convert text (numbers) to stroke paths for G-code.

    Args:
        text: String to render (should be numeric like "42" or "18.5")
        font_size_mm: Height of characters in mm
        center_x, center_y: Center position for the text in mm

    Returns:
        List of strokes, where each stroke is a list of (x, y) points in mm
    """
    strokes = []

    # Calculate total width of text
    total_width = 0
    for char in text:
        total_width += CHAR_WIDTHS.get(char, DEFAULT_CHAR_WIDTH)
    total_width += CHAR_SPACING * (len(text) - 1) if len(text) > 1 else 0
    total_width *= font_size_mm

    # Starting X position (left edge of first character)
    start_x = center_x - total_width / 2
    current_x = start_x

    # Y offset to center vertically
    y_offset = center_y - (font_size_mm * CHAR_HEIGHT) / 2

    for char in text:
        if char in STROKE_FONT:
            char_strokes = STROKE_FONT[char]
            char_width = CHAR_WIDTHS.get(char, DEFAULT_CHAR_WIDTH) * font_size_mm

            for stroke in char_strokes:
                transformed_stroke = []
                for px, py in stroke:
                    # Scale and translate
                    x = current_x + px * font_size_mm
                    y = y_offset + py * font_size_mm
                    transformed_stroke.append((x, y))
                strokes.append(transformed_stroke)

            current_x += char_width + CHAR_SPACING * font_size_mm
        else:
            # Skip unknown characters, but advance position
            current_x += DEFAULT_CHAR_WIDTH * font_size_mm + CHAR_SPACING * font_size_mm

    return strokes


# =============================================================================
# CIRCLE LINEARIZATION
# =============================================================================

def linearize_circle(cx, cy, radius, segments=72):
    """
    Convert a circle to a list of line segment points.

    Args:
        cx, cy: Center coordinates in mm
        radius: Radius in mm
        segments: Number of line segments (default 72 = 5-degree steps)

    Returns:
        List of (x, y) points forming a closed polygon
    """
    points = []
    for i in range(segments + 1):
        angle = 2 * math.pi * i / segments
        x = cx + radius * math.cos(angle)
        y = cy + radius * math.sin(angle)
        points.append((x, y))
    return points


# =============================================================================
# STAR PATH CONVERSION
# =============================================================================

def parse_svg_path_to_points(path_d):
    """
    Parse an SVG path 'd' attribute to extract points.
    Handles M (move), L (line), and Z (close) commands.

    This is a simplified parser for the star paths generated by svg_engine.

    Args:
        path_d: SVG path 'd' attribute string

    Returns:
        List of (x, y) points
    """
    points = []
    # Remove commas and split by spaces
    path_d = path_d.replace(',', ' ')
    tokens = path_d.split()

    i = 0
    current_x, current_y = 0, 0
    start_x, start_y = 0, 0

    while i < len(tokens):
        token = tokens[i]

        if token == 'M':
            # Move to absolute
            current_x = float(tokens[i + 1])
            current_y = float(tokens[i + 2])
            start_x, start_y = current_x, current_y
            points.append((current_x, current_y))
            i += 3
        elif token == 'm':
            # Move to relative
            current_x += float(tokens[i + 1])
            current_y += float(tokens[i + 2])
            start_x, start_y = current_x, current_y
            points.append((current_x, current_y))
            i += 3
        elif token == 'L':
            # Line to absolute
            current_x = float(tokens[i + 1])
            current_y = float(tokens[i + 2])
            points.append((current_x, current_y))
            i += 3
        elif token == 'l':
            # Line to relative
            current_x += float(tokens[i + 1])
            current_y += float(tokens[i + 2])
            points.append((current_x, current_y))
            i += 3
        elif token == 'Z' or token == 'z':
            # Close path - return to start
            if points and (current_x != start_x or current_y != start_y):
                points.append((start_x, start_y))
            i += 1
        elif token == 'H':
            # Horizontal line absolute
            current_x = float(tokens[i + 1])
            points.append((current_x, current_y))
            i += 2
        elif token == 'h':
            # Horizontal line relative
            current_x += float(tokens[i + 1])
            points.append((current_x, current_y))
            i += 2
        elif token == 'V':
            # Vertical line absolute
            current_y = float(tokens[i + 1])
            points.append((current_x, current_y))
            i += 2
        elif token == 'v':
            # Vertical line relative
            current_y += float(tokens[i + 1])
            points.append((current_x, current_y))
            i += 2
        else:
            # Try to parse as coordinate pair (implicit L command)
            try:
                x = float(token)
                y = float(tokens[i + 1])
                current_x, current_y = x, y
                points.append((current_x, current_y))
                i += 2
            except (ValueError, IndexError):
                i += 1

    return points


# =============================================================================
# G-CODE GENERATION
# =============================================================================

def _generate_star_points(cx, cy, outer_r, inner_r, num_points, shape_factor=0.0):
    """
    Generate star/dart pattern points for G-code.

    This matches the calculate_star_path logic in svg_engine.py but returns
    a list of (x, y) points instead of an SVG path string.

    Args:
        cx, cy: Center coordinates
        outer_r: Outer (tip) radius
        inner_r: Inner (valley) radius
        num_points: Number of points (tips) on the star
        shape_factor: 0.0 = sine wave, 1.0 = square wave

    Returns:
        List of (x, y) points forming a closed star shape
    """
    points = []
    steps_per_segment = 8  # Smoothness of curves between tips

    for i in range(num_points):
        for j in range(steps_per_segment):
            # Angle within this segment (0 to 1)
            t = j / steps_per_segment
            # Overall angle
            angle = 2 * math.pi * (i + t) / num_points

            # Calculate radius using sine wave, modified by shape_factor
            # t=0 is tip (outer_r), t=0.5 is valley (inner_r)
            sine_val = math.cos(2 * math.pi * t)  # 1 at t=0, -1 at t=0.5

            if shape_factor > 0:
                # Blend toward square wave
                if sine_val > 0:
                    square_val = 1.0
                else:
                    square_val = -1.0
                blended = sine_val * (1 - shape_factor) + square_val * shape_factor
            else:
                blended = sine_val

            # Map from [-1, 1] to [inner_r, outer_r]
            r = inner_r + (outer_r - inner_r) * (blended + 1) / 2

            x = cx + r * math.cos(angle)
            y = cy + r * math.sin(angle)
            points.append((x, y))

    # Close the shape
    if points:
        points.append(points[0])

    return points


def generate_gcode_header(bounds_min_x, bounds_min_y, bounds_max_x, bounds_max_y):
    """Generate G-code file header."""
    lines = [
        "; Generated by Stohrer Sax Shop Companion",
        "; GRBL device profile, absolute coords",
        f"; Bounds: X{bounds_min_x:.2f} Y{bounds_min_y:.2f} to X{bounds_max_x:.2f} Y{bounds_max_y:.2f}",
        "G00 G17 G40 G21 G54",
        "G90",
        "M4",
    ]
    return lines


def generate_gcode_footer():
    """Generate G-code file footer."""
    lines = [
        "M9",
        "G1 S0",
        "M5",
        "G90",
        "; return to origin",
        "G0 X0Y0",
        "M2",
    ]
    return lines


def generate_gcode_layer(strokes, speed_mm_min, power_percent, layer_name):
    """
    Generate G-code for a layer (set of strokes).

    Args:
        strokes: List of strokes, where each stroke is a list of (x, y) points
        speed_mm_min: Feed rate in mm/min
        power_percent: Laser power as percentage (0-100)
        layer_name: Layer name for comment

    Returns:
        List of G-code lines
    """
    if not strokes:
        return []

    lines = []

    # Layer header
    lines.append(f"; Cut @ {speed_mm_min} mm/min, {power_percent}% power")
    lines.append("M8")

    # Power value for S parameter (percentage * 10 for 0-1000 scale)
    s_value = int(power_percent * 10)

    first_move = True

    for stroke in strokes:
        if len(stroke) < 2:
            continue

        # Rapid move to start of stroke
        x0, y0 = stroke[0]
        lines.append(f"G0 X{x0:.3f}Y{y0:.3f}")

        if first_move:
            lines.append(f"; Layer {layer_name}")
            first_move = False

        # Cut moves
        for i, (x, y) in enumerate(stroke[1:], 1):
            if i == 1:
                # First cut move includes S and F parameters
                lines.append(f"G1 X{x:.3f}Y{y:.3f}S{s_value}F{speed_mm_min}")
            else:
                lines.append(f"G1 X{x:.3f}Y{y:.3f}")

    return lines


def generate_gcode(pads, material, sheet_width_mm, sheet_height_mm, filename,
                   hole_dia, settings, polygon=None):
    """
    Generate G-code file for laser cutting pads.

    This mirrors the generate_svg function but outputs G-code instead.

    Args:
        pads: List of pad dictionaries with 'size' and 'qty' keys
        material: Material type ('felt', 'card', 'leather', 'leather_topgrain')
        sheet_width_mm, sheet_height_mm: Sheet dimensions in mm
        filename: Output filename
        hole_dia: Center hole diameter in mm (0 for no hole)
        settings: App settings dictionary
        polygon: Optional custom polygon shape
    """
    from svg_engine import _nest_discs, get_felt_thickness_mm

    # Use the same nesting as SVG generation
    placed, _, _ = _nest_discs(pads, material, sheet_width_mm, sheet_height_mm, settings, polygon=polygon)

    if not placed:
        return

    # Flip Y coordinates for G-code: SVG uses Y=0 at top, G-code uses Y=0 at bottom
    placed = [(pad_size, cx, sheet_height_mm - cy, radius) for (pad_size, cx, cy, radius) in placed]

    # Get G-code settings for this material
    gcode_settings = settings.get("gcode_settings", {})
    mat_settings = gcode_settings.get(material, gcode_settings.get("felt", {}))

    engraving_speed = mat_settings.get("engraving_speed", 1500)
    engraving_power = mat_settings.get("engraving_power", 10)
    hole_speed = mat_settings.get("hole_speed", 400)
    hole_power = mat_settings.get("hole_power", 25)
    cut_speed = mat_settings.get("cut_speed", 900)
    cut_power = mat_settings.get("cut_power", 35)

    # Kerf compensation (applied to cuts, not engraving) - per material
    kerf_width = mat_settings.get("kerf_width", 0.0)
    kerf_offset = kerf_width / 2  # Half kerf applied to each side

    # Font size for engraving (use engraving_font_size from settings)
    engraving_font_sizes = settings.get("engraving_font_size", {})
    font_size = engraving_font_sizes.get(material, 3.0)

    # Collect strokes for each layer
    engraving_strokes = []
    hole_strokes = []
    cut_strokes = []

    # Track bounds
    all_x = []
    all_y = []

    # Check if this material uses star patterns
    use_stars = material in ('leather', 'leather_topgrain') and settings.get("darts_enabled", True)
    dart_threshold = settings.get("dart_threshold", 18.0)

    for pad_size, cx, cy, radius in placed:
        all_x.extend([cx - radius, cx + radius])
        all_y.extend([cy - radius, cy + radius])

        # Determine if this is a star/dart pad
        is_dart_pad = use_stars and pad_size < dart_threshold

        # Engraving - use same logic as svg_engine.py
        should_engrave = False
        engraving_settings = None

        if is_dart_pad:
            # Use star-specific engraving settings
            if settings.get("dart_engraving_on", True):
                engraving_settings = settings.get("dart_engraving_loc", {"mode": "from_outside", "value": 2.5})
                should_engrave = True
        else:
            # Use standard engraving settings
            if settings.get("engraving_on", True):
                engraving_settings = settings.get("engraving_location", {}).get(material, {"mode": "centered", "value": 0})
                should_engrave = True

        # Safety check: don't engrave if text would be too large
        if should_engrave and font_size >= radius * 0.8:
            should_engrave = False

        if should_engrave:
            mode = engraving_settings.get('mode', 'centered')
            value = engraving_settings.get('value', 0)

            # Calculate engraving Y position (matching svg_engine.py exactly)
            if mode == 'from_outside':
                engraving_y = cy - (radius - value)
            elif mode == 'from_inside':
                hole_r = hole_dia / 2 if hole_dia > 0 else 0
                engraving_y = cy - (hole_r + value)
            else:  # centered
                hole_r = hole_dia / 2 if hole_dia > 0 else 1.75
                offset_from_center = (radius + hole_r) / 2
                engraving_y = cy - offset_from_center

            # Vertical adjustment for text baseline (same as SVG)
            vertical_adjust = font_size * 0.35
            label_y = engraving_y + vertical_adjust

            text = f"{pad_size:.1f}".rstrip('0').rstrip('.')
            text_strokes = get_text_strokes(text, font_size, cx, label_y)
            engraving_strokes.extend(text_strokes)

        # Center hole (respect min_hole_size setting)
        # Kerf compensation: shrink hole radius so final hole is correct size
        min_hole_size = settings.get("min_hole_size", 16.5)
        if hole_dia > 0 and pad_size >= min_hole_size:
            hole_radius = (hole_dia / 2) - kerf_offset
            if hole_radius > 0:
                hole_points = linearize_circle(cx, cy, hole_radius, segments=36)
                hole_strokes.append(hole_points)

        # Outer cut - circle or star pattern
        # Kerf compensation: expand outer cuts so final size is correct
        if use_stars and pad_size < dart_threshold:
            # Generate star path using same logic as svg_engine
            felt_thick = get_felt_thickness_mm(settings)
            overwrap = settings.get("dart_overwrap", 0.5)

            # Inner radius (valley)
            felt_r = (pad_size - settings.get("felt_offset", 0.75)) / 2
            inner_r = felt_r + felt_thick + overwrap

            # Outer radius is the full radius from nesting
            outer_r = radius

            if inner_r >= outer_r:
                inner_r = outer_r - 0.2

            # Apply kerf compensation - shift entire star outward
            outer_r += kerf_offset
            inner_r += kerf_offset

            # Calculate number of points
            circumference = 2 * math.pi * inner_r
            freq_mult = settings.get("dart_frequency_multiplier", 1.0)
            num_points = int((circumference / 3.5) * freq_mult)
            if num_points < 12:
                num_points = 12
            if num_points % 2 != 0:
                num_points += 1

            shape_factor = settings.get("dart_shape_factor", 0.0)

            # Generate star points directly
            star_points = _generate_star_points(cx, cy, outer_r, inner_r, num_points, shape_factor)
            cut_strokes.append(star_points)
        else:
            cut_radius = radius + kerf_offset
            cut_points = linearize_circle(cx, cy, cut_radius, segments=72)
            cut_strokes.append(cut_points)

    # Generate G-code
    gcode_lines = []

    # Header
    bounds_min_x = min(all_x) if all_x else 0
    bounds_min_y = min(all_y) if all_y else 0
    bounds_max_x = max(all_x) if all_x else sheet_width_mm
    bounds_max_y = max(all_y) if all_y else sheet_height_mm
    gcode_lines.extend(generate_gcode_header(bounds_min_x, bounds_min_y, bounds_max_x, bounds_max_y))

    # Layer names based on material (matching LightBurn convention)
    layer_names = {
        'felt': ('C10', 'C09', 'C00'),
        'card': ('C15', 'C14', 'C01'),
        'leather': ('C05', 'C03', 'C02'),
        'leather_topgrain': ('C05', 'C03', 'C02'),
    }
    eng_layer, hole_layer, cut_layer = layer_names.get(material, ('C00', 'C01', 'C02'))

    # Engraving layer (first)
    if engraving_strokes:
        gcode_lines.extend(generate_gcode_layer(engraving_strokes, engraving_speed, engraving_power, eng_layer))

    # Center hole layer (second)
    if hole_strokes:
        gcode_lines.extend(generate_gcode_layer(hole_strokes, hole_speed, hole_power, hole_layer))

    # Outer cut layer (last)
    if cut_strokes:
        gcode_lines.extend(generate_gcode_layer(cut_strokes, cut_speed, cut_power, cut_layer))

    # Footer
    gcode_lines.extend(generate_gcode_footer())

    # Write file
    with open(filename, 'w') as f:
        f.write('\n'.join(gcode_lines))


def can_generate_gcode(material):
    """Check if G-code generation is supported for a material."""
    return material in ('felt', 'card', 'leather', 'leather_topgrain')


def generate_gcode_from_placed(placed, material, sheet_width_mm, sheet_height_mm, filename,
                                hole_dia, settings, polygon=None):
    """
    Generate G-code file from pre-computed placed discs.

    This is used by scrap mode where nesting is done separately via try_nest_partial().
    The function generates G-code for the placed discs without re-running nesting.

    Args:
        placed: List of (pad_size, cx, cy, radius) tuples from nesting
        material: Material type ('felt', 'card', 'leather', 'leather_topgrain')
        sheet_width_mm, sheet_height_mm: Sheet dimensions in mm
        filename: Output filename
        hole_dia: Center hole diameter in mm (0 for no hole)
        settings: App settings dictionary
        polygon: Optional (unused, for API consistency)
    """
    from svg_engine import get_felt_thickness_mm

    if not placed:
        return

    # Flip Y coordinates for G-code: SVG uses Y=0 at top, G-code uses Y=0 at bottom
    placed = [(pad_size, cx, sheet_height_mm - cy, radius) for (pad_size, cx, cy, radius) in placed]

    # Get G-code settings for this material
    gcode_settings = settings.get("gcode_settings", {})
    mat_settings = gcode_settings.get(material, gcode_settings.get("felt", {}))

    engraving_speed = mat_settings.get("engraving_speed", 1500)
    engraving_power = mat_settings.get("engraving_power", 10)
    hole_speed = mat_settings.get("hole_speed", 400)
    hole_power = mat_settings.get("hole_power", 25)
    cut_speed = mat_settings.get("cut_speed", 900)
    cut_power = mat_settings.get("cut_power", 35)

    # Kerf compensation (applied to cuts, not engraving) - per material
    kerf_width = mat_settings.get("kerf_width", 0.0)
    kerf_offset = kerf_width / 2

    # Font size for engraving
    engraving_font_sizes = settings.get("engraving_font_size", {})
    font_size = engraving_font_sizes.get(material, 3.0)

    # Collect strokes for each layer
    engraving_strokes = []
    hole_strokes = []
    cut_strokes = []

    # Track bounds
    all_x = []
    all_y = []

    # Check if this material uses star patterns
    use_stars = material in ('leather', 'leather_topgrain') and settings.get("darts_enabled", True)
    dart_threshold = settings.get("dart_threshold", 18.0)

    for pad_size, cx, cy, radius in placed:
        all_x.extend([cx - radius, cx + radius])
        all_y.extend([cy - radius, cy + radius])

        # Determine if this is a star/dart pad
        is_dart_pad = use_stars and pad_size < dart_threshold

        # Engraving
        should_engrave = False
        engraving_settings = None

        if is_dart_pad:
            if settings.get("dart_engraving_on", True):
                engraving_settings = settings.get("dart_engraving_loc", {"mode": "from_outside", "value": 2.5})
                should_engrave = True
        else:
            if settings.get("engraving_on", True):
                engraving_settings = settings.get("engraving_location", {}).get(material, {"mode": "centered", "value": 0})
                should_engrave = True

        if should_engrave and font_size >= radius * 0.8:
            should_engrave = False

        if should_engrave:
            mode = engraving_settings.get('mode', 'centered')
            value = engraving_settings.get('value', 0)

            if mode == 'from_outside':
                engraving_y = cy - (radius - value)
            elif mode == 'from_inside':
                hole_r = hole_dia / 2 if hole_dia > 0 else 0
                engraving_y = cy - (hole_r + value)
            else:  # centered
                hole_r = hole_dia / 2 if hole_dia > 0 else 1.75
                offset_from_center = (radius + hole_r) / 2
                engraving_y = cy - offset_from_center

            vertical_adjust = font_size * 0.35
            label_y = engraving_y + vertical_adjust

            text = f"{pad_size:.1f}".rstrip('0').rstrip('.')
            text_strokes = get_text_strokes(text, font_size, cx, label_y)
            engraving_strokes.extend(text_strokes)

        # Center hole
        min_hole_size = settings.get("min_hole_size", 16.5)
        if hole_dia > 0 and pad_size >= min_hole_size:
            hole_radius = (hole_dia / 2) - kerf_offset
            if hole_radius > 0:
                hole_points = linearize_circle(cx, cy, hole_radius, segments=36)
                hole_strokes.append(hole_points)

        # Outer cut - circle or star pattern
        if use_stars and pad_size < dart_threshold:
            felt_thick = get_felt_thickness_mm(settings)
            overwrap = settings.get("dart_overwrap", 0.5)

            felt_r = (pad_size - settings.get("felt_offset", 0.75)) / 2
            inner_r = felt_r + felt_thick + overwrap
            outer_r = radius

            if inner_r >= outer_r:
                inner_r = outer_r - 0.2

            outer_r += kerf_offset
            inner_r += kerf_offset

            circumference = 2 * math.pi * inner_r
            freq_mult = settings.get("dart_frequency_multiplier", 1.0)
            num_points = int((circumference / 3.5) * freq_mult)
            if num_points < 12:
                num_points = 12
            if num_points % 2 != 0:
                num_points += 1

            shape_factor = settings.get("dart_shape_factor", 0.0)
            star_points = _generate_star_points(cx, cy, outer_r, inner_r, num_points, shape_factor)
            cut_strokes.append(star_points)
        else:
            cut_radius = radius + kerf_offset
            cut_points = linearize_circle(cx, cy, cut_radius, segments=72)
            cut_strokes.append(cut_points)

    # Generate G-code
    gcode_lines = []

    bounds_min_x = min(all_x) if all_x else 0
    bounds_min_y = min(all_y) if all_y else 0
    bounds_max_x = max(all_x) if all_x else sheet_width_mm
    bounds_max_y = max(all_y) if all_y else sheet_height_mm
    gcode_lines.extend(generate_gcode_header(bounds_min_x, bounds_min_y, bounds_max_x, bounds_max_y))

    layer_names = {
        'felt': ('C10', 'C09', 'C00'),
        'card': ('C15', 'C14', 'C01'),
        'leather': ('C05', 'C03', 'C02'),
        'leather_topgrain': ('C05', 'C03', 'C02'),
    }
    eng_layer, hole_layer, cut_layer = layer_names.get(material, ('C00', 'C01', 'C02'))

    if engraving_strokes:
        gcode_lines.extend(generate_gcode_layer(engraving_strokes, engraving_speed, engraving_power, eng_layer))

    if hole_strokes:
        gcode_lines.extend(generate_gcode_layer(hole_strokes, hole_speed, hole_power, hole_layer))

    if cut_strokes:
        gcode_lines.extend(generate_gcode_layer(cut_strokes, cut_speed, cut_power, cut_layer))

    gcode_lines.extend(generate_gcode_footer())

    with open(filename, 'w') as f:
        f.write('\n'.join(gcode_lines))
