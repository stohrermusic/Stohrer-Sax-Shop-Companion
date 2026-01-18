import math
import svgwrite
from config import DEFAULT_SETTINGS

# ==========================================
# CORE MATH & LOGIC
# ==========================================

def calculate_star_path(cx, cy, outer_r, inner_r, num_points=12, shape_factor=0.0):
    """
    Generates an SVG path string for a smooth Sine Wave (Flower) shape.
    shape_factor: 0.0 = Sine, 1.0 = Flattened (Square-ish)
    """
    path_data = []
    
    avg_r = (outer_r + inner_r) / 2.0
    amplitude = (outer_r - inner_r) / 2.0
    
    steps = int(num_points * 8) 
    if steps < 64: steps = 64
    
    angle_step = (2 * math.pi) / steps

    # Calculate power for shaping. 
    power = 1.0 - (0.9 * shape_factor)

    for i in range(steps + 1):
        theta = i * angle_step
        
        # Raw Sine Wave (-1 to 1)
        raw_wave = math.cos(num_points * theta)
        
        # Apply Shaping: sign * |raw|^power
        shaped_wave = (1 if raw_wave >= 0 else -1) * (abs(raw_wave) ** power)
        
        r = avg_r + amplitude * shaped_wave
        
        x = cx + r * math.cos(theta)
        y = cy + r * math.sin(theta)
        
        command = "M" if i == 0 else "L"
        path_data.append(f"{command} {x:.3f} {y:.3f}")
        
    path_data.append("Z") 
    return " ".join(path_data)

def leather_back_wrap(pad_size, multiplier, extra_base=0.0):
    base_wrap = 0
    if pad_size >= 45:
        base_wrap = 3.2
    elif pad_size >= 12:
        base_wrap = 1.2 + (pad_size - 12) * (2.0 / 33.0)
    elif pad_size >= 6:
        base_wrap = 1.0 + (pad_size - 6) * (0.2 / 6.0)
    else:
        base_wrap = 1.0
        
    # Apply the Dart Bonus (if any) before multiplier
    total_base = base_wrap + extra_base
    return total_base * multiplier

def get_felt_thickness_mm(settings):
    thickness = settings.get("felt_thickness", 3.175)
    if settings.get("felt_thickness_unit") == "in":
        return thickness * 25.4
    return thickness

def get_disc_diameter(pad_size, material, settings):
    if material == 'felt': return pad_size - settings["felt_offset"]
    if material == 'card': return pad_size - (settings["felt_offset"] + settings["card_to_felt_offset"])
    if material == 'exact_size': return pad_size
    
    if material == 'leather':
        threshold = settings.get("dart_threshold", 18.0)
        darts_enabled = settings.get("darts_enabled", True)
        
        if darts_enabled and pad_size < threshold:
            # DART BOOST: Add extra wrap for stars
            bonus = settings.get("dart_wrap_bonus", 0.75)
            wrap = leather_back_wrap(pad_size, settings["leather_wrap_multiplier"], extra_base=bonus)
        else:
            # Standard Wrap
            wrap = leather_back_wrap(pad_size, settings["leather_wrap_multiplier"])
            
        felt_thickness_mm = get_felt_thickness_mm(settings)
        diameter = pad_size + 2 * (felt_thickness_mm + wrap)
        return round(diameter * 2) / 2
        
    return 0

def should_have_center_hole(pad_size, hole_dia, settings):
    min_size = settings.get("min_hole_size", 16.5)
    return hole_dia > 0 and pad_size >= min_size

def check_for_oversized_engravings(pads, material_vars, settings):
    oversized = {}
    for material, var in material_vars.items():
        if not var.get(): continue
        
        font_size = settings["engraving_font_size"].get(material, 2.0)
        oversized_sizes = set()

        for pad in pads:
            pad_size = pad['size']
            diameter = get_disc_diameter(pad_size, material, settings)
            radius = diameter / 2
            if font_size >= radius * 0.8:
                oversized_sizes.add(pad_size)
        
        if oversized_sizes:
            oversized[material] = oversized_sizes
    return oversized

def _nest_discs(pads, material, width_mm, height_mm, settings, spacing_mm=1.0, polygon=None):
    """
    Greedy circle-packing algorithm. Returns list of placed discs as (pad_size, cx, cy, r).
    Discs that couldn't be placed are omitted from the result.

    If polygon is provided (list of (x,y) tuples in mm), uses polygon nesting instead of rectangle.
    """
    if polygon:
        return _nest_discs_polygon(pads, material, settings, polygon, spacing_mm)

    discs = []
    for pad in pads:
        pad_size, qty = pad['size'], pad['qty']
        diameter = get_disc_diameter(pad_size, material, settings)
        for _ in range(qty):
            discs.append((pad_size, diameter))

    discs.sort(key=lambda x: -x[1])  # Largest first
    placed = []

    for pad_size, dia in discs:
        r = dia / 2
        placed_successfully = False
        y = spacing_mm
        while y + dia + spacing_mm <= height_mm and not placed_successfully:
            x = spacing_mm
            while x + dia + spacing_mm <= width_mm:
                cx, cy = x + r, y + r
                is_collision = any((cx - px)**2 + (cy - py)**2 < (r + pr + spacing_mm)**2 for _, px, py, pr in placed)
                if not is_collision:
                    placed.append((pad_size, cx, cy, r))
                    placed_successfully = True
                    break
                x += 1
            y += 1

    return placed, len(discs)


# ==========================================
# POLYGON NESTING HELPERS
# ==========================================

def _point_in_polygon(x, y, polygon):
    """
    Ray casting algorithm to check if point (x,y) is inside polygon.
    Polygon is a list of (x, y) tuples.
    """
    n = len(polygon)
    inside = False

    j = n - 1
    for i in range(n):
        xi, yi = polygon[i]
        xj, yj = polygon[j]

        if ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / (yj - yi) + xi):
            inside = not inside
        j = i

    return inside


def _distance_point_to_segment(px, py, x1, y1, x2, y2):
    """
    Calculate the minimum distance from point (px, py) to line segment (x1,y1)-(x2,y2).
    """
    dx = x2 - x1
    dy = y2 - y1
    length_sq = dx * dx + dy * dy

    if length_sq == 0:
        # Segment is a point
        return math.sqrt((px - x1) ** 2 + (py - y1) ** 2)

    # Parameter t for projection onto segment
    t = max(0, min(1, ((px - x1) * dx + (py - y1) * dy) / length_sq))

    # Closest point on segment
    closest_x = x1 + t * dx
    closest_y = y1 + t * dy

    return math.sqrt((px - closest_x) ** 2 + (py - closest_y) ** 2)


def _circle_fits_in_polygon(cx, cy, radius, polygon, spacing_mm=1.0):
    """
    Check if a circle with center (cx, cy) and given radius fits inside the polygon.
    The circle must be at least spacing_mm away from all edges.
    """
    # First check if center is inside
    if not _point_in_polygon(cx, cy, polygon):
        return False

    # Check distance to each edge
    n = len(polygon)
    for i in range(n):
        x1, y1 = polygon[i]
        x2, y2 = polygon[(i + 1) % n]

        dist = _distance_point_to_segment(cx, cy, x1, y1, x2, y2)
        if dist < radius + spacing_mm:
            return False

    return True


def _nest_discs_polygon(pads, material, settings, polygon, spacing_mm=1.0):
    """
    Greedy circle-packing algorithm for polygon boundaries.
    Returns list of placed discs as (pad_size, cx, cy, r).
    """
    discs = []
    for pad in pads:
        pad_size, qty = pad['size'], pad['qty']
        diameter = get_disc_diameter(pad_size, material, settings)
        for _ in range(qty):
            discs.append((pad_size, diameter))

    discs.sort(key=lambda x: -x[1])  # Largest first
    placed = []

    # Get bounding box of polygon for search limits
    min_x = min(p[0] for p in polygon)
    max_x = max(p[0] for p in polygon)
    min_y = min(p[1] for p in polygon)
    max_y = max(p[1] for p in polygon)

    for pad_size, dia in discs:
        r = dia / 2
        placed_successfully = False

        # Scan from bottom-left of bounding box
        y = min_y + spacing_mm
        while y + r <= max_y and not placed_successfully:
            x = min_x + spacing_mm
            while x + r <= max_x:
                cx, cy = x + r, y + r

                # Check if circle fits in polygon
                if _circle_fits_in_polygon(cx, cy, r, polygon, spacing_mm):
                    # Check collision with already placed discs
                    is_collision = any(
                        (cx - px) ** 2 + (cy - py) ** 2 < (r + pr + spacing_mm) ** 2
                        for _, px, py, pr in placed
                    )
                    if not is_collision:
                        placed.append((pad_size, cx, cy, r))
                        placed_successfully = True
                        break
                x += 1
            y += 1

    return placed, len(discs)


def can_all_pads_fit(pads, material, width_mm, height_mm, settings, polygon=None):
    placed, total = _nest_discs(pads, material, width_mm, height_mm, settings, polygon=polygon)
    return len(placed) == total


def generate_svg(pads, material, width_mm, height_mm, filename, hole_dia_preset, settings, polygon=None):
    placed, _ = _nest_discs(pads, material, width_mm, height_mm, settings, polygon=polygon)

    compatibility_mode = settings.get("compatibility_mode", False)

    # Both modes use viewBox to ensure path coordinates (which can't have units) work correctly
    if compatibility_mode:
        dwg = svgwrite.Drawing(filename, size=(f"{width_mm}mm", f"{height_mm}mm"), viewBox=f"0 0 {width_mm} {height_mm}")
        stroke_w = 0.1
    else:
        dwg = svgwrite.Drawing(filename, size=(f"{width_mm}mm", f"{height_mm}mm"), viewBox=f"0 0 {width_mm} {height_mm}", profile='tiny')
        stroke_w = '0.1mm'

    layer_colors = settings.get("layer_colors", DEFAULT_SETTINGS["layer_colors"])

    for pad_size, cx, cy, r in placed:
        
        threshold = settings.get("dart_threshold", 18.0)
        darts_enabled = settings.get("darts_enabled", True)
        
        is_dart_pad = (material == 'leather' and darts_enabled and pad_size < threshold)
        
        if is_dart_pad:
            # --- STAR LOGIC ---
            felt_thick = get_felt_thickness_mm(settings)
            overwrap = settings.get("dart_overwrap", 0.5)
            
            # 1. Inner Radius (Valley) - Safe Zone
            felt_r = (pad_size - settings["felt_offset"]) / 2
            inner_r = felt_r + felt_thick + overwrap
            
            # 2. Outer Radius (Tip) - The Boosted Wrap
            # 'r' is the full Boosted radius from get_disc_diameter
            outer_r = r
            
            # Safety Check
            if inner_r >= outer_r:
                 inner_r = outer_r - 0.2 
            
            # 3. Dynamic Points & Shape
            circumference = 2 * math.pi * inner_r
            freq_mult = settings.get("dart_frequency_multiplier", 1.0)
            num_points = int((circumference / 3.5) * freq_mult)
            if num_points < 12: num_points = 12 
            if num_points % 2 != 0: num_points += 1 
            
            shape_factor = settings.get("dart_shape_factor", 0.0)
            
            path_d = calculate_star_path(cx, cy, outer_r, inner_r, num_points=num_points, shape_factor=shape_factor)
            
            dwg.add(dwg.path(d=path_d, stroke=layer_colors[f'{material}_outline'], fill='none', stroke_width=stroke_w))
        else:
            # --- STANDARD CIRCLE LOGIC ---
            if compatibility_mode:
                dwg.add(dwg.circle(center=(cx, cy), r=r, stroke=layer_colors[f'{material}_outline'], fill='none', stroke_width=stroke_w))
            else:
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{r}mm", stroke=layer_colors[f'{material}_outline'], fill='none', stroke_width=stroke_w))

        hole_dia = 0
        if should_have_center_hole(pad_size, hole_dia_preset, settings):
            hole_dia = hole_dia_preset

        if hole_dia > 0:
            if compatibility_mode:
                dwg.add(dwg.circle(center=(cx, cy), r=hole_dia / 2, stroke=layer_colors[f'{material}_center_hole'], fill='none', stroke_width=stroke_w))
            else:
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{hole_dia / 2}mm", stroke=layer_colors[f'{material}_center_hole'], fill='none', stroke_width=stroke_w))

        font_size = settings.get("engraving_font_size", {}).get(material, 2.0)
        
        # --- Determine Engraving Settings (Standard vs Star) ---
        should_engrave = False
        
        if is_dart_pad:
            # Use Star Specific Settings
            if settings.get("dart_engraving_on", True):
                engraving_settings = settings.get("dart_engraving_loc", {"mode": "from_outside", "value": 2.5})
                should_engrave = True
        else:
            # Use Standard Settings
            if settings.get("engraving_on", True):
                engraving_settings = settings["engraving_location"][material]
                should_engrave = True

        # Safety Check: Don't engrave if text is wider than the pad radius
        if should_engrave and (font_size >= r * 0.8):
            should_engrave = False
        
        if should_engrave:
            mode = engraving_settings['mode']
            value = engraving_settings['value']
            
            engraving_y = 0
            if mode == 'from_outside':
                engraving_y = cy - (r - value)
            elif mode == 'from_inside':
                hole_r = hole_dia / 2 if hole_dia > 0 else 0
                engraving_y = cy - (hole_r + value)
            else: # centered
                hole_r = hole_dia / 2 if hole_dia > 0 else 1.75
                offset_from_center = (r + hole_r) / 2
                engraving_y = cy - offset_from_center

            vertical_adjust = font_size * 0.35
            text_content = f"{pad_size:.1f}".rstrip('0').rstrip('.')
            
            if compatibility_mode:
                dwg.add(dwg.text(text_content,
                                 insert=(cx, engraving_y + vertical_adjust),
                                 text_anchor="middle",
                                 font_size=font_size,
                                 fill=layer_colors[f'{material}_engraving']))
            else:
                dwg.add(dwg.text(text_content,
                                 insert=(f"{cx}mm", f"{engraving_y + vertical_adjust}mm"),
                                 text_anchor="middle",
                                 font_size=f"{font_size}mm",
                                 fill=layer_colors[f'{material}_engraving']))
        
    dwg.save()
