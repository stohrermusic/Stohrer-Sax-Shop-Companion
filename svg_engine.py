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
    if material == 'die_ring':
        return 70.0 if pad_size >= 40.0 else 50.0
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

    Supports 'max' quantity: fixed-qty pads are placed first, then max pads fill remaining space.
    """
    if polygon:
        return _nest_discs_polygon(pads, material, settings, polygon, spacing_mm)

    # Separate fixed and max pads
    fixed_pads = [p for p in pads if p['qty'] != 'max']
    max_pads = [p for p in pads if p['qty'] == 'max']

    # Build disc list from fixed pads
    discs = []
    for pad in fixed_pads:
        pad_size, qty = pad['size'], pad['qty']
        diameter = get_disc_diameter(pad_size, material, settings)
        for _ in range(qty):
            discs.append((pad_size, diameter))

    discs.sort(key=lambda x: -x[1])  # Largest first
    placed = []
    fixed_total = len(discs)
    fixed_placed = 0

    # Place fixed pads
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
                    fixed_placed += 1
                    break
                x += 1
            y += 1

    # Fill remaining space with max pad (if any)
    if max_pads:
        max_pad = max_pads[0]
        max_size = max_pad['size']
        max_dia = get_disc_diameter(max_size, material, settings)
        max_r = max_dia / 2

        while True:
            placed_successfully = False
            y = spacing_mm
            while y + max_dia + spacing_mm <= height_mm and not placed_successfully:
                x = spacing_mm
                while x + max_dia + spacing_mm <= width_mm:
                    cx, cy = x + max_r, y + max_r
                    is_collision = any((cx - px)**2 + (cy - py)**2 < (max_r + pr + spacing_mm)**2 for _, px, py, pr in placed)
                    if not is_collision:
                        placed.append((max_size, cx, cy, max_r))
                        placed_successfully = True
                        break
                    x += 1
                y += 1
            if not placed_successfully:
                break  # No more room for max pads

    return placed, fixed_placed, fixed_total


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


def _distance_to_nearest_edge(cx, cy, polygon):
    """Calculate minimum distance from point (cx, cy) to any polygon edge."""
    min_dist = float('inf')
    n = len(polygon)
    for i in range(n):
        x1, y1 = polygon[i]
        x2, y2 = polygon[(i + 1) % n]
        dist = _distance_point_to_segment(cx, cy, x1, y1, x2, y2)
        min_dist = min(min_dist, dist)
    return min_dist


def _distance_to_nearest_vertex(cx, cy, polygon):
    """Calculate minimum distance from point (cx, cy) to any polygon vertex (corner)."""
    min_dist = float('inf')
    for vx, vy in polygon:
        dist = math.sqrt((cx - vx) ** 2 + (cy - vy) ** 2)
        min_dist = min(min_dist, dist)
    return min_dist


def _find_longest_edge(polygon):
    """
    Find the longest edge of a polygon.
    Returns (start_point, end_point, length, edge_index).
    """
    n = len(polygon)
    longest = None
    longest_len = 0
    longest_idx = 0

    for i in range(n):
        x1, y1 = polygon[i]
        x2, y2 = polygon[(i + 1) % n]
        length = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
        if length > longest_len:
            longest_len = length
            longest = ((x1, y1), (x2, y2))
            longest_idx = i

    return longest[0], longest[1], longest_len, longest_idx


def _nest_discs_polygon(pads, material, settings, polygon, spacing_mm=1.0):
    """
    Smart circle-packing algorithm for polygon boundaries.

    - Large discs: prefer positions closest to centroid (center-out)
    - Small discs: prefer edges, corners, and snug fits against other discs

    Supports 'max' quantity: fixed-qty pads are placed first, then max pads fill remaining space.

    Returns list of placed discs as (pad_size, cx, cy, r).
    """
    # Separate fixed and max pads
    fixed_pads = [p for p in pads if p['qty'] != 'max']
    max_pads = [p for p in pads if p['qty'] == 'max']

    discs = []
    for pad in fixed_pads:
        pad_size, qty = pad['size'], pad['qty']
        diameter = get_disc_diameter(pad_size, material, settings)
        for _ in range(qty):
            discs.append((pad_size, diameter))

    discs.sort(key=lambda x: -x[1])  # Largest first
    placed = []
    fixed_total = len(discs)
    fixed_placed = 0

    # Get bounding box of polygon for search limits
    min_x = min(p[0] for p in polygon)
    max_x = max(p[0] for p in polygon)
    min_y = min(p[1] for p in polygon)
    max_y = max(p[1] for p in polygon)

    # Calculate polygon centroid
    n = len(polygon)
    centroid_x = sum(p[0] for p in polygon) / n
    centroid_y = sum(p[1] for p in polygon) / n

    # Size threshold - small pads use edge-seeking behavior
    size_threshold = settings.get("dart_threshold", 18.0)

    # Use 1mm grid steps for accuracy (worth the extra time for scrap efficiency)
    step = 1

    def calc_snugness(cx, cy, r, placed_discs):
        """
        Calculate how snugly a disc fits - lower = more snug (better).
        Measures gap to nearest placed disc (0 = touching at spacing distance).
        """
        if not placed_discs:
            return 0  # No penalty if no discs placed yet

        min_gap = float('inf')
        for _, px, py, pr in placed_discs:
            dist = math.sqrt((cx - px) ** 2 + (cy - py) ** 2)
            gap = dist - (r + pr + spacing_mm)  # Gap beyond minimum required
            min_gap = min(min_gap, gap)
        return min_gap

    def find_best_position_large(r, placed_discs):
        """Find position closest to centroid (for large discs)."""
        best_pos = None
        best_score = float('inf')

        y = min_y + spacing_mm
        while y + r <= max_y:
            x = min_x + spacing_mm
            while x + r <= max_x:
                cx, cy = x + r, y + r

                # Score = distance to centroid (lower = better)
                score = math.sqrt((cx - centroid_x) ** 2 + (cy - centroid_y) ** 2)

                if score < best_score:
                    if _circle_fits_in_polygon(cx, cy, r, polygon, spacing_mm):
                        is_collision = any(
                            (cx - px) ** 2 + (cy - py) ** 2 < (r + pr + spacing_mm) ** 2
                            for _, px, py, pr in placed_discs
                        )
                        if not is_collision:
                            best_score = score
                            best_pos = (cx, cy)
                x += step
            y += step

        return best_pos

    def find_best_position_small(r, placed_discs):
        """Find position near edges/corners with snug fit (for small discs)."""
        best_pos = None
        best_score = float('inf')

        y = min_y + spacing_mm
        while y + r <= max_y:
            x = min_x + spacing_mm
            while x + r <= max_x:
                cx, cy = x + r, y + r

                # Check validity first (fast rejection)
                if not _circle_fits_in_polygon(cx, cy, r, polygon, spacing_mm):
                    x += step
                    continue

                is_collision = any(
                    (cx - px) ** 2 + (cy - py) ** 2 < (r + pr + spacing_mm) ** 2
                    for _, px, py, pr in placed_discs
                )
                if is_collision:
                    x += step
                    continue

                # Calculate score for small disc (lower = better)
                # 1. Distance to nearest edge (prefer close to edges)
                edge_dist = _distance_to_nearest_edge(cx, cy, polygon)
                edge_gap = edge_dist - r - spacing_mm  # Gap beyond disc radius

                # 2. Distance to nearest corner/vertex (prefer corners)
                vertex_dist = _distance_to_nearest_vertex(cx, cy, polygon)

                # 3. Snugness with other discs (prefer tight packing)
                snugness = calc_snugness(cx, cy, r, placed_discs)

                # Combined score: weight edge proximity and corners heavily
                # Lower score = better position
                score = edge_gap * 1.0 + vertex_dist * 0.3 + snugness * 0.5

                if score < best_score:
                    best_score = score
                    best_pos = (cx, cy)

                x += step
            y += step

        return best_pos

    # Cache longest edge for the entire nesting operation
    (longest_ex1, longest_ey1), (longest_ex2, longest_ey2), _, _ = _find_longest_edge(polygon)

    def find_best_position_longest_edge(r, placed_discs):
        """
        Fill from longest edge inward.
        Prioritizes positions closest to the longest edge, with snug packing.
        """
        best_pos = None
        best_score = float('inf')

        # Use 1mm grid for accuracy (early-exit optimization keeps it fast)
        edge_step = 1

        y = min_y + spacing_mm
        while y + r <= max_y:
            x = min_x + spacing_mm
            while x + r <= max_x:
                cx, cy = x + r, y + r

                # Quick score check - skip if can't possibly be better
                dist_to_edge = _distance_point_to_segment(cx, cy, longest_ex1, longest_ey1, longest_ex2, longest_ey2)
                if dist_to_edge >= best_score:
                    x += edge_step
                    continue

                # Check validity
                if not _circle_fits_in_polygon(cx, cy, r, polygon, spacing_mm):
                    x += edge_step
                    continue

                is_collision = any(
                    (cx - px) ** 2 + (cy - py) ** 2 < (r + pr + spacing_mm) ** 2
                    for _, px, py, pr in placed_discs
                )
                if is_collision:
                    x += edge_step
                    continue

                # Full score with snugness
                snugness = 0
                if placed_discs:
                    min_gap = float('inf')
                    for _, px, py, pr in placed_discs:
                        dist = math.sqrt((cx - px) ** 2 + (cy - py) ** 2)
                        gap = dist - (r + pr + spacing_mm)
                        min_gap = min(min_gap, gap)
                    snugness = min_gap

                score = dist_to_edge + snugness * 0.3

                if score < best_score:
                    best_score = score
                    best_pos = (cx, cy)

                x += edge_step
            y += edge_step

        return best_pos

    # Place fixed pads
    for pad_size, dia in discs:
        r = dia / 2
        if pad_size >= size_threshold:
            best_pos = find_best_position_large(r, placed)
        else:
            best_pos = find_best_position_small(r, placed)

        if best_pos:
            placed.append((pad_size, best_pos[0], best_pos[1], r))
            fixed_placed += 1

    # Fill remaining space with max pad (if any)
    if max_pads:
        max_pad = max_pads[0]
        max_size = max_pad['size']
        max_dia = get_disc_diameter(max_size, material, settings)
        max_r = max_dia / 2

        # Choose fill strategy based on setting
        max_fill_style = settings.get("max_fill_style", "center_out")
        if max_fill_style == "longest_edge":
            find_fn = find_best_position_longest_edge
        else:
            find_fn = find_best_position_large  # center_out

        while True:
            best_pos = find_fn(max_r, placed)
            if best_pos:
                placed.append((max_size, best_pos[0], best_pos[1], max_r))
            else:
                break  # No more room for max pads

    return placed, fixed_placed, fixed_total


def can_all_pads_fit(pads, material, width_mm, height_mm, settings, polygon=None):
    placed, fixed_placed, fixed_total = _nest_discs(pads, material, width_mm, height_mm, settings, polygon=polygon)
    # Check if all fixed-quantity pads fit (max pads are flexible by definition)
    return fixed_placed == fixed_total


def _render_svg_discs(dwg, placed, material, hole_dia_preset, settings, compatibility_mode, stroke_w):
    """Render placed discs (outlines, holes, engravings) into an SVG drawing."""
    layer_colors = settings.get("layer_colors", DEFAULT_SETTINGS["layer_colors"])

    for pad_size, cx, cy, r in placed:
        threshold = settings.get("dart_threshold", 18.0)
        darts_enabled = settings.get("darts_enabled", True)

        is_dart_pad = (material == 'leather' and darts_enabled and pad_size < threshold)

        if is_dart_pad:
            # --- STAR LOGIC ---
            felt_thick = get_felt_thickness_mm(settings)
            overwrap = settings.get("dart_overwrap", 0.5)

            felt_r = (pad_size - settings["felt_offset"]) / 2
            inner_r = felt_r + felt_thick + overwrap
            outer_r = r

            if inner_r >= outer_r:
                inner_r = outer_r - 0.2

            circumference = 2 * math.pi * inner_r
            freq_mult = settings.get("dart_frequency_multiplier", 1.0)
            num_points = int((circumference / 3.5) * freq_mult)
            if num_points < 12:
                num_points = 12
            if num_points % 2 != 0:
                num_points += 1

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
            if settings.get("dart_engraving_on", True):
                engraving_settings = settings.get("dart_engraving_loc", {"mode": "from_outside", "value": 2.5})
                should_engrave = True
        else:
            if settings.get("engraving_on", True):
                engraving_settings = settings["engraving_location"][material]
                should_engrave = True

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
            else:  # centered
                hole_r = hole_dia / 2 if hole_dia > 0 else 1.75
                offset_from_center = (r + hole_r) / 2
                engraving_y = cy - offset_from_center

            vertical_adjust = font_size * 0.35
            label_y = engraving_y + vertical_adjust
            text_content = f"{pad_size:.1f}".rstrip('0').rstrip('.')

            # Auto-fit: shift text toward center if it impinges on disc edge
            min_clearance = 0.5
            text_half_h = font_size / 2
            text_half_w = sum(0.3 if c == '.' else 0.6 for c in text_content)
            text_half_w += 0.1 * (len(text_content) - 1) if len(text_content) > 1 else 0
            text_half_w = text_half_w * font_size / 2

            # Check farthest corner of text bbox from disc center
            corners = [(cx - text_half_w, label_y - text_half_h),
                       (cx + text_half_w, label_y - text_half_h),
                       (cx - text_half_w, label_y + text_half_h),
                       (cx + text_half_w, label_y + text_half_h)]
            max_dist = max(math.sqrt((px - cx)**2 + (py - cy)**2) for px, py in corners)

            safe_radius = r - min_clearance
            if max_dist > safe_radius > 0:
                # Prefer shifting toward center over shrinking for readability
                centered_max_dist = math.sqrt(text_half_w**2 + text_half_h**2)
                if centered_max_dist <= safe_radius:
                    # Fits at full size — shift toward center just enough
                    max_offset = math.sqrt(safe_radius**2 - text_half_w**2) - text_half_h
                    dy = label_y - cy
                    if abs(dy) > max_offset:
                        label_y = cy + math.copysign(max_offset, dy)
                else:
                    # Too large even centered — center and scale as last resort
                    label_y = cy
                    scale = safe_radius / centered_max_dist
                    font_size *= scale

            if compatibility_mode:
                dwg.add(dwg.text(text_content,
                                 insert=(cx, label_y),
                                 text_anchor="middle",
                                 font_size=font_size,
                                 fill=layer_colors[f'{material}_engraving']))
            else:
                dwg.add(dwg.text(text_content,
                                 insert=(f"{cx}mm", f"{label_y}mm"),
                                 text_anchor="middle",
                                 font_size=f"{font_size}mm",
                                 fill=layer_colors[f'{material}_engraving']))


def _create_svg_drawing(filename, width_mm, height_mm, settings):
    """Create an SVG drawing with the correct settings."""
    compatibility_mode = settings.get("compatibility_mode", False)
    if compatibility_mode:
        dwg = svgwrite.Drawing(filename, size=(f"{width_mm}mm", f"{height_mm}mm"), viewBox=f"0 0 {width_mm} {height_mm}")
        stroke_w = 0.1
    else:
        dwg = svgwrite.Drawing(filename, size=(f"{width_mm}mm", f"{height_mm}mm"), viewBox=f"0 0 {width_mm} {height_mm}", profile='tiny')
        stroke_w = '0.1mm'
    return dwg, compatibility_mode, stroke_w


def generate_svg(pads, material, width_mm, height_mm, filename, hole_dia_preset, settings, polygon=None):
    placed, _, _ = _nest_discs(pads, material, width_mm, height_mm, settings, polygon=polygon)
    dwg, compatibility_mode, stroke_w = _create_svg_drawing(filename, width_mm, height_mm, settings)
    _render_svg_discs(dwg, placed, material, hole_dia_preset, settings, compatibility_mode, stroke_w)
    dwg.save()


# ==========================================
# SCRAP MODE HELPERS
# ==========================================

def compute_remaining_pads(original_pads, placed):
    """
    Compute remaining pads after partial placement.

    Args:
        original_pads: List of {'size': float, 'qty': int|'max'} dicts
        placed: List of (pad_size, cx, cy, r) tuples from _nest_discs

    Returns:
        List of {'size': float, 'qty': int} for unplaced pads (excludes 'max' qty)
    """
    # Count what was placed by size
    placed_counts = {}
    for pad_size, cx, cy, r in placed:
        placed_counts[pad_size] = placed_counts.get(pad_size, 0) + 1

    # Subtract from original (only fixed-qty pads, not 'max')
    remaining = []
    for pad in original_pads:
        if pad['qty'] == 'max':
            continue  # Max pads don't carry over between scraps

        size = pad['size']
        original_qty = pad['qty']
        placed_qty = placed_counts.get(size, 0)

        # Calculate remaining for this size
        remaining_qty = original_qty - placed_qty

        if remaining_qty > 0:
            remaining.append({'size': size, 'qty': remaining_qty})

        # Reduce placed_counts so we don't double-count
        if placed_qty > 0:
            placed_counts[size] = max(0, placed_qty - original_qty)

    return remaining


def try_nest_partial(pads, material, width_mm, height_mm, settings, polygon=None):
    """
    Attempt to place as many pads as possible, return placed and remaining.

    This is the main entry point for scrap mode - it tries to fit pads on a scrap
    and returns both what was placed and what's left for the next scrap.

    Args:
        pads: List of {'size': float, 'qty': int|'max'} dicts
        material: Material type string
        width_mm, height_mm: Scrap dimensions in mm
        settings: App settings dict
        polygon: Optional polygon coordinates for irregular shapes

    Returns:
        (placed, remaining_pads, any_placed)
        - placed: [(pad_size, cx, cy, r), ...] - what was placed
        - remaining_pads: [{'size': float, 'qty': int}, ...] - what's left
        - any_placed: bool - True if at least one pad was placed
    """
    placed, fixed_placed, fixed_total = _nest_discs(
        pads, material, width_mm, height_mm, settings, polygon=polygon
    )
    remaining = compute_remaining_pads(pads, placed)
    any_placed = len(placed) > 0
    return placed, remaining, any_placed


def generate_svg_from_placed(placed, material, width_mm, height_mm, filename, hole_dia_preset, settings, polygon=None):
    """
    Generate SVG from pre-computed placed discs.

    This is used by scrap mode where nesting is done separately via try_nest_partial().
    The function draws all the placed discs without re-running the nesting algorithm.

    Args:
        placed: List of (pad_size, cx, cy, r) tuples
        material: Material type string
        width_mm, height_mm: Sheet dimensions in mm
        filename: Output file path
        hole_dia_preset: Hole diameter setting
        settings: App settings dict
        polygon: Optional (unused, for API consistency)
    """
    dwg, compatibility_mode, stroke_w = _create_svg_drawing(filename, width_mm, height_mm, settings)
    _render_svg_discs(dwg, placed, material, hole_dia_preset, settings, compatibility_mode, stroke_w)
    dwg.save()


# ==========================================
# DIE INSERT SVG GENERATION
# ==========================================

def _render_svg_dies(dwg, placed, settings, compatibility_mode, stroke_w):
    """Render placed die rings (outer cut, inner cut, engravings) into an SVG drawing."""
    layer_colors = settings.get("layer_colors", DEFAULT_SETTINGS["layer_colors"])
    tooling = settings.get("tooling_settings", {})
    engrave_ring = tooling.get("engrave_ring", True)
    engrave_cutout = tooling.get("engrave_cutout", True)
    engraving_mode = tooling.get("engraving_mode", "filled")

    # Get kerf from acrylic gcode settings for cutout size calculation
    gcode_settings = settings.get("gcode_settings", {})
    acrylic_settings = gcode_settings.get("acrylic", {})
    kerf_width = acrylic_settings.get("kerf_width", 0.15)

    outer_cut_color = layer_colors.get('die_outer_cut', '#FF0000')
    inner_cut_color = layer_colors.get('die_inner_cut', '#0000FF')
    engraving_color = layer_colors.get('die_engraving', '#00E000')
    cutout_eng_color = layer_colors.get('die_cutout_engraving', '#FF8000')

    ring_font_size = tooling.get("ring_font_size", 3.5)
    cutout_font_size = tooling.get("cutout_font_size", 3.5)
    ring_location = tooling.get("ring_engraving_location", "centered")
    ring_offset = tooling.get("ring_engraving_offset", 0.0)

    for pad_size, cx, cy, r in placed:
        inner_r = pad_size / 2

        # Outer cut circle
        if compatibility_mode:
            dwg.add(dwg.circle(center=(cx, cy), r=r, stroke=outer_cut_color, fill='none', stroke_width=stroke_w))
        else:
            dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{r}mm", stroke=outer_cut_color, fill='none', stroke_width=stroke_w))

        # Inner cut circle (pad-sized hole)
        if compatibility_mode:
            dwg.add(dwg.circle(center=(cx, cy), r=inner_r, stroke=inner_cut_color, fill='none', stroke_width=stroke_w))
        else:
            dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{inner_r}mm", stroke=inner_cut_color, fill='none', stroke_width=stroke_w))

        # Engraving on ring (pad size label between inner and outer circles)
        if engrave_ring:
            text_content = f"{pad_size:.1f}".rstrip('0').rstrip('.')
            font_size = ring_font_size
            ring_width = r - inner_r
            # Clamp font size to ring width
            if font_size > ring_width * 0.9:
                font_size = ring_width * 0.9
            if font_size < 1.0:
                font_size = 1.0

            if ring_location == "from_outside":
                label_y = cy - (r - ring_offset) + font_size * 0.35
            else:  # centered
                ring_center_r = (r + inner_r) / 2
                label_y = cy - ring_center_r + font_size * 0.35

            if compatibility_mode:
                dwg.add(dwg.text(text_content, insert=(cx, label_y),
                                 text_anchor="middle", font_size=font_size,
                                 fill=engraving_color))
            else:
                dwg.add(dwg.text(text_content, insert=(f"{cx}mm", f"{label_y}mm"),
                                 text_anchor="middle", font_size=f"{font_size}mm",
                                 fill=engraving_color))

        # Engraving on inner cutout disc (actual size after kerf)
        if engrave_cutout:
            actual_size = pad_size - kerf_width
            text_content = f"{actual_size:.2f}".rstrip('0').rstrip('.')
            font_size = cutout_font_size
            # Only engrave if inner disc is large enough
            if inner_r > font_size:
                label_y = cy + font_size * 0.35
                if compatibility_mode:
                    dwg.add(dwg.text(text_content, insert=(cx, label_y),
                                     text_anchor="middle", font_size=font_size,
                                     fill=cutout_eng_color))
                else:
                    dwg.add(dwg.text(text_content, insert=(f"{cx}mm", f"{label_y}mm"),
                                     text_anchor="middle", font_size=f"{font_size}mm",
                                     fill=cutout_eng_color))


def generate_die_svg(pads, width_mm, height_mm, filename, settings):
    """Generate SVG for die insert rings."""
    placed, _, _ = _nest_discs(pads, 'die_ring', width_mm, height_mm, settings)
    dwg, compatibility_mode, stroke_w = _create_svg_drawing(filename, width_mm, height_mm, settings)
    _render_svg_dies(dwg, placed, settings, compatibility_mode, stroke_w)
    dwg.save()
    return placed


def generate_die_svg_from_placed(placed, width_mm, height_mm, filename, settings):
    """Generate SVG for die inserts from pre-computed placements (scrap mode)."""
    dwg, compatibility_mode, stroke_w = _create_svg_drawing(filename, width_mm, height_mm, settings)
    _render_svg_dies(dwg, placed, settings, compatibility_mode, stroke_w)
    dwg.save()


# ==========================================
# DIE HOLDER SVG GENERATION
# ==========================================

# Die holder constants
HOLDER_OUTER_R = 42.5        # 85mm outer diameter
HOLDER_MAGNET_HOLE_R = 1.75  # 3.5mm magnet hole
HOLDER_PIN_HOLE_R = 3.0      # 6.0mm pin hole
HOLDER_PIN_OFFSET = 25.0     # Pin hole distance from center
HOLDER_LARGE_INNER_R = 35.0  # 70mm inner for large holder
HOLDER_SMALL_INNER_R = 25.0  # 50mm inner for small holder

def generate_holder_svg(variant, filename, settings):
    """
    Generate SVG for die holder pieces.

    Args:
        variant: "large", "small", or "both"
        filename: Output file path
        settings: App settings dict
    """
    layer_colors = settings.get("layer_colors", DEFAULT_SETTINGS["layer_colors"])
    cut_color = layer_colors.get('die_holder_cut', '#FF0000')
    hole_color = layer_colors.get('die_holder_hole', '#0000FF')

    spacing = 5.0
    outer_d = HOLDER_OUTER_R * 2

    # Determine pieces to generate
    # Each holder variant needs: solid disc, magnet disc, pin disc, retaining ring
    pieces = []
    if variant in ('large', 'both'):
        pieces.append(('solid', None))
        pieces.append(('magnet', None))
        pieces.append(('pin', None))
        pieces.append(('ring', HOLDER_LARGE_INNER_R))
    if variant in ('small', 'both'):
        if variant == 'both':
            # Shared layers already added, just add the small retaining ring
            pieces.append(('ring', HOLDER_SMALL_INNER_R))
        else:
            pieces.append(('solid', None))
            pieces.append(('magnet', None))
            pieces.append(('pin', None))
            pieces.append(('ring', HOLDER_SMALL_INNER_R))

    # Layout: arrange in a row
    num_pieces = len(pieces)
    width_mm = num_pieces * outer_d + (num_pieces + 1) * spacing
    height_mm = outer_d + 2 * spacing

    dwg, compatibility_mode, stroke_w = _create_svg_drawing(filename, width_mm, height_mm, settings)

    for i, (piece_type, inner_r) in enumerate(pieces):
        cx = spacing + HOLDER_OUTER_R + i * (outer_d + spacing)
        cy = spacing + HOLDER_OUTER_R

        # Outer circle (always present)
        if compatibility_mode:
            dwg.add(dwg.circle(center=(cx, cy), r=HOLDER_OUTER_R, stroke=cut_color, fill='none', stroke_width=stroke_w))
        else:
            dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{HOLDER_OUTER_R}mm", stroke=cut_color, fill='none', stroke_width=stroke_w))

        if piece_type == 'magnet':
            # Center hole for magnet
            if compatibility_mode:
                dwg.add(dwg.circle(center=(cx, cy), r=HOLDER_MAGNET_HOLE_R, stroke=hole_color, fill='none', stroke_width=stroke_w))
            else:
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{HOLDER_MAGNET_HOLE_R}mm", stroke=hole_color, fill='none', stroke_width=stroke_w))

        elif piece_type == 'pin':
            # Center alignment hole + offset pin hole
            if compatibility_mode:
                dwg.add(dwg.circle(center=(cx, cy), r=HOLDER_MAGNET_HOLE_R, stroke=hole_color, fill='none', stroke_width=stroke_w))
                dwg.add(dwg.circle(center=(cx, cy + HOLDER_PIN_OFFSET), r=HOLDER_PIN_HOLE_R, stroke=hole_color, fill='none', stroke_width=stroke_w))
            else:
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{HOLDER_MAGNET_HOLE_R}mm", stroke=hole_color, fill='none', stroke_width=stroke_w))
                pin_y = cy + HOLDER_PIN_OFFSET
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{pin_y}mm"), r=f"{HOLDER_PIN_HOLE_R}mm", stroke=hole_color, fill='none', stroke_width=stroke_w))

        elif piece_type == 'ring':
            # Retaining ring: inner circle
            if compatibility_mode:
                dwg.add(dwg.circle(center=(cx, cy), r=inner_r, stroke=cut_color, fill='none', stroke_width=stroke_w))
            else:
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{inner_r}mm", stroke=cut_color, fill='none', stroke_width=stroke_w))

    dwg.save()
