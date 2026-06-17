import math
import os
import sys
import svgwrite
from config import (DEFAULT_SETTINGS, get_dart_settings_for_size, get_sizing_for_size,
                    get_engraving_settings_for_size, get_engraving_placement_for_size)

# numpy is optional — used only to vectorize the radial-bias nesting scan.
# The macOS Intel build ships without numpy; in that case we fall back to the
# pure-Python implementation (slower but identical results).
try:
    import numpy as _np
    _HAS_NUMPY = True
except ImportError:  # pragma: no cover
    _np = None
    _HAS_NUMPY = False

# ==========================================
# CORE MATH & LOGIC
# ==========================================

_SQUARE_POWER = 0.01  # small power → sign(c) * |c|^p approaches a square wave


def _wave_value(raw_cos, shape_factor):
    """Map a raw cosine value (-1..1) to a shaped wave value (-1..1).

    shape_factor spans triangle → sine → square:
      0.0  = Triangle (linear ramps between peak and valley)
      0.5  = Sine (the raw cosine itself)
      1.0  = Square (sharp transitions between flats at +/-1)
    Values in between blend smoothly — 0..0.5 blends triangle→sine,
    0.5..1 blends sine→square.
    """
    s = max(0.0, min(1.0, shape_factor))

    # Clamp for numerical safety before arcsin
    raw_c = max(-1.0, min(1.0, raw_cos))

    triangle = (2.0 / math.pi) * math.asin(raw_c)
    sine = raw_c
    sign = 1.0 if raw_c >= 0 else -1.0
    square = sign * (abs(raw_c) ** _SQUARE_POWER)

    if s <= 0.5:
        t = s * 2.0
        return (1.0 - t) * triangle + t * sine
    t = (s - 0.5) * 2.0
    return (1.0 - t) * sine + t * square


def calculate_star_path(cx, cy, outer_r, inner_r, num_points=12, shape_factor=0.5):
    """
    Generates an SVG path string for a darted (geared) leather pad shape.
    shape_factor spans 0.0 (Triangle) → 0.5 (Sine) → 1.0 (Square).
    """
    path_data = []

    avg_r = (outer_r + inner_r) / 2.0
    amplitude = (outer_r - inner_r) / 2.0

    steps = int(num_points * 8)
    if steps < 64:
        steps = 64

    angle_step = (2 * math.pi) / steps

    for i in range(steps + 1):
        theta = i * angle_step
        raw_wave = math.cos(num_points * theta)
        shaped_wave = _wave_value(raw_wave, shape_factor)
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

def get_felt_thickness_mm(settings, sizing=None):
    src = sizing or settings
    thickness = src.get("felt_thickness", 3.175)
    if src.get("felt_thickness_unit") == "in":
        return thickness * 25.4
    return thickness

def get_disc_diameter(pad_size, material, settings):
    if material == 'die_ring':
        return 70.0 if pad_size >= 40.0 else 50.0

    sizing = get_sizing_for_size(pad_size, settings)

    if material == 'felt': return pad_size - sizing["felt_offset"]
    if material == 'card': return pad_size - (sizing["felt_offset"] + sizing["card_to_felt_offset"])
    if material == 'exact_size': return pad_size

    if material == 'leather':
        dart_cfg = get_dart_settings_for_size(pad_size, settings)
        if dart_cfg:
            bonus = dart_cfg.get("wrap_bonus", 0.75)
            wrap = leather_back_wrap(pad_size, sizing["leather_wrap_multiplier"], extra_base=bonus)
        else:
            wrap = leather_back_wrap(pad_size, sizing["leather_wrap_multiplier"])

        felt_thickness_mm = get_felt_thickness_mm(settings, sizing)
        diameter = pad_size + 2 * (felt_thickness_mm + wrap)
        return round(diameter * 2) / 2

    return 0

def should_have_center_hole(pad_size, hole_dia, settings):
    sizing = get_sizing_for_size(pad_size, settings)
    min_size = sizing.get("min_hole_size", 16.5)
    return hole_dia > 0 and pad_size >= min_size

def check_for_oversized_engravings(pads, material_vars, settings):
    oversized = {}
    for material, var in material_vars.items():
        if not var.get(): continue

        oversized_sizes = set()
        for pad in pads:
            pad_size = pad['size']
            eng_cfg = get_engraving_settings_for_size(pad_size, settings)
            font_size = eng_cfg.get("engraving_font_size", {}).get(material, 2.0)
            diameter = get_disc_diameter(pad_size, material, settings)
            radius = diameter / 2
            if font_size >= radius * 0.8:
                oversized_sizes.add(pad_size)

        if oversized_sizes:
            oversized[material] = oversized_sizes
    return oversized

def _scan_radial_python(dia, r_val, placed_list, target_x, target_y,
                         width_mm, height_mm, spacing_mm):
    """Reference Python implementation of the radial (closest-to-target) scan.

    Iterates the full sheet at 1 mm steps starting from spacing_mm, keeping
    the closest valid (non-colliding) cell to the target point. Iteration
    order is y-outer, x-inner so that on ties the earlier-visited cell wins.

    Kept as a fallback for environments without numpy and as the parity
    reference for tests.
    """
    best_pos = None
    best_dist = float('inf')
    y = spacing_mm
    while y + dia + spacing_mm <= height_mm:
        x = spacing_mm
        while x + dia + spacing_mm <= width_mm:
            cx, cy = x + r_val, y + r_val
            dist = math.sqrt((cx - target_x) ** 2 + (cy - target_y) ** 2)
            if dist < best_dist:
                is_collision = any(
                    (cx - px)**2 + (cy - py)**2 < (r_val + pr + spacing_mm)**2
                    for _, px, py, pr in placed_list)
                if not is_collision:
                    best_dist = dist
                    best_pos = (cx, cy)
            x += 1
        y += 1
    return best_pos


def _scan_radial_numpy(dia, r_val, placed_list, target_x, target_y,
                        width_mm, height_mm, spacing_mm):
    """Vectorized radial scan. Identical results to _scan_radial_python.

    Builds the candidate grid as a numpy array, computes squared distance
    to the target in one bulk op, argsorts (stable so y-then-x tie-break
    is preserved), then walks the sorted candidates and returns the first
    non-colliding cell — which is by construction the closest valid cell.
    """
    # The Python loop visits x in [spacing_mm, spacing_mm+1, ...] while
    # x + dia + spacing_mm <= width_mm. So the largest x is x_max =
    # width_mm - dia - spacing_mm. The number of integer steps from
    # spacing_mm to x_max inclusive is floor(x_max - spacing_mm) + 1.
    x_max = width_mm - dia - spacing_mm
    y_max = height_mm - dia - spacing_mm
    eps = 1e-9
    if x_max < spacing_mm - eps or y_max < spacing_mm - eps:
        return None
    n_x = int(math.floor(x_max - spacing_mm + eps)) + 1
    n_y = int(math.floor(y_max - spacing_mm + eps)) + 1
    if n_x <= 0 or n_y <= 0:
        return None

    xs = spacing_mm + _np.arange(n_x, dtype=_np.float64)
    ys = spacing_mm + _np.arange(n_y, dtype=_np.float64)
    cxs = xs + r_val  # candidate disc-center x for each column
    cys = ys + r_val  # candidate disc-center y for each row
    cxs0 = float(cxs[0])
    cys0 = float(cys[0])

    # Build occupancy mask: True where a candidate disc would collide with
    # an already-placed pad. Each placed pad only affects cells inside the
    # bounding box of its exclusion zone, so the per-pad work scales with
    # pad area, not full-sheet area.
    occupied = _np.zeros((n_y, n_x), dtype=bool)
    for _, px, py, pr in placed_list:
        rsum = r_val + pr + spacing_mm
        xi_min = int(math.ceil((px - rsum - cxs0) - eps))
        xi_max = int(math.floor((px + rsum - cxs0) + eps))
        yi_min = int(math.ceil((py - rsum - cys0) - eps))
        yi_max = int(math.floor((py + rsum - cys0) + eps))
        if xi_min < 0:
            xi_min = 0
        if yi_min < 0:
            yi_min = 0
        if xi_max > n_x - 1:
            xi_max = n_x - 1
        if yi_max > n_y - 1:
            yi_max = n_y - 1
        if xi_max < xi_min or yi_max < yi_min:
            continue
        sub_cx = cxs[xi_min:xi_max + 1]
        sub_cy = cys[yi_min:yi_max + 1]
        sub_d2 = (sub_cx[None, :] - px) ** 2 + (sub_cy[:, None] - py) ** 2
        occupied[yi_min:yi_max + 1, xi_min:xi_max + 1] |= sub_d2 < rsum * rsum

    # Distance² from each candidate to the target. Invalid (occupied) cells
    # become +inf so argmin skips them. argmin's first-occurrence tie-break
    # matches the y-outer/x-inner flat order of the Python reference.
    dx = cxs - target_x
    dy = cys - target_y
    dist_sq = dy[:, None] ** 2 + dx[None, :] ** 2
    if occupied.any():
        masked = _np.where(occupied, _np.inf, dist_sq)
    else:
        masked = dist_sq

    flat_idx = int(_np.argmin(masked))
    if not _np.isfinite(masked.flat[flat_idx]):
        return None
    yi = flat_idx // n_x
    xi = flat_idx % n_x
    return (float(cxs[xi]), float(cys[yi]))


def _scan_radial(dia, r_val, placed_list, target_x, target_y,
                  width_mm, height_mm, spacing_mm):
    """Dispatch to the numpy-vectorized scan if available, else Python."""
    if _HAS_NUMPY:
        return _scan_radial_numpy(dia, r_val, placed_list,
                                   target_x, target_y,
                                   width_mm, height_mm, spacing_mm)
    return _scan_radial_python(dia, r_val, placed_list,
                                target_x, target_y,
                                width_mm, height_mm, spacing_mm)


def _nest_discs(pads, material, width_mm, height_mm, settings, spacing_mm=1.0, polygon=None, _discs_override=None):
    """
    Greedy circle-packing algorithm. Returns list of placed discs as (pad_size, cx, cy, r).
    Discs that couldn't be placed are omitted from the result.

    If polygon is provided (list of (x,y) tuples in mm), uses polygon nesting instead of rectangle.

    Supports 'max' quantity: fixed-qty pads are placed first, then max pads fill remaining space.

    ``_discs_override`` (internal): if provided, skip the default build-+
    -sort step and use this pre-ordered list of (pad_size, diameter)
    tuples for the fixed-pad placement. Used by the multistart optimizer
    (try_nest_partial(optimize=True)) to try the same set of pads in
    different orderings and pick the best result.
    """
    if polygon:
        return _nest_discs_polygon(pads, material, settings, polygon,
                                    spacing_mm,
                                    _discs_override=_discs_override)

    # Separate fixed and max pads
    fixed_pads = [p for p in pads if p['qty'] != 'max']
    max_pads = [p for p in pads if p['qty'] == 'max']

    if _discs_override is not None:
        discs = list(_discs_override)
    else:
        # Build disc list from fixed pads
        discs = []
        for pad in fixed_pads:
            pad_size, qty = pad['size'], pad['qty']
            diameter = get_disc_diameter(pad_size, material, settings)
            for _ in range(qty):
                discs.append((pad_size, diameter))

    placed = []
    fixed_total = len(discs)
    fixed_placed = 0

    # Edge bias scan direction
    edge_bias = settings.get("edge_bias", "center")

    if _discs_override is None:
        # Corner bias: smallest first (small discs nestle into corners efficiently).
        # All others: largest first (standard greedy circle packing).
        if edge_bias in ("nw", "ne", "sw", "se"):
            discs.sort(key=lambda x: x[1])
        else:
            discs.sort(key=lambda x: -x[1])
    scan_y_reversed = edge_bias in ("s", "se", "sw")
    scan_x_reversed = edge_bias in ("e", "ne", "se")
    is_radial = edge_bias in ("center", "ne", "nw", "se", "sw")
    x_primary = edge_bias in ("w", "e")

    # Radial targets for distance-based placement
    _radial_targets = {
        "center": (width_mm / 2, height_mm / 2),
        "nw": (0, 0),
        "ne": (width_mm, 0),
        "sw": (0, height_mm),
        "se": (width_mm, height_mm),
    }

    def _scan_place(dia, r_val, placed_list):
        """Scan the sheet for a valid placement respecting edge bias direction."""

        if is_radial:
            target_x, target_y = _radial_targets[edge_bias]
            return _scan_radial(dia, r_val, placed_list, target_x, target_y,
                                width_mm, height_mm, spacing_mm)

        # Cardinal directions: linear scan
        if scan_y_reversed:
            y_start = height_mm - spacing_mm - dia
            y_ok = lambda yv: yv >= spacing_mm
            y_step = -1
        else:
            y_start = spacing_mm
            y_ok = lambda yv: yv + dia + spacing_mm <= height_mm
            y_step = 1

        if scan_x_reversed:
            x_start = width_mm - spacing_mm - dia
            x_ok = lambda xv: xv >= spacing_mm
            x_step = -1
        else:
            x_start = spacing_mm
            x_ok = lambda xv: xv + dia + spacing_mm <= width_mm
            x_step = 1

        if x_primary:
            x = x_start
            while x_ok(x):
                y = y_start
                while y_ok(y):
                    cx, cy = x + r_val, y + r_val
                    is_collision = any((cx - px)**2 + (cy - py)**2 < (r_val + pr + spacing_mm)**2 for _, px, py, pr in placed_list)
                    if not is_collision:
                        return (cx, cy)
                    y += y_step
                x += x_step
        else:
            y = y_start
            while y_ok(y):
                x = x_start
                while x_ok(x):
                    cx, cy = x + r_val, y + r_val
                    is_collision = any((cx - px)**2 + (cy - py)**2 < (r_val + pr + spacing_mm)**2 for _, px, py, pr in placed_list)
                    if not is_collision:
                        return (cx, cy)
                    x += x_step
                y += y_step
        return None

    # Place fixed pads
    for pad_size, dia in discs:
        r = dia / 2
        pos = _scan_place(dia, r, placed)
        if pos:
            placed.append((pad_size, pos[0], pos[1], r))
            fixed_placed += 1

    # Fill remaining space with max pad (if any)
    if max_pads:
        max_pad = max_pads[0]
        max_size = max_pad['size']
        max_dia = get_disc_diameter(max_size, material, settings)
        max_r = max_dia / 2

        while True:
            pos = _scan_place(max_dia, max_r, placed)
            if pos:
                placed.append((max_size, pos[0], pos[1], max_r))
            else:
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


# ==========================================
# VECTORIZED POLYGON HELPERS (numpy)
# ==========================================
# These mirror the scalar polygon helpers above but operate on a 2D grid of
# candidate cells in a single bulk op. Used by the vectorized polygon scan.
# Each returns a numpy array shape (n_y, n_x); the per-pad polygon-nesting
# loops then mask + argmin to find the best valid cell.

def _points_in_polygon_grid(cxs, cys, polygon):
    """Vectorized ray-cast point-in-polygon for a candidate grid.

    cxs: 1D numpy array of candidate x values (length n_x).
    cys: 1D numpy array of candidate y values (length n_y).
    Returns: bool array shape (n_y, n_x), True where (cx, cy) is inside.

    Matches the per-point algorithm in _point_in_polygon exactly: for each
    edge (j -> i), toggle inside where the edge crosses the cell's y AND
    the cell's x is to the left of the intersection.
    """
    cx_grid = cxs[None, :].astype(_np.float64)
    cy_grid = cys[:, None].astype(_np.float64)
    inside = _np.zeros((cy_grid.shape[0], cx_grid.shape[1]), dtype=bool)
    n = len(polygon)
    with _np.errstate(divide='ignore', invalid='ignore'):
        for i in range(n):
            xi, yi = polygon[i]
            xj, yj = polygon[(i - 1) % n]
            cross = (yi > cy_grid) != (yj > cy_grid)
            # For horizontal edges (yi == yj) the division is by zero; the
            # `cross` mask is False for all cells there (both sides equal),
            # so the NaN/inf produced in x_intersect never feeds the xor.
            x_intersect = (xj - xi) * (cy_grid - yi) / (yj - yi) + xi
            below = cx_grid < x_intersect
            inside ^= cross & below
    return inside


def _distances_grid_to_segment(cx_grid, cy_grid, x1, y1, x2, y2):
    """Distance from each cell in (cx_grid, cy_grid) to a single line segment.

    cx_grid, cy_grid: 2D broadcasting-shaped numpy arrays (or compatible).
    Returns: 2D array of distances, same shape as the broadcast of inputs.
    """
    dx = x2 - x1
    dy = y2 - y1
    length_sq = dx * dx + dy * dy
    if length_sq < 1e-12:
        # Degenerate segment — distance to the point.
        return _np.sqrt((cx_grid - x1) ** 2 + (cy_grid - y1) ** 2)
    t = ((cx_grid - x1) * dx + (cy_grid - y1) * dy) / length_sq
    t = _np.clip(t, 0.0, 1.0)
    proj_x = x1 + t * dx
    proj_y = y1 + t * dy
    return _np.sqrt((cx_grid - proj_x) ** 2 + (cy_grid - proj_y) ** 2)


def _distances_grid_to_polygon_edges(cxs, cys, polygon):
    """Min distance from each cell to any polygon edge. Returns (n_y, n_x)."""
    cx_grid = cxs[None, :].astype(_np.float64)
    cy_grid = cys[:, None].astype(_np.float64)
    min_dist = _np.full((cy_grid.shape[0], cx_grid.shape[1]), _np.inf,
                         dtype=_np.float64)
    n = len(polygon)
    for i in range(n):
        x1, y1 = polygon[i]
        x2, y2 = polygon[(i + 1) % n]
        d = _distances_grid_to_segment(cx_grid, cy_grid, x1, y1, x2, y2)
        _np.minimum(min_dist, d, out=min_dist)
    return min_dist


def _distances_grid_to_vertices(cxs, cys, polygon):
    """Min distance from each cell to any polygon vertex. Returns (n_y, n_x)."""
    cx_grid = cxs[None, :].astype(_np.float64)
    cy_grid = cys[:, None].astype(_np.float64)
    min_dist = _np.full((cy_grid.shape[0], cx_grid.shape[1]), _np.inf,
                         dtype=_np.float64)
    for vx, vy in polygon:
        d = _np.sqrt((cx_grid - vx) ** 2 + (cy_grid - vy) ** 2)
        _np.minimum(min_dist, d, out=min_dist)
    return min_dist


def _build_occupancy_grid(cxs, cys, r_val, placed_discs, spacing_mm):
    """Boolean occupancy mask (n_y, n_x) — True where a disc would collide.

    Per-pad work is limited to the bounding box of its exclusion zone so
    cost scales with pad area rather than full-sheet area.
    """
    n_x = len(cxs)
    n_y = len(cys)
    occupied = _np.zeros((n_y, n_x), dtype=bool)
    if not placed_discs:
        return occupied
    cxs0 = float(cxs[0])
    cys0 = float(cys[0])
    # The grid steps by 1 in this codebase. If that ever changes, derive
    # the step from cxs/cys instead of hard-coding it here.
    eps = 1e-9
    for _, px, py, pr in placed_discs:
        rsum = r_val + pr + spacing_mm
        xi_min = int(math.ceil((px - rsum - cxs0) - eps))
        xi_max = int(math.floor((px + rsum - cxs0) + eps))
        yi_min = int(math.ceil((py - rsum - cys0) - eps))
        yi_max = int(math.floor((py + rsum - cys0) + eps))
        if xi_min < 0:
            xi_min = 0
        if yi_min < 0:
            yi_min = 0
        if xi_max > n_x - 1:
            xi_max = n_x - 1
        if yi_max > n_y - 1:
            yi_max = n_y - 1
        if xi_max < xi_min or yi_max < yi_min:
            continue
        sub_cx = cxs[xi_min:xi_max + 1]
        sub_cy = cys[yi_min:yi_max + 1]
        sub_d2 = (sub_cx[None, :] - px) ** 2 + (sub_cy[:, None] - py) ** 2
        occupied[yi_min:yi_max + 1, xi_min:xi_max + 1] |= sub_d2 < rsum * rsum
    return occupied


# ==========================================
# POLYGON SCAN — PYTHON REFERENCE & NUMPY
# ==========================================
# Both implementations match the original nested-function behavior exactly:
# y-outer / x-inner iteration, strict `score < best_score` tie-break (the
# first cell to reach the minimum wins), 1mm grid step starting at
# (min_x + spacing_mm, min_y + spacing_mm).

def _find_best_polygon_large_python(r, placed_discs, polygon,
                                     min_x, max_x, min_y, max_y,
                                     centroid_x, centroid_y,
                                     bias_target_x, bias_target_y, bias_weight,
                                     spacing_mm, step):
    """Reference Python scan for large-disc polygon placement."""
    best_pos = None
    best_score = float('inf')
    y = min_y + spacing_mm
    while y + r <= max_y:
        x = min_x + spacing_mm
        while x + r <= max_x:
            cx, cy = x + r, y + r
            if bias_weight == 0.0:
                bias = 0.0
            else:
                bias = bias_weight * math.sqrt(
                    (cx - bias_target_x) ** 2 + (cy - bias_target_y) ** 2)
            score = math.sqrt(
                (cx - centroid_x) ** 2 + (cy - centroid_y) ** 2) + bias
            if score < best_score:
                if _circle_fits_in_polygon(cx, cy, r, polygon, spacing_mm):
                    is_collision = any(
                        (cx - px) ** 2 + (cy - py) ** 2
                        < (r + pr + spacing_mm) ** 2
                        for _, px, py, pr in placed_discs)
                    if not is_collision:
                        best_score = score
                        best_pos = (cx, cy)
            x += step
        y += step
    return best_pos


def _find_best_polygon_small_python(r, placed_discs, polygon,
                                     min_x, max_x, min_y, max_y,
                                     bias_target_x, bias_target_y, bias_weight,
                                     spacing_mm, step):
    """Reference Python scan for small-disc polygon placement (edge-seeking)."""
    best_pos = None
    best_score = float('inf')
    y = min_y + spacing_mm
    while y + r <= max_y:
        x = min_x + spacing_mm
        while x + r <= max_x:
            cx, cy = x + r, y + r
            if not _circle_fits_in_polygon(cx, cy, r, polygon, spacing_mm):
                x += step
                continue
            is_collision = any(
                (cx - px) ** 2 + (cy - py) ** 2 < (r + pr + spacing_mm) ** 2
                for _, px, py, pr in placed_discs)
            if is_collision:
                x += step
                continue

            edge_dist = _distance_to_nearest_edge(cx, cy, polygon)
            edge_gap = edge_dist - r - spacing_mm
            vertex_dist = _distance_to_nearest_vertex(cx, cy, polygon)
            if not placed_discs:
                snugness = 0
            else:
                min_gap = float('inf')
                for _, px, py, pr in placed_discs:
                    dist = math.sqrt((cx - px) ** 2 + (cy - py) ** 2)
                    gap = dist - (r + pr + spacing_mm)
                    if gap < min_gap:
                        min_gap = gap
                snugness = min_gap
            if bias_weight == 0.0:
                bias = 0.0
            else:
                bias = bias_weight * math.sqrt(
                    (cx - bias_target_x) ** 2 + (cy - bias_target_y) ** 2)
            score = edge_gap * 1.0 + vertex_dist * 0.3 + snugness * 0.5 + bias
            if score < best_score:
                best_score = score
                best_pos = (cx, cy)
            x += step
        y += step
    return best_pos


def _polygon_candidate_grid(r, min_x, max_x, min_y, max_y, spacing_mm, step):
    """Build the same candidate grid as the Python polygon scan.

    Returns (cxs, cys) numpy arrays, or (None, None) if no candidates fit.
    """
    eps = 1e-9
    x_start = min_x + spacing_mm
    y_start = min_y + spacing_mm
    x_max = max_x - r
    y_max = max_y - r
    if x_start > x_max + eps or y_start > y_max + eps:
        return None, None
    n_x = int(math.floor((x_max - x_start) / step + eps)) + 1
    n_y = int(math.floor((y_max - y_start) / step + eps)) + 1
    if n_x <= 0 or n_y <= 0:
        return None, None
    xs = x_start + _np.arange(n_x, dtype=_np.float64) * step
    ys = y_start + _np.arange(n_y, dtype=_np.float64) * step
    cxs = xs + r
    cys = ys + r
    return cxs, cys


def _find_best_polygon_large_numpy(r, placed_discs, polygon,
                                    min_x, max_x, min_y, max_y,
                                    centroid_x, centroid_y,
                                    bias_target_x, bias_target_y, bias_weight,
                                    spacing_mm, step):
    """Vectorized large-disc polygon scan. Identical results to the Python ref."""
    cxs, cys = _polygon_candidate_grid(r, min_x, max_x, min_y, max_y,
                                        spacing_mm, step)
    if cxs is None:
        return None

    inside = _points_in_polygon_grid(cxs, cys, polygon)
    edge_dist = _distances_grid_to_polygon_edges(cxs, cys, polygon)
    valid = inside & (edge_dist >= r + spacing_mm)

    occupied = _build_occupancy_grid(cxs, cys, r, placed_discs, spacing_mm)
    valid &= ~occupied

    if not valid.any():
        return None

    score = _np.sqrt((cxs[None, :] - centroid_x) ** 2
                     + (cys[:, None] - centroid_y) ** 2)
    if bias_weight > 0.0:
        score = score + bias_weight * _np.sqrt(
            (cxs[None, :] - bias_target_x) ** 2
            + (cys[:, None] - bias_target_y) ** 2)

    masked = _np.where(valid, score, _np.inf)
    flat_idx = int(_np.argmin(masked))
    if not _np.isfinite(masked.flat[flat_idx]):
        return None
    n_x = cxs.shape[0]
    yi = flat_idx // n_x
    xi = flat_idx % n_x
    return (float(cxs[xi]), float(cys[yi]))


def _find_best_polygon_small_numpy(r, placed_discs, polygon,
                                    min_x, max_x, min_y, max_y,
                                    bias_target_x, bias_target_y, bias_weight,
                                    spacing_mm, step):
    """Vectorized small-disc polygon scan. Identical results to the Python ref."""
    cxs, cys = _polygon_candidate_grid(r, min_x, max_x, min_y, max_y,
                                        spacing_mm, step)
    if cxs is None:
        return None

    inside = _points_in_polygon_grid(cxs, cys, polygon)
    edge_dist = _distances_grid_to_polygon_edges(cxs, cys, polygon)
    valid = inside & (edge_dist >= r + spacing_mm)

    occupied = _build_occupancy_grid(cxs, cys, r, placed_discs, spacing_mm)
    valid &= ~occupied

    if not valid.any():
        return None

    vertex_dist = _distances_grid_to_vertices(cxs, cys, polygon)
    edge_gap = edge_dist - r - spacing_mm

    if placed_discs:
        snugness = _np.full((cys.shape[0], cxs.shape[0]), _np.inf,
                             dtype=_np.float64)
        for _, px, py, pr in placed_discs:
            d = _np.sqrt((cxs[None, :] - px) ** 2 + (cys[:, None] - py) ** 2)
            gap = d - (r + pr + spacing_mm)
            _np.minimum(snugness, gap, out=snugness)
    else:
        snugness = _np.zeros((cys.shape[0], cxs.shape[0]), dtype=_np.float64)

    score = edge_gap * 1.0 + vertex_dist * 0.3 + snugness * 0.5
    if bias_weight > 0.0:
        score = score + bias_weight * _np.sqrt(
            (cxs[None, :] - bias_target_x) ** 2
            + (cys[:, None] - bias_target_y) ** 2)

    masked = _np.where(valid, score, _np.inf)
    flat_idx = int(_np.argmin(masked))
    if not _np.isfinite(masked.flat[flat_idx]):
        return None
    n_x = cxs.shape[0]
    yi = flat_idx // n_x
    xi = flat_idx % n_x
    return (float(cxs[xi]), float(cys[yi]))


def _find_best_polygon_large(*args, **kwargs):
    """Dispatch to vectorized scan if numpy is available."""
    if _HAS_NUMPY:
        return _find_best_polygon_large_numpy(*args, **kwargs)
    return _find_best_polygon_large_python(*args, **kwargs)


def _find_best_polygon_small(*args, **kwargs):
    """Dispatch to vectorized scan if numpy is available."""
    if _HAS_NUMPY:
        return _find_best_polygon_small_numpy(*args, **kwargs)
    return _find_best_polygon_small_python(*args, **kwargs)


def _nest_discs_polygon(pads, material, settings, polygon, spacing_mm=1.0, _discs_override=None):
    """
    Smart circle-packing algorithm for polygon boundaries.

    - Large discs: prefer positions closest to centroid (center-out)
    - Small discs: prefer edges, corners, and snug fits against other discs

    Supports 'max' quantity: fixed-qty pads are placed first, then max pads fill remaining space.

    Returns list of placed discs as (pad_size, cx, cy, r).

    ``_discs_override``: see _nest_discs.
    """
    # Separate fixed and max pads
    fixed_pads = [p for p in pads if p['qty'] != 'max']
    max_pads = [p for p in pads if p['qty'] == 'max']

    if _discs_override is not None:
        discs = list(_discs_override)
    else:
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
    if settings.get("dart_range_mode", "universal") == "range":
        ranges = settings.get("dart_ranges", [])
        size_threshold = max((r["max_size"] for r in ranges), default=18.0) if ranges else 18.0
    else:
        size_threshold = settings.get("dart_threshold", 18.0)

    # Edge bias: compute a target point that pulls discs toward the biased edge/corner
    edge_bias = settings.get("edge_bias", "center")
    # Map bias direction to a target point on/near the polygon bounding box
    _bias_targets = {
        "center": (centroid_x, centroid_y),
        "n":  (centroid_x, min_y),
        "ne": (max_x, min_y),
        "e":  (max_x, centroid_y),
        "se": (max_x, max_y),
        "s":  (centroid_x, max_y),
        "sw": (min_x, max_y),
        "w":  (min_x, centroid_y),
        "nw": (min_x, min_y),
    }
    bias_target_x, bias_target_y = _bias_targets.get(edge_bias, (centroid_x, centroid_y))
    # Bias weight controls how strongly packing favors the biased direction
    # Scale relative to polygon size so it works for any shape
    bias_weight = 0.0 if edge_bias == "center" else 2.0

    # Use 1mm grid steps for accuracy (worth the extra time for scrap efficiency)
    step = 1

    # Place fixed pads via dispatch helpers (numpy vectorized when available,
    # falls back to the bit-identical Python reference otherwise).
    for pad_size, dia in discs:
        r = dia / 2
        if pad_size >= size_threshold:
            best_pos = _find_best_polygon_large(
                r, placed, polygon,
                min_x, max_x, min_y, max_y,
                centroid_x, centroid_y,
                bias_target_x, bias_target_y, bias_weight,
                spacing_mm, step)
        else:
            best_pos = _find_best_polygon_small(
                r, placed, polygon,
                min_x, max_x, min_y, max_y,
                bias_target_x, bias_target_y, bias_weight,
                spacing_mm, step)

        if best_pos:
            placed.append((pad_size, best_pos[0], best_pos[1], r))
            fixed_placed += 1

    # Fill remaining space with max pad (if any) — uses the large-disc scan.
    if max_pads:
        max_pad = max_pads[0]
        max_size = max_pad['size']
        max_dia = get_disc_diameter(max_size, material, settings)
        max_r = max_dia / 2

        while True:
            best_pos = _find_best_polygon_large(
                max_r, placed, polygon,
                min_x, max_x, min_y, max_y,
                centroid_x, centroid_y,
                bias_target_x, bias_target_y, bias_weight,
                spacing_mm, step)
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
        sizing = get_sizing_for_size(pad_size, settings)
        eng_cfg = get_engraving_settings_for_size(pad_size, settings)
        plc_cfg = get_engraving_placement_for_size(pad_size, settings)
        dart_cfg = get_dart_settings_for_size(pad_size, settings) if material == 'leather' else None
        is_dart_pad = dart_cfg is not None

        if is_dart_pad:
            # --- STAR LOGIC ---
            felt_thick = get_felt_thickness_mm(settings, sizing)
            overwrap = dart_cfg.get("overwrap", 0.5)

            felt_r = (pad_size - sizing["felt_offset"]) / 2
            inner_r = felt_r + felt_thick + overwrap
            outer_r = r

            if inner_r >= outer_r:
                inner_r = outer_r - 0.2

            circumference = 2 * math.pi * inner_r
            freq_mult = dart_cfg.get("frequency_multiplier", 1.0)
            num_points = int((circumference / 3.5) * freq_mult)
            if num_points < 12:
                num_points = 12
            if num_points % 2 != 0:
                num_points += 1

            shape_factor = dart_cfg.get("shape_factor", 0.5)
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

        font_size = eng_cfg.get("engraving_font_size", {}).get(material, 2.0)

        # --- Determine Engraving Settings (Standard vs Star) ---
        should_engrave = False

        if is_dart_pad:
            if dart_cfg.get("engraving_on", True):
                engraving_settings = plc_cfg.get("engraving_location", {}).get("darted_leather", {"mode": "from_outside", "value": 2.5})
                should_engrave = True
        else:
            if eng_cfg.get("engraving_on", True):
                engraving_settings = plc_cfg.get("engraving_location", {}).get(material, {"mode": "centered", "value": 0})
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


def nest_pads(pads, material, width_mm, height_mm, settings, polygon=None):
    """Run the nesting algorithm and return placed disc positions.

    Returns list of (pad_size, cx, cy, r) tuples. Used by the preview
    window to show the layout before committing to file generation.
    """
    placed, _, _ = _nest_discs(pads, material, width_mm, height_mm, settings, polygon=polygon)
    return placed


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


def try_nest_partial(pads, material, width_mm, height_mm, settings, polygon=None, optimize=False):
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
        optimize: If True, run multistart greedy — try several disc
            orderings and return the best result. Costs ~5x compute
            (typically 5-30s for ≥75 pads) but often fits 5-15% more
            pads per scrap. Used by the "large batch optimization"
            opt-in flow in scrap mode.

    Returns:
        (placed, remaining_pads, any_placed)
        - placed: [(pad_size, cx, cy, r), ...] - what was placed
        - remaining_pads: [{'size': float, 'qty': int}, ...] - what's left
        - any_placed: bool - True if at least one pad was placed
    """
    if optimize:
        placed = _multistart_nest(
            pads, material, width_mm, height_mm, settings, polygon=polygon)
    else:
        placed, _fixed_placed, _fixed_total = _nest_discs(
            pads, material, width_mm, height_mm, settings, polygon=polygon
        )
    remaining = compute_remaining_pads(pads, placed)
    any_placed = len(placed) > 0
    return placed, remaining, any_placed


def _multistart_nest(pads, material, width_mm, height_mm, settings, polygon=None):
    """Multistart greedy nesting: try the default ordering plus several
    alternatives, return whichever fit the most pads.

    The greedy nester is a "local-search" algorithm — each disc placement
    is locally optimal but the OVERALL packing can be suboptimal because
    earlier placements constrain later ones. Trying a handful of input
    orderings is the cheapest way to escape local optima: each ordering
    explores a different region of the solution space, and on large pad
    sets the best of 5 typically beats the default by 5-15%.

    Orderings tried:
      1. Largest first (the current default; near-optimal for most cases).
      2. Smallest first (sometimes wins when scrap shape favors corner-
         packing or has many small features).
      3-5. Random shuffles with a fixed seed (reproducible across runs).

    Returns the placed list with the most discs; ties broken by the
    first ordering that achieved that count.
    """
    import random

    fixed_pads = [p for p in pads if p['qty'] != 'max']

    # Build the base disc list once; orderings are permutations of this.
    base_discs = []
    for pad in fixed_pads:
        pad_size, qty = pad['size'], pad['qty']
        diameter = get_disc_diameter(pad_size, material, settings)
        for _ in range(qty):
            base_discs.append((pad_size, diameter))

    orderings = [
        sorted(base_discs, key=lambda d: -d[1]),  # largest first
        sorted(base_discs, key=lambda d: d[1]),   # smallest first
    ]
    rng = random.Random(0)  # reproducible
    for _ in range(3):
        shuffled = list(base_discs)
        rng.shuffle(shuffled)
        orderings.append(shuffled)

    best_placed = None
    for ordering in orderings:
        placed, _fp, _ft = _nest_discs(
            pads, material, width_mm, height_mm, settings,
            polygon=polygon, _discs_override=ordering)
        if best_placed is None or len(placed) > len(best_placed):
            best_placed = placed
    return best_placed or []


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

        # "NOY" arced opposite the size number (Phil Noy credit on each
        # die insert). Only render if the ring size label was rendered too —
        # otherwise there's nothing to be opposite of.
        # Rendered as polylines via the same gcode arc helpers (filled or
        # stroke depending on engraving_mode) for SVG/laser parity.
        if engrave_ring:
            from gcode_engine import (get_text_strokes_arc,
                                       get_filled_text_strokes_arc)
            credit_font = ring_font_size
            ring_width_local = r - inner_r
            if credit_font > ring_width_local * 0.9:
                credit_font = ring_width_local * 0.9
            if credit_font < 1.0:
                credit_font = 1.0
            ring_center_r = (r + inner_r) / 2
            if engraving_mode == "filled":
                filled_spacing = acrylic_settings.get("filled_line_spacing", 0.15)
                noy_strokes = get_filled_text_strokes_arc(
                    "NOY", credit_font,
                    cx, 0.0, ring_center_r, -math.pi / 2.0,
                    filled_spacing, side='bottom')
            else:
                noy_strokes = get_text_strokes_arc(
                    "NOY", credit_font,
                    cx, 0.0, ring_center_r, -math.pi / 2.0, side='bottom')
            for stroke in noy_strokes:
                # Y-up gcode -> SVG Y-down. Raw numbers (no unit suffix)
                # so svgwrite's Tiny 1.2 polyline validator accepts them.
                pts = [(x, cy - y) for (x, y) in stroke]
                dwg.add(dwg.polyline(points=pts, stroke=engraving_color,
                                     fill='none', stroke_width=stroke_w))

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
HOLDER_MAGNET_HOLE_R = 3.25  # 6.5mm magnet hole (layer 2)
HOLDER_PIN_HOLE_R = 1.75     # 3.5mm pin/alignment hole (layers 3-5)
HOLDER_LARGE_INNER_R = 35.0  # 70mm inner for large holder
HOLDER_SMALL_INNER_R = 25.0  # 50mm inner for small holder
HOLDER_LAYERS = 6            # Total layers in a holder stack
HOLDER_THICKNESS_MM = HOLDER_LAYERS * 3.0  # 18mm total

def _holder_pieces_for(variant, layer_count):
    """Return the piece list for a holder variant and layer count.

    Each piece is (type, inner_r). 5-layer = 2x pin, 6-layer = 3x pin.
    "both" = two complete independent holders (no shared layers).
    """
    if layer_count not in (5, 6):
        raise ValueError(f"layer_count must be 5 or 6, got {layer_count}")
    if variant not in ("large", "small", "both"):
        raise ValueError(f"variant must be 'large', 'small', or 'both', got {variant!r}")

    pin_count = layer_count - 3  # solid + magnet + Npin + ring = layer_count

    def one_holder(inner_r):
        return [('solid', None), ('magnet', None)] + \
               [('pin', None)] * pin_count + \
               [('ring', inner_r)]

    if variant == 'large':
        return one_holder(HOLDER_LARGE_INNER_R)
    if variant == 'small':
        return one_holder(HOLDER_SMALL_INNER_R)
    return one_holder(HOLDER_LARGE_INNER_R) + one_holder(HOLDER_SMALL_INNER_R)


def _pack_holder_grid(num_pieces, sheet_w_mm, sheet_h_mm,
                      outer_d=HOLDER_OUTER_R * 2, spacing=5.0):
    """Pick a grid (cols, rows) that fits num_pieces uniform OD circles into the sheet.

    Returns (cols, rows) for the layout closest to square that fits, or None
    if the sheet is too small. Pieces sit on a 5mm gutter from the top-left.
    """
    if num_pieces <= 0:
        return None
    pitch = outer_d + spacing
    max_cols = int((sheet_w_mm - spacing) // pitch)
    max_rows = int((sheet_h_mm - spacing) // pitch)
    if max_cols < 1 or max_rows < 1:
        return None
    if max_cols * max_rows < num_pieces:
        return None
    best = None
    for cols in range(1, max_cols + 1):
        rows = (num_pieces + cols - 1) // cols
        if rows > max_rows:
            continue
        score = (abs(cols - rows), cols * rows)
        if best is None or score < best[0]:
            best = (score, cols, rows)
    return (best[1], best[2]) if best else None


def _min_holder_sheet(num_pieces, outer_d=HOLDER_OUTER_R * 2, spacing=5.0):
    """Smallest near-square sheet (mm) that fits num_pieces uniform circles."""
    cols = max(1, math.ceil(math.sqrt(num_pieces)))
    rows = (num_pieces + cols - 1) // cols
    return (cols * outer_d + (cols + 1) * spacing,
            rows * outer_d + (rows + 1) * spacing)


def generate_holder_svg(variant, filename, settings, *,
                        layer_count=6,
                        sheet_width_mm=None, sheet_height_mm=None):
    """
    Generate SVG for die holder pieces.

    Args:
        variant: "large", "small", or "both" (both = two complete holders)
        filename: Output file path
        settings: App settings dict
        layer_count: 5 (2x pin) or 6 (3x pin). Default 6.
        sheet_width_mm, sheet_height_mm: User sheet size. If None, auto-size
            into a near-square grid (legacy behavior). If supplied and pieces
            don't fit, raises ValueError.
    """
    layer_colors = settings.get("layer_colors", DEFAULT_SETTINGS["layer_colors"])
    cut_color = layer_colors.get('die_holder_cut', '#FF0000')
    hole_color = layer_colors.get('die_holder_hole', '#0000FF')
    # Match the engraving mode used by the gcode renderer so the SVG
    # preview shows the same fill the laser will actually do.
    tooling_settings = settings.get("tooling_settings", {})
    engraving_mode = tooling_settings.get("engraving_mode", "filled")
    gcode_settings_local = settings.get("gcode_settings", {})
    acrylic_local = gcode_settings_local.get("acrylic", {})
    filled_line_spacing_svg = acrylic_local.get("filled_line_spacing", 0.15)

    spacing = 5.0
    outer_d = HOLDER_OUTER_R * 2

    pieces = _holder_pieces_for(variant, layer_count)
    num_pieces = len(pieces)

    if sheet_width_mm is not None and sheet_height_mm is not None:
        layout = _pack_holder_grid(num_pieces, sheet_width_mm, sheet_height_mm,
                                   outer_d, spacing)
        if layout is None:
            min_w, min_h = _min_holder_sheet(num_pieces, outer_d, spacing)
            raise ValueError(
                f"{num_pieces} holder pieces don't fit on a "
                f"{sheet_width_mm:.0f} × {sheet_height_mm:.0f} mm sheet. "
                f"Need at least {min_w:.0f} × {min_h:.0f} mm."
            )
        cols, rows = layout
        width_mm = sheet_width_mm
        height_mm = sheet_height_mm
    else:
        # Legacy auto-size: near-square grid sized exactly to the pieces.
        cols = max(1, math.ceil(math.sqrt(num_pieces)))
        rows = (num_pieces + cols - 1) // cols
        width_mm = cols * outer_d + (cols + 1) * spacing
        height_mm = rows * outer_d + (rows + 1) * spacing

    dwg, compatibility_mode, stroke_w = _create_svg_drawing(filename, width_mm, height_mm, settings)

    for i, (piece_type, inner_r) in enumerate(pieces):
        col = i % cols
        row = i // cols
        cx = spacing + HOLDER_OUTER_R + col * (outer_d + spacing)
        cy = spacing + HOLDER_OUTER_R + row * (outer_d + spacing)

        # Outer circle (always present)
        if compatibility_mode:
            dwg.add(dwg.circle(center=(cx, cy), r=HOLDER_OUTER_R, stroke=cut_color, fill='none', stroke_width=stroke_w))
        else:
            dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{HOLDER_OUTER_R}mm", stroke=cut_color, fill='none', stroke_width=stroke_w))

        if piece_type == 'magnet':
            # Center hole for magnet (6.5mm dia)
            if compatibility_mode:
                dwg.add(dwg.circle(center=(cx, cy), r=HOLDER_MAGNET_HOLE_R, stroke=hole_color, fill='none', stroke_width=stroke_w))
            else:
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{HOLDER_MAGNET_HOLE_R}mm", stroke=hole_color, fill='none', stroke_width=stroke_w))

        elif piece_type == 'pin':
            # Center alignment hole (3.5mm dia)
            if compatibility_mode:
                dwg.add(dwg.circle(center=(cx, cy), r=HOLDER_PIN_HOLE_R, stroke=hole_color, fill='none', stroke_width=stroke_w))
            else:
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{HOLDER_PIN_HOLE_R}mm", stroke=hole_color, fill='none', stroke_width=stroke_w))

        elif piece_type == 'ring':
            # Retaining ring: inner circle
            if compatibility_mode:
                dwg.add(dwg.circle(center=(cx, cy), r=inner_r, stroke=cut_color, fill='none', stroke_width=stroke_w))
            else:
                dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{inner_r}mm", stroke=cut_color, fill='none', stroke_width=stroke_w))

            eng_color = layer_colors.get('die_engraving', '#00E000')
            tooling = settings.get("tooling_settings", {})
            font_size = tooling.get("ring_font_size", 3.5)
            ring_width = HOLDER_OUTER_R - inner_r
            if font_size > ring_width * 0.9:
                font_size = ring_width * 0.9
            ring_center_r = (HOLDER_OUTER_R + inner_r) / 2

            # "DESIGNED BY PHIL NOY" arced along the top of the ring annulus.
            # Phil Noy gave away this method for free; this credit honors him.
            # Replaces the previous size-range label (which was redundant
            # with the variant name shown in the UI).
            #
            # Rendered as polylines (not textPath) so the SVG preview matches
            # the laser-cut path exactly, and so the same code path works
            # under svgwrite's Tiny 1.2 profile (which doesn't support
            # textPath). The arc helpers compute Y-up gcode coords; we flip
            # to SVG Y-down by negating the Y component around cy.
            from gcode_engine import (get_text_strokes_arc,
                                       get_filled_text_strokes_arc)
            if engraving_mode == "filled":
                credit_strokes = get_filled_text_strokes_arc(
                    "DESIGNED BY PHIL NOY", font_size,
                    cx, 0.0, ring_center_r, math.pi / 2.0,
                    filled_line_spacing_svg, side='top')
            else:
                credit_strokes = get_text_strokes_arc(
                    "DESIGNED BY PHIL NOY", font_size,
                    cx, 0.0, ring_center_r, math.pi / 2.0, side='top')
            for stroke in credit_strokes:
                # Y-up gcode -> SVG Y-down. Polyline points are raw numbers
                # in viewBox units (which are mm here); svgwrite Tiny 1.2
                # rejects unit suffixes inside polyline `points`.
                pts = [(x, cy - y) for (x, y) in stroke]
                dwg.add(dwg.polyline(points=pts, stroke=eng_color,
                                     fill='none', stroke_width=stroke_w))

    dwg.save()


# ==========================================
# KERF TEST SVG GENERATION
# ==========================================

KERF_TEST_DIAMETERS = [10.0, 20.0, 30.0]

def generate_die_organizer_svg(variant, filename, settings=None):
    """Copy the bundled die-organizer template SVG to ``filename``.

    Static design (Matt's CAD output): three Upper plates plus one Lower
    glue together with the four corner alignment holes. The asset ships
    as-is — users open the file in LightBurn (or similar) to cut.
    ``settings`` is accepted for API symmetry but unused.
    """
    if variant not in ('upper', 'lower'):
        raise ValueError(f"variant must be 'upper' or 'lower', got {variant!r}")
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(__file__)
    asset = os.path.join(base, 'tooling_assets', f'die_organizer_{variant}.svg')
    if not os.path.exists(asset):
        raise FileNotFoundError(f"Die organizer template missing: {asset}")
    import shutil
    shutil.copyfile(asset, filename)


def generate_kerf_test_svg(material_name, filename, settings):
    """
    Generate an SVG kerf test pattern: 3 circles at known diameters with
    engraved size labels. No kerf compensation applied.
    """
    layer_colors = settings.get("layer_colors", DEFAULT_SETTINGS["layer_colors"])
    cut_color = layer_colors.get('die_outer_cut', '#FF0000')
    eng_color = layer_colors.get('die_engraving', '#00E000')

    spacing = 5.0
    max_r = max(KERF_TEST_DIAMETERS) / 2
    title_height = 8.0
    label_height = 6.0

    # Calculate positions
    positions = []
    x_cursor = spacing
    for dia in KERF_TEST_DIAMETERS:
        r = dia / 2
        cx = x_cursor + max_r
        positions.append((dia, cx, r))
        x_cursor += max_r * 2 + spacing

    width_mm = x_cursor
    cy = title_height + label_height + max_r + spacing
    height_mm = cy + max_r + spacing

    dwg, compat, stroke_w = _create_svg_drawing(filename, width_mm, height_mm, settings)

    # Title
    title_text = f"Kerf Test \u2014 {material_name}"
    title_y = title_height * 0.7
    if compat:
        dwg.add(dwg.text(title_text, insert=(width_mm / 2, title_y),
                         text_anchor="middle", font_size=4.0, fill=eng_color,
                         font_weight="bold"))
    else:
        dwg.add(dwg.text(title_text, insert=(f"{width_mm / 2}mm", f"{title_y}mm"),
                         text_anchor="middle", font_size="4.0mm", fill=eng_color,
                         font_weight="bold"))

    # Circles and labels
    for dia, cx, r in positions:
        label_text = f"{dia:.0f}"
        label_y = title_height + label_height * 0.5
        if compat:
            dwg.add(dwg.text(label_text, insert=(cx, label_y),
                             text_anchor="middle", font_size=3.5, fill=eng_color))
            dwg.add(dwg.circle(center=(cx, cy), r=r, stroke=cut_color,
                               fill='none', stroke_width=stroke_w))
        else:
            dwg.add(dwg.text(label_text, insert=(f"{cx}mm", f"{label_y}mm"),
                             text_anchor="middle", font_size="3.5mm", fill=eng_color))
            dwg.add(dwg.circle(center=(f"{cx}mm", f"{cy}mm"), r=f"{r}mm",
                               stroke=cut_color, fill='none', stroke_width=stroke_w))

    dwg.save()


# ==========================================
# FEEDS & SPEEDS TESTER
# ==========================================

FEEDS_SPEEDS_SPACING_MM = 3.0      # edge-to-edge spacing between discs within a block
FEEDS_SPEEDS_SHEET_MARGIN_MM = 3.0  # margin from sheet edge to outermost discs

# When a test piece has a center hole (washer / shim), its ID can't sit at the
# center anymore — that's cut away. It moves into the lower ring. These bound
# the auto-fit so the number stays on solid annular material and stays legible.
FEEDS_SPEEDS_LABEL_MIN_FONT_MM = 2.0      # don't shrink the ID below this
FEEDS_SPEEDS_LABEL_RING_CLEARANCE_MM = 1.0  # total radial clearance (0.5mm each side)


def feeds_speeds_label_geometry(diameter_mm, inner_diameter_mm=0.0):
    """Return (label_dy_mm, font_size_mm) for a Speed & Power test disc's ID.

    label_dy_mm is how far the ID label is shifted toward the BOTTOM of the
    disc from its center, in mm:
      - Solid disc (inner_diameter_mm <= 0): 0.0 — the label stays centered and
        the font matches the historical formula exactly (no behavior change).
      - Washer (inner_diameter_mm > 0): the label moves into the lower ring,
        centered on the mid-annulus radius, with the font auto-fit to the ring
        width (floored at FEEDS_SPEEDS_LABEL_MIN_FONT_MM so it stays readable).

    gcode_engine (engraving strokes) and the preview dialog both call this so
    the drawn label and the cut label land in the same place.
    """
    solid_font = max(min(diameter_mm * 0.3, 5.0), 2.5)
    if not inner_diameter_mm or inner_diameter_mm <= 0:
        return 0.0, solid_font
    outer_r = diameter_mm / 2.0
    inner_r = inner_diameter_mm / 2.0
    annulus = outer_r - inner_r
    label_dy = (inner_r + outer_r) / 2.0  # mid-ring radius (toward the bottom)
    fit_font = annulus - FEEDS_SPEEDS_LABEL_RING_CLEARANCE_MM
    font = min(min(diameter_mm * 0.3, 5.0), fit_font)
    font = max(font, FEEDS_SPEEDS_LABEL_MIN_FONT_MM)
    return label_dy, font


def feeds_speeds_ring_fits_label(diameter_mm, inner_diameter_mm):
    """True if a holed disc's ring is wide enough to hold a legible ID label.

    Used by the UI to warn (not block) when the hole is so large the ID would
    have to overflow the ring. Mirrors feeds_speeds_label_geometry's fit math.
    """
    if not inner_diameter_mm or inner_diameter_mm <= 0:
        return True
    annulus = (diameter_mm - inner_diameter_mm) / 2.0
    return annulus - FEEDS_SPEEDS_LABEL_RING_CLEARANCE_MM >= FEEDS_SPEEDS_LABEL_MIN_FONT_MM


def _grid_pack_discs(diameter_mm, cols, rows, num_blocks, sheet_w_mm, sheet_h_mm,
                     spacing_mm=FEEDS_SPEEDS_SPACING_MM,
                     margin_mm=FEEDS_SPEEDS_SHEET_MARGIN_MM,
                     block_gap_mm=None):
    """
    Pack equal-diameter discs into a grid of blocks on a sheet.

    Layout: num_blocks side-by-side horizontally. Each block is a cols x rows
    grid with `spacing_mm` between disc edges. Blocks are separated by
    `block_gap_mm` (defaults to ~2.5 x spacing).

    Returns (positions, total_w_mm, total_h_mm) where positions is a list of
    (cx, cy) tuples in placement order: for block b in [0..num_blocks),
    for row r in [0..rows), for col c in [0..cols), append the disc center.
    The caller's matrix-expansion order must match this iteration order so
    the i-th matrix triple lands on the i-th position.

    Raises ValueError if the matrix doesn't fit, with the minimum required
    sheet size in the message.
    """
    cols = max(1, int(cols))
    rows = max(1, int(rows))
    num_blocks = max(1, int(num_blocks))
    if block_gap_mm is None:
        block_gap_mm = max(spacing_mm * 2.5, diameter_mm * 0.4)

    pitch = diameter_mm + spacing_mm
    block_w = cols * diameter_mm + max(cols - 1, 0) * spacing_mm
    total_w = (num_blocks * block_w
               + max(num_blocks - 1, 0) * block_gap_mm
               + 2 * margin_mm)
    total_h = rows * diameter_mm + max(rows - 1, 0) * spacing_mm + 2 * margin_mm

    if total_w > sheet_w_mm + 1e-6 or total_h > sheet_h_mm + 1e-6:
        raise ValueError(
            f"Test matrix doesn't fit on sheet. Need at least "
            f"{total_w:.1f} × {total_h:.1f} mm; got "
            f"{sheet_w_mm:.1f} × {sheet_h_mm:.1f} mm."
        )

    positions = []
    for block_idx in range(num_blocks):
        block_x0 = margin_mm + block_idx * (block_w + block_gap_mm)
        for row_idx in range(rows):
            cy = margin_mm + diameter_mm / 2 + row_idx * pitch
            for col_idx in range(cols):
                cx = block_x0 + diameter_mm / 2 + col_idx * pitch
                positions.append((cx, cy))
    return positions, total_w, total_h


def _min_feeds_speeds_sheet(diameter_mm, cols, rows, num_blocks,
                             spacing_mm=FEEDS_SPEEDS_SPACING_MM,
                             margin_mm=FEEDS_SPEEDS_SHEET_MARGIN_MM,
                             block_gap_mm=None):
    """Return (min_w_mm, min_h_mm) needed to fit the given matrix."""
    cols = max(1, int(cols))
    rows = max(1, int(rows))
    num_blocks = max(1, int(num_blocks))
    if block_gap_mm is None:
        block_gap_mm = max(spacing_mm * 2.5, diameter_mm * 0.4)
    block_w = cols * diameter_mm + max(cols - 1, 0) * spacing_mm
    min_w = (num_blocks * block_w
             + max(num_blocks - 1, 0) * block_gap_mm
             + 2 * margin_mm)
    min_h = rows * diameter_mm + max(rows - 1, 0) * spacing_mm + 2 * margin_mm
    return min_w, min_h


def feeds_speeds_linspace(start, end, stops):
    """Evenly-spaced integer values from start to end inclusive over `stops` points.

    Stops <= 1 returns a single-element list with `start` rounded.
    Start == end with stops > 1 returns a list of duplicates (caller is
    expected to decide if that's meaningful).
    """
    stops = max(1, int(stops))
    if stops == 1:
        return [int(round(start))]
    step = (end - start) / (stops - 1)
    return [int(round(start + i * step)) for i in range(stops)]


def build_feeds_speeds_matrix(speed_cfg, power_cfg, passes_cfg):
    """
    Expand three sweep configs into ordered (speed, power, passes) triples
    and a (cols, rows, num_blocks) grid shape.

    Each config is a dict with keys:
      'sweep' (bool), 'value' (number, used when not swept),
      'start', 'end' (numbers, used when swept), 'stops' (int).

    Iteration order (outermost → innermost) is passes, power, speed,
    matching _grid_pack_discs' (block, row, col) iteration so the i-th
    triple lands on the i-th packed position.

    Clamps: passes >= 1, power in [1, 100], speed >= 1, stops in [2, 10].
    Raises ValueError if any required number is missing or non-numeric.
    """
    def values_for(name, cfg, *, kind):
        try:
            if cfg.get('sweep'):
                start = float(cfg['start'])
                end = float(cfg['end'])
                stops = int(cfg.get('stops', 4))
                stops = max(2, min(10, stops))
                return feeds_speeds_linspace(start, end, stops)
            else:
                return [int(round(float(cfg['value'])))]
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"{name}: {exc}")

    speeds = values_for("Speed", speed_cfg, kind='speed')
    powers = values_for("Power", power_cfg, kind='power')
    passes_list = values_for("Passes", passes_cfg, kind='passes')

    speeds = [max(1, s) for s in speeds]
    powers = [max(1, min(100, p)) for p in powers]
    passes_list = [max(1, p) for p in passes_list]

    triples = []
    for p_count in passes_list:
        for power in powers:
            for speed in speeds:
                triples.append((speed, power, p_count))

    return triples, len(speeds), len(powers), len(passes_list)


