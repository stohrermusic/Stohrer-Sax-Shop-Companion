"""
One-time utility to extract glyph outlines from a TTF font for use as
FILLED_FONT data in gcode_engine.py.

Uses fonttools (MIT license) to extract contours from Roboto (Apache 2.0).
Output is a Python dict literal ready to paste into gcode_engine.py.

Usage:
    python tools/extract_font_outlines.py
"""

from fontTools.ttLib import TTFont
from fontTools.pens.recordingPen import RecordingPen
from fontTools.pens.pointPen import SegmentToPointPen
import math
import os

FONT_PATH = os.path.join(os.path.dirname(__file__), "Roboto-Regular.ttf")
CHARS = "0123456789.-"

# Target coordinate system: origin at bottom-left, ~0.6 wide, 1.0 tall
# (matching existing STROKE_FONT convention)
TARGET_HEIGHT = 1.0


def flatten_curve(points, segments_per_curve=8):
    """Flatten a cubic or quadratic bezier curve to line segments."""
    if len(points) == 3:
        # Quadratic bezier: start, control, end
        p0, p1, p2 = points
        result = []
        for i in range(1, segments_per_curve + 1):
            t = i / segments_per_curve
            x = (1-t)**2 * p0[0] + 2*(1-t)*t * p1[0] + t**2 * p2[0]
            y = (1-t)**2 * p0[1] + 2*(1-t)*t * p1[1] + t**2 * p2[1]
            result.append((x, y))
        return result
    elif len(points) == 4:
        # Cubic bezier: start, control1, control2, end
        p0, p1, p2, p3 = points
        result = []
        for i in range(1, segments_per_curve + 1):
            t = i / segments_per_curve
            x = (1-t)**3*p0[0] + 3*(1-t)**2*t*p1[0] + 3*(1-t)*t**2*p2[0] + t**3*p3[0]
            y = (1-t)**3*p0[1] + 3*(1-t)**2*t*p1[1] + 3*(1-t)*t**2*p2[1] + t**3*p3[1]
            result.append((x, y))
        return result
    else:
        return points[1:]  # Fallback: just use the points


def extract_contours(font, glyph_name):
    """Extract flattened contours from a glyph."""
    gs = font.getGlyphSet()
    pen = RecordingPen()
    gs[glyph_name].draw(pen)

    contours = []
    current_contour = []
    current_pos = (0, 0)

    for op, args in pen.value:
        if op == 'moveTo':
            if current_contour:
                contours.append(current_contour)
            current_contour = [args[0]]
            current_pos = args[0]
        elif op == 'lineTo':
            current_contour.append(args[0])
            current_pos = args[0]
        elif op == 'qCurveTo':
            # Quadratic bezier - may have multiple control points (TrueType)
            # Process as sequence of quadratic segments
            points = list(args)
            if len(points) == 2:
                # Simple quadratic: 1 control + 1 end
                curve_pts = flatten_curve([current_pos, points[0], points[1]])
                current_contour.extend(curve_pts)
                current_pos = points[-1]
            else:
                # Multiple control points - implied on-curve points between them
                for i in range(len(points) - 1):
                    if i < len(points) - 2:
                        # Implied on-curve point is midpoint between consecutive controls
                        mid_x = (points[i][0] + points[i+1][0]) / 2
                        mid_y = (points[i][1] + points[i+1][1]) / 2
                        end = (mid_x, mid_y)
                    else:
                        end = points[i + 1]
                    curve_pts = flatten_curve([current_pos, points[i], end])
                    current_contour.extend(curve_pts)
                    current_pos = end
        elif op == 'curveTo':
            # Cubic bezier
            curve_pts = flatten_curve([current_pos] + list(args))
            current_contour.extend(curve_pts)
            current_pos = args[-1]
        elif op == 'closePath' or op == 'endPath':
            if current_contour:
                # Close the contour
                if current_contour[0] != current_contour[-1]:
                    current_contour.append(current_contour[0])
                contours.append(current_contour)
                current_contour = []

    if current_contour:
        contours.append(current_contour)

    return contours


def normalize_contours(contours, scale, ref_min_y, margin=0.05):
    """Normalize contours using a shared scale and baseline.

    Args:
        contours: List of contour point lists
        scale: Shared scale factor (from digit reference height)
        ref_min_y: Shared baseline Y coordinate (from digit reference)
        margin: Left margin in normalized coords
    """
    if not contours:
        return [], 0.0

    # Get bounding box (for width and x offset)
    all_points = [p for c in contours for p in c]
    min_x = min(p[0] for p in all_points)
    max_x = max(p[0] for p in all_points)

    normalized = []
    for contour in contours:
        norm_contour = []
        for x, y in contour:
            nx = (x - min_x) * scale + margin
            ny = (y - ref_min_y) * scale
            norm_contour.append((round(nx, 3), round(ny, 3)))
        normalized.append(norm_contour)

    # Calculate width
    width = (max_x - min_x) * scale + margin * 2

    return normalized, round(width, 2)


def main():
    font = TTFont(FONT_PATH)
    cmap = font.getBestCmap()

    # First pass: extract all contours and find digit reference height
    all_contours = {}
    digit_chars = "0123456789"

    for char in CHARS:
        code = ord(char)
        glyph_name = cmap.get(code)
        if not glyph_name:
            print(f"    # WARNING: No glyph for {repr(char)}", file=__import__('sys').stderr)
            continue
        all_contours[char] = extract_contours(font, glyph_name)

    # Find reference height from digits (max height across all digit glyphs)
    ref_min_y = float('inf')
    ref_max_y = float('-inf')
    for char in digit_chars:
        if char not in all_contours:
            continue
        for contour in all_contours[char]:
            for x, y in contour:
                ref_min_y = min(ref_min_y, y)
                ref_max_y = max(ref_max_y, y)

    ref_height = ref_max_y - ref_min_y
    if ref_height == 0:
        ref_height = 1
    scale = TARGET_HEIGHT / ref_height

    # Second pass: normalize all characters using shared scale
    print("# Extracted from Roboto-Regular.ttf (Apache 2.0 License)")
    print("# Generated by tools/extract_font_outlines.py")
    print()
    print("FILLED_FONT = {")

    widths = {}

    for char in CHARS:
        if char not in all_contours:
            continue

        normalized, width = normalize_contours(all_contours[char], scale, ref_min_y)
        widths[char] = width

        print(f"    {repr(char)}: [")
        for contour in normalized:
            points_str = ", ".join(f"({x}, {y})" for x, y in contour)
            print(f"        [{points_str}],")
        print(f"    ],")

    print("}")
    print()
    print("FILLED_CHAR_WIDTHS = {")
    for char in CHARS:
        if char in widths:
            print(f"    {repr(char)}: {widths[char]},")
    print("}")

    font.close()


if __name__ == "__main__":
    main()
