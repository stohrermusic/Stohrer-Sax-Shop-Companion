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


# =============================================================================
# FILLED FONT (outline contours for raster fill engraving)
# =============================================================================
# Extracted from Roboto-Regular.ttf (Apache 2.0 License)
# Each character is a list of closed contours (outer boundary + holes).
# The even-odd fill rule handles holes naturally during scan-line fill.
# Coordinate system matches STROKE_FONT: origin at bottom-left, ~0.6-0.8 wide, 1.0 tall.

FILLED_FONT = {
    '0': [
        [(0.664, 0.428), (0.663, 0.376), (0.66, 0.327), (0.654, 0.281), (0.646, 0.239), (0.635, 0.201), (0.623, 0.165), (0.607, 0.134), (0.59, 0.106), (0.57, 0.081), (0.548, 0.059), (0.523, 0.041), (0.495, 0.026), (0.465, 0.015), (0.432, 0.007), (0.396, 0.002), (0.358, 0.0), (0.32, 0.002), (0.285, 0.006), (0.253, 0.015), (0.223, 0.026), (0.195, 0.04), (0.17, 0.058), (0.148, 0.079), (0.128, 0.103), (0.11, 0.131), (0.095, 0.161), (0.082, 0.195), (0.071, 0.232), (0.062, 0.272), (0.056, 0.316), (0.052, 0.362), (0.05, 0.412), (0.05, 0.577), (0.051, 0.629), (0.055, 0.677), (0.06, 0.722), (0.069, 0.764), (0.079, 0.802), (0.092, 0.837), (0.107, 0.868), (0.124, 0.896), (0.144, 0.92), (0.167, 0.941), (0.192, 0.959), (0.219, 0.974), (0.25, 0.985), (0.283, 0.993), (0.319, 0.998), (0.357, 1.0), (0.395, 0.998), (0.43, 0.994), (0.463, 0.986), (0.493, 0.975), (0.521, 0.961), (0.546, 0.943), (0.568, 0.923), (0.588, 0.899), (0.606, 0.873), (0.621, 0.842), (0.634, 0.809), (0.644, 0.771), (0.653, 0.731), (0.659, 0.687), (0.663, 0.64), (0.664, 0.589), (0.664, 0.428)],
        [(0.541, 0.598), (0.54, 0.636), (0.538, 0.671), (0.534, 0.703), (0.53, 0.733), (0.523, 0.761), (0.516, 0.785), (0.507, 0.807), (0.497, 0.827), (0.485, 0.843), (0.471, 0.858), (0.456, 0.87), (0.44, 0.88), (0.421, 0.888), (0.401, 0.894), (0.38, 0.897), (0.357, 0.898), (0.334, 0.897), (0.313, 0.894), (0.293, 0.888), (0.275, 0.881), (0.258, 0.87), (0.243, 0.858), (0.23, 0.844), (0.218, 0.827), (0.208, 0.808), (0.199, 0.786), (0.192, 0.762), (0.186, 0.736), (0.181, 0.707), (0.177, 0.676), (0.175, 0.643), (0.174, 0.607), (0.174, 0.409), (0.174, 0.371), (0.177, 0.335), (0.18, 0.302), (0.185, 0.272), (0.192, 0.244), (0.199, 0.219), (0.209, 0.196), (0.219, 0.176), (0.232, 0.159), (0.245, 0.143), (0.26, 0.13), (0.277, 0.12), (0.295, 0.112), (0.315, 0.106), (0.336, 0.102), (0.358, 0.101), (0.38, 0.102), (0.401, 0.105), (0.42, 0.111), (0.438, 0.119), (0.454, 0.129), (0.469, 0.141), (0.482, 0.155), (0.494, 0.172), (0.505, 0.191), (0.514, 0.212), (0.522, 0.236), (0.528, 0.263), (0.533, 0.292), (0.537, 0.324), (0.539, 0.358), (0.541, 0.395), (0.541, 0.598)],
    ],
    '1': [
        [(0.424, 0.013), (0.299, 0.013), (0.299, 0.838), (0.05, 0.746), (0.05, 0.858), (0.404, 0.991), (0.424, 0.991), (0.424, 0.013)],
    ],
    '2': [
        [(0.706, 0.013), (0.069, 0.013), (0.069, 0.102), (0.406, 0.477), (0.424, 0.497), (0.44, 0.517), (0.455, 0.536), (0.469, 0.554), (0.481, 0.57), (0.492, 0.586), (0.501, 0.601), (0.509, 0.615), (0.516, 0.628), (0.521, 0.641), (0.526, 0.655), (0.53, 0.669), (0.533, 0.683), (0.536, 0.696), (0.537, 0.71), (0.537, 0.725), (0.537, 0.743), (0.534, 0.761), (0.531, 0.778), (0.526, 0.794), (0.519, 0.809), (0.511, 0.823), (0.502, 0.837), (0.491, 0.85), (0.479, 0.861), (0.466, 0.871), (0.452, 0.879), (0.437, 0.886), (0.422, 0.892), (0.405, 0.895), (0.387, 0.898), (0.368, 0.898), (0.346, 0.898), (0.325, 0.895), (0.305, 0.891), (0.286, 0.885), (0.269, 0.878), (0.253, 0.869), (0.238, 0.858), (0.225, 0.846), (0.213, 0.832), (0.202, 0.817), (0.194, 0.801), (0.186, 0.783), (0.181, 0.764), (0.177, 0.744), (0.174, 0.723), (0.174, 0.7), (0.05, 0.7), (0.051, 0.733), (0.055, 0.764), (0.062, 0.793), (0.072, 0.821), (0.084, 0.848), (0.099, 0.872), (0.116, 0.896), (0.137, 0.917), (0.159, 0.937), (0.184, 0.953), (0.21, 0.968), (0.238, 0.979), (0.268, 0.988), (0.299, 0.995), (0.333, 0.999), (0.368, 1.0), (0.401, 0.999), (0.432, 0.996), (0.462, 0.99), (0.49, 0.982), (0.516, 0.972), (0.54, 0.96), (0.562, 0.945), (0.583, 0.929), (0.601, 0.91), (0.617, 0.89), (0.631, 0.869), (0.642, 0.846), (0.651, 0.821), (0.657, 0.796), (0.66, 0.768), (0.662, 0.739), (0.659, 0.703), (0.65, 0.664), (0.636, 0.624), (0.616, 0.582), (0.59, 0.538), (0.559, 0.493), (0.521, 0.446), (0.478, 0.397), (0.218, 0.114), (0.706, 0.114), (0.706, 0.013)],
    ],
    '3': [
        [(0.248, 0.56), (0.341, 0.56), (0.362, 0.561), (0.382, 0.564), (0.401, 0.567), (0.419, 0.572), (0.436, 0.579), (0.451, 0.587), (0.465, 0.596), (0.478, 0.606), (0.49, 0.618), (0.5, 0.631), (0.509, 0.644), (0.516, 0.659), (0.522, 0.675), (0.525, 0.691), (0.528, 0.709), (0.529, 0.727), (0.526, 0.767), (0.518, 0.802), (0.505, 0.832), (0.486, 0.856), (0.462, 0.874), (0.433, 0.888), (0.398, 0.896), (0.358, 0.898), (0.339, 0.898), (0.32, 0.896), (0.303, 0.892), (0.286, 0.887), (0.271, 0.881), (0.256, 0.873), (0.243, 0.863), (0.23, 0.853), (0.219, 0.841), (0.209, 0.828), (0.201, 0.814), (0.194, 0.799), (0.189, 0.784), (0.185, 0.767), (0.183, 0.75), (0.182, 0.731), (0.059, 0.731), (0.06, 0.76), (0.064, 0.787), (0.071, 0.813), (0.08, 0.837), (0.092, 0.861), (0.106, 0.883), (0.123, 0.904), (0.143, 0.923), (0.165, 0.941), (0.188, 0.957), (0.213, 0.97), (0.239, 0.981), (0.267, 0.989), (0.296, 0.995), (0.326, 0.999), (0.358, 1.0), (0.392, 0.999), (0.423, 0.995), (0.453, 0.99), (0.481, 0.982), (0.507, 0.972), (0.531, 0.959), (0.554, 0.944), (0.574, 0.927), (0.592, 0.908), (0.608, 0.887), (0.622, 0.865), (0.633, 0.84), (0.641, 0.814), (0.647, 0.786), (0.651, 0.756), (0.652, 0.725), (0.652, 0.709), (0.65, 0.693), (0.646, 0.678), (0.642, 0.662), (0.636, 0.647), (0.629, 0.632), (0.621, 0.617), (0.611, 0.602), (0.6, 0.587), (0.589, 0.574), (0.576, 0.561), (0.563, 0.55), (0.548, 0.539), (0.533, 0.529), (0.516, 0.521), (0.499, 0.513), (0.519, 0.506), (0.537, 0.498), (0.554, 0.489), (0.57, 0.479), (0.585, 0.468), (0.599, 0.456), (0.612, 0.443), (0.623, 0.428), (0.633, 0.413), (0.642, 0.397), (0.65, 0.381), (0.656, 0.363), (0.661, 0.345), (0.664, 0.326), (0.666, 0.306), (0.667, 0.285), (0.666, 0.253), (0.662, 0.223), (0.655, 0.194), (0.646, 0.167), (0.634, 0.142), (0.619, 0.119), (0.601, 0.097), (0.581, 0.077), (0.559, 0.059), (0.535, 0.043), (0.51, 0.03), (0.483, 0.019), (0.454, 0.011), (0.424, 0.005), (0.392, 0.001), (0.359, 0.0), (0.325, 0.001), (0.294, 0.005), (0.263, 0.01), (0.235, 0.019), (0.208, 0.029), (0.182, 0.042), (0.158, 0.057), (0.136, 0.074), (0.116, 0.093), (0.098, 0.114), (0.084, 0.137), (0.071, 0.16), (0.062, 0.185), (0.055, 0.212), (0.051, 0.24), (0.05, 0.27), (0.174, 0.27), (0.175, 0.251), (0.177, 0.234), (0.181, 0.217), (0.187, 0.201), (0.194, 0.186), (0.203, 0.172), (0.213, 0.159), (0.224, 0.147), (0.238, 0.136), (0.252, 0.127), (0.267, 0.119), (0.283, 0.112), (0.3, 0.107), (0.319, 0.104), (0.338, 0.102), (0.359, 0.101), (0.381, 0.102), (0.401, 0.104), (0.42, 0.108), (0.438, 0.113), (0.454, 0.119), (0.469, 0.127), (0.483, 0.137), (0.496, 0.148), (0.507, 0.16), (0.517, 0.174), (0.525, 0.189), (0.531, 0.205), (0.537, 0.222), (0.54, 0.241), (0.543, 0.261), (0.543, 0.282), (0.543, 0.303), (0.54, 0.322), (0.536, 0.34), (0.53, 0.357), (0.523, 0.373), (0.514, 0.387), (0.503, 0.4), (0.491, 0.412), (0.477, 0.423), (0.462, 0.432), (0.446, 0.44), (0.428, 0.447), (0.408, 0.452), (0.387, 0.456), (0.365, 0.458), (0.341, 0.459), (0.248, 0.459), (0.248, 0.56)],
    ],
    '4': [
        [(0.618, 0.34), (0.753, 0.34), (0.753, 0.239), (0.618, 0.239), (0.618, 0.013), (0.493, 0.013), (0.493, 0.239), (0.05, 0.239), (0.05, 0.312), (0.486, 0.987), (0.618, 0.987), (0.618, 0.34)],
        [(0.19, 0.34), (0.493, 0.34), (0.493, 0.818), (0.478, 0.791), (0.19, 0.34)],
    ],
    '5': [
        [(0.085, 0.501), (0.134, 0.987), (0.633, 0.987), (0.633, 0.872), (0.239, 0.872), (0.21, 0.607), (0.228, 0.617), (0.247, 0.625), (0.266, 0.633), (0.286, 0.639), (0.307, 0.643), (0.328, 0.646), (0.35, 0.648), (0.372, 0.649), (0.405, 0.648), (0.435, 0.644), (0.464, 0.637), (0.492, 0.627), (0.517, 0.615), (0.541, 0.6), (0.563, 0.582), (0.583, 0.561), (0.602, 0.538), (0.618, 0.513), (0.631, 0.487), (0.642, 0.458), (0.651, 0.427), (0.657, 0.395), (0.66, 0.36), (0.662, 0.324), (0.66, 0.287), (0.657, 0.252), (0.65, 0.22), (0.641, 0.189), (0.63, 0.16), (0.616, 0.134), (0.599, 0.109), (0.58, 0.087), (0.559, 0.066), (0.536, 0.049), (0.51, 0.034), (0.483, 0.022), (0.454, 0.012), (0.422, 0.005), (0.389, 0.001), (0.353, 0.0), (0.322, 0.001), (0.292, 0.004), (0.263, 0.01), (0.236, 0.018), (0.211, 0.028), (0.187, 0.04), (0.164, 0.055), (0.143, 0.072), (0.124, 0.09), (0.107, 0.111), (0.092, 0.133), (0.079, 0.157), (0.068, 0.182), (0.06, 0.21), (0.054, 0.239), (0.05, 0.269), (0.167, 0.269), (0.17, 0.249), (0.175, 0.23), (0.181, 0.212), (0.188, 0.196), (0.196, 0.181), (0.205, 0.167), (0.215, 0.155), (0.226, 0.143), (0.239, 0.133), (0.252, 0.125), (0.267, 0.118), (0.282, 0.112), (0.299, 0.107), (0.316, 0.104), (0.334, 0.102), (0.353, 0.101), (0.374, 0.102), (0.394, 0.105), (0.413, 0.109), (0.43, 0.116), (0.447, 0.124), (0.462, 0.134), (0.476, 0.146), (0.489, 0.16), (0.5, 0.175), (0.51, 0.192), (0.519, 0.21), (0.526, 0.23), (0.531, 0.251), (0.535, 0.273), (0.537, 0.297), (0.538, 0.322), (0.537, 0.346), (0.535, 0.369), (0.53, 0.39), (0.525, 0.41), (0.517, 0.429), (0.508, 0.447), (0.497, 0.464), (0.485, 0.479), (0.471, 0.493), (0.456, 0.505), (0.44, 0.515), (0.423, 0.523), (0.405, 0.53), (0.385, 0.534), (0.365, 0.537), (0.343, 0.538), (0.324, 0.538), (0.305, 0.536), (0.288, 0.533), (0.271, 0.529), (0.256, 0.524), (0.242, 0.518), (0.229, 0.511), (0.216, 0.503), (0.184, 0.476), (0.085, 0.501)],
    ],
    '6': [
        [(0.528, 0.987), (0.528, 0.882), (0.505, 0.882), (0.47, 0.88), (0.437, 0.876), (0.405, 0.869), (0.376, 0.86), (0.348, 0.848), (0.322, 0.833), (0.298, 0.816), (0.275, 0.797), (0.255, 0.775), (0.237, 0.751), (0.221, 0.725), (0.208, 0.697), (0.197, 0.667), (0.188, 0.634), (0.181, 0.6), (0.176, 0.564), (0.196, 0.584), (0.218, 0.602), (0.242, 0.617), (0.267, 0.63), (0.294, 0.639), (0.323, 0.646), (0.354, 0.65), (0.386, 0.652), (0.417, 0.65), (0.447, 0.646), (0.474, 0.639), (0.5, 0.629), (0.525, 0.617), (0.548, 0.601), (0.569, 0.583), (0.589, 0.562), (0.607, 0.539), (0.622, 0.514), (0.635, 0.488), (0.646, 0.46), (0.654, 0.43), (0.66, 0.398), (0.664, 0.365), (0.665, 0.331), (0.664, 0.294), (0.66, 0.259), (0.653, 0.227), (0.645, 0.196), (0.633, 0.166), (0.619, 0.139), (0.602, 0.114), (0.583, 0.09), (0.562, 0.069), (0.539, 0.051), (0.514, 0.035), (0.487, 0.023), (0.459, 0.013), (0.429, 0.006), (0.397, 0.001), (0.364, 0.0), (0.329, 0.002), (0.297, 0.007), (0.266, 0.015), (0.237, 0.027), (0.209, 0.042), (0.184, 0.06), (0.159, 0.082), (0.137, 0.107), (0.117, 0.135), (0.099, 0.165), (0.084, 0.196), (0.072, 0.23), (0.062, 0.266), (0.055, 0.303), (0.051, 0.342), (0.05, 0.384), (0.05, 0.431), (0.052, 0.496), (0.057, 0.558), (0.066, 0.615), (0.079, 0.668), (0.095, 0.718), (0.114, 0.763), (0.138, 0.804), (0.165, 0.842), (0.195, 0.875), (0.229, 0.904), (0.266, 0.929), (0.307, 0.949), (0.352, 0.965), (0.4, 0.977), (0.451, 0.984), (0.506, 0.987), (0.528, 0.987)],
        [(0.366, 0.549), (0.35, 0.548), (0.334, 0.546), (0.319, 0.543), (0.305, 0.539), (0.29, 0.534), (0.276, 0.527), (0.262, 0.52), (0.249, 0.511), (0.236, 0.501), (0.224, 0.49), (0.213, 0.479), (0.203, 0.468), (0.194, 0.455), (0.187, 0.443), (0.18, 0.429), (0.174, 0.415), (0.174, 0.37), (0.175, 0.34), (0.178, 0.312), (0.182, 0.286), (0.188, 0.261), (0.195, 0.238), (0.205, 0.215), (0.216, 0.195), (0.228, 0.176), (0.242, 0.159), (0.257, 0.144), (0.273, 0.131), (0.289, 0.121), (0.307, 0.113), (0.325, 0.107), (0.344, 0.103), (0.364, 0.102), (0.384, 0.103), (0.403, 0.106), (0.421, 0.111), (0.438, 0.118), (0.454, 0.126), (0.469, 0.137), (0.482, 0.149), (0.495, 0.164), (0.506, 0.18), (0.516, 0.197), (0.524, 0.215), (0.531, 0.235), (0.536, 0.256), (0.54, 0.277), (0.542, 0.301), (0.543, 0.325), (0.542, 0.349), (0.54, 0.373), (0.536, 0.395), (0.531, 0.416), (0.524, 0.435), (0.515, 0.454), (0.506, 0.471), (0.494, 0.487), (0.482, 0.501), (0.468, 0.514), (0.453, 0.525), (0.438, 0.533), (0.421, 0.54), (0.404, 0.545), (0.385, 0.548), (0.366, 0.549)],
    ],
    '7': [
        [(0.708, 0.917), (0.305, 0.013), (0.175, 0.013), (0.577, 0.885), (0.05, 0.885), (0.05, 0.987), (0.708, 0.987), (0.708, 0.917)],
    ],
    '8': [
        [(0.646, 0.733), (0.646, 0.715), (0.644, 0.697), (0.641, 0.68), (0.637, 0.664), (0.631, 0.648), (0.625, 0.632), (0.617, 0.617), (0.608, 0.603), (0.598, 0.589), (0.587, 0.576), (0.575, 0.564), (0.563, 0.552), (0.549, 0.542), (0.535, 0.532), (0.52, 0.522), (0.504, 0.514), (0.522, 0.505), (0.54, 0.496), (0.557, 0.485), (0.572, 0.474), (0.587, 0.461), (0.6, 0.448), (0.613, 0.434), (0.625, 0.418), (0.635, 0.402), (0.644, 0.386), (0.652, 0.369), (0.658, 0.351), (0.663, 0.333), (0.666, 0.315), (0.668, 0.296), (0.669, 0.276), (0.668, 0.245), (0.664, 0.216), (0.657, 0.189), (0.648, 0.163), (0.636, 0.138), (0.621, 0.116), (0.604, 0.094), (0.584, 0.075), (0.562, 0.057), (0.538, 0.042), (0.512, 0.029), (0.485, 0.019), (0.456, 0.011), (0.426, 0.005), (0.393, 0.001), (0.359, 0.0), (0.325, 0.001), (0.293, 0.005), (0.262, 0.011), (0.233, 0.019), (0.206, 0.029), (0.18, 0.042), (0.157, 0.058), (0.135, 0.075), (0.115, 0.095), (0.098, 0.116), (0.083, 0.139), (0.071, 0.163), (0.062, 0.189), (0.055, 0.216), (0.051, 0.245), (0.05, 0.276), (0.051, 0.296), (0.053, 0.315), (0.056, 0.333), (0.061, 0.351), (0.067, 0.369), (0.074, 0.386), (0.083, 0.402), (0.093, 0.418), (0.104, 0.434), (0.117, 0.448), (0.13, 0.462), (0.145, 0.474), (0.16, 0.486), (0.176, 0.496), (0.194, 0.506), (0.212, 0.515), (0.197, 0.523), (0.182, 0.532), (0.168, 0.542), (0.154, 0.553), (0.142, 0.564), (0.131, 0.577), (0.12, 0.59), (0.11, 0.604), (0.101, 0.618), (0.094, 0.633), (0.087, 0.648), (0.082, 0.664), (0.078, 0.681), (0.075, 0.697), (0.073, 0.715), (0.073, 0.733), (0.074, 0.763), (0.078, 0.791), (0.084, 0.818), (0.092, 0.843), (0.104, 0.867), (0.117, 0.889), (0.133, 0.909), (0.152, 0.928), (0.172, 0.945), (0.194, 0.959), (0.218, 0.972), (0.243, 0.982), (0.27, 0.99), (0.298, 0.995), (0.328, 0.999), (0.359, 1.0), (0.391, 0.999), (0.421, 0.995), (0.449, 0.99), (0.476, 0.982), (0.501, 0.972), (0.524, 0.959), (0.546, 0.945), (0.567, 0.928), (0.586, 0.909), (0.602, 0.889), (0.615, 0.867), (0.626, 0.843), (0.635, 0.818), (0.641, 0.791), (0.645, 0.763), (0.646, 0.733)],
        [(0.545, 0.279), (0.545, 0.299), (0.542, 0.318), (0.538, 0.336), (0.532, 0.353), (0.525, 0.369), (0.516, 0.384), (0.506, 0.398), (0.494, 0.412), (0.48, 0.424), (0.466, 0.434), (0.45, 0.443), (0.434, 0.45), (0.416, 0.456), (0.398, 0.46), (0.379, 0.462), (0.358, 0.463), (0.338, 0.462), (0.318, 0.46), (0.3, 0.456), (0.283, 0.451), (0.267, 0.443), (0.251, 0.435), (0.237, 0.424), (0.224, 0.412), (0.212, 0.399), (0.202, 0.385), (0.193, 0.37), (0.186, 0.354), (0.181, 0.336), (0.177, 0.318), (0.174, 0.299), (0.174, 0.279), (0.174, 0.259), (0.177, 0.24), (0.181, 0.222), (0.186, 0.205), (0.193, 0.189), (0.201, 0.174), (0.211, 0.161), (0.223, 0.148), (0.236, 0.137), (0.25, 0.128), (0.265, 0.119), (0.282, 0.113), (0.299, 0.108), (0.318, 0.104), (0.338, 0.102), (0.359, 0.101), (0.381, 0.102), (0.401, 0.104), (0.419, 0.108), (0.437, 0.113), (0.453, 0.12), (0.469, 0.128), (0.483, 0.138), (0.496, 0.149), (0.507, 0.161), (0.517, 0.175), (0.526, 0.189), (0.533, 0.205), (0.538, 0.222), (0.542, 0.24), (0.545, 0.259), (0.545, 0.279)],
        [(0.359, 0.898), (0.342, 0.898), (0.325, 0.896), (0.309, 0.892), (0.294, 0.887), (0.279, 0.881), (0.266, 0.873), (0.253, 0.864), (0.242, 0.853), (0.231, 0.841), (0.222, 0.829), (0.214, 0.815), (0.208, 0.8), (0.203, 0.784), (0.199, 0.767), (0.197, 0.749), (0.196, 0.731), (0.197, 0.713), (0.199, 0.695), (0.203, 0.679), (0.208, 0.664), (0.214, 0.649), (0.221, 0.635), (0.23, 0.623), (0.241, 0.611), (0.252, 0.6), (0.265, 0.591), (0.278, 0.583), (0.293, 0.576), (0.308, 0.571), (0.324, 0.568), (0.341, 0.566), (0.359, 0.565), (0.378, 0.566), (0.395, 0.568), (0.411, 0.571), (0.426, 0.576), (0.441, 0.583), (0.454, 0.591), (0.467, 0.6), (0.478, 0.611), (0.489, 0.623), (0.498, 0.635), (0.505, 0.649), (0.511, 0.664), (0.516, 0.679), (0.52, 0.695), (0.522, 0.713), (0.523, 0.731), (0.522, 0.749), (0.52, 0.766), (0.516, 0.782), (0.511, 0.798), (0.505, 0.813), (0.497, 0.826), (0.487, 0.839), (0.476, 0.852), (0.465, 0.863), (0.452, 0.872), (0.438, 0.88), (0.424, 0.887), (0.409, 0.892), (0.393, 0.895), (0.377, 0.898), (0.359, 0.898)],
    ],
    '9': [
        [(0.538, 0.441), (0.528, 0.43), (0.518, 0.419), (0.507, 0.409), (0.495, 0.4), (0.484, 0.391), (0.471, 0.382), (0.459, 0.374), (0.445, 0.367), (0.432, 0.36), (0.418, 0.355), (0.404, 0.35), (0.389, 0.346), (0.374, 0.343), (0.359, 0.341), (0.343, 0.339), (0.327, 0.339), (0.307, 0.34), (0.287, 0.341), (0.267, 0.345), (0.249, 0.349), (0.231, 0.355), (0.213, 0.362), (0.197, 0.371), (0.181, 0.38), (0.166, 0.391), (0.151, 0.403), (0.138, 0.416), (0.125, 0.43), (0.114, 0.445), (0.103, 0.462), (0.093, 0.479), (0.084, 0.497), (0.076, 0.516), (0.069, 0.536), (0.063, 0.556), (0.059, 0.576), (0.055, 0.597), (0.052, 0.619), (0.051, 0.641), (0.05, 0.663), (0.051, 0.687), (0.052, 0.711), (0.055, 0.734), (0.059, 0.756), (0.064, 0.778), (0.071, 0.799), (0.078, 0.819), (0.087, 0.839), (0.097, 0.858), (0.108, 0.876), (0.119, 0.892), (0.132, 0.908), (0.146, 0.922), (0.16, 0.936), (0.176, 0.948), (0.192, 0.959), (0.21, 0.968), (0.228, 0.977), (0.247, 0.984), (0.266, 0.99), (0.286, 0.994), (0.307, 0.997), (0.329, 0.999), (0.351, 1.0), (0.387, 0.998), (0.42, 0.993), (0.451, 0.985), (0.481, 0.973), (0.508, 0.958), (0.534, 0.939), (0.557, 0.917), (0.579, 0.892), (0.599, 0.864), (0.615, 0.833), (0.63, 0.8), (0.641, 0.764), (0.651, 0.726), (0.657, 0.686), (0.661, 0.643), (0.662, 0.598), (0.662, 0.561), (0.661, 0.493), (0.655, 0.429), (0.647, 0.37), (0.634, 0.316), (0.618, 0.266), (0.599, 0.222), (0.576, 0.182), (0.55, 0.147), (0.52, 0.116), (0.487, 0.089), (0.45, 0.067), (0.409, 0.048), (0.365, 0.033), (0.317, 0.022), (0.266, 0.016), (0.211, 0.013), (0.187, 0.013), (0.187, 0.117), (0.213, 0.117), (0.25, 0.119), (0.285, 0.123), (0.318, 0.129), (0.348, 0.138), (0.377, 0.149), (0.403, 0.163), (0.427, 0.179), (0.448, 0.197), (0.468, 0.218), (0.485, 0.241), (0.5, 0.267), (0.512, 0.297), (0.522, 0.328), (0.53, 0.363), (0.535, 0.401), (0.538, 0.441)],
        [(0.347, 0.441), (0.363, 0.442), (0.378, 0.444), (0.393, 0.447), (0.407, 0.451), (0.421, 0.456), (0.435, 0.463), (0.449, 0.47), (0.462, 0.479), (0.475, 0.489), (0.487, 0.499), (0.497, 0.51), (0.507, 0.522), (0.517, 0.534), (0.525, 0.547), (0.532, 0.56), (0.539, 0.574), (0.539, 0.623), (0.538, 0.653), (0.535, 0.681), (0.531, 0.708), (0.525, 0.733), (0.518, 0.757), (0.509, 0.78), (0.498, 0.801), (0.486, 0.821), (0.472, 0.839), (0.458, 0.854), (0.442, 0.867), (0.426, 0.878), (0.409, 0.886), (0.391, 0.892), (0.372, 0.896), (0.352, 0.897), (0.332, 0.896), (0.313, 0.893), (0.295, 0.888), (0.279, 0.881), (0.263, 0.873), (0.248, 0.862), (0.234, 0.849), (0.221, 0.835), (0.21, 0.818), (0.199, 0.801), (0.191, 0.782), (0.184, 0.762), (0.179, 0.741), (0.175, 0.718), (0.172, 0.695), (0.172, 0.67), (0.172, 0.645), (0.175, 0.622), (0.178, 0.6), (0.184, 0.579), (0.19, 0.559), (0.199, 0.54), (0.208, 0.522), (0.219, 0.506), (0.232, 0.491), (0.245, 0.477), (0.26, 0.466), (0.275, 0.457), (0.292, 0.45), (0.309, 0.445), (0.328, 0.442), (0.347, 0.441)],
    ],
    '.': [
        [(0.05, 0.078), (0.05, 0.086), (0.051, 0.094), (0.053, 0.101), (0.055, 0.108), (0.057, 0.114), (0.061, 0.12), (0.065, 0.126), (0.069, 0.132), (0.074, 0.137), (0.08, 0.141), (0.086, 0.145), (0.093, 0.148), (0.1, 0.15), (0.108, 0.152), (0.117, 0.153), (0.126, 0.153), (0.135, 0.153), (0.144, 0.152), (0.152, 0.15), (0.16, 0.148), (0.167, 0.145), (0.173, 0.141), (0.179, 0.137), (0.184, 0.132), (0.189, 0.126), (0.193, 0.12), (0.196, 0.114), (0.199, 0.108), (0.201, 0.101), (0.203, 0.094), (0.203, 0.086), (0.204, 0.078), (0.203, 0.071), (0.203, 0.063), (0.201, 0.057), (0.199, 0.05), (0.196, 0.044), (0.193, 0.038), (0.189, 0.032), (0.184, 0.027), (0.179, 0.022), (0.173, 0.018), (0.167, 0.014), (0.16, 0.011), (0.152, 0.009), (0.144, 0.007), (0.135, 0.006), (0.126, 0.006), (0.117, 0.006), (0.108, 0.007), (0.1, 0.009), (0.093, 0.011), (0.086, 0.014), (0.08, 0.018), (0.074, 0.022), (0.069, 0.027), (0.065, 0.032), (0.061, 0.038), (0.057, 0.044), (0.055, 0.05), (0.053, 0.057), (0.051, 0.063), (0.05, 0.071), (0.05, 0.078)],
    ],
}

FILLED_CHAR_WIDTHS = {
    '0': 0.71, '1': 0.47, '2': 0.76, '3': 0.72, '4': 0.8, '5': 0.71,
    '6': 0.71, '7': 0.76, '8': 0.72, '9': 0.71, '.': 0.25,
}


def _scanline_intersections(contours, y):
    """Find all X intersections of a horizontal line at y with polygon contours.

    Uses the same edge-crossing math as _point_in_polygon in svg_engine.py,
    but collects intersection X values instead of toggling inside/outside.
    """
    intersections = []
    for contour in contours:
        n = len(contour)
        for i in range(n - 1):
            x1, y1 = contour[i]
            x2, y2 = contour[i + 1]

            # Skip horizontal edges
            if y1 == y2:
                continue

            # Check if scan line crosses this edge
            if (y1 <= y < y2) or (y2 <= y < y1):
                # Linear interpolation to find X at intersection
                t = (y - y1) / (y2 - y1)
                x = x1 + t * (x2 - x1)
                intersections.append(x)

    intersections.sort()
    return intersections


def get_filled_text_strokes(text, font_size_mm, center_x, center_y, line_spacing_mm):
    """
    Convert text to filled (raster) strokes for G-code using scan-line fill.

    Args:
        text: String to render (numeric like "42" or "18.5")
        font_size_mm: Height of characters in mm
        center_x, center_y: Center position for the text in mm
        line_spacing_mm: Distance between scan lines in mm

    Returns:
        List of strokes, where each stroke is a list of (x, y) points in mm.
        Compatible with generate_gcode_layer().
    """
    if line_spacing_mm <= 0:
        line_spacing_mm = 0.15

    strokes = []

    # Calculate total width of text using FILLED_CHAR_WIDTHS
    total_width = 0
    for char in text:
        total_width += FILLED_CHAR_WIDTHS.get(char, DEFAULT_CHAR_WIDTH)
    total_width += CHAR_SPACING * (len(text) - 1) if len(text) > 1 else 0
    total_width *= font_size_mm

    # Starting X position (left edge of first character)
    start_x = center_x - total_width / 2
    current_x = start_x

    # Y offset to center vertically
    y_offset = center_y - (font_size_mm * CHAR_HEIGHT) / 2

    # Collect all transformed contours for all characters
    all_contours = []
    for char in text:
        if char in FILLED_FONT:
            char_contours = FILLED_FONT[char]
            char_width = FILLED_CHAR_WIDTHS.get(char, DEFAULT_CHAR_WIDTH) * font_size_mm

            for contour in char_contours:
                transformed = []
                for px, py in contour:
                    x = current_x + px * font_size_mm
                    y = y_offset + py * font_size_mm
                    transformed.append((x, y))
                all_contours.append(transformed)

            current_x += char_width + CHAR_SPACING * font_size_mm
        else:
            current_x += DEFAULT_CHAR_WIDTH * font_size_mm + CHAR_SPACING * font_size_mm

    if not all_contours:
        return strokes

    # Find bounding box of all contours
    all_points = [p for c in all_contours for p in c]
    min_y = min(p[1] for p in all_points)
    max_y = max(p[1] for p in all_points)

    # Generate scan lines
    scan_y = min_y + line_spacing_mm / 2
    line_index = 0

    while scan_y < max_y:
        intersections = _scanline_intersections(all_contours, scan_y)

        # Pair up intersections using even-odd rule
        for i in range(0, len(intersections) - 1, 2):
            x_start = intersections[i]
            x_end = intersections[i + 1]
            if x_end - x_start > 0.01:  # Skip degenerate segments
                # Alternate direction for efficiency (reduces rapid moves)
                if line_index % 2 == 0:
                    strokes.append([(x_start, scan_y), (x_end, scan_y)])
                else:
                    strokes.append([(x_end, scan_y), (x_start, scan_y)])

        line_index += 1
        scan_y += line_spacing_mm

    return strokes


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

    Uses the same math as calculate_star_path() in svg_engine.py but returns
    a list of (x, y) points instead of an SVG path string.

    Args:
        cx, cy: Center coordinates
        outer_r: Outer (tip) radius
        inner_r: Inner (valley) radius
        num_points: Number of points (tips) on the star
        shape_factor: 0.0 = sine wave, 1.0 = flattened (square-ish)

    Returns:
        List of (x, y) points forming a closed star shape
    """
    points = []

    avg_r = (outer_r + inner_r) / 2.0
    amplitude = (outer_r - inner_r) / 2.0

    steps = int(num_points * 8)
    if steps < 64:
        steps = 64

    angle_step = (2 * math.pi) / steps
    power = 1.0 - (0.9 * shape_factor)

    for i in range(steps + 1):
        theta = i * angle_step

        raw_wave = math.cos(num_points * theta)
        shaped_wave = (1 if raw_wave >= 0 else -1) * (abs(raw_wave) ** power)

        r = avg_r + amplitude * shaped_wave

        x = cx + r * math.cos(theta)
        y = cy + r * math.sin(theta)
        points.append((x, y))

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


def generate_gcode_layer(strokes, speed_mm_min, power_percent, layer_name, overscan_mm=0):
    """
    Generate G-code for a layer (set of strokes).

    Args:
        strokes: List of strokes, where each stroke is a list of (x, y) points
        speed_mm_min: Feed rate in mm/min
        power_percent: Laser power as percentage (0-100)
        layer_name: Layer name for comment
        overscan_mm: If > 0, extend horizontal 2-point scan lines by this amount
                     on each side with laser off, so the head is at full speed
                     when it reaches the actual fill area.

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

        # Check if this is a horizontal scan line eligible for overscan
        is_scanline = (overscan_mm > 0 and len(stroke) == 2
                       and abs(stroke[0][1] - stroke[1][1]) < 0.001)

        if is_scanline:
            x0, y0 = stroke[0]
            x1, y1 = stroke[1]
            # Determine scan direction and extend
            if x1 > x0:
                approach_x = x0 - overscan_mm
                exit_x = x1 + overscan_mm
            else:
                approach_x = x0 + overscan_mm
                exit_x = x1 - overscan_mm

            # Rapid to overscan start
            lines.append(f"G0 X{approach_x:.3f}Y{y0:.3f}")
            if first_move:
                lines.append(f"; Layer {layer_name}")
                first_move = False
            # Approach at engrave speed, laser off
            lines.append(f"G1 X{x0:.3f}Y{y0:.3f}S0F{speed_mm_min}")
            # Engrave through fill area, laser on
            lines.append(f"G1 X{x1:.3f}S{s_value}")
            # Exit overscan, laser off
            lines.append(f"G1 X{exit_x:.3f}S0")
        else:
            # Standard stroke: rapid to start, then cut
            x0, y0 = stroke[0]
            lines.append(f"G0 X{x0:.3f}Y{y0:.3f}")

            if first_move:
                lines.append(f"; Layer {layer_name}")
                first_move = False

            # Cut moves
            for i, (x, y) in enumerate(stroke[1:], 1):
                if i == 1:
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
        material: Material type ('felt', 'card', 'leather')
        sheet_width_mm, sheet_height_mm: Sheet dimensions in mm
        filename: Output filename
        hole_dia: Center hole diameter in mm (0 for no hole)
        settings: App settings dictionary
        polygon: Optional custom polygon shape
    """
    from svg_engine import _nest_discs

    # Use the same nesting as SVG generation
    placed, _, _ = _nest_discs(pads, material, sheet_width_mm, sheet_height_mm, settings, polygon=polygon)

    if not placed:
        return

    generate_gcode_from_placed(placed, material, sheet_width_mm, sheet_height_mm, filename,
                                hole_dia, settings, polygon=polygon)


def can_generate_gcode(material):
    """Check if G-code generation is supported for a material."""
    return material in ('felt', 'card', 'leather')


def generate_gcode_from_placed(placed, material, sheet_width_mm, sheet_height_mm, filename,
                                hole_dia, settings, polygon=None):
    """
    Generate G-code file from pre-computed placed discs.

    This is used by scrap mode where nesting is done separately via try_nest_partial().
    The function generates G-code for the placed discs without re-running nesting.

    Args:
        placed: List of (pad_size, cx, cy, radius) tuples from nesting
        material: Material type ('felt', 'card', 'leather')
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

    engraving_mode = mat_settings.get("engraving_mode", "line")
    if engraving_mode == "filled":
        engraving_speed = mat_settings.get("filled_engraving_speed", 1000)
        engraving_power = mat_settings.get("filled_engraving_power", 12)
        filled_line_spacing = mat_settings.get("filled_line_spacing", 0.15)
    else:
        engraving_speed = mat_settings.get("engraving_speed", 1500)
        engraving_power = mat_settings.get("engraving_power", 10)

    # Overscan: extend filled scan lines so the head is at speed at the character edge
    overscan_mm = 0
    if engraving_mode == "filled" and settings.get("filled_overscan_enabled", False):
        overscan_mm = settings.get("filled_overscan_mm", 1.5)
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
    use_stars = material == 'leather' and settings.get("darts_enabled", True)
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
            if engraving_mode == "filled":
                text_strokes = get_filled_text_strokes(text, font_size, cx, label_y, filled_line_spacing)
            else:
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
    }
    eng_layer, hole_layer, cut_layer = layer_names.get(material, ('C00', 'C01', 'C02'))

    if engraving_strokes:
        gcode_lines.extend(generate_gcode_layer(engraving_strokes, engraving_speed, engraving_power, eng_layer,
                                                 overscan_mm=overscan_mm))

    if hole_strokes:
        gcode_lines.extend(generate_gcode_layer(hole_strokes, hole_speed, hole_power, hole_layer))

    if cut_strokes:
        gcode_lines.extend(generate_gcode_layer(cut_strokes, cut_speed, cut_power, cut_layer))

    gcode_lines.extend(generate_gcode_footer())

    with open(filename, 'w') as f:
        f.write('\n'.join(gcode_lines))
