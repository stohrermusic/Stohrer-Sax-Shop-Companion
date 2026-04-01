// Strobe wheel GPU shader — analytical per-pixel computation.
//
// Renders 12 stroboscopic tuner wheels as instanced quads.
// The fragment shader computes ring patterns, rotation, wedge masking,
// and per-ring brightness entirely per-pixel — no geometry needed.

// ── Data structures ─────────────────────────────────────────────────

struct WheelData {
    center: vec2<f32>,              // pixel coordinates
    radius: f32,                    // pixel radius
    _pad0: f32,                     // alignment padding
    ring_phases: array<f32, 8>,     // per-ring rotation offsets (radians, 7 used + 1 pad)
    ring_mags: array<f32, 8>,       // per-ring brightness (7 used + 1 pad)
    stripe_color: vec4<f32>,        // rgb + unused
    faceplate_color: vec4<f32>,     // rgb + unused
    brightness: f32,                // overall wheel brightness (gamma-applied)
    wedge_center: f32,              // wedge center angle (radians, 0 = top)
    ring_brightness_blend: f32,     // 0 = uniform, 1 = full per-ring
    overall_brightness: f32,        // master brightness scale
};

struct Globals {
    screen_size: vec2<f32>,
    _pad: vec2<f32>,
};

@group(0) @binding(0) var<storage, read> wheels: array<WheelData, 12>;
@group(0) @binding(1) var<uniform> globals: Globals;

struct VertexOutput {
    @builtin(position) clip_pos: vec4<f32>,
    @location(0) uv: vec2<f32>,
    @location(1) @interpolate(flat) wheel_idx: u32,
};

// ── Vertex shader ───────────────────────────────────────────────────

@vertex
fn vs_main(
    @builtin(vertex_index) vid: u32,
    @builtin(instance_index) iid: u32,
) -> VertexOutput {
    let wheel = wheels[iid];

    // Quad corners via triangle strip: 0=TL, 1=TR, 2=BL, 3=BR
    let qx = select(-1.0, 1.0, (vid & 1u) != 0u);
    let qy = select(-1.0, 1.0, (vid & 2u) != 0u);

    // Small margin so anti-aliased edges aren't clipped
    let margin = 1.03;

    // Pixel position of this vertex
    let px = wheel.center.x + qx * wheel.radius * margin;
    let py = wheel.center.y + qy * wheel.radius * margin;

    // Convert pixel coords → NDC
    let ndc_x = px / globals.screen_size.x * 2.0 - 1.0;
    let ndc_y = -(py / globals.screen_size.y * 2.0 - 1.0);

    // UV with Y-up convention (for angle math in fragment shader)
    let uv = vec2(qx, -qy) * margin;

    var out: VertexOutput;
    out.clip_pos = vec4(ndc_x, ndc_y, 0.0, 1.0);
    out.uv = uv;
    out.wheel_idx = iid;
    return out;
}

// ── Fragment shader constants ───────────────────────────────────────

const PI: f32 = 3.14159265358979;
const TWO_PI: f32 = 6.28318530717959;

const CENTER_GAP: f32 = 0.12;          // center hole as fraction of radius
const RING_GAP: f32 = 0.015;           // inter-ring gap as fraction of ring width
const WEDGE_HALF_ANGLE: f32 = 0.6981;  // 40° in radians (80° total opening)
const NUM_RINGS: f32 = 7.0;
const DIM_MULTIPLIER: f32 = 0.08;      // minimum brightness for visible stripes

// Segment counts per ring (octaves 1-7, doubling frequency each ring)
fn ring_segments(ring: u32) -> f32 {
    switch ring {
        case 0u: { return 4.0; }
        case 1u: { return 8.0; }
        case 2u: { return 16.0; }
        case 3u: { return 32.0; }
        case 4u: { return 64.0; }
        case 5u: { return 128.0; }
        case 6u: { return 256.0; }
        default: { return 256.0; }
    }
}

// Shortest angular distance on a circle
fn angle_distance(a: f32, b: f32) -> f32 {
    var d = a - b;
    d = d - TWO_PI * round(d / TWO_PI);
    return abs(d);
}

// ── Fragment shader ─────────────────────────────────────────────────

@fragment
fn fs_main(in: VertexOutput) -> @location(0) vec4<f32> {
    let wheel = wheels[in.wheel_idx];
    let fp = vec4(wheel.faceplate_color.rgb, 1.0);

    // Distance from wheel center in UV space (0 to ~1.03)
    let dist = length(in.uv);

    // Outside disc — discard to avoid overwriting overlapping wheels
    if dist > 1.0 {
        discard;
    }

    // Inside center gap — faceplate
    if dist < CENTER_GAP {
        return fp;
    }

    // Anti-aliased outer edge
    let edge_aa = 1.0 - smoothstep(0.99, 1.0, dist);

    // Angle from top, clockwise: atan2(x, y) gives 0 at +Y (top)
    let angle = atan2(in.uv.x, in.uv.y);

    // Wedge mask — only render within the visible opening
    let wd = angle_distance(angle, wheel.wedge_center);
    if wd > WEDGE_HALF_ANGLE + 0.02 {
        return fp;
    }
    let wedge_aa = 1.0 - smoothstep(WEDGE_HALF_ANGLE - 0.01, WEDGE_HALF_ANGLE, wd);

    // Ring index (0 = innermost, 6 = outermost)
    let ring_zone = 1.0 - CENTER_GAP;
    let ring_f = (dist - CENTER_GAP) / ring_zone * NUM_RINGS;
    let ring_idx = min(u32(ring_f), 6u);
    let ring_local = fract(ring_f);

    // Inter-ring gap (thin dark line between rings)
    if ring_local < RING_GAP || ring_local > (1.0 - RING_GAP) {
        return fp;
    }

    // Stripe alternation: sin() of rotated angle gives smooth alternation,
    // step() converts to binary stripe. segments/2 full sine cycles = segments zones.
    // Each ring uses its own phase — independent frequency tracking per octave,
    // showing real inharmonicity across octaves.
    let segs = ring_segments(ring_idx);
    let rotated = angle - wheel.ring_phases[ring_idx];
    let stripe = step(0.0, sin(rotated * segs * 0.5));

    // Per-ring vs uniform brightness blending
    let ring_mag = wheel.ring_mags[ring_idx];
    // DIM_MULTIPLIER keeps stripes visible in always-spinning mode.
    // When brightness is exactly 0, the wheel is inactive — don't force dim.
    let uniform_b = select(max(DIM_MULTIPLIER, wheel.brightness), 0.0, wheel.brightness <= 0.0);
    let per_ring_b = select(max(DIM_MULTIPLIER, ring_mag), 0.0, wheel.brightness <= 0.0);
    let blended_b = mix(uniform_b, per_ring_b, wheel.ring_brightness_blend);
    let final_b = blended_b * wheel.overall_brightness;

    // Stripe regions get the stripe color at computed brightness;
    // non-stripe regions are faceplate (dark)
    let lit = wheel.stripe_color.rgb * final_b;
    let color = mix(fp.rgb, lit, stripe);

    return vec4(color, edge_aa * wedge_aa);
}
