//! wgpu renderer for 12 stroboscopic tuner wheels.
//!
//! All wheel geometry is computed analytically in the fragment shader —
//! the CPU only uploads 12 wheel parameter structs per frame.

use bytemuck::{Pod, Zeroable};
use wgpu;

use crate::platform;

const MAX_WHEELS: usize = 12;
const BRIGHTNESS_GAMMA: f32 = 0.45;
const MAGNITUDE_THRESHOLD: f32 = 0.02;

// ── GPU data structures (must match shader.wgsl layout exactly) ─────

/// Per-wheel parameters uploaded to the storage buffer each frame.
/// Layout: 128 bytes, aligned to 16 (matching WGSL struct).
#[repr(C)]
#[derive(Copy, Clone, Debug, Pod, Zeroable)]
struct GpuWheelData {
    center: [f32; 2],           // offset  0: pixel coords
    radius: f32,                // offset  8: pixel radius
    _pad0: f32,                 // offset 12: alignment padding
    ring_phases: [f32; 8],      // offset 16: per-ring rotation (radians, 7+1 pad)
    ring_mags: [f32; 8],        // offset 48: per-ring brightness (7+1 pad)
    stripe_color: [f32; 4],     // offset 80: rgb + pad
    faceplate_color: [f32; 4],  // offset 96: rgb + pad
    brightness: f32,            // offset 112: overall brightness (gamma-applied)
    wedge_center: f32,          // offset 116: wedge center angle (radians)
    ring_brightness_blend: f32, // offset 120: 0=uniform, 1=per-ring
    overall_brightness: f32,    // offset 124: master scale
}

// Compile-time layout verification
const _: () = assert!(std::mem::size_of::<GpuWheelData>() == 128);

/// Global uniforms (screen size for pixel→NDC conversion in vertex shader).
#[repr(C)]
#[derive(Copy, Clone, Debug, Pod, Zeroable)]
struct Globals {
    screen_size: [f32; 2],
    _pad: [f32; 2],
}

const _: () = assert!(std::mem::size_of::<Globals>() == 16);

// ── Color helpers ───────────────────────────────────────────────────

fn srgb_to_linear(c: f32) -> f32 {
    if c <= 0.04045 {
        c / 12.92
    } else {
        ((c + 0.055) / 1.055).powf(2.4)
    }
}

fn hex_to_rgb(hex: &str) -> [f32; 3] {
    let hex = hex.trim_start_matches('#');
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f32 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f32 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f32 / 255.0;
    [r, g, b]
}

fn convert_color(hex: &str, to_linear: bool) -> [f32; 3] {
    let rgb = hex_to_rgb(hex);
    if to_linear {
        [
            srgb_to_linear(rgb[0]),
            srgb_to_linear(rgb[1]),
            srgb_to_linear(rgb[2]),
        ]
    } else {
        rgb
    }
}

// ── Renderer ────────────────────────────────────────────────────────

pub struct Renderer {
    device: wgpu::Device,
    queue: wgpu::Queue,
    surface: wgpu::Surface<'static>,
    config: wgpu::SurfaceConfiguration,
    pipeline: wgpu::RenderPipeline,
    wheel_buffer: wgpu::Buffer,
    globals_buffer: wgpu::Buffer,
    bind_group: wgpu::BindGroup,

    // State
    layouts: Vec<(f32, f32, f32, bool)>, // (cx, cy, radius, is_up) in pixels
    stripe_color: [f32; 3],
    faceplate_color: [f32; 3],
    is_srgb: bool,
    width: u32,
    height: u32,
}

impl Renderer {
    pub fn new(
        window_handle: isize,
        width: u32,
        height: u32,
    ) -> Result<Self, Box<dyn std::error::Error>> {
        let (raw_window, raw_display) = platform::raw_handles_from_winfo_id(window_handle);

        let instance = wgpu::Instance::new(wgpu::InstanceDescriptor {
            backends: wgpu::Backends::all(),
            ..Default::default()
        });

        // SAFETY: the Python side must keep the tkinter widget alive
        // as long as this Renderer exists.
        let surface = unsafe {
            instance.create_surface_unsafe(wgpu::SurfaceTargetUnsafe::RawHandle {
                raw_display_handle: raw_display,
                raw_window_handle: raw_window,
            })?
        };

        let adapter = pollster::block_on(instance.request_adapter(&wgpu::RequestAdapterOptions {
            compatible_surface: Some(&surface),
            power_preference: wgpu::PowerPreference::LowPower,
            ..Default::default()
        }))
        .ok_or("No suitable GPU adapter found")?;

        let (device, queue) = pollster::block_on(adapter.request_device(
            &wgpu::DeviceDescriptor {
                label: Some("tuner"),
                required_features: wgpu::Features::empty(),
                required_limits: wgpu::Limits::downlevel_defaults(),
                ..Default::default()
            },
            None,
        ))?;

        // Prefer sRGB for correct color handling
        let surface_caps = surface.get_capabilities(&adapter);
        let format = surface_caps
            .formats
            .iter()
            .find(|f| f.is_srgb())
            .copied()
            .unwrap_or(surface_caps.formats[0]);
        let is_srgb = format.is_srgb();

        let config = wgpu::SurfaceConfiguration {
            usage: wgpu::TextureUsages::RENDER_ATTACHMENT,
            format,
            width: width.max(1),
            height: height.max(1),
            present_mode: wgpu::PresentMode::Fifo, // VSync
            alpha_mode: surface_caps.alpha_modes[0],
            view_formats: vec![],
            desired_maximum_frame_latency: 2,
        };
        surface.configure(&device, &config);

        // ── Buffers ──

        let wheel_buffer = device.create_buffer(&wgpu::BufferDescriptor {
            label: Some("wheel_data"),
            size: (MAX_WHEELS * std::mem::size_of::<GpuWheelData>()) as u64,
            usage: wgpu::BufferUsages::STORAGE | wgpu::BufferUsages::COPY_DST,
            mapped_at_creation: false,
        });

        let globals_buffer = device.create_buffer(&wgpu::BufferDescriptor {
            label: Some("globals"),
            size: std::mem::size_of::<Globals>() as u64,
            usage: wgpu::BufferUsages::UNIFORM | wgpu::BufferUsages::COPY_DST,
            mapped_at_creation: false,
        });

        // Upload initial globals
        queue.write_buffer(
            &globals_buffer,
            0,
            bytemuck::bytes_of(&Globals {
                screen_size: [width as f32, height as f32],
                _pad: [0.0; 2],
            }),
        );

        // ── Bind group ──

        let bind_group_layout =
            device.create_bind_group_layout(&wgpu::BindGroupLayoutDescriptor {
                label: Some("tuner_bgl"),
                entries: &[
                    wgpu::BindGroupLayoutEntry {
                        binding: 0,
                        visibility: wgpu::ShaderStages::VERTEX | wgpu::ShaderStages::FRAGMENT,
                        ty: wgpu::BindingType::Buffer {
                            ty: wgpu::BufferBindingType::Storage { read_only: true },
                            has_dynamic_offset: false,
                            min_binding_size: None,
                        },
                        count: None,
                    },
                    wgpu::BindGroupLayoutEntry {
                        binding: 1,
                        visibility: wgpu::ShaderStages::VERTEX,
                        ty: wgpu::BindingType::Buffer {
                            ty: wgpu::BufferBindingType::Uniform,
                            has_dynamic_offset: false,
                            min_binding_size: None,
                        },
                        count: None,
                    },
                ],
            });

        let bind_group = device.create_bind_group(&wgpu::BindGroupDescriptor {
            label: Some("tuner_bg"),
            layout: &bind_group_layout,
            entries: &[
                wgpu::BindGroupEntry {
                    binding: 0,
                    resource: wheel_buffer.as_entire_binding(),
                },
                wgpu::BindGroupEntry {
                    binding: 1,
                    resource: globals_buffer.as_entire_binding(),
                },
            ],
        });

        // ── Pipeline ──

        let shader = device.create_shader_module(wgpu::ShaderModuleDescriptor {
            label: Some("wheel_shader"),
            source: wgpu::ShaderSource::Wgsl(include_str!("shader.wgsl").into()),
        });

        let pipeline_layout = device.create_pipeline_layout(&wgpu::PipelineLayoutDescriptor {
            label: Some("tuner_pl"),
            bind_group_layouts: &[&bind_group_layout],
            push_constant_ranges: &[],
        });

        let pipeline = device.create_render_pipeline(&wgpu::RenderPipelineDescriptor {
            label: Some("tuner_rp"),
            layout: Some(&pipeline_layout),
            vertex: wgpu::VertexState {
                module: &shader,
                entry_point: Some("vs_main"),
                buffers: &[],
                compilation_options: Default::default(),
            },
            fragment: Some(wgpu::FragmentState {
                module: &shader,
                entry_point: Some("fs_main"),
                targets: &[Some(wgpu::ColorTargetState {
                    format: config.format,
                    blend: Some(wgpu::BlendState::ALPHA_BLENDING),
                    write_mask: wgpu::ColorWrites::ALL,
                })],
                compilation_options: Default::default(),
            }),
            primitive: wgpu::PrimitiveState {
                topology: wgpu::PrimitiveTopology::TriangleStrip,
                strip_index_format: None,
                front_face: wgpu::FrontFace::Ccw,
                cull_mode: None, // no culling for 2D
                ..Default::default()
            },
            depth_stencil: None,
            multisample: wgpu::MultisampleState::default(),
            multiview: None,
            cache: None,
        });

        // Default colors (green stripes, dark faceplate)
        let stripe_color = convert_color("#00FF00", is_srgb);
        let faceplate_color = convert_color("#1A1A1A", is_srgb);

        Ok(Self {
            device,
            queue,
            surface,
            config,
            pipeline,
            wheel_buffer,
            globals_buffer,
            bind_group,
            layouts: Vec::new(),
            stripe_color,
            faceplate_color,
            is_srgb,
            width: width.max(1),
            height: height.max(1),
        })
    }

    pub fn resize(&mut self, width: u32, height: u32) {
        let w = width.max(1);
        let h = height.max(1);
        if w == self.width && h == self.height {
            return;
        }
        self.width = w;
        self.height = h;
        self.config.width = w;
        self.config.height = h;
        self.surface.configure(&self.device, &self.config);

        // Update globals
        self.queue.write_buffer(
            &self.globals_buffer,
            0,
            bytemuck::bytes_of(&Globals {
                screen_size: [w as f32, h as f32],
                _pad: [0.0; 2],
            }),
        );
    }

    pub fn set_layout(&mut self, positions: Vec<(f32, f32, f32, bool)>) {
        self.layouts = positions;
    }

    pub fn set_stripe_color(&mut self, hex: &str) {
        self.stripe_color = convert_color(hex, self.is_srgb);
    }

    pub fn set_faceplate_color(&mut self, hex: &str) {
        self.faceplate_color = convert_color(hex, self.is_srgb);
    }

    pub fn render(
        &mut self,
        ring_phases_deg: &[Vec<f32>],
        magnitudes: &[f32],
        ring_magnitudes: &[Vec<f32>],
        ring_brightness_pct: f32,
        overall_brightness_pct: f32,
    ) -> Result<(), Box<dyn std::error::Error>> {
        let num_wheels = self.layouts.len().min(MAX_WHEELS);
        if num_wheels == 0 {
            return Ok(());
        }

        // Build per-wheel GPU data
        let mut gpu_data = [GpuWheelData::zeroed(); MAX_WHEELS];

        for (i, &(cx, cy, radius, is_up)) in self.layouts.iter().enumerate().take(num_wheels) {
            let mag = magnitudes.get(i).copied().unwrap_or(0.0);
            let rm = ring_magnitudes.get(i);
            let rp = ring_phases_deg.get(i);

            // Gamma-corrected brightness
            let brightness = if mag > MAGNITUDE_THRESHOLD {
                mag.powf(BRIGHTNESS_GAMMA).min(1.0)
            } else {
                0.0
            };

            // Per-ring phases (degrees → radians)
            let mut ring_phases = [0.0f32; 8];
            for j in 0..7 {
                ring_phases[j] = rp
                    .and_then(|r| r.get(j).copied())
                    .unwrap_or(0.0)
                    .to_radians();
            }

            // Per-ring magnitudes with gamma
            let mut ring_mags = [0.0f32; 8];
            for j in 0..7 {
                let v = rm
                    .and_then(|r| r.get(j).copied())
                    .unwrap_or(mag);
                ring_mags[j] = if v > MAGNITUDE_THRESHOLD {
                    v.powf(BRIGHTNESS_GAMMA).min(1.0)
                } else {
                    0.0
                };
            }

            let wedge_center = if is_up {
                0.0 // top = angle 0 in our atan2(x,y) convention
            } else {
                std::f32::consts::PI // bottom
            };

            gpu_data[i] = GpuWheelData {
                center: [cx, cy],
                radius,
                _pad0: 0.0,
                ring_phases,
                ring_mags,
                stripe_color: [
                    self.stripe_color[0],
                    self.stripe_color[1],
                    self.stripe_color[2],
                    0.0,
                ],
                faceplate_color: [
                    self.faceplate_color[0],
                    self.faceplate_color[1],
                    self.faceplate_color[2],
                    0.0,
                ],
                brightness,
                wedge_center,
                ring_brightness_blend: ring_brightness_pct / 100.0,
                overall_brightness: overall_brightness_pct / 100.0,
            };
        }

        // Upload wheel data
        self.queue
            .write_buffer(&self.wheel_buffer, 0, bytemuck::cast_slice(&gpu_data));

        // Render
        let output = self.surface.get_current_texture()?;
        let view = output
            .texture
            .create_view(&wgpu::TextureViewDescriptor::default());

        let fp = &self.faceplate_color;
        let mut encoder = self
            .device
            .create_command_encoder(&wgpu::CommandEncoderDescriptor {
                label: Some("tuner_enc"),
            });

        {
            let mut pass = encoder.begin_render_pass(&wgpu::RenderPassDescriptor {
                label: Some("tuner_pass"),
                color_attachments: &[Some(wgpu::RenderPassColorAttachment {
                    view: &view,
                    resolve_target: None,
                    ops: wgpu::Operations {
                        load: wgpu::LoadOp::Clear(wgpu::Color {
                            r: fp[0] as f64,
                            g: fp[1] as f64,
                            b: fp[2] as f64,
                            a: 1.0,
                        }),
                        store: wgpu::StoreOp::Store,
                    },
                })],
                depth_stencil_attachment: None,
                ..Default::default()
            });

            pass.set_pipeline(&self.pipeline);
            pass.set_bind_group(0, &self.bind_group, &[]);
            pass.draw(0..4, 0..num_wheels as u32);
        }

        self.queue.submit(std::iter::once(encoder.finish()));
        output.present();

        Ok(())
    }
}
