//! Python extension module for GPU-accelerated strobe tuner rendering.
//!
//! Build: `cd tuner_renderer && maturin develop --release`
//! Use:   `import tuner_render` in Python

use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;

mod platform;
mod renderer;

#[pyclass]
struct TunerRenderer {
    inner: renderer::Renderer,
}

#[pymethods]
impl TunerRenderer {
    /// Create a new GPU renderer attached to a tkinter widget's native window.
    ///
    /// Args:
    ///     window_handle: Result of widget.winfo_id() (HWND / NSView / X11 id)
    ///     width: Initial surface width in pixels
    ///     height: Initial surface height in pixels
    #[new]
    fn new(window_handle: isize, width: u32, height: u32) -> PyResult<Self> {
        let inner = renderer::Renderer::new(window_handle, width, height)
            .map_err(|e| PyRuntimeError::new_err(format!("GPU init failed: {e}")))?;
        Ok(Self { inner })
    }

    /// Update the rendering surface size (call on widget resize).
    fn resize(&mut self, width: u32, height: u32) {
        self.inner.resize(width, height);
    }

    /// Set wheel positions. Call after resize or layout change.
    ///
    /// Args:
    ///     positions: List of (cx_pixels, cy_pixels, radius_pixels, is_up)
    fn set_layout(&mut self, positions: Vec<(f32, f32, f32, bool)>) {
        self.inner.set_layout(positions);
    }

    /// Render one frame with current wheel states.
    ///
    /// Args:
    ///     ring_phases: 12 lists of 7 per-ring phase offsets in degrees
    ///     magnitudes: 12 magnitude values, 0.0–1.0 (gain already applied)
    ///     ring_magnitudes: 12 lists of 7 per-ring magnitudes (gain applied)
    ///     ring_brightness_pct: 0–100, per-ring brightness blend
    ///     overall_brightness_pct: 10–150, master brightness
    fn render(
        &mut self,
        ring_phases: Vec<Vec<f32>>,
        magnitudes: Vec<f32>,
        ring_magnitudes: Vec<Vec<f32>>,
        ring_brightness_pct: f32,
        overall_brightness_pct: f32,
    ) -> PyResult<()> {
        self.inner
            .render(
                &ring_phases,
                &magnitudes,
                &ring_magnitudes,
                ring_brightness_pct,
                overall_brightness_pct,
            )
            .map_err(|e| PyRuntimeError::new_err(format!("Render failed: {e}")))
    }

    /// Set stripe color from hex string (e.g. "#00FF00").
    fn set_stripe_color(&mut self, hex_color: &str) {
        self.inner.set_stripe_color(hex_color);
    }

    /// Set faceplate (background) color from hex string (e.g. "#1A1A1A").
    fn set_faceplate_color(&mut self, hex_color: &str) {
        self.inner.set_faceplate_color(hex_color);
    }
}

#[pymodule]
fn tuner_render(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<TunerRenderer>()?;
    Ok(())
}
