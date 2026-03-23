//! Platform-specific window handle creation from tkinter's winfo_id().
//!
//! tkinter's `winfo_id()` returns:
//!   - Windows: HWND (isize)
//!   - macOS:   NSView pointer (isize)
//!   - Linux:   X11 window ID (isize, but really u32)

use raw_window_handle::{RawDisplayHandle, RawWindowHandle};

/// Create raw window and display handles from a tkinter widget's `winfo_id()`.
///
/// # Safety
/// The caller must ensure the handle remains valid for the lifetime of any
/// surface created from it. In practice this means the tkinter widget must
/// outlive the wgpu Surface.
#[cfg(target_os = "windows")]
pub fn raw_handles_from_winfo_id(handle: isize) -> (RawWindowHandle, RawDisplayHandle) {
    use raw_window_handle::{Win32WindowHandle, WindowsDisplayHandle};
    let wh = Win32WindowHandle::new(
        std::num::NonZeroIsize::new(handle).expect("null HWND from winfo_id()"),
    );
    let dh = WindowsDisplayHandle::new();
    (RawWindowHandle::Win32(wh), RawDisplayHandle::Windows(dh))
}

#[cfg(target_os = "macos")]
pub fn raw_handles_from_winfo_id(handle: isize) -> (RawWindowHandle, RawDisplayHandle) {
    use raw_window_handle::{AppKitWindowHandle, AppKitDisplayHandle};
    let ns_view = std::ptr::NonNull::new(handle as *mut std::ffi::c_void)
        .expect("null NSView from winfo_id()");
    let wh = AppKitWindowHandle::new(ns_view);
    let dh = AppKitDisplayHandle::new();
    (RawWindowHandle::AppKit(wh), RawDisplayHandle::AppKit(dh))
}

#[cfg(target_os = "linux")]
pub fn raw_handles_from_winfo_id(handle: isize) -> (RawWindowHandle, RawDisplayHandle) {
    use raw_window_handle::{XlibWindowHandle, XlibDisplayHandle};
    let window = std::num::NonZero::new(handle as u32)
        .expect("null X11 window from winfo_id()");
    let wh = XlibWindowHandle::new(window);
    // None = default display, 0 = default screen
    let dh = XlibDisplayHandle::new(None, 0);
    (RawWindowHandle::Xlib(wh), RawDisplayHandle::Xlib(dh))
}
