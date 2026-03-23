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

    // wgpu's Vulkan backend requires a real X11 Display pointer.
    // Since tkinter is already using X11, libX11 is loaded in the process.
    #[link(name = "X11")]
    extern "C" {
        fn XOpenDisplay(display_name: *const std::ffi::c_char) -> *mut std::ffi::c_void;
    }

    let display = unsafe { XOpenDisplay(std::ptr::null()) };
    let display_ptr = std::ptr::NonNull::new(display);

    let wh = XlibWindowHandle::new(handle as std::ffi::c_ulong);
    let dh = XlibDisplayHandle::new(display_ptr, 0);
    (RawWindowHandle::Xlib(wh), RawDisplayHandle::Xlib(dh))
}
