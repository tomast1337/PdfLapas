//! C ABI over `pdf_core::Renderer` for `@pdflapas/native` (bun:ffi), mirroring what `crates/pdf-wasm` exposes:
//! `pdflapas_renderer_new`, `pdflapas_add_font` (= `addFont`), `pdflapas_render` (= `renderHtml`), `pdflapas_report`
//! (`warnings()` and `missingFonts()` as one UTF-8 JSON buffer), `pdflapas_last_error`, `pdflapas_buffer_free` and
//! `pdflapas_renderer_free`.
//!
//! Every buffer a call writes to `out` belongs to the caller, who copies it and releases it with
//! `pdflapas_buffer_free`. A panic never crosses the boundary: each call runs under `catch_unwind` (the `native`
//! profile unwinds), returns `PDFLAPAS_ERROR`, and poisons the handle, whose state the panic may have left half-written.
use std::fmt::Write as _;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::ptr;

pub const PDFLAPAS_OK: i32 = 0;
/// The call failed; `pdflapas_last_error` holds the message (none for a null handle).
pub const PDFLAPAS_ERROR: i32 = 1;

/// An owned byte buffer handed to the caller.
#[repr(C)]
pub struct PdflapasBuffer {
    pub ptr: *mut u8,
    pub len: usize,
}

/// The opaque handle: one renderer, its fonts (registered once) and the last error message.
pub struct PdflapasRenderer {
    renderer: pdf_core::Renderer,
    last_error: String,
    poisoned: bool,
}

const POISONED: &str = "this native renderer panicked in an earlier call and cannot be used again; create another (and report the document that caused it)";

fn panic_message(payload: &(dyn std::any::Any + Send)) -> String {
    let detail = payload
        .downcast_ref::<&str>()
        .map(|text| (*text).to_owned())
        .or_else(|| payload.downcast_ref::<String>().cloned())
        .unwrap_or_default();
    format!("the native renderer panicked: {detail}; {POISONED}")
}

/// Bytes from a caller pointer: a null pointer is an empty slice only when `len` is 0.
///
/// # Safety
/// A non-null `data` must point to `len` readable bytes that stay valid for the returned lifetime.
unsafe fn caller_bytes<'a>(data: *const u8, len: usize) -> Result<&'a [u8], String> {
    if data.is_null() {
        return if len == 0 { Ok(&[]) } else { Err(format!("null pointer with length {len}")) };
    }
    // SAFETY: non-null, and the caller guarantees `len` readable bytes.
    Ok(unsafe { std::slice::from_raw_parts(data, len) })
}

/// Hands `bytes` to the caller through `out`.
///
/// # Safety
/// `out` must be a valid, writable `PdflapasBuffer`.
unsafe fn give(out: *mut PdflapasBuffer, bytes: Vec<u8>) {
    let boxed = bytes.into_boxed_slice();
    let len = boxed.len();
    // SAFETY: the caller guarantees `out` is writable.
    unsafe {
        out.write(PdflapasBuffer {
            ptr: Box::into_raw(boxed).cast::<u8>(),
            len,
        });
    }
}

/// Runs `call` on the handle under `catch_unwind` and writes its bytes to `out`, or its error to `last_error`.
///
/// # Safety
/// `handle` is null or from `pdflapas_renderer_new`; `out` is null or a valid, writable `PdflapasBuffer`.
unsafe fn run(handle: *mut PdflapasRenderer, out: *mut PdflapasBuffer, call: impl FnOnce(&mut PdflapasRenderer) -> Result<Vec<u8>, String>) -> i32 {
    // SAFETY: the caller guarantees `handle` is null or live and not used by another thread during the call.
    let Some(state) = (unsafe { handle.as_mut() }) else {
        return PDFLAPAS_ERROR;
    };
    if out.is_null() {
        state.last_error = "null output buffer".to_owned();
        return PDFLAPAS_ERROR;
    }
    if state.poisoned {
        state.last_error = POISONED.to_owned();
        return PDFLAPAS_ERROR;
    }
    match catch_unwind(AssertUnwindSafe(|| call(state))) {
        Ok(Ok(bytes)) => {
            // SAFETY: `out` is non-null and the caller guarantees it is writable.
            unsafe { give(out, bytes) };
            PDFLAPAS_OK
        }
        Ok(Err(message)) => {
            state.last_error = message;
            PDFLAPAS_ERROR
        }
        Err(payload) => {
            state.poisoned = true;
            state.last_error = panic_message(payload.as_ref());
            PDFLAPAS_ERROR
        }
    }
}

/// Appends `text` as a JSON string literal.
fn json_string(json: &mut String, text: &str) {
    json.push('"');
    for c in text.chars() {
        match c {
            '"' => json.push_str("\\\""),
            '\\' => json.push_str("\\\\"),
            '\n' => json.push_str("\\n"),
            '\r' => json.push_str("\\r"),
            '\t' => json.push_str("\\t"),
            // Writing to a String cannot fail.
            c if u32::from(c) < 0x20 => {
                let _ = write!(json, "\\u{:04x}", u32::from(c));
            }
            c => json.push(c),
        }
    }
    json.push('"');
}

/// `{"warnings":[...],"missingFonts":[{"face","used"?,"objects":[...]}]}` of the last render.
fn report_json(renderer: &pdf_core::Renderer) -> String {
    let mut json = String::from("{\"warnings\":[");
    for (index, warning) in renderer.warnings().iter().enumerate() {
        if index > 0 {
            json.push(',');
        }
        json_string(&mut json, warning);
    }
    json.push_str("],\"missingFonts\":[");
    for (index, missing) in renderer.missing_fonts().iter().enumerate() {
        if index > 0 {
            json.push(',');
        }
        json.push_str("{\"face\":");
        json_string(&mut json, &missing.face);
        if let Some(used) = &missing.used {
            json.push_str(",\"used\":");
            json_string(&mut json, used);
        }
        json.push_str(",\"objects\":[");
        for (position, object) in missing.objects.iter().enumerate() {
            if position > 0 {
                json.push(',');
            }
            json_string(&mut json, object);
        }
        json.push_str("]}");
    }
    json.push_str("]}");
    json
}

/// A new renderer with no fonts, or null if creating it panicked. Release it with `pdflapas_renderer_free`.
#[unsafe(no_mangle)]
pub extern "C" fn pdflapas_renderer_new() -> *mut PdflapasRenderer {
    catch_unwind(|| {
        Box::into_raw(Box::new(PdflapasRenderer {
            renderer: pdf_core::Renderer::new(),
            last_error: String::new(),
            poisoned: false,
        }))
    })
    .unwrap_or(ptr::null_mut())
}

/// Registers a TTF/OTF/TTC file (copied) and returns its number of faces; 0 for a file with no usable face, a null
/// handle or pointer, or a panic (which poisons the handle).
///
/// # Safety
/// `handle` is null or from `pdflapas_renderer_new`; a non-null `data` points to `len` readable bytes.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_add_font(handle: *mut PdflapasRenderer, data: *const u8, len: usize) -> u32 {
    // SAFETY: the caller guarantees `handle` is null or live.
    let Some(state) = (unsafe { handle.as_mut() }) else {
        return 0;
    };
    // SAFETY: the caller guarantees `len` readable bytes at a non-null `data`.
    let Ok(bytes) = (unsafe { caller_bytes(data, len) }) else {
        return 0;
    };
    if state.poisoned {
        return 0;
    }
    let renderer = &mut state.renderer;
    match catch_unwind(AssertUnwindSafe(|| renderer.add_font(bytes.to_vec()))) {
        Ok(faces) => faces,
        Err(payload) => {
            state.poisoned = true;
            state.last_error = panic_message(payload.as_ref());
            0
        }
    }
}

/// Renders a PdfLapas HTML profile document (UTF-8, `toHtml().html`) into PDF bytes written to `out`.
///
/// # Safety
/// `handle` is null or from `pdflapas_renderer_new`; a non-null `html` points to `len` readable bytes; `out` is null
/// or a valid, writable `PdflapasBuffer`.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_render(handle: *mut PdflapasRenderer, html: *const u8, len: usize, out: *mut PdflapasBuffer) -> i32 {
    // SAFETY: forwarded from this function's contract.
    unsafe {
        run(handle, out, |state| {
            // SAFETY: the caller guarantees `len` readable bytes at a non-null `html`.
            let bytes = caller_bytes(html, len)?;
            let text = std::str::from_utf8(bytes).map_err(|error| format!("the HTML is not UTF-8: {error}"))?;
            state.renderer.render_html(text)
        })
    }
}

/// Writes the last render's warnings and missing fonts to `out` as UTF-8 JSON (see `report_json`).
///
/// # Safety
/// As `pdflapas_render`.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_report(handle: *mut PdflapasRenderer, out: *mut PdflapasBuffer) -> i32 {
    // SAFETY: forwarded from this function's contract.
    unsafe { run(handle, out, |state| Ok(report_json(&state.renderer).into_bytes())) }
}

/// Writes the message of the last failed call to `out` (UTF-8). Answers even on a poisoned handle.
///
/// # Safety
/// As `pdflapas_render`.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_last_error(handle: *mut PdflapasRenderer, out: *mut PdflapasBuffer) -> i32 {
    // SAFETY: the caller guarantees `handle` is null or live.
    let Some(state) = (unsafe { handle.as_ref() }) else {
        return PDFLAPAS_ERROR;
    };
    if out.is_null() {
        return PDFLAPAS_ERROR;
    }
    // SAFETY: `out` is non-null and the caller guarantees it is writable.
    unsafe { give(out, state.last_error.clone().into_bytes()) };
    PDFLAPAS_OK
}

/// Releases a buffer a call wrote to `out`. Null is a no-op.
///
/// # Safety
/// `ptr` and `len` are exactly a `PdflapasBuffer` this library returned, released once.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_buffer_free(ptr: *mut u8, len: usize) {
    if !ptr.is_null() {
        // SAFETY: the caller guarantees this is a boxed slice of `len` bytes from `give`, not yet released.
        drop(unsafe { Box::from_raw(ptr::slice_from_raw_parts_mut(ptr, len)) });
    }
}

/// Releases a renderer and its fonts. Null is a no-op.
///
/// # Safety
/// `handle` is null or from `pdflapas_renderer_new`, released once and not used afterwards.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_renderer_free(handle: *mut PdflapasRenderer) {
    if !handle.is_null() {
        // SAFETY: the caller guarantees `handle` came from `Box::into_raw` in `pdflapas_renderer_new`.
        drop(unsafe { Box::from_raw(handle) });
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn take(buffer: &PdflapasBuffer) -> String {
        let text = String::from_utf8(unsafe { std::slice::from_raw_parts(buffer.ptr, buffer.len) }.to_vec()).unwrap();
        unsafe { pdflapas_buffer_free(buffer.ptr, buffer.len) };
        text
    }

    #[test]
    fn invalid_html_is_an_error_status_with_its_message() {
        let handle = pdflapas_renderer_new();
        let mut out = PdflapasBuffer { ptr: ptr::null_mut(), len: 0 };
        let html = b"<p>not a profile document</p>";

        assert_eq!(unsafe { pdflapas_render(handle, html.as_ptr(), html.len(), &mut out) }, PDFLAPAS_ERROR);
        assert!(out.ptr.is_null());
        assert_eq!(unsafe { pdflapas_last_error(handle, &mut out) }, PDFLAPAS_OK);
        assert!(take(&out).starts_with("invalid PdfLapas HTML"));
        assert_eq!(unsafe { pdflapas_report(handle, &mut out) }, PDFLAPAS_OK);
        assert_eq!(take(&out), r#"{"warnings":[],"missingFonts":[]}"#);
        unsafe { pdflapas_renderer_free(handle) };
    }

    #[test]
    fn bad_arguments_are_errors_not_crashes() {
        let handle = pdflapas_renderer_new();
        let mut out = PdflapasBuffer { ptr: ptr::null_mut(), len: 0 };
        let invalid_utf8 = [0xff_u8, 0xfe];

        assert_eq!(unsafe { pdflapas_render(ptr::null_mut(), ptr::null(), 0, &mut out) }, PDFLAPAS_ERROR);
        assert_eq!(unsafe { pdflapas_render(handle, ptr::null(), 3, &mut out) }, PDFLAPAS_ERROR);
        assert_eq!(unsafe { pdflapas_render(handle, invalid_utf8.as_ptr(), 2, ptr::null_mut()) }, PDFLAPAS_ERROR);
        assert_eq!(unsafe { pdflapas_render(handle, invalid_utf8.as_ptr(), 2, &mut out) }, PDFLAPAS_ERROR);
        assert_eq!(unsafe { pdflapas_last_error(handle, &mut out) }, PDFLAPAS_OK);
        assert!(take(&out).starts_with("the HTML is not UTF-8"));
        assert_eq!(unsafe { pdflapas_add_font(handle, b"not a font".as_ptr(), 10) }, 0);
        assert_eq!(unsafe { pdflapas_add_font(handle, ptr::null(), 10) }, 0);
        unsafe { pdflapas_renderer_free(handle) };
        unsafe { pdflapas_renderer_free(ptr::null_mut()) };
        unsafe { pdflapas_buffer_free(ptr::null_mut(), 0) };
    }

    #[test]
    fn a_panic_poisons_the_handle_instead_of_unwinding_out() {
        let handle = pdflapas_renderer_new();
        let mut out = PdflapasBuffer { ptr: ptr::null_mut(), len: 0 };

        assert_eq!(unsafe { run(handle, &mut out, |_| panic!("boom")) }, PDFLAPAS_ERROR);
        assert_eq!(unsafe { pdflapas_last_error(handle, &mut out) }, PDFLAPAS_OK);
        assert!(take(&out).starts_with("the native renderer panicked: boom;"));
        assert_eq!(unsafe { pdflapas_report(handle, &mut out) }, PDFLAPAS_ERROR);
        unsafe { pdflapas_renderer_free(handle) };
    }

    #[test]
    fn report_strings_are_json_escaped() {
        let mut json = String::new();

        json_string(&mut json, "a\"b\\c\nd\u{1}é");
        assert_eq!(json, r#""a\"b\\c\nd\u0001é""#);
    }
}
