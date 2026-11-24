//! Several renderers on their own threads behind one handle, for `createNativeRendererPool()`.
//!
//! The font files are read once into `SharedFont`s (bytes behind an `Arc`) and every thread registers the same bytes in
//! its own `pdf_core::Renderer`. What stays per thread is what a render mutates: the parsed font collection and
//! parley's layout caches, krilla's per-font PDF objects, the scratch buffers and the document being drawn.
//!
//! Jobs go to the threads through one channel, first free thread first. A finished job goes to a completion queue
//! and rings `wake`, a function the caller provides (a threadsafe bun:ffi `JSCallback`) that carries no data: the
//! caller then drains the queue with `pdflapas_pool_next` on its own thread. Rings coalesce, so the caller is not called
//! once per job, and nothing the queue owns crosses into JS until the caller asks for it. A panicking render fails
//! that job only: its thread replaces its renderer and goes on.
use std::collections::VecDeque;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::ptr;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc::{Receiver, Sender, channel};
use std::sync::{Arc, Mutex, PoisonError};
use std::thread::JoinHandle;

use pdf_core::SharedFont;

use crate::{PDFLAPAS_ERROR, PDFLAPAS_OK, PdflapasBuffer, caller_bytes, owned, panic_detail, report_json};

/// A render thread's stack: the size of a process's main-thread stack on Linux, which the single renderer runs on.
const THREAD_STACK_BYTES: usize = 8 * 1024 * 1024;
/// `report_json` of a render that panicked before it could report.
const EMPTY_REPORT: &str = r#"{"warnings":[],"missingFonts":[]}"#;

/// Called from a render thread when the completion queue goes from drained to not.
pub type Wake = extern "C" fn();

struct Job {
    id: u64,
    html: String,
}

struct Done {
    id: u64,
    pdf: Result<Vec<u8>, String>,
    report: String,
}

struct Completions {
    queue: VecDeque<Done>,
    /// Set when `wake` has rung and the caller has not yet drained the queue empty.
    rung: bool,
}

/// What the render threads share.
struct Shared {
    completions: Mutex<Completions>,
    /// Set by `pdflapas_pool_free`: threads skip the jobs still queued.
    closing: AtomicBool,
    wake: Option<Wake>,
}

/// The opaque pool handle.
pub struct PdflapasPool {
    threads: u32,
    fonts: Vec<SharedFont>,
    shared: Arc<Shared>,
    jobs: Option<Sender<Job>>,
    workers: Vec<JoinHandle<()>>,
}

/// One finished job, as `pdflapas_pool_next` writes it. `status` is `PDFLAPAS_OK` with the PDF in `pdf`, or
/// `PDFLAPAS_ERROR` with the message in `error`; `report` is `pdflapas_report`'s JSON for that render either way.
/// The caller releases the three buffers with `pdflapas_buffer_free`.
#[repr(C)]
pub struct PdflapasCompletion {
    pub id: u64,
    pub status: u64,
    pub pdf: PdflapasBuffer,
    pub report: PdflapasBuffer,
    pub error: PdflapasBuffer,
}

fn renderer_with(fonts: &[SharedFont]) -> pdf_core::Renderer {
    let mut renderer = pdf_core::Renderer::new();
    for font in fonts {
        renderer.add_shared_font(font.clone());
    }
    renderer
}

/// A render thread: takes jobs until the channel closes.
fn work(fonts: Vec<SharedFont>, jobs: Arc<Mutex<Receiver<Job>>>, shared: Arc<Shared>) {
    let mut renderer = renderer_with(&fonts);
    loop {
        let next = jobs.lock().unwrap_or_else(PoisonError::into_inner).recv();
        let Ok(job) = next else {
            return;
        };
        if shared.closing.load(Ordering::Relaxed) {
            continue;
        }
        let done = match catch_unwind(AssertUnwindSafe(|| {
            let pdf = renderer.render_html(&job.html);
            (pdf, report_json(&renderer))
        })) {
            Ok((pdf, report)) => Done { id: job.id, pdf, report },
            Err(payload) => {
                // The panic may have left the renderer half-written: this job fails, the next gets a fresh one.
                renderer = renderer_with(&fonts);
                Done {
                    id: job.id,
                    pdf: Err(format!(
                        "the native renderer panicked: {}; this document failed, the pool goes on (report the document that caused it)",
                        panic_detail(payload.as_ref())
                    )),
                    report: EMPTY_REPORT.to_owned(),
                }
            }
        };
        let ring = {
            let mut completions = shared.completions.lock().unwrap_or_else(PoisonError::into_inner);
            completions.queue.push_back(done);
            !std::mem::replace(&mut completions.rung, true)
        };
        if let Some(wake) = shared.wake.filter(|_| ring) {
            wake();
        }
    }
}

/// A pool of `threads` render threads (not started yet), or null for 0 threads. `wake` may be null (the caller polls).
#[unsafe(no_mangle)]
pub extern "C" fn pdflapas_pool_new(threads: u32, wake: Option<Wake>) -> *mut PdflapasPool {
    if threads == 0 {
        return ptr::null_mut();
    }
    Box::into_raw(Box::new(PdflapasPool {
        threads,
        fonts: Vec::new(),
        shared: Arc::new(Shared {
            completions: Mutex::new(Completions {
                queue: VecDeque::new(),
                rung: false,
            }),
            closing: AtomicBool::new(false),
            wake,
        }),
        jobs: None,
        workers: Vec::new(),
    }))
}

/// Reads a TTF/OTF/TTC file (copied once, shared by every thread) and returns its number of faces; 0 for a file with
/// no usable face, a null handle or pointer, a pool already started, or a panic.
///
/// # Safety
/// `pool` is null or from `pdflapas_pool_new`; a non-null `data` points to `len` readable bytes.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_pool_add_font(pool: *mut PdflapasPool, data: *const u8, len: usize) -> u32 {
    // SAFETY: the caller guarantees `pool` is null or live.
    let Some(pool) = (unsafe { pool.as_mut() }) else {
        return 0;
    };
    // SAFETY: the caller guarantees `len` readable bytes at a non-null `data`.
    let Ok(bytes) = (unsafe { caller_bytes(data, len) }) else {
        return 0;
    };
    if pool.jobs.is_some() {
        return 0;
    }
    let font = SharedFont::from(bytes.to_vec());
    // Counted on a throwaway renderer: the threads register the font when they start.
    let faces = catch_unwind(AssertUnwindSafe(|| pdf_core::Renderer::new().add_shared_font(font.clone()))).unwrap_or(0);
    if faces > 0 {
        pool.fonts.push(font);
    }
    faces
}

/// Starts the threads, each registering the pool's fonts. `PDFLAPAS_ERROR` when it is null, already started, or the
/// system refuses a thread (those already started stop with the pool).
///
/// # Safety
/// `pool` is null or from `pdflapas_pool_new`.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_pool_start(pool: *mut PdflapasPool) -> i32 {
    // SAFETY: the caller guarantees `pool` is null or live.
    let Some(pool) = (unsafe { pool.as_mut() }) else {
        return PDFLAPAS_ERROR;
    };
    if pool.jobs.is_some() {
        return PDFLAPAS_ERROR;
    }
    let (sender, receiver) = channel();
    let receiver = Arc::new(Mutex::new(receiver));
    pool.jobs = Some(sender);
    for index in 0..pool.threads {
        let (fonts, jobs, shared) = (pool.fonts.clone(), Arc::clone(&receiver), Arc::clone(&pool.shared));
        let spawned = std::thread::Builder::new()
            .name(format!("pdflapas-render-{index}"))
            .stack_size(THREAD_STACK_BYTES)
            .spawn(move || work(fonts, jobs, shared));
        match spawned {
            Ok(worker) => pool.workers.push(worker),
            Err(_) => return PDFLAPAS_ERROR,
        }
    }
    PDFLAPAS_OK
}

/// Queues a PdfLapas HTML profile document (UTF-8) as job `id`; its result comes out of `pdflapas_pool_next`.
/// `PDFLAPAS_ERROR` when the pool is null or not started, or the HTML is not UTF-8.
///
/// # Safety
/// `pool` is null or from `pdflapas_pool_new`; a non-null `html` points to `len` readable bytes.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_pool_submit(pool: *mut PdflapasPool, id: u64, html: *const u8, len: usize) -> i32 {
    // SAFETY: the caller guarantees `pool` is null or live.
    let Some(pool) = (unsafe { pool.as_ref() }) else {
        return PDFLAPAS_ERROR;
    };
    // SAFETY: the caller guarantees `len` readable bytes at a non-null `html`.
    let Ok(text) = (unsafe { caller_bytes(html, len) }).and_then(|bytes| std::str::from_utf8(bytes).map_err(|error| error.to_string())) else {
        return PDFLAPAS_ERROR;
    };
    match &pool.jobs {
        Some(jobs) if jobs.send(Job { id, html: text.to_owned() }).is_ok() => PDFLAPAS_OK,
        _ => PDFLAPAS_ERROR,
    }
}

/// Writes the oldest finished job to `out` and returns 1, or returns 0 when none is waiting; a 0 re-arms `wake`.
///
/// # Safety
/// `pool` is null or from `pdflapas_pool_new`; `out` is null or a valid, writable `PdflapasCompletion`.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_pool_next(pool: *mut PdflapasPool, out: *mut PdflapasCompletion) -> i32 {
    // SAFETY: the caller guarantees `pool` is null or live.
    let Some(pool) = (unsafe { pool.as_ref() }) else {
        return 0;
    };
    if out.is_null() {
        return 0;
    }
    let next = {
        let mut completions = pool.shared.completions.lock().unwrap_or_else(PoisonError::into_inner);
        let next = completions.queue.pop_front();
        // Re-armed under the lock, so a job finishing after this drain rings again.
        if next.is_none() {
            completions.rung = false;
        }
        next
    };
    let Some(done) = next else {
        return 0;
    };
    let (status, pdf, error) = match done.pdf {
        Ok(pdf) => (PDFLAPAS_OK, pdf, Vec::new()),
        Err(message) => (PDFLAPAS_ERROR, Vec::new(), message.into_bytes()),
    };
    // SAFETY: `out` is non-null and the caller guarantees it is writable.
    unsafe {
        out.write(PdflapasCompletion {
            id: done.id,
            status: u64::from(status.unsigned_abs()),
            pdf: owned(pdf),
            report: owned(done.report.into_bytes()),
            error: owned(error),
        });
    }
    1
}

/// Stops the pool: jobs still queued are dropped unrendered, renders in flight finish, then the threads are joined
/// (so `wake` is never called after this returns) and everything is released. Null is a no-op.
///
/// # Safety
/// `pool` is null or from `pdflapas_pool_new`, released once and not used afterwards.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn pdflapas_pool_free(pool: *mut PdflapasPool) {
    if pool.is_null() {
        return;
    }
    // SAFETY: the caller guarantees `pool` came from `Box::into_raw` in `pdflapas_pool_new`.
    let mut pool = unsafe { Box::from_raw(pool) };
    pool.shared.closing.store(true, Ordering::Relaxed);
    // Closing the channel ends each thread's loop once the queue is empty.
    pool.jobs = None;
    for worker in pool.workers.drain(..) {
        // A thread only ends by returning: every render runs under catch_unwind.
        let _ = worker.join();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pdflapas_buffer_free;

    fn take(buffer: &PdflapasBuffer) -> String {
        let text = String::from_utf8(unsafe { std::slice::from_raw_parts(buffer.ptr, buffer.len) }.to_vec()).unwrap();
        unsafe { pdflapas_buffer_free(buffer.ptr, buffer.len) };
        text
    }

    #[test]
    fn every_job_completes_once_with_its_own_result() {
        let pool = pdflapas_pool_new(3, None);

        assert_eq!(unsafe { pdflapas_pool_start(pool) }, PDFLAPAS_OK);
        assert_eq!(unsafe { pdflapas_pool_start(pool) }, PDFLAPAS_ERROR);
        let html = b"<p>not a profile document</p>";
        for id in 0..20 {
            assert_eq!(unsafe { pdflapas_pool_submit(pool, id, html.as_ptr(), html.len()) }, PDFLAPAS_OK);
        }
        assert_eq!(unsafe { pdflapas_pool_submit(pool, 99, [0xff_u8].as_ptr(), 1) }, PDFLAPAS_ERROR);
        let mut seen = Vec::new();
        let mut out = PdflapasCompletion {
            id: 0,
            status: 0,
            pdf: PdflapasBuffer { ptr: ptr::null_mut(), len: 0 },
            report: PdflapasBuffer { ptr: ptr::null_mut(), len: 0 },
            error: PdflapasBuffer { ptr: ptr::null_mut(), len: 0 },
        };
        while seen.len() < 20 {
            if unsafe { pdflapas_pool_next(pool, &mut out) } == 1 {
                assert_eq!(out.status, 1);
                assert!(take(&out.error).starts_with("invalid PdfLapas HTML"));
                assert_eq!(take(&out.report), EMPTY_REPORT);
                assert_eq!(take(&out.pdf), "");
                seen.push(out.id);
            } else {
                std::thread::yield_now();
            }
        }
        seen.sort_unstable();
        assert_eq!(seen, (0..20).collect::<Vec<_>>());
        unsafe { pdflapas_pool_free(pool) };
    }

    #[test]
    fn bad_arguments_are_errors_not_crashes() {
        assert!(pdflapas_pool_new(0, None).is_null());
        let pool = pdflapas_pool_new(1, None);

        assert_eq!(unsafe { pdflapas_pool_submit(pool, 1, b"x".as_ptr(), 1) }, PDFLAPAS_ERROR);
        assert_eq!(unsafe { pdflapas_pool_add_font(pool, b"not a font".as_ptr(), 10) }, 0);
        assert_eq!(unsafe { pdflapas_pool_add_font(ptr::null_mut(), b"x".as_ptr(), 1) }, 0);
        assert_eq!(unsafe { pdflapas_pool_next(pool, ptr::null_mut()) }, 0);
        assert_eq!(unsafe { pdflapas_pool_start(ptr::null_mut()) }, PDFLAPAS_ERROR);
        unsafe { pdflapas_pool_free(pool) };
        unsafe { pdflapas_pool_free(ptr::null_mut()) };
    }
}
