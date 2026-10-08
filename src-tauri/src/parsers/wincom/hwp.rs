//! Windows Hancom COM(HWPFrame.HwpObject) fallback for DRM-protected HWP.
//!
//! Used only after Kordoc fails on an HWP in a way that indicates DRM / a foreign wrapper
//! (see `parsers::hwp_com_fallback_eligible`). It never copies, decrypts, registers bypass
//! modules, or auto-approves prompts.
//!
//! Order:
//! 1. **Attach (read-only)** to a Hancom instance the user already has open (Running Object
//!    Table). Used only when that instance's active document IS this file; only `Path`,
//!    `PageCount` and `GetPageText` are called. Never Open/Clear/Quit on it, never switch
//!    the user's active document.
//! 2. **Own instance**: create `HWPFrame.HwpObject` and open the ORIGINAL path through
//!    Hancom's normal automation/security flow (the user's consent prompt stays intact).
//!    A process snapshot proves the instance is a NEW Hwp.exe; if COM handed back a
//!    pre-existing (user) instance, nothing is opened/cleared/quit on it.
//!    Anything-owned instances are recorded (pid + creation time) and cleaned up on
//!    Quit, or at application shutdown if one is still alive.

use std::collections::HashSet;
use std::path::Path;
use std::sync::Mutex;

use windows::core::Interface;
use windows::Win32::System::Com::{
    CoTaskMemFree, CreateBindCtx, GetRunningObjectTable, IDispatch, IMoniker,
};

use super::{path_arg, to_parse_error, v_i32, v_string, var_i32, var_str, ComApartment, Obj};
use crate::parsers::{
    chunk_text, DocumentMetadata, ParseError, ParsedDocument, DEFAULT_CHUNK_OVERLAP,
    DEFAULT_CHUNK_SIZE, MAX_FILE_SIZE,
};

const HWP_EXE: &str = "hwp.exe";
/// ROT item name Hancom registers for a running HwpObject (e.g. `!HwpObject.130.1`).
const ROT_MARKER: &str = "hwpobject";

/// Hwp.exe processes started BY Anything's COM fallback: (pid, creation time FILETIME).
static OWNED: Mutex<Vec<(u32, u64)>> = Mutex::new(Vec::new());

pub fn parse(path: &Path) -> Result<ParsedDocument, ParseError> {
    if let Ok(metadata) = std::fs::metadata(path) {
        if metadata.len() > MAX_FILE_SIZE {
            return Err(ParseError::ParseError(format!(
                "HWP 파일 크기 초과: {}MB (최대 {}MB)",
                metadata.len() / 1024 / 1024,
                MAX_FILE_SIZE / 1024 / 1024
            )));
        }
    }

    let _com = ComApartment::init();

    tracing::info!("HWP COM attach/start attempt: {}", path.display());
    match attach_open_document(path) {
        Some(Ok(doc)) => return Ok(doc),
        Some(Err(e)) => tracing::warn!("HWP COM attach read failed, trying own instance: {e}"),
        None => tracing::info!(
            "HWP COM: document not active in a user Hancom instance; starting own instance"
        ),
    }
    parse_with_own_instance(path)
}

// ----------------------------------------------------------------------------
// 1. Read-only attach to an already-open document
// ----------------------------------------------------------------------------

/// `None` = no running Hancom instance has this file as its active document.
fn attach_open_document(path: &Path) -> Option<Result<ParsedDocument, ParseError>> {
    let target = normalize_path(&path_arg(path));
    for hwp in running_hwp_objects() {
        let active = normalize_path(&hwp.get_string("Path", &[]));
        if active.is_empty() || active != target {
            continue; // different (or no) active document: never switch the user's document
        }
        tracing::info!("HWP COM document matched (attached to open Hancom, read-only)");
        let result = extract(&hwp, path);
        drop(hwp); // release only - the user's Hancom stays exactly as it was
        return Some(result);
    }
    None
}

fn running_hwp_objects() -> Vec<Obj> {
    let mut out = Vec::new();
    unsafe {
        let (Ok(rot), Ok(ctx)) = (GetRunningObjectTable(0), CreateBindCtx(0)) else {
            return out;
        };
        let Ok(monikers) = rot.EnumRunning() else {
            return out;
        };
        loop {
            let mut slot: [Option<IMoniker>; 1] = [None];
            let mut fetched = 0u32;
            let hr = monikers.Next(&mut slot, Some(&mut fetched as *mut u32));
            if hr.is_err() || fetched == 0 {
                break;
            }
            let Some(moniker) = slot[0].take() else {
                break;
            };
            let Ok(name_ptr) = moniker.GetDisplayName(&ctx, None::<&IMoniker>) else {
                continue;
            };
            let name = name_ptr.to_string().unwrap_or_default();
            CoTaskMemFree(Some(name_ptr.0 as *const _));
            if !name.to_lowercase().contains(ROT_MARKER) {
                continue;
            }
            if let Ok(unknown) = rot.GetObject(&moniker) {
                if let Ok(disp) = unknown.cast::<IDispatch>() {
                    out.push(Obj(disp));
                }
            }
        }
    }
    out
}

pub(crate) fn normalize_path(p: &str) -> String {
    let p = p.trim();
    let p = p.strip_prefix(r"\\?\").unwrap_or(p);
    p.replace('/', "\\").to_lowercase()
}

// ----------------------------------------------------------------------------
// 2. Anything-owned instance
// ----------------------------------------------------------------------------

fn parse_with_own_instance(path: &Path) -> Result<ParsedDocument, ParseError> {
    let before = hwp_processes();
    let hwp = Obj::create("HWPFrame.HwpObject").map_err(|e| to_parse_error("Hancom HWP", &e))?;
    let started = new_processes(&before, &hwp_processes());
    let Some(owned) = single_owned(&started) else {
        // COM returned a pre-existing (user-owned) server or ownership is ambiguous:
        // do not open, clear, or quit anything in it.
        drop(hwp);
        return Err(ParseError::ParseError(format!(
            "Hancom COM 실행본의 소유를 확인할 수 없어 사용하지 않음 (새 Hwp.exe {}개)",
            started.len()
        )));
    };
    if let Ok(mut v) = OWNED.lock() {
        v.push(owned);
    }
    tracing::info!("HWP COM own instance started (pid {})", owned.0);

    let path_str = path_arg(path);
    let mut opened = false;
    let result = (|| {
        // Keep Hancom's own security prompt intact. The user must explicitly allow access.
        // No FilePathChecker module is registered and no temp copy is made.
        let open_result = hwp
            .call(
                "Open",
                &[
                    var_str(&path_str),
                    var_str(""),
                    var_str("suspendpassword:TRUE;forceopen:TRUE;versionwarning:FALSE"),
                ],
            )
            .map_err(|e| to_parse_error("Hancom HWP Open", &e))?;
        opened = v_i32(&open_result) != 0;
        if !opened {
            return Err(ParseError::ParseError(format!(
                "Hancom HWP COM이 문서를 열지 못했습니다: {}",
                path.display()
            )));
        }
        tracing::info!("HWP COM document matched (opened in own instance)");
        extract(&hwp, path)
    })();

    // Cleanup ONLY on the instance Anything started. Clear(1) means discard changes.
    if opened {
        let _ = hwp.call("Clear", &[var_i32(1)]);
    }
    let _ = hwp.call("Quit", &[]);
    drop(hwp);
    if wait_exit(owned, 5_000) {
        if let Ok(mut v) = OWNED.lock() {
            v.retain(|p| *p != owned);
        }
    } else {
        tracing::warn!(
            "HWP COM own instance still running after Quit (pid {}); cleaned up at shutdown",
            owned.0
        );
    }
    result
}

/// Exactly one new Hwp.exe appeared -> it is Anything's. Zero (existing server reused) or
/// several (user started Hancom at the same moment) -> not provably owned.
pub(crate) fn single_owned(started: &[(u32, u64)]) -> Option<(u32, u64)> {
    match started {
        [one] => Some(*one),
        _ => None,
    }
}

pub(crate) fn new_processes(before: &[(u32, u64)], after: &[(u32, u64)]) -> Vec<(u32, u64)> {
    let known: HashSet<(u32, u64)> = before.iter().copied().collect();
    after
        .iter()
        .copied()
        .filter(|p| !known.contains(p))
        .collect()
}

/// Shared read-only extraction (PageCount + GetPageText) - no document state is changed.
fn extract(hwp: &Obj, path: &Path) -> Result<ParsedDocument, ParseError> {
    let page_count = hwp.get_i32("PageCount", &[]).max(0) as usize;
    let mut pages = Vec::with_capacity(page_count);
    for page in 1..=page_count {
        let text_var = hwp
            .call("GetPageText", &[var_i32(page as i32), var_i32(0)])
            .map_err(|e| to_parse_error("Hancom HWP GetPageText", &e))?;
        let text = v_string(&text_var)
            .replace("\r\n", "\n")
            .replace('\r', "\n");
        if !text.trim().is_empty() {
            pages.push(text);
        }
    }

    let content = pages.join("\n\n");
    if content.trim().is_empty() {
        return Err(ParseError::ParseError(format!(
            "Hancom HWP COM 텍스트 추출 결과가 비어 있습니다: {}",
            path.display()
        )));
    }

    let chunks = chunk_text(&content, DEFAULT_CHUNK_SIZE, DEFAULT_CHUNK_OVERLAP);
    Ok(ParsedDocument {
        content,
        metadata: DocumentMetadata {
            title: path.file_stem().and_then(|s| s.to_str()).map(String::from),
            author: None,
            created_at: None,
            page_count: if page_count > 0 {
                Some(page_count)
            } else {
                None
            },
        },
        chunks,
        garbled_hint: false,
    })
}

// ----------------------------------------------------------------------------
// Process ownership (ToolHelp snapshot + creation time; guards against PID reuse)
// ----------------------------------------------------------------------------

fn hwp_processes() -> Vec<(u32, u64)> {
    use windows_sys::Win32::Foundation::{CloseHandle, INVALID_HANDLE_VALUE};
    use windows_sys::Win32::System::Diagnostics::ToolHelp::{
        CreateToolhelp32Snapshot, Process32FirstW, Process32NextW, PROCESSENTRY32W,
        TH32CS_SNAPPROCESS,
    };
    let mut out = Vec::new();
    unsafe {
        let snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if snap == INVALID_HANDLE_VALUE || snap.is_null() {
            return out;
        }
        let mut entry: PROCESSENTRY32W = std::mem::zeroed();
        entry.dwSize = std::mem::size_of::<PROCESSENTRY32W>() as u32;
        let mut ok = Process32FirstW(snap, &mut entry) != 0;
        while ok {
            let len = entry
                .szExeFile
                .iter()
                .position(|&c| c == 0)
                .unwrap_or(entry.szExeFile.len());
            let exe = String::from_utf16_lossy(&entry.szExeFile[..len]);
            if exe.eq_ignore_ascii_case(HWP_EXE) {
                if let Some(created) = creation_time(entry.th32ProcessID) {
                    out.push((entry.th32ProcessID, created));
                }
            }
            ok = Process32NextW(snap, &mut entry) != 0;
        }
        CloseHandle(snap);
    }
    out
}

fn creation_time(pid: u32) -> Option<u64> {
    use windows_sys::Win32::Foundation::{CloseHandle, FILETIME};
    use windows_sys::Win32::System::Threading::{
        GetProcessTimes, OpenProcess, PROCESS_QUERY_LIMITED_INFORMATION,
    };
    unsafe {
        let h = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, pid);
        if h.is_null() {
            return None;
        }
        let zero = FILETIME {
            dwLowDateTime: 0,
            dwHighDateTime: 0,
        };
        let (mut c, mut e, mut k, mut u) = (zero, zero, zero, zero);
        let ok = GetProcessTimes(h, &mut c, &mut e, &mut k, &mut u) != 0;
        CloseHandle(h);
        if ok {
            Some(((c.dwHighDateTime as u64) << 32) | c.dwLowDateTime as u64)
        } else {
            None
        }
    }
}

fn wait_exit(owned: (u32, u64), timeout_ms: u64) -> bool {
    let deadline = std::time::Instant::now() + std::time::Duration::from_millis(timeout_ms);
    loop {
        if !hwp_processes().contains(&owned) {
            return true;
        }
        if std::time::Instant::now() >= deadline {
            return false;
        }
        std::thread::sleep(std::time::Duration::from_millis(200));
    }
}

/// Application shutdown: end only Hwp.exe instances that Anything itself started and that are
/// still alive (pid AND creation time must match - a user's Hancom is never touched).
pub(crate) fn shutdown_owned() {
    use windows_sys::Win32::Foundation::CloseHandle;
    use windows_sys::Win32::System::Threading::{OpenProcess, TerminateProcess, PROCESS_TERMINATE};
    let owned: Vec<(u32, u64)> = OWNED
        .lock()
        .map(|mut v| std::mem::take(&mut *v))
        .unwrap_or_default();
    if owned.is_empty() {
        return;
    }
    let alive = hwp_processes();
    for p in owned.into_iter().filter(|p| alive.contains(p)) {
        unsafe {
            let h = OpenProcess(PROCESS_TERMINATE, 0, p.0);
            if !h.is_null() {
                let _ = TerminateProcess(h, 0);
                CloseHandle(h);
            }
        }
        tracing::info!("HWP COM own instance ended at shutdown (pid {})", p.0);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn module_compiles_without_hancom_runtime() {
        // Unit tests must not require Hancom to be installed or launch COM.
        let _ = parse as fn(&Path) -> Result<ParsedDocument, ParseError>;
    }

    #[test]
    fn ownership_requires_exactly_one_new_hwp_process() {
        let user = (8392, 111);
        // COM reused the user's Hancom -> not ours
        assert_eq!(single_owned(&new_processes(&[user], &[user])), None);
        let mine = (9000, 222);
        assert_eq!(single_owned(&new_processes(&[user], &[user, mine])), Some(mine));
        // ambiguous (two new processes) -> not ours
        assert_eq!(single_owned(&new_processes(&[], &[(1, 1), (2, 2)])), None);
        // same pid, different creation time = a different process (pid reuse) -> new
        assert_eq!(new_processes(&[(9000, 1)], &[(9000, 2)]), vec![(9000, 2)]);
    }

    #[test]
    fn shutdown_without_owned_instances_touches_nothing() {
        OWNED.lock().unwrap().clear();
        shutdown_owned(); // no-op: never enumerates or terminates user processes
        assert!(OWNED.lock().unwrap().is_empty());
    }

    #[test]
    fn path_matching_is_normalized() {
        let long = normalize_path(r"\\?\C:\Docs\A.hwp");
        assert_eq!(long, normalize_path("c:/docs/a.HWP"));
        assert_ne!(long, normalize_path(r"C:\Docs\B.hwp"));
    }
}
