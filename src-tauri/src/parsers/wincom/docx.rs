//! Windows COM(Word.Application) 기반 DOCX 파서 — Windows 전용.
//!
//! `parsers::docx` (Rust ZIP/XML 파서) 실패 시 `wincom_fallback_docx` 에서 호출.
//! 설치된 Word 로 문서를 열어 `Content.Text` 를 추출한다.

use std::path::{Path, PathBuf};
use std::sync::{mpsc, Mutex, OnceLock};
use std::time::Duration;

use super::{
    quit_if_idle, to_parse_error, v_string, var_bool, var_i32, var_str, ComApartment, Obj,
    ScopedAppSettings,
};
use crate::parsers::{
    DocumentChunk, DocumentMetadata, ParseError, ParsedDocument, DEFAULT_CHUNK_OVERLAP,
    DEFAULT_CHUNK_SIZE, MAX_FILE_SIZE,
};

/// `WdSaveOptions::wdDoNotSaveChanges` — Close 시 저장하지 않음.
const WD_DO_NOT_SAVE_CHANGES: i32 = 0;
/// `WdAlertLevel::wdAlertsNone` — 경고 대화상자 비활성.
const WD_ALERTS_NONE: i32 = 0;

/// DOCX 파일을 COM(Word.Application) 으로 파싱.
///
/// `parse_with_timeout` 래퍼를 통해 별도 STA 스레드에서 실행되며,
/// 타임아웃·패닉 방어가 적용된다.
pub fn parse(path: &Path) -> Result<ParsedDocument, ParseError> {
    // 파일 크기 체크 (대용량 파일 메모리 보호)
    if let Ok(metadata) = std::fs::metadata(path) {
        if metadata.len() > MAX_FILE_SIZE {
            return Err(ParseError::ParseError(format!(
                "DOCX 파일 크기 초과: {}MB (최대 {}MB)",
                metadata.len() / 1024 / 1024,
                MAX_FILE_SIZE / 1024 / 1024
            )));
        }
    }

    let full_text = request_text(path)?;
    let pages = split_pages(&full_text);
    if pages.is_empty() {
        tracing::warn!("DOCX(COM) file has no text content: {:?}", path);
    }
    let total_text = pages
        .iter()
        .map(|p| p.text.as_str())
        .collect::<Vec<_>>()
        .join("\n");
    let chunks = chunk_pages(&pages, DEFAULT_CHUNK_SIZE, DEFAULT_CHUNK_OVERLAP);
    let page_count = pages.len();

    Ok(ParsedDocument {
        content: total_text,
        metadata: DocumentMetadata {
            title: path.file_stem().and_then(|s| s.to_str()).map(String::from),
            author: None,
            created_at: None,
            page_count: if page_count > 1 {
                Some(page_count)
            } else {
                None
            },
        },
        chunks,
        garbled_hint: false,
    })
}

/// Word 를 COM 자동화로 구동해 문서 전체 텍스트 추출.
///
/// Application 객체 생성 → 보안 설정(Visible/DisplayAlerts/ScreenUpdating/AutomationSecurity)
/// → `Documents.Open` → `Content.Text` 읽기 → `Close` 순서로 동작.
/// Open 시 가짜 암호를 전달해 보호된 문서의 모달 다이얼로그를 차단한다.
fn extract_text_with_app(app: &Obj, path: &Path) -> windows::core::Result<String> {
    let mut settings = ScopedAppSettings::new(app);
    settings.put("Visible", var_bool(false));
    settings.put("DisplayAlerts", var_i32(WD_ALERTS_NONE));
    settings.put("ScreenUpdating", var_bool(false));
    settings.put(
        "AutomationSecurity",
        var_i32(super::MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
    );

    let documents = app.get_obj("Documents", &[])?;
    let path_str = super::path_arg(path);
    let doc_var = documents.call(
        "Open",
        &[
            var_str(&path_str),
            var_bool(false),
            var_bool(true),
            var_bool(false),
            var_str(super::BOGUS_PASSWORD),
        ],
    )?;
    let doc = super::as_obj(&doc_var)?;

    let result = (|| {
        let content = doc.get_obj("Content", &[])?;
        let text_var = content.get("Text", &[])?;
        Ok(v_string(&text_var))
    })();

    // Always close the document even when extraction failed.
    let _ = doc.call("Close", &[var_i32(WD_DO_NOT_SAVE_CHANGES)]);
    result
}

struct WordRequest {
    path: PathBuf,
    reply: mpsc::Sender<Result<String, ParseError>>,
}

static WORD_WORKER: OnceLock<Mutex<Option<mpsc::Sender<WordRequest>>>> = OnceLock::new();

fn word_worker_slot() -> &'static Mutex<Option<mpsc::Sender<WordRequest>>> {
    WORD_WORKER.get_or_init(|| Mutex::new(None))
}

fn start_word_worker() -> mpsc::Sender<WordRequest> {
    let (tx, rx) = mpsc::channel::<WordRequest>();
    std::thread::Builder::new()
        .name("anything-word-com".into())
        .spawn(move || {
            let _com = ComApartment::init();
            let app = match Obj::create("Word.Application") {
                Ok(app) => app,
                Err(e) => {
                    let err = to_parse_error("Word.Application", &e);
                    while let Ok(req) = rx.recv() {
                        let _ = req.reply.send(Err(ParseError::ParseError(err.to_string())));
                    }
                    return;
                }
            };

            tracing::info!("Word COM persistent session started");
            while let Ok(req) = rx.recv() {
                let result =
                    extract_text_with_app(&app, &req.path).map_err(|e| to_parse_error("Word", &e));
                let _ = req.reply.send(result);
            }
            quit_if_idle(&app, "Documents");
            tracing::info!("Word COM persistent session stopped");
        })
        .expect("failed to start Word COM worker");
    tx
}

fn request_text(path: &Path) -> Result<String, ParseError> {
    let (reply_tx, reply_rx) = mpsc::channel();
    let req = WordRequest {
        path: path.to_path_buf(),
        reply: reply_tx,
    };

    let sender = {
        let mut slot = word_worker_slot().lock().unwrap_or_else(|e| e.into_inner());
        if slot.is_none() {
            *slot = Some(start_word_worker());
        }
        slot.as_ref().unwrap().clone()
    };

    if sender.send(req).is_err() {
        let mut slot = word_worker_slot().lock().unwrap_or_else(|e| e.into_inner());
        *slot = None;
        return Err(ParseError::ParseError(
            "Word COM 세션이 종료되었습니다. 다음 파일에서 자동 재시작합니다.".into(),
        ));
    }

    match reply_rx.recv_timeout(Duration::from_secs(25)) {
        Ok(result) => result,
        Err(_) => {
            // Do not queue more work behind a hung COM call.  Detach this
            // worker; the next request starts a fresh STA/Application session.
            let mut slot = word_worker_slot().lock().unwrap_or_else(|e| e.into_inner());
            *slot = None;
            Err(ParseError::ParseError(format!(
                "Word COM 세션 응답 타임아웃 (25초): {}",
                path.display()
            )))
        }
    }
}

/// 페이지별 텍스트 정보 (페이지 번호 + 텍스트 + 전문 기준 시작 오프셋).
struct PageText {
    page_number: usize,
    text: String,
    start_offset: usize,
}

/// 폼 피드(`\u{000C}`) 기준 페이지 분할.
/// Word `Content.Text` 는 페이지 경계를 폼 피드 문자로 구분한다.
fn split_pages(full: &str) -> Vec<PageText> {
    let mut pages: Vec<PageText> = Vec::new();
    let mut running_offset = 0usize;

    for (idx, raw_page) in full.split('\u{000C}').enumerate() {
        let text = normalize_paragraphs(raw_page);
        if text.is_empty() {
            continue;
        }
        let char_len = text.chars().count();
        pages.push(PageText {
            page_number: idx + 1,
            text,
            start_offset: running_offset,
        });
        running_offset += char_len + 1;
    }

    pages
}

/// 단락 구분자 정규화 — CR/LF/수직탭/벨/Shift-In 등을 `\n`으로 통일하고
/// 빈 줄을 제거한다.
fn normalize_paragraphs(page: &str) -> String {
    page.split(['\r', '\n', '\u{000B}', '\u{0007}', '\u{000E}'])
        .map(|line| line.trim())
        .filter(|line| !line.is_empty())
        .collect::<Vec<_>>()
        .join("\n")
}

/// 페이지별 청크 분할 (페이지 번호·위치 힌트 유지).
/// 각 페이지 내부에서 `chunk_size` 크기로 분할하고 `overlap` 만큼 겹친다.
fn chunk_pages(pages: &[PageText], chunk_size: usize, overlap: usize) -> Vec<DocumentChunk> {
    let mut chunks = Vec::new();

    for page in pages {
        let chars: Vec<char> = page.text.chars().collect();
        let total_len = chars.len();

        if total_len == 0 {
            continue;
        }

        let step = chunk_size.saturating_sub(overlap).max(1);
        let mut start = 0;

        while start < total_len {
            let end = (start + chunk_size).min(total_len);
            let chunk_content: String = chars[start..end].iter().collect();

            chunks.push(DocumentChunk {
                content: chunk_content,
                start_offset: page.start_offset + start,
                end_offset: page.start_offset + end,
                page_number: Some(page.page_number),
                page_end: Some(page.page_number),
                location_hint: Some(format!("페이지 {}", page.page_number)),
            });

            start += step;

            if end >= total_len {
                break;
            }
        }
    }

    chunks
}
