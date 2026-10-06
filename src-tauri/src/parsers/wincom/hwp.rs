//! Windows Hancom COM(HWPFrame.HwpObject) fallback for DRM-protected HWP.
//!
//! This path is used only after Kordoc explicitly reports DRM protection.
//! It opens the ORIGINAL path through Hancom's normal automation/security flow;
//! it does not copy, decrypt, register bypass modules, or auto-approve prompts.

use std::path::Path;

use super::{path_arg, to_parse_error, v_i32, v_string, var_i32, var_str, ComApartment, Obj};
use crate::parsers::{
    chunk_text, DocumentMetadata, ParseError, ParsedDocument, DEFAULT_CHUNK_OVERLAP,
    DEFAULT_CHUNK_SIZE, MAX_FILE_SIZE,
};

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
    let hwp = Obj::create("HWPFrame.HwpObject").map_err(|e| to_parse_error("Hancom HWP", &e))?;
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
                    var_str("suspendpassword:TRUE;versionwarning:FALSE"),
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
    })();

    // Best-effort cleanup on every path. Clear(1) means discard changes.
    if opened {
        let _ = hwp.call("Clear", &[var_i32(1)]);
    }
    let _ = hwp.call("Quit", &[]);

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn module_compiles_without_hancom_runtime() {
        // Unit tests must not require Hancom to be installed or launch COM.
        let _ = parse as fn(&Path) -> Result<ParsedDocument, ParseError>;
    }
}
