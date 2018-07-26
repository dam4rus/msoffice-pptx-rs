extern crate quick_xml;
extern crate zip;

#[macro_use]
mod macros;
pub mod docprops;
pub mod relationship;
pub mod pml;
pub mod drawingml;
pub mod document;

#[cfg(test)]
mod tests {
    use drawingml;
    use document::*;
    use docprops::*;
    use std::path::{ Path, PathBuf };

    #[test]
    fn test_sample_pptx() {
        let mut test_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));

        let mut sample_pptx_path = test_dir.join("tests/samplepptx.pptx");
        assert_eq!(sample_pptx_path.to_str().unwrap(), "/home/kalmar/coding/rust/office-rs/msoffice-pptx/tests/samplepptx.pptx");

        let document = match Document::from_file(&sample_pptx_path) {
            Ok(info) => info,
            Err(err) => panic!(err),
        };

        if let Some(app_info) = document.app {
            assert_eq!(app_info.app_name.unwrap(), "Microsoft Office PowerPoint");
            assert_eq!(app_info.app_version.unwrap(), "12.0000");
        }

        if let Some(core) = document.core {
            assert_eq!(core.title.unwrap(), "Sample PowerPoint File");
            assert_eq!(core.creator.unwrap(), "James Falkofske");
            assert_eq!(core.last_modified_by.unwrap(), "James Falkofske");
            assert_eq!(core.revision.unwrap(), 2);
            assert_eq!(core.created_time.unwrap(), "2009-05-06T22:06:09Z");
            assert_eq!(core.modified_time.unwrap(), "2009-05-06T22:13:30Z");
        }
    }
}
