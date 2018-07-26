use std::collections::{ HashMap };
use std::path::{ Path, PathBuf };
use std::fs;
use std::io::{ Read };
use std::str;

use pml;
use drawingml;
use docprops::{ AppInfo, Core };

use quick_xml;

use zip;

/// Document
pub struct Document {
    pub app: Option<AppInfo>,
    pub core: Option<Core>,
    pub presentation: Option<pml::Presentation>,
    pub theme_map: HashMap<PathBuf, drawingml::OfficeStyleSheet>,
    pub slide_master_map: HashMap<PathBuf, pml::SlideMaster>,
    pub slide_layout_map: HashMap<PathBuf, pml::SlideLayout>,
    pub slide_map: HashMap<PathBuf, pml::Slide>,
    pub images: Vec<PathBuf>,
    pub videos: Vec<PathBuf>,
}

impl Document {
    fn new() -> Document {
        Document {
            app: None,
            core: None,
            presentation: None,
            theme_map: HashMap::new(),
            slide_master_map: HashMap::new(),
            slide_layout_map: HashMap::new(),
            slide_map: HashMap::new(),
            images: Vec::new(),
            videos: Vec::new(),
        }
    }

    pub fn from_file(pptx_path: &Path) -> Result<Document, String> {

        let pptx_file = match fs::File::open(&pptx_path) {
            Ok(f) => f,
            Err(err) => return Err(err.to_string()),
        };

        let mut zipper = match zip::ZipArchive::new(&pptx_file) {
            Ok(z) => z,
            Err(err) => return Err(err.to_string()),
        };

        let mut document = Document::new();

        document.app = AppInfo::from_zip(&mut zipper);
        document.core = Core::from_zip(&mut zipper);
        document.presentation = pml::Presentation;

        Ok(document)
    }
}