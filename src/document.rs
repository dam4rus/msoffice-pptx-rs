use ::std::collections::{ HashMap };
use ::std::path::{Path, PathBuf};
use ::std::fs::File;
use ::zip::ZipArchive;
use ::docprops::{AppInfo, Core};


/// Document
pub struct Document {
    pub app: Option<AppInfo>,
    pub core: Option<Core>,
    pub presentation: Option<::pml::Presentation>,
    pub theme_map: HashMap<PathBuf, ::drawingml::OfficeStyleSheet>,
    pub slide_master_map: HashMap<PathBuf, ::pml::SlideMaster>,
    pub slide_layout_map: HashMap<PathBuf, ::pml::SlideLayout>,
    pub slide_map: HashMap<PathBuf, ::pml::Slide>,
    pub images: Vec<PathBuf>,
    pub videos: Vec<PathBuf>,
}

impl Document {
    fn new() -> Self {
        Self {
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

    pub fn from_file(pptx_path: &Path) -> Result<Document, Box<::std::error::Error>> {
        let pptx_file = File::open(&pptx_path)?;
        let mut zipper = ZipArchive::new(&pptx_file)?;

        let mut document = Document::new();

        document.app = AppInfo::from_zip(&mut zipper);
        document.core = Core::from_zip(&mut zipper);
        document.presentation = match ::pml::Presentation::from_zip(&mut zipper) {
            Ok(p) => Some(p),
            Err(err) => {
                println!("Failed to load presentation.xml. Error: {}", err);
                None
            }
        };

        for i in 0..zipper.len() {
            let zip_file = match zipper.by_index(i) {
                Ok(f) => f,
                Err(err) => {
                    println!("Failed to get zip file by index. Index: {}, error: {}", i, err);
                    continue;
                }
            };

            if Path::new(zip_file.name()).starts_with("ppt/theme") {
                //let theme = ::drawingml::OfficeStyleSheet::
            }
        }

        //::drawingml::OfficeStyleSheet::

        Ok(document)
    }
}