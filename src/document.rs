use ::std::collections::{ HashMap };
use ::std::path::{Path, PathBuf};
use ::std::fs::File;
use ::std::io::Read;
use ::zip::ZipArchive;
use ::docprops::{AppInfo, Core};
use ::xml::XmlNode;


/// Document
pub struct Document {
    pub file_path: PathBuf,
    pub app: Option<AppInfo>,
    pub core: Option<Core>,
    pub presentation: Option<::pml::Presentation>,
    pub theme_map: HashMap<PathBuf, ::drawingml::OfficeStyleSheet>,
    pub slide_master_map: HashMap<PathBuf, ::pml::SlideMaster>,
    pub slide_layout_map: HashMap<PathBuf, ::pml::SlideLayout>,
    pub slide_map: HashMap<PathBuf, ::pml::Slide>,
    pub medias: Vec<PathBuf>,
}

impl Document {
    pub fn from_file(pptx_path: &Path) -> Result<Document, Box<::std::error::Error>> {
        let pptx_file = File::open(&pptx_path)?;
        let mut zipper = ZipArchive::new(&pptx_file)?;

        let app = AppInfo::from_zip(&mut zipper).ok();
        let core = Core::from_zip(&mut zipper).ok();
        println!("parsing presentation.xml");
        let presentation = ::pml::Presentation::from_zip(&mut zipper).ok();
        let mut theme_map = HashMap::new();
        let mut slide_master_map = HashMap::new();
        let mut slide_layout_map = HashMap::new();
        let mut slide_map = HashMap::new();
        let mut medias = Vec::new();

        println!("");

        for i in 0..zipper.len() {
            let mut zip_file = match zipper.by_index(i) {
                Ok(f) => f,
                Err(err) => {
                    println!("Failed to get zip file by index. Index: {}, error: {}", i, err);
                    continue;
                }
            };

            let file_path = PathBuf::from(zip_file.name());
            if file_path.starts_with("ppt/theme") {
                println!("parsing theme file: {}", zip_file.name());
                match ::drawingml::OfficeStyleSheet::from_zip_file(&mut zip_file) {
                    Ok(theme) => {
                        theme_map.insert(file_path, theme);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/media") {
                medias.push(file_path);
            }
        }

        //::drawingml::OfficeStyleSheet::

        Ok(Self {
            file_path: PathBuf::from(pptx_path),
            app,
            core,
            presentation,
            theme_map,
            slide_master_map,
            slide_layout_map,
            slide_map,
            medias,
        })
    }
}