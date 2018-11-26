use ::std::collections::{ HashMap };
use ::std::path::{Path, PathBuf};
use ::std::fs::File;
use ::std::io::Read;
use ::zip::ZipArchive;
use ::docprops::{AppInfo, Core};
use ::xml::XmlNode;


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
        println!("parsing presentation.xml");
        document.presentation = match ::pml::Presentation::from_zip(&mut zipper) {
            Ok(p) => Some(p),
            Err(err) => {
                println!("Failed to load presentation.xml. Error: {}", err);
                None
            }
        };

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
                let mut xml_string = String::new();
                if let Err(err) = zip_file.read_to_string(&mut xml_string) {
                    println!("{}", err);
                    continue;
                }

                let xml_node = match XmlNode::from_str(xml_string.as_str()) {
                    Ok(node) => node,
                    Err(err) => {
                        println!("{}", err);
                        continue;
                    }
                };

                match ::drawingml::OfficeStyleSheet::from_xml_element(&xml_node) {
                    Ok(v) => {
                        document.theme_map.insert(file_path, v);
                    }
                    Err(err) => println!("{}", err),
                }
            }
        }

        //::drawingml::OfficeStyleSheet::

        Ok(document)
    }
}