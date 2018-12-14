use ::std::collections::{ HashMap };
use ::std::path::{Path, PathBuf};
use ::std::ffi::{OsStr};
use ::std::fs::File;
use ::zip::ZipArchive;
use ::docprops::{AppInfo, Core};


/// Document
pub struct PPTXDocument {
    pub file_path: PathBuf,
    pub app: Option<Box<AppInfo>>,
    pub core: Option<Box<Core>>,
    pub presentation: Option<Box<::pml::Presentation>>,
    pub theme_map: HashMap<PathBuf, Box<::drawingml::OfficeStyleSheet>>,
    pub slide_master_map: HashMap<PathBuf, Box<::pml::SlideMaster>>,
    pub slide_layout_map: HashMap<PathBuf, Box<::pml::SlideLayout>>,
    pub slide_map: HashMap<PathBuf, Box<::pml::Slide>>,
    pub medias: Vec<PathBuf>,
}

impl PPTXDocument {
    pub fn from_file(pptx_path: &Path) -> Result<Self, Box<::std::error::Error>> {
        let pptx_file = File::open(&pptx_path)?;
        let mut zipper = ZipArchive::new(&pptx_file)?;

        println!("parsing docProps/app.xml");
        let app = AppInfo::from_zip(&mut zipper).map(|val| val.into()).ok();
        println!("parsing docProps/core.xml");
        let core = Core::from_zip(&mut zipper).map(|val| val.into()).ok();
        println!("parsing ppt/presentation.xml");
        let presentation = ::pml::Presentation::from_zip(&mut zipper).map(|val| val.into()).ok();
        let mut theme_map = HashMap::new();
        let mut slide_master_map = HashMap::new();
        let mut slide_layout_map = HashMap::new();
        let mut slide_map = HashMap::new();
        let mut medias = Vec::new();

        println!();

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
                match ::drawingml::OfficeStyleSheet::from_zip_file(&mut zip_file).map(|val| val.into()) {
                    Ok(theme) => {
                        theme_map.insert(file_path, theme);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/slideMasters") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "xml" {
                    continue;
                }

                println!("parsing slide master file: {}", zip_file.name());

                match ::pml::SlideMaster::from_zip_file(&mut zip_file).map(|val| val.into()) {
                    Ok(slide) => {
                        slide_master_map.insert(file_path, slide);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/slideLayouts") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "xml" {
                    continue;
                }

                println!("parsing slide layout file: {}", zip_file.name());

                match ::pml::SlideLayout::from_zip_file(&mut zip_file).map(|val| val.into()) {
                    Ok(slide) => {
                        slide_layout_map.insert(file_path, slide);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/slides") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "xml" {
                    continue;
                }

                println!("parsing slide file: {}", zip_file.name());

                match ::pml::Slide::from_zip_file(&mut zip_file).map(|val| val.into()) {
                    Ok(slide) => {
                        slide_map.insert(file_path, slide);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/media") {
                medias.push(file_path);
            }
        }

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