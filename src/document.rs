use crate::docprops::{AppInfo, Core};
use std::collections::HashMap;
use std::fs::File;
use std::path::{Path, PathBuf};
use zip::ZipArchive;
use log::{info};

/// Document
pub struct PPTXDocument {
    pub file_path: PathBuf,
    pub app: Option<Box<AppInfo>>,
    pub core: Option<Box<Core>>,
    pub presentation: Option<Box<crate::pml::Presentation>>,
    pub theme_map: HashMap<PathBuf, Box<msoffice_shared::drawingml::OfficeStyleSheet>>,
    pub slide_master_map: HashMap<PathBuf, Box<crate::pml::SlideMaster>>,
    pub slide_layout_map: HashMap<PathBuf, Box<crate::pml::SlideLayout>>,
    pub slide_map: HashMap<PathBuf, Box<crate::pml::Slide>>,
    pub slide_master_rels_map: HashMap<PathBuf, Vec<msoffice_shared::relationship::Relationship>>,
    pub slide_layout_rels_map: HashMap<PathBuf, Vec<msoffice_shared::relationship::Relationship>>,
    pub slide_rels_map: HashMap<PathBuf, Vec<msoffice_shared::relationship::Relationship>>,
    pub medias: Vec<PathBuf>,
}

impl PPTXDocument {
    pub fn from_file(pptx_path: &Path) -> Result<Self, Box<dyn (::std::error::Error)>> {
        let pptx_file = File::open(&pptx_path)?;
        let mut zipper = ZipArchive::new(&pptx_file)?;

        info!("parsing docProps/app.xml");
        let app = AppInfo::from_zip(&mut zipper).map(|val| val.into()).ok();
        info!("parsing docProps/core.xml");
        let core = Core::from_zip(&mut zipper).map(|val| val.into()).ok();
        info!("parsing ppt/presentation.xml");
        let presentation = crate::pml::Presentation::from_zip(&mut zipper)
            .map(|val| val.into())
            .ok();
        let mut theme_map = HashMap::new();
        let mut slide_master_map = HashMap::new();
        let mut slide_layout_map = HashMap::new();
        let mut slide_map = HashMap::new();
        let mut slide_master_rels_map = HashMap::new();
        let mut slide_layout_rels_map = HashMap::new();
        let mut slide_rels_map = HashMap::new();
        let mut medias = Vec::new();

        for i in 0..zipper.len() {
            let mut zip_file = zipper.by_index(i)?;

            let file_path = PathBuf::from(zip_file.name());
            if file_path.starts_with("ppt/theme") {
                info!("parsing theme file: {}", zip_file.name());
                theme_map.insert(
                    file_path,
                    Box::new(msoffice_shared::drawingml::OfficeStyleSheet::from_zip_file(&mut zip_file)?),
                );
            } else if file_path.starts_with("ppt/slideMasters/_rels") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "rels" {
                    continue;
                }

                info!("parsing slide master relationship file: {}", zip_file.name());
                slide_master_rels_map.insert(
                    file_path,
                    msoffice_shared::relationship::relationships_from_zip_file(&mut zip_file)?,
                );
            } else if file_path.starts_with("ppt/slideMasters") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "xml" {
                    continue;
                }

                info!("parsing slide master file: {}", zip_file.name());
                slide_master_map.insert(
                    file_path,
                    Box::new(crate::pml::SlideMaster::from_zip_file(&mut zip_file)?),
                );
            } else if file_path.starts_with("ppt/slideLayouts/_rels") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "rels" {
                    continue;
                }

                info!("parsing slide layout relationship file: {}", zip_file.name());
                slide_layout_rels_map.insert(
                    file_path,
                    msoffice_shared::relationship::relationships_from_zip_file(&mut zip_file)?,
                );
            } else if file_path.starts_with("ppt/slideLayouts") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "xml" {
                    continue;
                }

                info!("parsing slide layout file: {}", zip_file.name());
                slide_layout_map.insert(
                    file_path,
                    Box::new(crate::pml::SlideLayout::from_zip_file(&mut zip_file)?),
                );
            } else if file_path.starts_with("ppt/slides/_rels") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "rels" {
                    continue;
                }

                info!("parsing slide relationship file: {}", zip_file.name());
                slide_rels_map.insert(
                    file_path,
                    msoffice_shared::relationship::relationships_from_zip_file(&mut zip_file)?,
                );
            } else if file_path.starts_with("ppt/slides") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "xml" {
                    continue;
                }

                info!("parsing slide file: {}", zip_file.name());
                slide_map.insert(file_path, Box::new(crate::pml::Slide::from_zip_file(&mut zip_file)?));
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
            slide_master_rels_map,
            slide_layout_rels_map,
            slide_rels_map,
            medias,
        })
    }
}

#[cfg(test)]
#[test]
fn test_sample_pptx() {
    simple_logger::init().unwrap();
    let test_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let sample_pptx_path = test_dir.join("tests/samplepptx.pptx");

    PPTXDocument::from_file(&sample_pptx_path).unwrap();
}
