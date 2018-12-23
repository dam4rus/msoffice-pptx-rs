use crate::docprops::{AppInfo, Core};
use std::collections::HashMap;
use std::fs::File;
use std::path::{Path, PathBuf};
use zip::ZipArchive;

/// Document
pub struct PPTXDocument {
    pub file_path: PathBuf,
    pub app: Option<Box<AppInfo>>,
    pub core: Option<Box<Core>>,
    pub presentation: Option<Box<crate::pml::Presentation>>,
    pub theme_map: HashMap<PathBuf, Box<crate::drawingml::OfficeStyleSheet>>,
    pub slide_master_map: HashMap<PathBuf, Box<crate::pml::SlideMaster>>,
    pub slide_layout_map: HashMap<PathBuf, Box<crate::pml::SlideLayout>>,
    pub slide_map: HashMap<PathBuf, Box<crate::pml::Slide>>,
    pub slide_master_rels_map: HashMap<PathBuf, Vec<crate::relationship::Relationship>>,
    pub slide_layout_rels_map: HashMap<PathBuf, Vec<crate::relationship::Relationship>>,
    pub slide_rels_map: HashMap<PathBuf, Vec<crate::relationship::Relationship>>,
    pub medias: Vec<PathBuf>,
}

impl PPTXDocument {
    pub fn from_file(pptx_path: &Path) -> Result<Self, Box<dyn (::std::error::Error)>> {
        let pptx_file = File::open(&pptx_path)?;
        let mut zipper = ZipArchive::new(&pptx_file)?;

        println!("parsing docProps/app.xml");
        let app = AppInfo::from_zip(&mut zipper).map(|val| val.into()).ok();
        println!("parsing docProps/core.xml");
        let core = Core::from_zip(&mut zipper).map(|val| val.into()).ok();
        println!("parsing ppt/presentation.xml");
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
                match crate::drawingml::OfficeStyleSheet::from_zip_file(&mut zip_file).map(|val| val.into()) {
                    Ok(theme) => {
                        theme_map.insert(file_path, theme);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/slideMasters/_rels") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "rels" {
                    continue;
                }

                println!("parsing slide master relationship file: {}", zip_file.name());
                match crate::relationship::relationships_from_zip_file(&mut zip_file) {
                    Ok(vec) => {
                        slide_master_rels_map.insert(file_path, vec);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/slideMasters") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "xml" {
                    continue;
                }

                println!("parsing slide master file: {}", zip_file.name());

                match crate::pml::SlideMaster::from_zip_file(&mut zip_file).map(|val| val.into()) {
                    Ok(slide) => {
                        slide_master_map.insert(file_path, slide);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/slideLayouts/_rels") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "rels" {
                    continue;
                }

                println!("parsing slide layout relationship file: {}", zip_file.name());
                match crate::relationship::relationships_from_zip_file(&mut zip_file) {
                    Ok(vec) => {
                        slide_layout_rels_map.insert(file_path, vec);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/slideLayouts") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "xml" {
                    continue;
                }

                println!("parsing slide layout file: {}", zip_file.name());

                match crate::pml::SlideLayout::from_zip_file(&mut zip_file).map(|val| val.into()) {
                    Ok(slide) => {
                        slide_layout_map.insert(file_path, slide);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/slides/_rels") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "rels" {
                    continue;
                }

                println!("parsing slide relationship file: {}", zip_file.name());
                match crate::relationship::relationships_from_zip_file(&mut zip_file) {
                    Ok(vec) => {
                        slide_rels_map.insert(file_path, vec);
                    }
                    Err(err) => println!("{}", err),
                }
            } else if file_path.starts_with("ppt/slides") {
                if file_path.extension().unwrap_or_else(|| "".as_ref()) != "xml" {
                    continue;
                }

                println!("parsing slide file: {}", zip_file.name());

                match crate::pml::Slide::from_zip_file(&mut zip_file).map(|val| val.into()) {
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
    let test_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let sample_pptx_path = test_dir.join("tests/samplepptx.pptx");

    let document = PPTXDocument::from_file(&sample_pptx_path).unwrap();

    if let Some(ref app_info) = document.app {
        assert_eq!(app_info.app_name.as_ref().unwrap(), "Microsoft Office PowerPoint");
        assert_eq!(app_info.app_version.as_ref().unwrap(), "12.0000");
    }

    if let Some(ref core) = document.core {
        assert_eq!(core.title.as_ref().unwrap(), "Sample PowerPoint File");
        assert_eq!(core.creator.as_ref().unwrap(), "James Falkofske");
        assert_eq!(core.last_modified_by.as_ref().unwrap(), "James Falkofske");
        assert_eq!(*core.revision.as_ref().unwrap(), 2);
        assert_eq!(core.created_time.as_ref().unwrap(), "2009-05-06T22:06:09Z");
        assert_eq!(core.modified_time.as_ref().unwrap(), "2009-05-06T22:13:30Z");
    }

    // presentation test
    if let Some(ref presentation) = document.presentation {
        let master_id = presentation.slide_master_id_list.get(0).unwrap();
        assert_eq!(master_id.id.unwrap(), 2147483684);
        assert_eq!(master_id.relationship_id, "rId1");

        let slide_id_0 = presentation.slide_id_list.get(0).unwrap();
        assert_eq!(slide_id_0.id, 256);
        assert_eq!(slide_id_0.relationship_id, "rId2");

        let slide_id_1 = presentation.slide_id_list.get(1).unwrap();
        assert_eq!(slide_id_1.id, 257);
        assert_eq!(slide_id_1.relationship_id, "rId3");

        let slide_size = presentation.slide_size.as_ref().unwrap();
        assert_eq!(slide_size.width, 9144000);
        assert_eq!(slide_size.height, 6858000);
        assert_eq!(
            *slide_size.size_type.as_ref().unwrap(),
            crate::pml::SlideSizeType::Screen4x3
        );

        let notes_size = presentation.notes_size.as_ref().unwrap();
        assert_eq!(notes_size.width, 6858000);
        assert_eq!(notes_size.height, 9144000);

        let def_text_style = presentation.default_text_style.as_ref().unwrap();
        let def_par_props = def_text_style.def_paragraph_props.as_ref().unwrap();
        assert_eq!(
            def_par_props
                .default_run_properties
                .as_ref()
                .unwrap()
                .language
                .as_ref()
                .unwrap(),
            "en-US"
        );

        let lvl1_ppr = def_text_style.lvl1_paragraph_props.as_ref().unwrap();
        assert_eq!(lvl1_ppr.margin_left.unwrap(), 0);
        assert_eq!(*lvl1_ppr.align.as_ref().unwrap(), crate::drawingml::TextAlignType::Left);
        assert_eq!(lvl1_ppr.default_tab_size.unwrap(), 914400);
        assert_eq!(lvl1_ppr.rtl.unwrap(), false);
        assert_eq!(lvl1_ppr.east_asian_line_break.unwrap(), true);
        assert_eq!(lvl1_ppr.latin_line_break.unwrap(), false);
        assert_eq!(lvl1_ppr.hanging_punctuations.unwrap(), true);

        let lvl1_def_rpr = lvl1_ppr.default_run_properties.as_ref().unwrap();
        assert_eq!(lvl1_def_rpr.font_size.unwrap(), 1800);
        assert_eq!(lvl1_def_rpr.kerning.unwrap(), 1200);

        assert_eq!(lvl1_def_rpr.latin_font.as_ref().unwrap().typeface, "+mn-lt");
        assert_eq!(lvl1_def_rpr.east_asian_font.as_ref().unwrap().typeface, "+mn-ea");
        assert_eq!(lvl1_def_rpr.complex_script_font.as_ref().unwrap().typeface, "+mn-cs");
    }

    for (path, theme) in &document.theme_map {
        assert_eq!("ppt/theme/theme1.xml", path.to_str().unwrap());

        // color scheme test
        let color_scheme = &theme.theme_elements.color_scheme;
        assert_eq!(color_scheme.name, "Default Design 1");

        match color_scheme.dark1 {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0x000066),
            _ => panic!("theme1 dk1 color type mismatch"),
        }

        match color_scheme.light1 {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0xFFFFFF),
            _ => panic!("theme1 lt1 color type mismatch"),
        }

        match color_scheme.dark2 {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0x003366),
            _ => panic!("theme1 dk2 color type mismatch"),
        }

        match color_scheme.light2 {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0xFFFFFF),
            _ => panic!("theme1 lt2 color type mismatch"),
        }

        match color_scheme.accent1 {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0x8EB3C8),
            _ => panic!("theme1 accent1 color type mismatch"),
        }

        match color_scheme.accent2 {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0x6F97B3),
            _ => panic!("theme1 accent2 color type mismatch"),
        }

        match color_scheme.accent3 {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0xAAADB8),
            _ => panic!("theme1 accent3 color type mismatch"),
        }

        match color_scheme.accent4 {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0xDADADA),
            _ => panic!("theme1 accent4 color type mismatch"),
        }

        match color_scheme.accent5 {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0xC6D6E0),
            _ => panic!("theme1 accent5 color type mismatch"),
        }

        match color_scheme.accent6 {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0x6488A2),
            _ => panic!("theme1 accent6 color type mismatch"),
        }

        match color_scheme.hyperlink {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0x556575),
            _ => panic!("theme1 hyperlink color type mismatch"),
        }

        match color_scheme.followed_hyperlink {
            crate::drawingml::Color::SRgbColor(ref clr) => assert_eq!(clr.value, 0x3D556F),
            _ => panic!("theme1 followhyperlink type mismatch"),
        }

        // font_scheme test
        let font_scheme = &theme.theme_elements.font_scheme;
        assert_eq!(font_scheme.name, "Default Design");

        let major_font = &font_scheme.major_font;
        assert_eq!(major_font.latin.typeface, "Tahoma");
        assert_eq!(major_font.east_asian.typeface, "");
        assert_eq!(major_font.complex_script.typeface, "");

        let minor_font = &font_scheme.minor_font;
        assert_eq!(minor_font.latin.typeface, "Tahoma");
        assert_eq!(minor_font.east_asian.typeface, "");
        assert_eq!(minor_font.complex_script.typeface, "");

        let format_scheme = &theme.theme_elements.format_scheme;
        assert_eq!(format_scheme.name.as_ref().unwrap(), "Office");

        // first fill style test
        let fill_style = &format_scheme.fill_style_list[0];
        match fill_style {
            crate::drawingml::FillProperties::SolidFill(ref color) => match color {
                crate::drawingml::Color::SchemeColor(ref clr) => {
                    assert_eq!(clr.value, crate::drawingml::SchemeColorVal::PlaceholderColor)
                }
                _ => panic!("fill[0] is invalid"),
            },
            _ => panic!("fill[0] is invalid"),
        }

        // second fill style test
        let fill_style = &format_scheme.fill_style_list[1];
        match fill_style {
            crate::drawingml::FillProperties::GradientFill(ref gradient) => {
                let stop = &gradient.gradient_stop_list[0];
                assert_eq!(stop.position, 0.0);
                match stop.color {
                    crate::drawingml::Color::SchemeColor(ref clr) => {
                        assert_eq!(clr.value, crate::drawingml::SchemeColorVal::PlaceholderColor);
                        match clr.color_transforms[0] {
                            crate::drawingml::ColorTransform::Tint(val) => assert_eq!(val, 50_000.0),
                            _ => panic!("color transform is not tint"),
                        }

                        match clr.color_transforms[1] {
                            crate::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 300_000.0),
                            _ => panic!("color transform is not satMod"),
                        }
                    }
                    _ => panic!("stop color is not schemeColor"),
                }

                let stop = &gradient.gradient_stop_list[1];
                assert_eq!(stop.position, 35_000.0);
                match stop.color {
                    crate::drawingml::Color::SchemeColor(ref clr) => {
                        assert_eq!(clr.value, crate::drawingml::SchemeColorVal::PlaceholderColor);
                        match clr.color_transforms[0] {
                            crate::drawingml::ColorTransform::Tint(val) => assert_eq!(val, 37_000.0),
                            _ => panic!("color transform is not tint"),
                        }

                        match clr.color_transforms[1] {
                            crate::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 300_000.0),
                            _ => panic!("color transform is not satMod"),
                        }
                    }
                    _ => panic!("stop color is not schemeColor"),
                }

                let stop = &gradient.gradient_stop_list[2];
                assert_eq!(stop.position, 100_000.0);
                match stop.color {
                    crate::drawingml::Color::SchemeColor(ref clr) => {
                        assert_eq!(clr.value, crate::drawingml::SchemeColorVal::PlaceholderColor);
                        match clr.color_transforms[0] {
                            crate::drawingml::ColorTransform::Tint(val) => assert_eq!(val, 15_000.0),
                            _ => panic!("color transform is not tint"),
                        }

                        match clr.color_transforms[1] {
                            crate::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 350_000.0),
                            _ => panic!("color transform is not satMod"),
                        }
                    }
                    _ => panic!("stop color is not schemeColor"),
                }

                match gradient.shade_properties {
                    Some(crate::drawingml::ShadeProperties::Linear(ref shade_props)) => {
                        assert_eq!(shade_props.angle, Some(16_200_000));
                        assert_eq!(shade_props.scaled, Some(true));
                    }
                    _ => panic!("shape is not linear"),
                }
            }
            _ => panic!("fill[1] is invalid"),
        }

        // outline style test
        let ln_style = &format_scheme.line_style_list[0];
        assert_eq!(ln_style.width, Some(9_525));
        assert_eq!(ln_style.cap, Some(crate::drawingml::LineCap::Flat));
        assert_eq!(ln_style.compound, Some(crate::drawingml::CompoundLine::Single));
        assert_eq!(ln_style.pen_alignment, Some(crate::drawingml::PenAlignment::Center));

        match ln_style.fill_properties {
            Some(crate::drawingml::LineFillProperties::SolidFill(ref clr)) => match clr {
                crate::drawingml::Color::SchemeColor(ref scheme_clr) => {
                    assert_eq!(scheme_clr.value, crate::drawingml::SchemeColorVal::PlaceholderColor);
                    match scheme_clr.color_transforms[0] {
                        crate::drawingml::ColorTransform::Shade(val) => assert_eq!(val, 95_000.0),
                        _ => panic!("ColorTransform is not Shade"),
                    }

                    match scheme_clr.color_transforms[1] {
                        crate::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 105_000.0),
                        _ => panic!("ColorTransform is not SatMode"),
                    }
                }
                _ => panic!("Scheme color is not PlaceholderColor"),
            },
            _ => panic!("ln_style.fill_properties is not SolidFill"),
        }

        match ln_style.dash_properties {
            Some(crate::drawingml::LineDashProperties::PresetDash(ref dash)) => {
                assert_eq!(*dash, crate::drawingml::PresetLineDashVal::Solid)
            }
            _ => panic!("ln_style.dash_properties is not PresetDash"),
        }

        // bg fill style test
        let bg_fill_style = &format_scheme.bg_fill_style_list[1];
        match bg_fill_style {
            crate::drawingml::FillProperties::GradientFill(ref gradient) => {
                assert_eq!(gradient.rotate_with_shape, Some(true));

                let stop = &gradient.gradient_stop_list[0];
                assert_eq!(stop.position, 0.0);
                match stop.color {
                    crate::drawingml::Color::SchemeColor(ref scheme_clr) => {
                        assert_eq!(scheme_clr.value, crate::drawingml::SchemeColorVal::PlaceholderColor);

                        match scheme_clr.color_transforms[0] {
                            crate::drawingml::ColorTransform::Tint(val) => assert_eq!(val, 40_000.0),
                            _ => panic!("invalid color transform"),
                        }

                        match scheme_clr.color_transforms[1] {
                            crate::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 350_000.0),
                            _ => panic!("invalid color transform"),
                        }
                    }
                    _ => panic!("stop color is not scheme color"),
                }

                let stop = &gradient.gradient_stop_list[1];
                assert_eq!(stop.position, 40_000.0);
                match stop.color {
                    crate::drawingml::Color::SchemeColor(ref scheme_clr) => {
                        assert_eq!(scheme_clr.value, crate::drawingml::SchemeColorVal::PlaceholderColor);

                        match scheme_clr.color_transforms[0] {
                            crate::drawingml::ColorTransform::Tint(val) => assert_eq!(val, 45_000.0),
                            _ => panic!("invalid color transform"),
                        }

                        match scheme_clr.color_transforms[1] {
                            crate::drawingml::ColorTransform::Shade(val) => assert_eq!(val, 99_000.0),
                            _ => panic!("invalid color transform"),
                        }

                        match scheme_clr.color_transforms[2] {
                            crate::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 350_000.0),
                            _ => panic!("invalid color transform"),
                        }
                    }
                    _ => panic!("stop color is not scheme color"),
                }

                let stop = &gradient.gradient_stop_list[2];
                assert_eq!(stop.position, 100_000.0);
                match stop.color {
                    crate::drawingml::Color::SchemeColor(ref scheme_clr) => {
                        assert_eq!(scheme_clr.value, crate::drawingml::SchemeColorVal::PlaceholderColor);

                        match scheme_clr.color_transforms[0] {
                            crate::drawingml::ColorTransform::Shade(val) => assert_eq!(val, 20_000.0),
                            _ => panic!("invalid color transform"),
                        }

                        match scheme_clr.color_transforms[1] {
                            crate::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 255_000.0),
                            _ => panic!("invalid color transform"),
                        }
                    }
                    _ => panic!("stop color is not scheme color"),
                }

                match gradient.shade_properties {
                    Some(crate::drawingml::ShadeProperties::Path(ref path)) => {
                        assert_eq!(path.path, Some(crate::drawingml::PathShadeType::Circle));
                        match path.fill_to_rect {
                            Some(ref rect) => {
                                assert_eq!(rect.left, Some(50_000.0));
                                assert_eq!(rect.top, Some(-80_000.0));
                                assert_eq!(rect.right, Some(50_000.0));
                                assert_eq!(rect.bottom, Some(180_000.0));
                            }
                            None => panic!("fill_to_rect is None"),
                        }
                    }
                    _ => panic!("gradient shade properties is not path"),
                }
            }
            _ => panic!("bg fill style is not gradient"),
        }
    }
}