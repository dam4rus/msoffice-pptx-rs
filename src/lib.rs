extern crate quick_xml;
extern crate zip;

#[macro_use]
mod macros;
mod xml;
pub mod error;
pub mod docprops;
pub mod relationship;
pub mod pml;
pub mod drawingml;
pub mod document;

#[cfg(test)]
mod tests {
    use document::*;
    use docprops::*;
    use xml::*;
    use std::fs::File;
    use std::io::{ Read };
    use std::path::{ Path, PathBuf };

    #[test]
    fn test_sample_pptx() {
        let test_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        let sample_pptx_path = test_dir.join("tests/samplepptx.pptx");

        let document = Document::from_file(&sample_pptx_path).unwrap();

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

        if let Some(presentation) = document.presentation {
            let master_id = presentation.slide_master_id_list.get(0).unwrap();
            assert_eq!(master_id.id.unwrap(), 2147483684);
            assert_eq!(master_id.relationship_id, "rId1");

            let slide_id_0 = presentation.slide_id_list.get(0).unwrap();
            assert_eq!(slide_id_0.id, 256);
            assert_eq!(slide_id_0.relationship_id, "rId2");

            let slide_id_1 = presentation.slide_id_list.get(1).unwrap();
            assert_eq!(slide_id_1.id, 257);
            assert_eq!(slide_id_1.relationship_id, "rId3");

            let slide_size = presentation.slide_size.unwrap();
            assert_eq!(slide_size.width, 9144000);
            assert_eq!(slide_size.height, 6858000);
            assert_eq!(slide_size.size_type.unwrap(), ::pml::SlideSizeType::Screen4x3);

            let notes_size = presentation.notes_size.unwrap();
            assert_eq!(notes_size.width, 6858000);
            assert_eq!(notes_size.height, 9144000);

            let def_text_style = presentation.default_text_style.unwrap();
            let def_par_props = def_text_style.def_paragraph_props.unwrap();
            assert_eq!(def_par_props.default_run_properties.unwrap().language.unwrap(), "en-US");

            let lvl1_ppr = def_text_style.lvl1_paragraph_props.unwrap();
            assert_eq!(lvl1_ppr.margin_left.unwrap(), 0);
            assert_eq!(lvl1_ppr.align.unwrap(), ::drawingml::TextAlignType::Left);
            assert_eq!(lvl1_ppr.default_tab_size.unwrap(), 914400);
            assert_eq!(lvl1_ppr.rtl.unwrap(), false);
            assert_eq!(lvl1_ppr.east_asian_line_break.unwrap(), true);
            assert_eq!(lvl1_ppr.latin_line_break.unwrap(), false);
            assert_eq!(lvl1_ppr.hanging_punctuations.unwrap(), true);

            let lvl1_def_rpr = lvl1_ppr.default_run_properties.unwrap();
            assert_eq!(lvl1_def_rpr.font_size.unwrap(), 1800);
            assert_eq!(lvl1_def_rpr.kerning.unwrap(), 1200);

            assert_eq!(lvl1_def_rpr.latin_font.unwrap().typeface, "+mn-lt");
            assert_eq!(lvl1_def_rpr.east_asian_font.unwrap().typeface, "+mn-ea");
            assert_eq!(lvl1_def_rpr.complex_script_font.unwrap().typeface, "+mn-cs");
        }

        for (path, theme) in document.theme_map {
            assert_eq!("ppt/theme/theme1.xml", path.to_str().unwrap());

            let color_scheme = theme.theme_elements.color_scheme;
            assert_eq!(color_scheme.name, "Default Design 1");

            match color_scheme.dark1 {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0x000066),
                _ => panic!("theme1 dk1 color doesn't equals to 0x00000066"),
            }
            
            match color_scheme.light1 {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0xFFFFFF),
                _ => panic!("theme1 lt1 color doesn't equals to 0xFFFFFF"),
            }
            
            match color_scheme.dark2 {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0x003366),
                _ => panic!("theme1 dk2 color doesn't equals to 0x003366"),
            }
            
            match color_scheme.light2 {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0xFFFFFF),
                _ => panic!("theme1 lt2 color doesn't equals to 0xFFFFFF"),
            }
            
            match color_scheme.accent1 {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0x8EB3C8),
                _ => panic!("theme1 accent1 color doesn't equals to 0x8EB3C8"),
            }
            
            match color_scheme.accent2 {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0x6F97B3),
                _ => panic!("theme1 accent2 color doesn't equals to 0x6F97B3"),
            }
            
            match color_scheme.accent3 {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0xAAADB8),
                _ => panic!("theme1 accent3 color doesn't equals to 0xAAADB8"),
            }
            
            match color_scheme.accent4 {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0xDADADA),
                _ => panic!("theme1 accent4 color doesn't equals to 0xDADADA"),
            }
            
            match color_scheme.accent5 {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0xC6D6E0),
                _ => panic!("theme1 accent5 color doesn't equals to 0xC6D6E0"),
            }
            
            match color_scheme.accent6 {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0x6488A2),
                _ => panic!("theme1 accent6 color doesn't equals to 0x6488A2"),
            }
            
            match color_scheme.hyperlink {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0x556575),
                _ => panic!("theme1 hyperlink color doesn't equals to 0x556575"),
            }
            
            match color_scheme.followed_hyperlink {
                ::drawingml::Color::SRgbColor(clr) => assert_eq!(clr.value, 0x3D556F),
                _ => panic!("theme1 followhyperlink color doesn't equals to 0x3D556F"),
            }

            // font_scheme test
            let font_scheme = theme.theme_elements.font_scheme;
            assert_eq!(font_scheme.name, "Default Design");

            let major_font = font_scheme.major_font;
            assert_eq!(major_font.latin.typeface, "Tahoma");
            assert_eq!(major_font.east_asian.typeface, "");
            assert_eq!(major_font.complex_script.typeface, "");

            let minor_font = font_scheme.minor_font;
            assert_eq!(minor_font.latin.typeface, "Tahoma");
            assert_eq!(minor_font.east_asian.typeface, "");
            assert_eq!(minor_font.complex_script.typeface, "");

            let format_scheme = theme.theme_elements.format_scheme;
            assert_eq!(format_scheme.name.unwrap(), "Office");

            let fill_style = &format_scheme.fill_style_list[0];
            match fill_style {
                ::drawingml::FillProperties::SolidFill(::drawingml::Color::SchemeColor(clr)) => assert_eq!(clr.value, ::drawingml::SchemeColorVal::PlaceholderColor),
                _ => panic!("fill[0] is invalid"),
            }

            let fill_style = &format_scheme.fill_style_list[1];
            match fill_style {
                ::drawingml::FillProperties::GradientFill(ref gradient) => {
                    let stop = &gradient.gradient_stop_list[0];
                    assert_eq!(stop.position, 0.0);
                    match stop.color {
                        ::drawingml::Color::SchemeColor(ref clr) => {
                            assert_eq!(clr.value, ::drawingml::SchemeColorVal::PlaceholderColor);
                            match clr.color_transforms[0] {
                                ::drawingml::ColorTransform::Tint(val) => assert_eq!(val, 50_000.0),
                                _ => panic!("color transform is not tint"),
                            }

                            match clr.color_transforms[1] {
                                ::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 300_000.0),
                                _ => panic!("color transform is not satMod"),
                            }
                        }
                        _ => panic!("stop color is not schemeColor"),
                    }

                    let stop = &gradient.gradient_stop_list[1];
                    assert_eq!(stop.position, 35_000.0);
                    match stop.color {
                        ::drawingml::Color::SchemeColor(ref clr) => {
                            assert_eq!(clr.value, ::drawingml::SchemeColorVal::PlaceholderColor);
                            match clr.color_transforms[0] {
                                ::drawingml::ColorTransform::Tint(val) => assert_eq!(val, 37_000.0),
                                _ => panic!("color transform is not tint"),
                            }

                            match clr.color_transforms[1] {
                                ::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 300_000.0),
                                _ => panic!("color transform is not satMod"),
                            }
                        }
                        _ => panic!("stop color is not schemeColor"),
                    }

                    let stop = &gradient.gradient_stop_list[2];
                    assert_eq!(stop.position, 100_000.0);
                    match stop.color {
                        ::drawingml::Color::SchemeColor(ref clr) => {
                            assert_eq!(clr.value, ::drawingml::SchemeColorVal::PlaceholderColor);
                            match clr.color_transforms[0] {
                                ::drawingml::ColorTransform::Tint(val) => assert_eq!(val, 15_000.0),
                                _ => panic!("color transform is not tint"),
                            }

                            match clr.color_transforms[1] {
                                ::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 350_000.0),
                                _ => panic!("color transform is not satMod"),
                            }
                        }
                        _ => panic!("stop color is not schemeColor"),
                    }

                    match gradient.shade_properties {
                        Some(::drawingml::ShadeProperties::Linear(ref shade_props)) => {
                            assert_eq!(shade_props.angle, Some(16_200_000));
                            assert_eq!(shade_props.scaled, Some(true));
                        }
                        _ => panic!("shape is not linear"),
                    }
                }
                _ => panic!("fill[1] is invalid"),
            }

            let ln_style = &format_scheme.line_style_list[0];
            assert_eq!(ln_style.width, Some(9_525));
            assert_eq!(ln_style.cap, Some(::drawingml::LineCap::Flat));
            assert_eq!(ln_style.compound, Some(::drawingml::CompoundLine::Single));
            assert_eq!(ln_style.pen_alignment, Some(::drawingml::PenAlignment::Center));
            
            match ln_style.fill_properties {
                Some(::drawingml::LineFillProperties::SolidFill(ref clr)) => {
                    match clr {
                        ::drawingml::Color::SchemeColor(ref scheme_clr) => {
                            assert_eq!(scheme_clr.value, ::drawingml::SchemeColorVal::PlaceholderColor);
                            match scheme_clr.color_transforms[0] {
                                ::drawingml::ColorTransform::Shade(val) => assert_eq!(val, 95_000.0),
                                _ => panic!("ColorTransform is not Shade"),
                            }

                            match scheme_clr.color_transforms[1] {
                                ::drawingml::ColorTransform::SaturationModulate(val) => assert_eq!(val, 105_000.0),
                                _ => panic!("ColorTransform is not SatMode"),
                            }
                        }
                        _ => panic!("Scheme color is not PlaceholderColor"),
                    }
                }
                _ => panic!("ln_style.fill_properties is not SolidFill"),
            }

            match ln_style.dash_properties {
                Some(::drawingml::LineDashProperties::PresetDash(ref dash)) => assert_eq!(*dash, ::drawingml::PresetLineDashVal::Solid),
                _ => panic!("ln_style.dash_properties is not PresetDash"),
            }
        }
    }

    #[test]
    fn text_xml_parser() {
        let test_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        let sample_xml_file = test_dir.join("tests/presentation.xml");
        let mut file = File::open(sample_xml_file).expect("Sample xml file not found");

        let mut file_content = String::new();
        file.read_to_string(&mut file_content).expect("Failed to read sample xml file to string");

        let root_node = XmlNode::from_str(file_content.as_str()).expect("Couldn't create XmlNode from string");
        assert_eq!(root_node.name, "p:presentation");
        assert_eq!(root_node.attribute("xmlns:a").unwrap(), "http://schemas.openxmlformats.org/drawingml/2006/main");

        assert_eq!(root_node.child_nodes[0].name, "p:sldMasterIdLst");
        assert_eq!(root_node.child_nodes[1].name, "p:sldIdLst");
        assert_eq!(root_node.child_nodes[2].name, "p:sldSz");
        assert_eq!(root_node.child_nodes[3].name, "p:notesSz");
        assert_eq!(root_node.child_nodes[4].name, "p:custDataLst");
        assert_eq!(root_node.child_nodes[5].name, "p:defaultTextStyle");
        assert_eq!(root_node.child_nodes[0].child_nodes[0].name, "p:sldMasterId");

        let slide_id_0_node = &root_node.child_nodes[1].child_nodes[0];
        assert_eq!(slide_id_0_node.name, "p:sldId");
        assert_eq!(slide_id_0_node.attribute("id").unwrap(), "256");
        assert_eq!(slide_id_0_node.attribute("r:id").unwrap(), "rId2");

        assert_eq!(root_node.child_nodes[1].child_nodes[1].name, "p:sldId");

        let lvl1_ppr_defrpr_node = &root_node.child_nodes[5].child_nodes[1].child_nodes[0];
        assert_eq!(lvl1_ppr_defrpr_node.attribute("sz").unwrap(), "1800");
        assert_eq!(lvl1_ppr_defrpr_node.attribute("kern").unwrap(), "1200");
    }
}
