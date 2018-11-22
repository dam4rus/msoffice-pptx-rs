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
    use drawingml;
    use pml;
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
            assert_eq!(slide_size.size_type.unwrap(), pml::SlideSizeType::Screen4x3);

            let notes_size = presentation.notes_size.unwrap();
            assert_eq!(notes_size.width, 6858000);
            assert_eq!(notes_size.height, 9144000);

            let def_text_style = presentation.default_text_style.unwrap();
            let def_par_props = def_text_style.def_paragraph_props.unwrap();
            assert_eq!(def_par_props.default_run_properties.unwrap().language.unwrap(), "en-US");

            let lvl1_ppr = def_text_style.lvl1_paragraph_props.unwrap();
            assert_eq!(lvl1_ppr.margin_left.unwrap(), 0);
            assert_eq!(lvl1_ppr.align.unwrap(), drawingml::TextAlignType::Left);
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
