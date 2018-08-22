extern crate quick_xml;
extern crate zip;

#[macro_use]
mod macros;
mod helpers;
mod xml;
pub mod errors;
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
    }

    #[test]
    fn text_xml_parser() {
        let test_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        let sample_xml_file = test_dir.join("tests/presentation.xml");
        let mut file = File::open(sample_xml_file).expect("Sample xml file not found");

        let mut file_content = String::new();
        file.read_to_string(&mut file_content).expect("Failed to read sample xml file to string");

        let root_node = XmlNode::from_str(file_content.as_str()).expect("Couldn't create XmlNode from string");
        assert_eq!(root_node.get_name(), "p:presentation");
        assert_eq!(root_node.get_attribute("xmlns:a"), "http://schemas.openxmlformats.org/drawingml/2006/main");

        assert_eq!(root_node.child_nodes[0].get_name(), "p:sldMasterIdLst");
        assert_eq!(root_node.child_nodes[1].get_name(), "p:sldIdLst");
        assert_eq!(root_node.child_nodes[2].get_name(), "p:sldSz");
        assert_eq!(root_node.child_nodes[3].get_name(), "p:notesSz");
        assert_eq!(root_node.child_nodes[4].get_name(), "p:custDataLst");
        assert_eq!(root_node.child_nodes[5].get_name(), "p:defaultTextStyle");
        assert_eq!(root_node.child_nodes[0].child_nodes[0].get_name(), "p:sldMasterId");

        let slide_id_0_node = &root_node.child_nodes[1].child_nodes[0];
        assert_eq!(slide_id_0_node.get_name(), "p:sldId");
        assert_eq!(slide_id_0_node.get_attribute("id"), "256");
        assert_eq!(slide_id_0_node.get_attribute("r:id"), "rId2");

        assert_eq!(root_node.child_nodes[1].child_nodes[1].get_name(), "p:sldId");

        let lvl1_ppr_defrpr_node = &root_node.child_nodes[5].child_nodes[1].child_nodes[0];
        assert_eq!(lvl1_ppr_defrpr_node.get_attribute("sz"), "1800");
        assert_eq!(lvl1_ppr_defrpr_node.get_attribute("kern"), "1200");
    }
}
