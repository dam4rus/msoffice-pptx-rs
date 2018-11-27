use std::io::{ Read, Seek };
use ::xml::XmlNode;

/// AppInfo
/// 
pub struct AppInfo {
    pub app_name: Option<String>,
    pub app_version: Option<String>,
}

impl AppInfo {
    fn new() -> AppInfo {
        AppInfo {
            app_name: None,
            app_version: None,
        }
    }

    pub fn from_zip<R>(zipper: &mut zip::ZipArchive<R>) -> Result<AppInfo, Box<::std::error::Error>>
    where
        R: Read + Seek
    {
        let mut app_xml_file = zipper.by_name("docProps/app.xml")?;
        
        let mut xml_string = String::new();
        app_xml_file.read_to_string(&mut xml_string)?;
        let root = XmlNode::from_str(&xml_string)?;

        let mut app_info = AppInfo::new();
        for child_node in &root.child_nodes {
            match child_node.local_name() {
                "Application" => app_info.app_name = child_node.text.as_ref().map(|ok| ok.clone()),
                "AppVersion" => app_info.app_version = child_node.text.as_ref().map(|ok| ok.clone()),
                _ => (),
            }
        }

        Ok(app_info)
    }
}

/// Core
/// 
pub struct Core {
    pub title: Option<String>,
    pub creator: Option<String>,
    pub last_modified_by: Option<String>,
    pub revision: Option<i32>,
    pub created_time: Option<String>,  // TODO: maybe store as some DateTime struct?
    pub modified_time: Option<String>,  // TODO: maybe store as some DateTime struct?
}

impl Core {
    fn new() -> Core {
        Core {
            title: None,
            creator: None,
            last_modified_by: None,
            revision: None,
            created_time: None,
            modified_time: None,
        }
    }

    pub fn from_zip<R>(zipper: &mut zip::ZipArchive<R>) -> Result<Core, Box<::std::error::Error>>
    where
        R: Read + Seek
    {
        let mut core_xml_file = zipper.by_name("docProps/core.xml")?;
        let mut xml_string = String::new();
        core_xml_file.read_to_string(&mut xml_string)?;
        let root = XmlNode::from_str(&xml_string)?;

        let mut core = Core::new();

        for child_node in &root.child_nodes {
            match child_node.local_name() {
                "title" => core.title = child_node.text.as_ref().map(|s| s.clone()),
                "creator" => core.creator = child_node.text.as_ref().map(|s| s.clone()),
                "lastModifiedBy" => core.last_modified_by = child_node.text.as_ref().map(|s| s.clone()),
                "revision" => core.revision = child_node.text.as_ref().and_then(|s| s.parse::<i32>().ok()),
                "created" => core.created_time = child_node.text.as_ref().map(|s| s.clone()),
                "modified" => core.modified_time = child_node.text.as_ref().map(|s| s.clone()),
                _ => (),
            }
        }

        Ok(core)
    }
}