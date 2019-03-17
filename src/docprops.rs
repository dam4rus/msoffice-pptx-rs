use msoffice_shared::xml::XmlNode;
use std::io::{Read, Seek};

pub struct AppInfo {
    pub app_name: Option<String>,
    pub app_version: Option<String>,
}

impl AppInfo {
    pub fn from_zip<R>(zipper: &mut zip::ZipArchive<R>) -> Result<Self, Box<dyn (::std::error::Error)>>
    where
        R: Read + Seek,
    {
        let mut app_xml_file = zipper.by_name("docProps/app.xml")?;

        let mut xml_string = String::new();
        app_xml_file.read_to_string(&mut xml_string)?;
        let root = XmlNode::from_str(&xml_string)?;

        let mut app_name = None;
        let mut app_version = None;
        for child_node in &root.child_nodes {
            match child_node.local_name() {
                "Application" => app_name = child_node.text.as_ref().cloned(),
                "AppVersion" => app_version = child_node.text.as_ref().cloned(),
                _ => (),
            }
        }

        Ok(Self { app_name, app_version })
    }
}

pub struct Core {
    pub title: Option<String>,
    pub creator: Option<String>,
    pub last_modified_by: Option<String>,
    pub revision: Option<i32>,
    pub created_time: Option<String>,  // TODO: maybe store as some DateTime struct?
    pub modified_time: Option<String>, // TODO: maybe store as some DateTime struct?
}

impl Core {
    pub fn from_zip<R>(zipper: &mut zip::ZipArchive<R>) -> Result<Self, Box<dyn (::std::error::Error)>>
    where
        R: Read + Seek,
    {
        let mut core_xml_file = zipper.by_name("docProps/core.xml")?;
        let mut xml_string = String::new();
        core_xml_file.read_to_string(&mut xml_string)?;
        let root = XmlNode::from_str(&xml_string)?;

        let mut title = None;
        let mut creator = None;
        let mut last_modified_by = None;
        let mut revision = None;
        let mut created_time = None;
        let mut modified_time = None;

        for child_node in &root.child_nodes {
            match child_node.local_name() {
                "title" => title = child_node.text.as_ref().cloned(),
                "creator" => creator = child_node.text.as_ref().cloned(),
                "lastModifiedBy" => last_modified_by = child_node.text.as_ref().cloned(),
                "revision" => revision = child_node.text.as_ref().and_then(|s| s.parse().ok()),
                "created" => created_time = child_node.text.as_ref().cloned(),
                "modified" => modified_time = child_node.text.as_ref().cloned(),
                _ => (),
            }
        }

        Ok(Self {
            title,
            creator,
            last_modified_by,
            revision,
            created_time,
            modified_time,
        })
    }
}
