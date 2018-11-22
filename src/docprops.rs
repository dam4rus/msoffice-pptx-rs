use std::io::{ Read, Seek };

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

    pub fn from_zip<R>(zipper: &mut zip::ZipArchive<R>) -> Option<AppInfo> where R: Read + Seek {
        let mut app_xml_file = match zipper.by_name("docProps/app.xml") {
            Ok(f) => f,
            Err(_) => return None,
        };
        
        let mut app_info = AppInfo::new();
        let mut xml_string = String::new();
        match app_xml_file.read_to_string(&mut xml_string) {
            Ok(_) => {
                let mut buffer = Vec::new();
                let mut xml_reader = quick_xml::Reader::from_str(xml_string.as_str());
                loop {
                    match xml_reader.read_event(&mut buffer) {
                        Ok(quick_xml::events::Event::Start(ref element)) => {
                            match element.local_name() {
                                b"Application" => app_info.app_name = Some(xml_reader.read_text(element.name(), &mut Vec::new()).unwrap()),
                                b"AppVersion" => app_info.app_version = Some(xml_reader.read_text(element.name(), &mut Vec::new()).unwrap()),
                                _ => (),
                            }
                        }
                        Ok(quick_xml::events::Event::Eof) => break,
                        _ => (),
                    }

                    buffer.clear();
                }
            },
            Err(_) => return None,
        }

        Some(app_info)
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

    pub fn from_zip<R>(zipper: &mut zip::ZipArchive<R>) -> Option<Core> where R: Read + Seek {
        let mut core_xml_file = match zipper.by_name("docProps/core.xml") {
            Ok(f) => f,
            Err(_) => return None,
        };

        let mut core = Core::new();
        let mut xml_string = String::new();
        match core_xml_file.read_to_string(&mut xml_string) {
            Ok(_) => {
                let mut buffer = Vec::new();
                let mut xml_reader = quick_xml::Reader::from_str(xml_string.as_str());
                loop {
                    match xml_reader.read_event(&mut buffer) {
                        Ok(quick_xml::events::Event::Start(ref element)) => {
                            match element.local_name() {
                                b"title" => core.title = Some(xml_reader.read_text(element.name(), &mut Vec::new()).unwrap()),
                                b"creator" => core.creator = Some(xml_reader.read_text(element.name(), &mut Vec::new()).unwrap()),
                                b"lastModifiedBy" => core.last_modified_by = Some(xml_reader.read_text(element.name(), &mut Vec::new()).unwrap()),
                                b"revision" => core.revision = Some(xml_reader.read_text(element.name(), &mut Vec::new()).unwrap().parse::<i32>().unwrap()),
                                b"created" => core.created_time = Some(xml_reader.read_text(element.name(), &mut Vec::new()).unwrap()),
                                b"modified" => core.modified_time = Some(xml_reader.read_text(element.name(), &mut Vec::new()).unwrap()),
                                _ => (),
                            }
                        }
                        Ok(quick_xml::events::Event::Eof) => break,
                        _ => (),
                    }

                    buffer.clear();
                }
            }
            Err(_) => return None,
        }

        Some(core)
    }
}