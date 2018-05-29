pub struct AppInfo {
    pub app_name: Option<String>,
    pub app_version: Option<String>,
}

pub struct Core {
    pub title: Option<String>,
    pub creator: Option<String>,
    pub last_modified_by: Option<String>,
    pub revision: Option<i32>,
    pub created_time: Option<String>,  // TODO: maybe store as some DateTime struct?
    pub modified_time: Option<String>,  // TODO: maybe store as some DateTime struct?
}