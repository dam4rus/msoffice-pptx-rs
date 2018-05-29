use docprops::{ AppInfo, Core };

pub struct Document {
    pub app: Option<AppInfo>,
    pub core: Option<Core>,
}