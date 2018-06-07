use pml;
use docprops::{ AppInfo, Core };
use std::collections::{ HashMap };

pub struct Document {
    pub app: Option<AppInfo>,
    pub core: Option<Core>,
    pub presentation: Option<pml::Presentation>,
    pub slide_master_map: HashMap<String, pml::SlideMaster>,
}