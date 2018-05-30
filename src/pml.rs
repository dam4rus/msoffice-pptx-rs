use drawingml;
use relationship;

pub type SlideId = u32; // TODO: 256 <= n <= 2147483648
pub type SlideLayoutId = u32; // TODO: 2147483648 <= n
pub type SlideMasterId = u32; // TODO: 2147483648 <= n
pub type Index = u32;
pub type TLTimeNodeId = u32;
pub type BookmarkIdSeed = u32; // TODO: 1 <= n <= 2147483648
pub type SlideSizeCoordinate = drawingml::PositiveCoordinate32; // TODO: 914400 <= n <= 51206400
pub type Name = String;

decl_oox_enum! {
    pub enum ConformanceClass {
        Strict = "strict",
        Transitional = "transitional",
    }
}

decl_oox_enum! {
    pub enum SlideLayoutType {
        Title = "title",
        Tx = "tx",
        TwoColTx = "twoColTx",
        Tbl = "tbl",
        TxAndChart = "txAndChart",
        ChartAndTx = "chartAndTx",
        Dgm = "dgm",
        Chart = "chart",
        TxAndClipArt = "txAndClipArt",
        ClipArtAndTx = "clipArtAndTx",
        TitleOnly = "titleOnly",
        Blank = "blank",
        TxAndObj = "txAndObj",
        ObjAndTx = "objAndTx",
        ObjOnly = "objOnly",
        Obj = "obj",
        TxAndMedia = "txAndMedia",
        MediaAndTx = "mediaAndTx",
        ObjOverTx = "objOverTx",
        TxOverObj = "txOverObj",
        TxAndTwoObj = "txAndTwoObj",
        TwoObjAndTx = "twoObjAndTx",
        TwoObjOverTx = "twoObjOverTx",
        FourObj = "fourObj",
        VertTx = "vertTx",
        ClipArtAndVertTx = "clipArtAndVertTx",
        VertTitleAndTx = "vertTitleAndTx",
        VertTitleAndTxOverChart = "vertTitleAndTxOverChart",
        TwoObj = "twoObj",
        ObjAndTwoObj = "objAndTwoObj",
        TwoObjAndObj = "twoObjAndObj",
        Cust = "cust",
        SecHead = "secHead",
        TwoTxTwoObj = "twoTxTwoObj",
        ObjTx = "objTx",
        PicTx = "picTx",
    }
}

decl_oox_enum! {
    pub enum PlaceholderType {
        Title = "title",
        Body = "body",
        CtrTitle = "ctrTitle",
        SubTitle = "subTitle",
        Dt = "dt",
        SldNum = "sldNum",
        Ftr = "ftr",
        Hdr = "hdr",
        Obj = "obj",
        Chart = "chart",
        Tbl = "tbl",
        ClipArt = "clipArt",
        Dgm = "dgm",
        Media = "media",
        SldImg = "sldImg",
        Pic = "pic",
    }
}

decl_oox_enum! {
    pub enum Direction {
        Horz = "horz",
        Vert = "vert",
    }
}

decl_oox_enum! {
    pub enum PlaceholderSize {
        Full = "full",
        Half = "half",
        Quarter = "quarter",
    }
}

decl_oox_enum! {
    pub enum SlideSizeType {
        Screen4x3 = "screen4x3",
        Letter = "letter",
        A4 = "a4",
        Mm35 = "mm35",
        Overhead = "overhead",
        Banner = "banner",
        Custom = "custom",
        Ledger = "ledger",
        A3 = "a3",
        B4ISO = "b4ISO",
        B5ISO = "b5ISO",
        B4JIS = "b4JIS",
        B5JIS = "b5JIS",
        HagakiCard = "hagakiCard",
        Screen16x9 = "screen16x9",
        Screen16x10 = "screen16x10",
    }
}

decl_oox_enum! {
    pub enum PhotoAlbumLayout {
        FitToSlide = "fitToSlide",
        Pic1 = "pic1",
        Pic2 = "pic2",
        Pic4 = "pic4",
        PicTitle1 = "picTitle1",
        PicTitle2 = "picTitle2",
        PicTitle4 = "picTitle4",
    }
}

decl_oox_enum! {
    pub enum PhotoAlbumFrameShape {
        FrameStyle1 = "frameStyle1",
        FrameStyle2 = "frameStyle2",
        FrameStyle3 = "frameStyle3",
        FrameStyle4 = "frameStyle4",
        FrameStyle5 = "frameStyle5",
        FrameStyle6 = "frameStyle6",
        FrameStyle7 = "frameStyle7",
    }
}

decl_oox_enum! {
    pub enum OleObjectFollowColorScheme {
        None = "none",
        Full = "full",
        TextAndBackground = "textAndBackground",
    }
}


pub struct SlideSize {
    pub width: SlideSizeCoordinate,
    pub height: SlideSizeCoordinate,
    pub size_type: Option<SlideSizeType>,
}


pub struct SlideIdListEntry {
    pub id: Option<SlideId>,
    pub relationship_id: relationship::RelationshipId,
}


pub struct SlideMasterIdListEntry {
    pub id: Option<SlideMasterId>,
    pub relationship_id: relationship::RelationshipId,
}


pub struct NotesMasterIdListEntry {
    pub relationship_id: relationship::RelationshipId,
}


pub struct HandoutMasterIdListEntry {
    pub relationship_id: relationship::RelationshipId,
}


pub struct EmbeddedFontListEntry {
    pub font: Option<drawingml::TextFont>,
    pub regular: Option<relationship::RelationshipId>,
    pub bold: Option<relationship::RelationshipId>,
    pub italic: Option<relationship::RelationshipId>,
    pub bold_italic: Option<relationship::RelationshipId>,
}


pub struct CustomShow {
    pub name: Name,
    pub id: u32,
    pub slide_list: Vec<relationship::RelationshipId>,
    //std::unique_ptr<ExtensionList> extLst;
}


pub struct PhotoAlbum {
    pub black_and_white: Option<bool>, // false
    pub show_captions: Option<bool>, // false
    pub layout: Option<PhotoAlbumLayout>, // PhotoAlbumLayout::FitToSlide
    pub frame: Option<PhotoAlbumFrameShape>, // PhotoAlbumFrameShape::FrameStyle1
    /*
    //std::unique_ptr<ExtensionList> extLst;
    */
}


pub struct Presentation {
    pub server_zoom: Option<drawingml::Percentage>, // 50%
    pub first_slide_num: Option<i32>, // 1
    pub show_special_pls_on_title_slide: Option<bool>, // true
    pub rtl: Option<bool>, // false
    pub remove_personal_info_on_save: Option<bool>, // false
    pub compatibility_mode: Option<bool>, // false
    pub strict_first_and_last_chars: Option<bool>, // true
    pub embed_true_type_fonts: Option<bool>, // false
    pub save_subset_fonts: Option<bool>, // false
    pub auto_compress_pictures: Option<bool>, // true
    pub bookmark_id_seed: Option<BookmarkIdSeed>, // 1
    pub conformance: ConformanceClass,
    pub slide_master_id_list: Vec<SlideMasterIdListEntry>,
    pub notes_master_id_list: Vec<NotesMasterIdListEntry>, // length = 1
    pub handout_master_id_list: Vec<HandoutMasterIdListEntry>, // length = 1
    pub slide_id_list: Vec<SlideIdListEntry>,
    pub slide_size: Option<SlideSize>,
    pub notes_size: Option<drawingml::PositiveSize2D>,
    //std::unique_ptr<SmartTags> smartTags;
    pub embedded_font_list: Vec<EmbeddedFontListEntry>,
    pub custom_show_list: Vec<CustomShow>,
    pub photo_album: Option<PhotoAlbum>,
    //std::unique_ptr<CustomerDataList> custDataLst;
    //std::unique_ptr<Kinsoku> kinsoku;
    pub default_text_style: Option<drawingml::TextListStyle>,
    //std::unique_ptr<ModifyVerifier> modifyVerifier;
    //std::unique_ptr<ExtensionList> extLst;
}