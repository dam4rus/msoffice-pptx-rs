use ::std::io::{ Read, Seek };
use ::relationship::RelationshipId;
use ::xml::{XmlNode, parse_xml_bool};
use ::error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError, XmlError, ParseBoolError};
use ::zip::read::ZipFile;
use ::zip::result::ZipError;

pub type SlideId = u32; // TODO: 256 <= n <= 2147483648
pub type SlideLayoutId = u32; // TODO: 2147483648 <= n
pub type SlideMasterId = u32; // TODO: 2147483648 <= n
pub type Index = u32;
pub type TLTimeNodeId = u32;
pub type BookmarkIdSeed = u32; // TODO: 1 <= n <= 2147483648
pub type SlideSizeCoordinate = ::drawingml::PositiveCoordinate32; // TODO: 914400 <= n <= 51206400
pub type Name = String;
pub type TLSubShapeId = ::drawingml::ShapeId;

pub type Result<T> = ::std::result::Result<T, Box<::std::error::Error>>;

decl_simple_type_enum! {
    pub enum ConformanceClass {
        Strict = "strict",
        Transitional = "transitional",
    }
}

decl_simple_type_enum! {
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

decl_simple_type_enum! {
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

decl_simple_type_enum! {
    pub enum Direction {
        Horz = "horz",
        Vert = "vert",
    }
}

decl_simple_type_enum! {
    pub enum PlaceholderSize {
        Full = "full",
        Half = "half",
        Quarter = "quarter",
    }
}

decl_simple_type_enum! {
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

decl_simple_type_enum! {
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

decl_simple_type_enum! {
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

decl_simple_type_enum! {
    pub enum OleObjectFollowColorScheme {
        None = "none",
        Full = "full",
        TextAndBackground = "textAndBackground",
    }
}

decl_simple_type_enum! {
    pub enum TransitionSideDirectionType {
        Left = "l",
        Up = "u",
        Right = "r",
        Down = "d",
    }
}

decl_simple_type_enum! {
    pub enum TransitionCornerDirectionType {
        LeftUp = "lu",
        RightUp = "ru",
        LeftDown = "ld",
        RightDown = "rd",
    }
}

decl_simple_type_enum! {
    pub enum TransitionEightDirectionType {
        Left = "l",
        Up = "u",
        Right = "r",
        Down = "d",
        LeftUp = "lu",
        RightUp = "ru",
        LeftDown = "ld",
        RightDown = "rd",
    }
}

decl_simple_type_enum! {
    pub enum TransitionInOutDirectionType {
        In = "in",
        Out = "out",
    }
}

decl_simple_type_enum! {
    pub enum TransitionSpeed {
        Slow = "slow",
        Medium = "med",
        Fast = "fast",
    }
}

decl_simple_type_enum! {
    pub enum TLChartSubelementType {
        GridLegend = "gridLegend",
        Series = "series",
        Category = "category",
        PointInSeries = "ptInSeries",
        PointInCategory = "ptInCategory",
    }
}

decl_simple_type_enum! {
    pub enum TLParaBuildType {
        AllAtOnce = "allAtOnce",
        Paragraph = "p",
        Custom = "cust",
        Whole = "whole",
    }
}

decl_simple_type_enum! {
    pub enum TLDiagramBuildType {
        Whole = "whole",
        DepthByNode = "depthByNode",
        DepthByBranch = "depthByBranch",
        BreadthByNode = "breadthByNode",
        BreadthByLevel = "breadthByLvl",
        Clockwise = "cw",
        ClockwiseIn = "cwIn",
        ClockwiseOut = "cwOut",
        CounterClockwise = "ccw",
        CounterClockwiseIn = "ccwIn",
        CounterClockwiseOut = "ccwOut",
        InByRing = "inByRing",
        OutByRing = "outByRing",
        Up = "up",
        Down = "down",
        AllAtOnce = "allAtOnce",
        Custom = "cust",
    }
}

decl_simple_type_enum! {
    pub enum TLOleChartBuildType {
        AllAtOnce = "allAtOnce",
        Series = "series",
        Category = "category",
        SeriesElement = "seriesEl",
        CategoryElement = "categoryEl",
    }
}

decl_simple_type_enum! {
    pub enum TLTriggerRuntimeNode {
        First = "first",
        Last = "last",
        All = "all",
    }
}

decl_simple_type_enum! {
    pub enum TLTriggerEvent {
        OnBegin = "onBegin",
        OnEnd = "onEnd",
        Begin = "begin",
        End = "end",
        OnClick = "onClick",
        OnDoubleClick = "onDblClick",
        OnMouseOver = "onMouseOver",
        OnMouseOut = "onMouseOut",
        OnNext = "onNext",
        OnPrev  = "onPrev",
        OnStopAudio = "onStopAudio",
    }
}

decl_simple_type_enum! {
    pub enum IterateType {
        Element = "el",
        Word = "wd",
        Letter = "lt",
    }
}


decl_simple_type_enum! {
    pub enum TLTimeNodePresetClassType {
        Entrance = "entr",
        Exit = "exit",
        Emphasis = "emph",
        Path = "path",
        Verb = "verb",
        Mediacall = "mediacall",
    }
}

decl_simple_type_enum! {
    pub enum TLTimeNodeRestartType {
        Always = "always",
        WhenNotActive = "whenNotActive",
        Never = "never",
    }
}

decl_simple_type_enum! {
    pub enum TLTimeNodeFillType {
        Remove = "remove",
        Freeze = "freeze",
        Hold = "hold",
        Transition = "transition",
    }
}

decl_simple_type_enum! {
    pub enum TLTimeNodeSyncType {
        CanSlip = "canSlip",
        Locked = "locked",
    }
}

decl_simple_type_enum! {
    pub enum TLTimeNodeMasterRelation {
        SameClick = "sameClick",
        LastClick = "lastClick",
        NextClick = "nextClick",
    }
}

decl_simple_type_enum! {
    pub enum TLTimeNodeType {
        ClickEffect = "clickEffect",
        WithEffect = "withEffect",
        AfterEffect = "afterEffect",
        MainSequence = "mainSequence",
        InteractiveSequence = "interactiveSeq",
        ClickParagraph = "clickPar",
        WithGroup = "withGroup",
        AfterGroup = "afterGroup",
        TimingRoot = "tmRoot",
    }
}

decl_simple_type_enum! {
    pub enum TLNextActionType {
        None = "none",
        Seek = "seek",
    }
}

decl_simple_type_enum! {
    pub enum TLPreviousActionType {
        None = "none",
        SkipTimed = "skipTimed",
    }
}

decl_simple_type_enum! {
    pub enum TLAnimateBehaviorCalcMode {
        Discrete = "discrete",
        Linear = "lin",
        Formula = "fmla",
    }
}

decl_simple_type_enum! {
    pub enum TLAnimateBehaviorValueType {
        String = "str",
        Number = "num",
        Color = "clr",
    }
}

decl_simple_type_enum! {
    pub enum TLBehaviorAdditiveType {
        Base = "base",
        Sum = "sum",
        Replace = "repl",
        Multiply = "mult",
        None = "none",
    }
}

decl_simple_type_enum! {
    pub enum TLBehaviorAccumulateType {
        None = "none",
        Always = "always",
    }
}

decl_simple_type_enum! {
    pub enum TLBehaviorTransformType {
        Point = "pt",
        Image = "img",
    }
}

decl_simple_type_enum! {
    pub enum TLBehaviorOverrideType {
        Normal = "normal",
        ChildStyle = "childStyle",
    }
}

decl_simple_type_enum! {
    pub enum TLAnimateColorSpace {
        Rgb = "rgb",
        Hsl = "hsl",
    }
}

decl_simple_type_enum! {
    pub enum TLAnimateColorDirection {
        Clockwise = "cw",
        CounterClockwise = "ccw",
    }
}

decl_simple_type_enum! {
    pub enum TLAnimateEffectTransition {
        In = "in",
        Out = "out",
        None = "none",
    }
}

decl_simple_type_enum! {
    pub enum TLAnimateMotionBehaviorOrigin {
        Parent = "parent",
        Layout = "layout",
    }
}

decl_simple_type_enum! {
    pub enum TLAnimateMotionPathEditMode {
        Relative = "relative",
        Fixed = "fixed",
    }
}

decl_simple_type_enum! {
    pub enum TLCommandType {
        Event = "evt",
        Call = "call",
        Verb = "verb",
    }
}

pub struct IndexRange {
    pub start: Index,
    pub end: Index,
}

pub struct BackgroundProperties {
    pub shade_to_title: Option<bool>, // false
    pub fill: ::drawingml::FillProperties,
    pub effect: Option<::drawingml::EffectProperties>,
}

impl BackgroundProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shade_to_title = None;
        let mut opt_fill = None;
        let mut effect = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "shadeToTitle" => shade_to_title = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            if ::drawingml::FillProperties::is_choice_member(child_node.local_name()) {
                opt_fill = Some(::drawingml::FillProperties::from_xml_element(child_node)?);
            }
            // TODO: implement EffectProperties
            // else if ::drawingml::EffectProperties::is_choice_member(child_node.local_name()) {
            //    effect = Some(::drawingml::EffectProperties::from_xml_element(child_node)?);
            //}
        }

        let fill = opt_fill.ok_or_else(|| MissingChildNodeError::new("EG_FillProperties"))?;

        Ok(Self {
            shade_to_title,
            fill,
            effect,
        })
    }
}

pub enum BackgroundGroup {
    Properties(BackgroundProperties),
    Reference(::drawingml::StyleMatrixReference),
}

impl BackgroundGroup {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "bgPr" | "bgRef" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bgPr" => Ok(BackgroundGroup::Properties(BackgroundProperties::from_xml_element(xml_node)?)),
            "bgRef" => Ok(BackgroundGroup::Reference(::drawingml::StyleMatrixReference::from_xml_element(xml_node)?)),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_Background").into()),
        }
    }
}

pub struct Background {
    pub background: BackgroundGroup,
    pub black_and_white_mode: Option<::drawingml::BlackWhiteMode>, // white
}

impl Background {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let opt_background = None;
        let black_and_white_mode = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "bwMode" => black_and_white_mode = Some(value.parse::<::drawingml::BlackWhiteMode>()?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            if BackgroundGroup::is_choice_member(child_node.local_name()) {
                opt_background = Some(BackgroundGroup::from_xml_element(child_node)?);
            }
        }

        let background = opt_background.ok_or_else(|| MissingChildNodeError::new("bgPr|bgRef"))?;

        Ok(Self{
            background,
            black_and_white_mode,
        })
    }
}

pub struct Placeholder {
    pub placeholder_type: Option<PlaceholderType>, // obj
    pub orientation: Option<Direction>, // horz
    pub size: Option<PlaceholderSize>, // full
    pub index: Option<u32>, // 0
    pub has_custom_prompt: Option<bool>, // false
}

impl Placeholder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let placeholder_type = None;
        let orientation = None;
        let size = None;
        let index = None;
        let has_custom_prompt = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => placeholder_type = Some(value.parse::<PlaceholderType>()?),
                "orient" => orientation = Some(value.parse::<Direction>()?),
                "sz" => size = Some(value.parse::<PlaceholderSize>()?),
                "idx" => index = Some(value.parse::<u32>()?),
                "hasCustomPrompt" => has_custom_prompt = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(Self {
            placeholder_type,
            orientation,
            size,
            index,
            has_custom_prompt,
        })
    }
}

pub struct ApplicationNonVisualDrawingProps {
    pub is_photo: Option<bool>, // false
    pub is_user_drawn: Option<bool>, // false
    pub placeholder: Option<Placeholder>,
    pub media: Option<::drawingml::Media>,
    //pub customer_data_list: Option<CustomerDataList>,
}

impl ApplicationNonVisualDrawingProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let is_photo = None;
        let is_user_drawn = None;
        let placeholder = None;
        let media = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "isPhoto" => is_photo = Some(parse_xml_bool(value)?),
                "userDrawn" => is_user_drawn = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            let local_name = child_node.local_name();
            // TODO implement
            // if ::drawingml::Media::is_choice_member(local_name) {
            //     media = ::drawingml::Media::from_xml_element(child_node)?;
            // } else {
            match child_node.local_name() {
                "ph" => placeholder = Some(Placeholder::from_xml_element(child_node)?),
                "custDataLst" => (), // TODO implement
                _ => (),
            }
            //}
        }

        Ok(Self {
            is_photo,
            is_user_drawn,
            placeholder,
            media,
        })
    }
}

pub enum ShapeGroup {
    Shape(Shape),
    GroupShape(GroupShape),
    GraphicFrame(GraphicalObjectFrame),
    Connector(Connector),
    Picture(Picture),
    ContentPart(RelationshipId),
}

pub struct Shape {
    pub use_bg_fill: Option<bool>, // false
    pub non_visual_props: ShapeNonVisual,
    pub shape_props: ::drawingml::ShapeProperties,
    pub style: Option<::drawingml::ShapeStyle>,
    pub text_body: Option<::drawingml::TextBody>,
}

impl Shape {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut use_bg_fill = None;
        let mut non_visual_props = None;
        let mut shape_props = None;
        let mut style = None;
        let mut text_body = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "useBgFill" => use_bg_fill = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvSpPr" => non_visual_props = Some(ShapeNonVisual::from_xml_element(child_node)?),
                "spPr" => shape_props = Some(::drawingml::ShapeProperties::from_xml_element(child_node)?),
                "style" => style = Some(::drawingml::ShapeStyle::from_xml_element(child_node)?),
                "txBody" => text_body = Some(::drawingml::TextBody::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let non_visual_props = non_visual_props.ok_or_else(|| MissingChildNodeError::new("nvSpPr"))?;
        let shape_props = shape_props.ok_or_else(|| MissingChildNodeError::new("spPr"))?;

        Ok(Self {
            use_bg_fill,
            non_visual_props,
            shape_props,
            style,
            text_body,
        })
    }
}

pub struct ShapeNonVisual {
    pub drawing_props: ::drawingml::NonVisualDrawingProps,
    pub shape_drawing_props: ::drawingml::NonVisualDrawingShapeProps,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl ShapeNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut shape_drawing_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => drawing_props = Some(::drawingml::NonVisualDrawingProps::from_xml_element(child_node)?),
                "cNvSpPr" => shape_drawing_props = Some(::drawingml::NonVisualDrawingShapeProps::from_xml_element(child_node)?),
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new("cNvPr"))?;
        let shape_drawing_props = shape_drawing_props.ok_or_else(|| MissingChildNodeError::new("cNvSpPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new("nvPr"))?;

        Ok(Self {
            drawing_props,
            shape_drawing_props,
            app_props,
        })
    }
}

pub struct GroupShape {
    pub non_visual_props: GroupShapeNonVisual,
    pub group_shape_props: ::drawingml::GroupShapeProperties,
    pub shape_array: Vec<ShapeGroup>,
}

impl GroupShape {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut opt_non_visual_props = None;
        let mut opt_group_shape_props = None;
        let mut shape_array = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvGrpSpPr" => opt_non_visual_props = Some(GroupShapeNonVisual::from_xml_element(child_node)?),
                "grpSpPr" => opt_group_shape_props = Some(::drawingml::GroupShapeProperties::from_xml_element(child_node)?),
                "sp" => shape_array.push(ShapeGroup::Shape(Shape::from_xml_element(child_node)?)),
                "grpSp" => shape_array.push(ShapeGroup::GroupShape(GroupShape::from_xml_element(child_node)?)),
                // TODO implement GraphicalObjectFrame
                //"graphicFrame" => shape_array.push(ShapeGroup::GraphicFrame(GraphicalObjectFrame::from_xml_element(child_node)?)),
                "cxnSp" => shape_array.push(ShapeGroup::Connector(Connector::from_xml_element(child_node)?)),
                "pic" => shape_array.push(ShapeGroup::Picture(Picture::from_xml_element(child_node)?)),
                "contentPart" => {
                    let attr = child_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                    shape_array.push(ShapeGroup::ContentPart(attr.clone()));
                }
                _ => (),
            }
        }

        let non_visual_props = opt_non_visual_props.ok_or_else(|| MissingChildNodeError::new("nvGrpSpPr"))?;
        let group_shape_props = opt_group_shape_props.ok_or_else(|| MissingChildNodeError::new("grpSpPr"))?;

        Ok(Self {
            non_visual_props,
            group_shape_props,
            shape_array,
        })
    }
}

pub struct GroupShapeNonVisual {
    pub drawing_props: ::drawingml::NonVisualDrawingProps,
    pub group_drawing_props: ::drawingml::NonVisualGroupDrawingShapeProps,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl GroupShapeNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut opt_drawing_props = None;
        let mut opt_group_drawing_props = None;
        let mut opt_app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => opt_drawing_props = Some(
                    ::drawingml::NonVisualDrawingProps::from_xml_element(child_node)?
                ),
                "cNvGrpSpPr" => opt_group_drawing_props = Some(
                    ::drawingml::NonVisualGroupDrawingShapeProps::from_xml_element(child_node)?
                ),
                "nvPr" => opt_app_props = Some(
                    ApplicationNonVisualDrawingProps::from_xml_element(child_node)?
                ),
                _ => (),
            }
        }

        let drawing_props = opt_drawing_props.ok_or_else(|| MissingChildNodeError::new("cNvPr"))?;
        let group_drawing_props = opt_group_drawing_props.ok_or_else(|| MissingChildNodeError::new("cNvGrpSpPr"))?;
        let app_props = opt_app_props.ok_or_else(|| MissingChildNodeError::new("nvPr"))?;
        Ok(Self {
            drawing_props,
            group_drawing_props,
            app_props,
        })
    }
}

pub struct GraphicalObjectFrame {
    pub non_visual_props: GraphicalObjectFrameNonVisual,
    pub transform: ::drawingml::Transform2D,
    pub graphic: ::drawingml::GraphicalObject,
    pub black_white_mode: Option<::drawingml::BlackWhiteMode>,
}

pub struct GraphicalObjectFrameNonVisual {
    pub drawing_props: ::drawingml::NonVisualDrawingProps,
    pub graphic_frame_props: ::drawingml::NonVisualGraphicFrameProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

pub struct Connector {
    pub non_visual_props: ConnectorNonVisual,
    pub shape_props: ::drawingml::ShapeProperties,
    pub shape_style: Option<::drawingml::ShapeStyle>,
}

impl Connector {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_props = None;
        let mut shape_props = None;
        let mut shape_style = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvCxnSpPr" => non_visual_props = Some(ConnectorNonVisual::from_xml_element(child_node)?),
                "spPr" => shape_props = Some(::drawingml::ShapeProperties::from_xml_element(child_node)?),
                "style" => shape_style = Some(::drawingml::ShapeStyle::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let non_visual_props = non_visual_props.ok_or_else(|| MissingChildNodeError::new("nvCxnSpPr"))?;
        let shape_props = shape_props.ok_or_else(|| MissingChildNodeError::new("nvCxnSpPr"))?;

        Ok(Self {
            non_visual_props,
            shape_props,
            shape_style,
        })
    }
}

pub struct ConnectorNonVisual {
    pub drawing_props: ::drawingml::NonVisualDrawingProps,
    pub connector_props: ::drawingml::NonVisualConnectorProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl ConnectorNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut connector_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => drawing_props = Some(::drawingml::NonVisualDrawingProps::from_xml_element(child_node)?),
                "cNvCxnSpPr" => connector_props = Some(::drawingml::NonVisualConnectorProperties::from_xml_element(child_node)?),
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new("cNvPr"))?;
        let connector_props = connector_props.ok_or_else(|| MissingChildNodeError::new("cNvCxnSpPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new("nvPr"))?;

        Ok(Self {
            drawing_props,
            connector_props,
            app_props,
        })
    }
}

pub struct Picture {
    pub non_visual_props: PictureNonVisual,
    pub blip_fill: ::drawingml::BlipFillProperties,
    pub shape_props: ::drawingml::ShapeProperties,
    pub shape_style: Option<::drawingml::ShapeStyle>,
}

impl Picture {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_props = None;
        let mut blip_fill = None;
        let mut shape_props = None;
        let mut shape_style = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvPicPr" => non_visual_props = Some(PictureNonVisual::from_xml_element(child_node)?),
                "blipFill" => blip_fill = Some(::drawingml::BlipFillProperties::from_xml_element(child_node)?),
                "spPr" => shape_props = Some(::drawingml::ShapeProperties::from_xml_element(child_node)?),
                "style" => shape_style = Some(::drawingml::ShapeStyle::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let non_visual_props = non_visual_props.ok_or_else(|| MissingChildNodeError::new("nvPicPr"))?;
        let blip_fill = blip_fill.ok_or_else(|| MissingChildNodeError::new("blipFill"))?;
        let shape_props = shape_props.ok_or_else(|| MissingChildNodeError::new("spPr"))?;

        Ok(Self {
            non_visual_props,
            blip_fill,
            shape_props,
            shape_style,
        })
    }
}

pub struct PictureNonVisual {
    pub drawing_props: ::drawingml::NonVisualDrawingProps,
    pub picture_props: ::drawingml::NonVisualPictureProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl PictureNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut picture_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => drawing_props = Some(::drawingml::NonVisualDrawingProps::from_xml_element(child_node)?),
                "cNvPicPr" => picture_props = Some(::drawingml::NonVisualPictureProperties::from_xml_element(child_node)?),
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new("cNvPr"))?;
        let picture_props = picture_props.ok_or_else(|| MissingChildNodeError::new("cNvPicPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new("nvPr"))?;

        Ok(Self {
            drawing_props,
            picture_props,
            app_props,
        })
    }
}

pub struct CommonSlideData {
    pub name: Option<String>,
    pub background: Option<Background>,
    pub shape_tree: GroupShape,
    pub customer_data_list: Option<CustomerDataList>,
    pub control_list: Vec<Control>,
}

impl CommonSlideData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut name = None;
        let mut background = None;
        let mut opt_shape_tree = None;
        let mut customer_data_list = None;
        let mut control_list = Vec::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "name" => name = Some(value.clone()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "bg" => background = Some(Background::from_xml_element(child_node)?),
                "spTree" => opt_shape_tree = Some(GroupShape::from_xml_element(child_node)?),
                "custDataList" => customer_data_list = Some(CustomerDataList::from_xml_element(child_node)?),
                // TODO implement
                // "controls" => {
                //     for control_node in child_node.child_nodes {
                //         match control_node.local_name() {
                //             "control" => control_list.push(Control::from_xml_element(control_node)?),
                //             _ => (),
                //         }
                //     }
                // }
                _ => (),
            }
        }

        let shape_tree = opt_shape_tree.ok_or_else(|| XmlError::from(MissingChildNodeError::new("spTree")))?;

        Ok(Self {
            name,
            background,
            shape_tree,
            customer_data_list,
            control_list,
        })
    }
}

/// CustomerDataList
pub struct CustomerDataList {
    pub customer_data_list: Vec<RelationshipId>,
    pub tags: Option<RelationshipId>,
}

impl CustomerDataList {
    fn from_xml_element(xml_node: &XmlNode) -> Result<CustomerDataList> {
        let mut customer_data_list = Vec::new();
        let mut tags = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "custData" => {
                    let id_attr = child_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                    customer_data_list.push(id_attr.clone());
                }
                "tags" => {
                    let id_attr = child_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                    tags = Some(id_attr.clone());
                }
                _ => (),
            }
        }

        Ok(Self {
            customer_data_list,
            tags,
        })
    }
}

pub struct Control {
    pub picture: Option<Picture>,
    pub ole_attributes: OleAttributes,
}

pub struct OleAttributes {
    pub shape_id: Option<::drawingml::ShapeId>,
    pub name: Option<String>, // ""
    pub show_as_icon: Option<bool>, // false
    pub id: Option<RelationshipId>,
    pub image_width: Option<::drawingml::PositiveCoordinate32>,
    pub image_height: Option<::drawingml::PositiveCoordinate32>,
}

pub struct SlideSize {
    pub width: SlideSizeCoordinate,
    pub height: SlideSizeCoordinate,
    pub size_type: Option<SlideSizeType>,
}

pub struct SlideIdListEntry {
    pub id: SlideId,
    pub relationship_id: RelationshipId,
}

pub struct SlideLayoutIdListEntry {
    pub id: Option<SlideLayoutId>,
    pub relationship_id: RelationshipId,
}

pub struct SlideMasterIdListEntry {
    pub id: Option<SlideMasterId>,
    pub relationship_id: RelationshipId,
}

pub struct NotesMasterIdListEntry {
    pub relationship_id: RelationshipId,
}

pub struct HandoutMasterIdListEntry {
    pub relationship_id: RelationshipId,
}

pub struct SlideMasterTextStyles {
    pub title_styles: Option<::drawingml::TextListStyle>,
    pub body_styles: Option<::drawingml::TextListStyle>,
    pub other_styles: Option<::drawingml::TextListStyle>,
}

pub struct OrientationTransition {
    pub direction: Option<Direction>, // horz
}

pub struct EightDirectionTransition {
    pub direction: Option<TransitionEightDirectionType>, // l
}

pub struct OptionalBlackTransition {
    pub through_black: Option<bool>, // false
}

pub struct SideDirectionTransition {
    pub direction: Option<TransitionSideDirectionType>, // l
}

pub struct SplitTransition {
    pub orientation: Option<Direction>, // horz
    pub direction: Option<TransitionInOutDirectionType>, // out
}

pub struct CornerDirectionTransition {
    pub direction: Option<TransitionCornerDirectionType>, // lu
}

pub struct WheelTransition {
    pub spokes: Option<u32>, // 4
}

pub struct InOutTransition {
    pub direction: Option<TransitionInOutDirectionType>, // out
}

pub enum SlideTransitionGroup {
    Blinds(OrientationTransition),
    Checker(OrientationTransition),
    Circle,
    Dissolve,
    Comb(OrientationTransition),
    Cover(EightDirectionTransition),
    Cut(OptionalBlackTransition),
    Diamond,
    Fade(OptionalBlackTransition),
    Newsflash,
    Plus,
    Pull(EightDirectionTransition),
    Push(SideDirectionTransition),
    Random,
    RandomBar(OrientationTransition),
    Split(SplitTransition),
    Strips(CornerDirectionTransition),
    Wedge,
    Wheel(WheelTransition),
    Wipe(SideDirectionTransition),
    Zoom(InOutTransition),
}

pub struct TransitionStartSoundAction {
    pub sound_file: ::drawingml::EmbeddedWAVAudioFile,
    pub is_looping:  Option<bool>, // false
}

pub enum TransitionSoundAction {
    StartSound(TransitionStartSoundAction),
    EndSound,
}

pub struct SlideTransition {
    pub transition_type: Option<SlideTransitionGroup>,
    pub sound_action: Option<TransitionSoundAction>,
    pub speed: Option<TransitionSpeed>, // fast
    pub advance_on_click: Option<bool>, // true
    pub advance_on_time: Option<u32>,
}

pub struct SlideTiming {
    pub time_node_list: Vec<TimeNodeGroup>,
    pub build_list: Vec<Build>,
}

pub enum Build {
    Paragraph(TLBuildParagraph),
    Diagram(TLBuildDiagram),
    OleChart(TLOleBuildChart),
    Graphic(TLGraphicalObjectBuild),
}

pub struct TLBuildParagraph {
    pub template_list: Vec<TLTemplate>, // size: 0-9
    pub build_common: TLBuildCommonAttributes,
    pub build_type: Option<TLParaBuildType>, // whole
    pub build_level: Option<u32>, // 1
    pub animate_bg: Option<bool>, // false
    pub auto_update_anim_bg: Option<bool>, // true
    pub reverse: Option<bool>, // false
    pub auto_advance_time: Option<TLTime>, // indefinite
}

pub struct TLPoint {
    pub x: ::drawingml::Percentage,
    pub y: ::drawingml::Percentage,
}

pub enum TLTime {
    TimePoint(u32),
    Indefinite,
}

pub struct TLTemplate {
    pub time_node_list: Vec<TimeNodeGroup>,
    pub level: Option<u32>, // 0
}

pub struct TLBuildCommonAttributes {
    pub shape_id: ::drawingml::DrawingElementId,
    pub group_id: u32,
    pub ui_expand: Option<bool>, // false
}

pub struct TLBuildDiagram {
    pub build_common: TLBuildCommonAttributes,
    pub build_type: Option<TLDiagramBuildType>, // whole
}

pub struct TLOleBuildChart {
    pub build_common: TLBuildCommonAttributes,
    pub build_type: Option<TLOleChartBuildType>, // allAtOnce
    pub animate_bg: Option<bool>, // true
}

pub struct TLGraphicalObjectBuild {
    pub build_choice: TLGraphicalObjectBuildChoice,
    pub build_common: TLBuildCommonAttributes,
}

pub enum TLGraphicalObjectBuildChoice {
    BuildAsOne,
    BuildSubElements(::drawingml::AnimationGraphicalObjectBuildProperties),
}

pub enum TimeNodeGroup {
    Parallel(TLCommonTimeNodeData),
    Sequence(TLTimeNodeSequence),
    Exclusive(TLCommonTimeNodeData),
    Animate(TLAnimateBehavior),
    AnimateColor(TLAnimateColorBehavior),
    AnimateEffect(TLAnimateEffectBehavior),
    AnimateMotion(TLAnimateMotionBehavior),
    AnimateRotation(TLAnimateRotationBehavior),
    AnimateScale(TLAnimateScaleBehavior),
    Command(TLCommandBehavior),
    Set(TLSetBehavior),
    Audio(TLMediaNodeAudio),
    Video(TLMediaNodeVideo),
}

pub struct TLTimeNodeSequence {
    pub common_time_node_data: TLCommonTimeNodeData,
    pub prev_condition_list: Vec<TLTimeCondition>,
    pub next_condition_list: Vec<TLTimeCondition>,
    pub concurrent: Option<bool>,
    pub prev_action_type: Option<TLPreviousActionType>,
    pub next_action_type: Option<TLNextActionType>,
}

pub struct TLAnimateBehavior {
    pub common_behavior_data: TLCommonBehaviorData,
    pub time_animate_value_list: Vec<TLTimeAnimateValue>,
    pub by: Option<String>,
    pub from: Option<String>,
    pub to: Option<String>,
    pub calc_mode: Option<TLAnimateBehaviorCalcMode>,
    pub value_type: Option<TLAnimateBehaviorValueType>,
}

pub struct TLAnimateColorBehavior {
    pub common_behavior_data: TLCommonBehaviorData,
    pub by: Option<TLByAnimateColorTransform>,
    pub from: Option<::drawingml::Color>,
    pub to: Option<::drawingml::Color>,
    pub color_space: Option<TLAnimateColorSpace>,
    pub direction: Option<TLAnimateColorDirection>,
}

pub struct TLAnimateEffectBehavior {
    pub common_behavior_data: TLCommonBehaviorData,
    pub progress: Option<TLAnimVariant>,
    pub transition: Option<TLAnimateEffectTransition>,
    pub filter: Option<String>,
    pub property_list: Option<String>,
}

pub struct TLAnimateMotionBehavior {
    pub common_behavior_data: TLCommonBehaviorData,
    pub by: Option<TLPoint>,
    pub from: Option<TLPoint>,
    pub to: Option<TLPoint>,
    pub rotation_center: Option<TLPoint>,
    pub origin: Option<TLAnimateMotionBehaviorOrigin>,
    pub path: Option<String>,
    pub path_edit_mode: Option<TLAnimateMotionPathEditMode>,
    pub rotate_angle: Option<::drawingml::Angle>,
    pub points_types: Option<String>,
}

pub struct TLAnimateRotationBehavior {
    pub common_behavior_data: TLCommonBehaviorData,
    pub by: Option<::drawingml::Angle>,
    pub from: Option<::drawingml::Angle>,
    pub to: Option<::drawingml::Angle>,
}

pub struct TLAnimateScaleBehavior {
    pub common_behavior_data: TLCommonBehaviorData,
    pub by: Option<TLPoint>,
    pub from: Option<TLPoint>,
    pub to: Option<TLPoint>,
    pub zoom_contents: Option<bool>,
}

pub struct TLCommandBehavior {
    pub common_behavior_data: TLCommonBehaviorData,
    pub command_type: Option<TLCommandType>,
    pub command: Option<String>,
}

pub struct TLSetBehavior {
    pub common_behavior_data: TLCommonBehaviorData,
    pub to: Option<TLAnimVariant>,
}

pub struct TLMediaNodeAudio {
    pub common_media_node_data: TLCommonMediaNodeData,
    pub is_narration: Option<bool>, // false
}

pub struct TLMediaNodeVideo {
    pub common_media_node_data: TLCommonMediaNodeData,
    pub fullscreen: Option<bool>, // false
}

pub struct TLTimeAnimateValue {
    pub value: Option<TLAnimVariant>,
    pub time: Option<TLTimeAnimateValueTime>, // indefinite
    pub formula: Option<String>, // ""
}

pub enum TLTimeAnimateValueTime {
    Percentage(::drawingml::PositiveFixedPercentage),
    Indefinite,
}

pub enum TLAnimVariant {
    Bool(bool),
    Int(i32),
    Float(f32),
    String(String),
    Color(::drawingml::Color),
}

pub struct TLCommonBehaviorData {
    pub common_time_node_data: TLCommonTimeNodeData,
    pub target_element: TLTimeTargetElement,
    pub attr_name_list: Vec<String>,
    pub additive: Option<TLBehaviorAdditiveType>,
    pub accumulate: Option<TLBehaviorAccumulateType>,
    pub transform_type: Option<TLBehaviorTransformType>,
    pub from: Option<String>,
    pub to: Option<String>,
    pub by: Option<String>,
    pub rctx: Option<String>,
    pub override_type: Option<TLBehaviorOverrideType>,
}

pub struct TLCommonMediaNodeData {
    pub common_time_node_data: TLCommonTimeNodeData,
    pub target_element: TLTimeTargetElement,
    pub volume: Option<::drawingml::PositiveFixedPercentage>, // 50000
    pub mute: Option<bool>, // false
    pub number_of_slides: Option<u32>, // 1
    pub show_when_stopped: Option<bool>, // true
}

pub enum TLTimeConditionTriggerGroup {
    TargetElement(TLTimeTargetElement),
    TimeNode(TLTimeNodeId),
    RuntimeNode(TLTriggerRuntimeNode),
}

pub enum TLTimeTargetElement {
    SlideTarget,
    SoundTarget(::drawingml::EmbeddedWAVAudioFile),
    ShapeTarget(TLShapeTargetElement),
    InkTarget(TLSubShapeId),
}

pub struct TLShapeTargetElement {
    pub target: Option<TLShapeTargetElementGroup>,
    pub shape_id: ::drawingml::DrawingElementId,
}

pub enum TLShapeTargetElementGroup {
    Background,
    SubShape(TLSubShapeId),
    OleChartElement(TLOleChartTargetElement),
    TextElement(Option<TLTextTargetElement>),
    GraphicElement(::drawingml::AnimationElementChoice),
}

pub struct TLOleChartTargetElement {
    pub element_type: TLChartSubelementType,
    pub level: Option<u32>, // 0
}

pub enum TLTextTargetElement {
    CharRange(IndexRange),
    ParagraphRange(IndexRange),
}

pub struct TLTimeCondition {
    pub trigger: Option<TLTimeConditionTriggerGroup>,
    pub trigger_event: Option<TLTriggerEvent>,
    pub delay: Option<TLTime>,
}

pub struct TLCommonTimeNodeData {
    pub start_condition_list: Vec<TLTimeCondition>,
    pub end_condition_list: Vec<TLTimeCondition>,
    pub end_sync: Option<TLTimeCondition>,
    pub iterate: Option<TLIterateData>,
    pub child_time_node_list: Vec<TimeNodeGroup>,
    pub sub_time_node_list: Vec<TimeNodeGroup>,
    pub id: Option<TLTimeNodeId>,
    pub preset_id: Option<i32>,
    pub preset_class: Option<TLTimeNodePresetClassType>,
    pub preset_subtype: Option<i32>,
    pub duration: Option<TLTime>,
    pub repeat_count: Option<TLTime>, // 1000
    pub repeat_duration: Option<TLTime>,
    pub speed: Option<::drawingml::Percentage>, // 100000
    pub acceleration: Option<::drawingml::PositiveFixedPercentage>, // 0
    pub deceleration: Option<::drawingml::PositiveFixedPercentage>, // 0
    pub auto_reverse: Option<bool>, // false
    pub restart_type: Option<TLTimeNodeRestartType>,
    pub fill_type: Option<TLTimeNodeFillType>,
    pub sync_behavior: Option<TLTimeNodeSyncType>,
    pub time_filter: Option<String>,
    pub event_filter: Option<String>,
    pub display: Option<bool>,
    pub master_relationship: Option<TLTimeNodeMasterRelation>,
    pub build_level: Option<i32>,
    pub group_id: Option<u32>,
    pub after_effect: Option<bool>,
    pub node_type: Option<TLTimeNodeType>,
    pub node_placeholder: Option<bool>,
}

pub enum TLIterateDataInterval {
    Absolute(TLTime),
    Percent(::drawingml::PositivePercentage),
}

pub struct TLIterateData {
    pub interval: TLIterateDataInterval,
    pub iterate_type: Option<IterateType>, // el
    pub backwards: Option<bool>, //false
}

pub enum TLByAnimateColorTransform {
    Rgb(TLByRgbColorTransform),
    Hsl(TLByHslColorTransform),
}

pub struct TLByRgbColorTransform {
    pub r: ::drawingml::FixedPercentage,
    pub g: ::drawingml::FixedPercentage,
    pub b: ::drawingml::FixedPercentage,
}

pub struct TLByHslColorTransform {
    pub h: ::drawingml::Angle,
    pub s: ::drawingml::FixedPercentage,
    pub l: ::drawingml::FixedPercentage,
}

pub struct ChildSlideData {
    pub color_mapping_override: Option<::drawingml::ColorMappingOverride>,
    pub show_master_shapes: Option<bool>, // true
    pub show_master_placholder_animations: Option<bool>, // true
}

pub struct SlideMaster {
    pub common_slide_data: CommonSlideData,
    pub color_mapping: ::drawingml::ColorMapping,
    pub slide_layout_id_list: Vec<SlideLayoutIdListEntry>,
    pub transition: Option<SlideTransition>,
    pub timing: Option<SlideTiming>,
    pub header_footer: Option<HeaderFooter>,
    pub text_styles: Option<SlideMasterTextStyles>,
    pub preserve: Option<bool>, // false
}

impl SlideMaster {
    pub fn from_zip_file(zip_file: &mut ZipFile) -> Result<Self> {
        let mut xml_string = String::new();
        zip_file.read_to_string(&mut xml_string)?;
        let xml_node = XmlNode::from_str(xml_string.as_str())?;

        Self::from_xml_element(&xml_node)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut opt_common_slide_data = None;
        let mut opt_color_mapping = None;
        let mut slide_layout_id_list = Vec::new();
        let mut transition = None;
        let mut timing = None;
        let mut header_footer = None;
        let mut text_styles = None;
        let mut preserve = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "preserve" => preserve = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cSld" => opt_common_slide_data = Some(CommonSlideData::from_xml_element(child_node)?),
                "clrMap" => opt_color_mapping = Some(::drawingml::ColorMapping::from_xml_element(child_node)?),
                "sldLayoutIdLst" => {
                    for slide_layout_id_node in child_node.child_nodes {
                        match slide_layout_id_node.local_name() {
                            "sldLayoutId" => slide_layout_id_list.push(SlideLayoutIdListEntry::from_xml_element(slide_layout_id_node)?),
                            _ => (),
                        }
                    }
                }
                "transition" => transition = Some(SlideTransition::from_xml_element(child_node)?),
                "timing" => timing = Some(SlideTiming::from_xml_element(child_node)?),
                "hf" => header_footer = Some(HeaderFooter::from_xml_element(child_node)?),
                "txStyles" => text_styles = Some(SlideMasterTextStyles::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_slide_data = opt_common_slide_data.ok_or_else(|| XmlError::from(MissingChildNodeError::new("cSld")))?;
        let color_mapping = opt_color_mapping.ok_or_else(|| XmlError::from(MissingChildNodeError::new("clrMap")))?;

        Ok(Self {
            common_slide_data,
            color_mapping,
            slide_layout_id_list,
            transition,
            timing,
            header_footer,
            text_styles,
            preserve,
        })
    }
}

pub struct SlideLayout {
    pub common_slide_data: CommonSlideData,
    pub child_slide_data: Option<ChildSlideData>,
    pub transition: Option<SlideTransition>,
    pub timing: Option<SlideTiming>,
    pub header_footer: Option<HeaderFooter>,
    pub matching_name: Option<String>, // ""
    pub slide_layout_type: Option<SlideLayoutType>, // cust
    pub preserve: Option<bool>, //false
    pub is_user_drawn: Option<bool>, // false
}

pub struct Slide {
    pub common_slide_data: CommonSlideData,
    pub child_slide_data: Option<ChildSlideData>,
    pub transition: Option<SlideTransition>,
    pub timing: Option<SlideTiming>,
    pub show: Option<bool>, // true
}

/// EmbeddedFontListEntry
pub struct EmbeddedFontListEntry {
    pub font: ::drawingml::TextFont,
    pub regular: Option<RelationshipId>,
    pub bold: Option<RelationshipId>,
    pub italic: Option<RelationshipId>,
    pub bold_italic: Option<RelationshipId>,
}

impl EmbeddedFontListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut font = None;
        let mut regular = None;
        let mut bold = None;
        let mut italic = None;
        let mut bold_italic = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "font" => font = Some(::drawingml::TextFont::from_xml_element(child_node)?),
                "regular" => {
                    let id_attr = child_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                    regular = Some(id_attr.clone());
                }
                "bold" => {
                    let id_attr = child_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                    bold = Some(id_attr.clone());
                }
                "italic" => {
                    let id_attr = child_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                    italic = Some(id_attr.clone());
                }
                "boldItalic" => {
                    let id_attr = child_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                    bold_italic = Some(id_attr.clone());
                }
                _ => (),
            }
        }

        let font = font.ok_or_else(|| MissingChildNodeError::new("font"))?;

        Ok(Self {
            font,
            regular,
            bold,
            italic,
            bold_italic,
        })
    }
}

/// CustomShow
pub struct CustomShow {
    pub name: Name,
    pub id: u32,
    pub slides: Vec<RelationshipId>,
}

impl CustomShow {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<CustomShow> {
        let mut name = None;
        let mut id = None;
        
        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "name" => name = Some(value.clone()),
                "id" => id = Some(value.parse::<u32>()?),
                _ => (),
            }
        }

        let mut slides = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "sldLst" => {
                    for slide_node in &child_node.child_nodes {
                        let id_attr = slide_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                        slides.push(id_attr.clone());
                    }
                }
                _ => (),
            }
        }

        let name = name.ok_or_else(|| MissingAttributeError::new("name"))?;
        let id = id.ok_or_else(|| MissingAttributeError::new("id"))?;

        Ok(Self {
            name,
            id,
            slides,
        })
    }
}

pub struct PhotoAlbum {
    pub black_and_white: Option<bool>, // false
    pub show_captions: Option<bool>, // false
    pub layout: Option<PhotoAlbumLayout>, // PhotoAlbumLayout::FitToSlide
    pub frame: Option<PhotoAlbumFrameShape>, // PhotoAlbumFrameShape::FrameStyle1
}

pub struct HeaderFooter {
    pub slide_number_enabled: Option<bool>, // true
    pub header_enabled: Option<bool>, // true
    pub footer_enabled: Option<bool>, // true
    pub date_time_enabled: Option<bool>, // true
}

pub struct Kinsoku {
    pub language: Option<String>,
    pub invalid_start_chars: String,
    pub invalid_end_chars: String,
}

pub struct ModifyVerifier {
    pub algorithm_name: Option<String>,
    pub hash_value: Option<String>,
    pub salt_value: Option<String>,
    pub spin_value: Option<u32>,
}

/// Presentation
pub struct Presentation {
    pub server_zoom: Option<::drawingml::Percentage>, // 50%
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
    pub conformance: Option<ConformanceClass>,
    pub slide_master_id_list: Vec<SlideMasterIdListEntry>,
    pub notes_master_id_list: Vec<NotesMasterIdListEntry>, // length = 1
    pub handout_master_id_list: Vec<HandoutMasterIdListEntry>, // length = 1
    pub slide_id_list: Vec<SlideIdListEntry>,
    pub slide_size: Option<SlideSize>,
    pub notes_size: Option<::drawingml::PositiveSize2D>,
    pub smart_tags: Option<RelationshipId>,
    pub embedded_font_list: Vec<EmbeddedFontListEntry>,
    pub custom_show_list: Vec<CustomShow>,
    pub photo_album: Option<PhotoAlbum>,
    pub customer_data_list: Option<CustomerDataList>,
    pub kinsoku: Option<Kinsoku>,
    pub default_text_style: Option<::drawingml::TextListStyle>,
    pub modify_verifier: Option<ModifyVerifier>,
}

impl Presentation {
    pub fn from_zip<R>(zipper: &mut zip::ZipArchive<R>) -> Result<Self>
    where
        R: Read + Seek
    {
        let mut presentation_file = zipper.by_name("ppt/presentation.xml")?;
        let mut xml_string = String::new();
        presentation_file.read_to_string(&mut xml_string)?;

        let root = XmlNode::from_str(xml_string.as_str())?;
        Self::from_xml_element(&root)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut server_zoom = None;
        let mut first_slide_num = None;
        let mut show_special_pls_on_title_slide = None;
        let mut rtl = None;
        let mut remove_personal_info_on_save = None;
        let mut compatibility_mode = None;
        let mut strict_first_and_last_chars = None;
        let mut embed_true_type_fonts = None;
        let mut save_subset_fonts = None;
        let mut auto_compress_pictures = None;
        let mut bookmark_id_seed = None;
        let mut conformance = None;

        for (attr, value) in &xml_node.attributes {
            println!("parsing presentation attribute: {}", attr);
            match attr.as_str() {
                "serverZoom" => server_zoom = Some(value.parse()?),
                "firstSlideNum" => first_slide_num = Some(value.parse()?),
                "showSpecialPlsOnTitleSld" => show_special_pls_on_title_slide = Some(parse_xml_bool(value)?),
                "rtl" => rtl = Some(parse_xml_bool(value)?),
                "removePersonalInfoOnSave" => remove_personal_info_on_save = Some(parse_xml_bool(value)?),
                "compatMode" => compatibility_mode = Some(parse_xml_bool(value)?),
                "strictFirstAndLastChars" => strict_first_and_last_chars = Some(parse_xml_bool(value)?),
                "embedTrueTypeFonts" => embed_true_type_fonts = Some(parse_xml_bool(value)?),
                "saveSubsetFonts" => save_subset_fonts = Some(parse_xml_bool(value)?),
                "autoCompressPictures" => auto_compress_pictures = Some(parse_xml_bool(value)?),
                "bookmarkIdSeed" => bookmark_id_seed = Some(value.parse()?),
                "conformance" => conformance = Some(value.parse()?),
                _ => (),
            }
        }
        
        let mut slide_master_id_list = Vec::new();
        let mut notes_master_id_list = Vec::new();
        let mut handout_master_id_list = Vec::new();
        let mut slide_id_list = Vec::new();
        let mut slide_size = None;
        let mut notes_size = None;
        let mut smart_tags = None;
        let mut embedded_font_list = Vec::new();
        let mut custom_show_list = Vec::new();
        let mut photo_album = None;
        let mut customer_data_list = None;
        let mut kinsoku = None;
        let mut default_text_style = None;
        let mut modify_verifier = None;

        for child_node in &xml_node.child_nodes {
            println!("parsing child_node: {}", child_node);
            match child_node.local_name() {
                "sldMasterIdLst" => {
                    for slide_master_id_node in &child_node.child_nodes {
                        let mut id = None;
                        let mut relationship_id = None;

                        for (attr, value) in &slide_master_id_node.attributes {
                            match attr.as_str() {
                                "id" => id = Some(value.parse::<u32>()?),
                                "r:id" => relationship_id =  Some(value.clone()),
                                _ => (),
                            }
                        }

                        // r:id attribute is required
                        let relationship_id = relationship_id.ok_or_else(|| MissingAttributeError::new("r:id"))?;
                        slide_master_id_list.push(SlideMasterIdListEntry {
                            id,
                            relationship_id,
                        });
                    }
                },
                "notesMasterIdLst" => {
                    for notes_master_id_node in &child_node.child_nodes {
                        let r_id_attr = notes_master_id_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                        notes_master_id_list.push(NotesMasterIdListEntry {
                            relationship_id: r_id_attr.clone(),
                        });
                    }
                },
                "handoutMasterIdLst" => {
                    for notes_master_id_node in &child_node.child_nodes {
                        let r_id_attr = notes_master_id_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                        handout_master_id_list.push(HandoutMasterIdListEntry {
                            relationship_id: r_id_attr.clone(),
                        });
                    }
                },
                "sldIdLst" => {
                    for slide_id_node in &child_node.child_nodes {
                        let mut id = None;
                        let mut relationship_id = None;

                        for (attr, value) in &slide_id_node.attributes {
                            match attr.as_str() {
                                "id" => id = Some(value.parse::<u32>()?),
                                "r:id" => relationship_id = Some(value.clone()),
                                _ => (),
                            }
                        }

                        let id = id.ok_or_else(|| MissingAttributeError::new("id"))?;
                        let relationship_id = relationship_id.ok_or_else(|| MissingAttributeError::new("r:id"))?;

                        slide_id_list.push(SlideIdListEntry {
                            id,
                            relationship_id,
                        });
                    }
                },
                "sldSz" => {
                    let mut width = None;
                    let mut height = None;
                    let mut size_type = None;

                    for (attr, value) in &child_node.attributes {
                        match attr.as_str() {
                            "cx" => width = Some(value.parse()?),
                            "cy" => height = Some(value.parse()?),
                            "type" => size_type = Some(value.parse()?),
                            _ => (),
                        }
                    }

                    let width = width.ok_or_else(|| MissingAttributeError::new("cx"))?;
                    let height = height.ok_or_else(|| MissingAttributeError::new("cy"))?;

                    slide_size = Some(SlideSize {
                        width,
                        height,
                        size_type,
                    })
                }
                "notesSz" => notes_size = Some(::drawingml::PositiveSize2D::from_xml_element(child_node)?),
                "smartTags" => {
                    let r_id_attr = child_node.attribute("r:id").ok_or_else(|| MissingAttributeError::new("r:id"))?;
                    smart_tags = Some(r_id_attr.clone());
                }
                "embeddedFontLst" => {
                    for embedded_font_node in &child_node.child_nodes {
                        embedded_font_list.push(EmbeddedFontListEntry::from_xml_element(embedded_font_node)?);
                    }
                }
                "custShowLst" => {
                    for custom_show_node in &child_node.child_nodes {
                        custom_show_list.push(CustomShow::from_xml_element(custom_show_node)?);
                    }
                }
                "photoAlbum" => {
                    let mut black_and_white = None;
                    let mut frame = None;
                    let mut layout = None;
                    let mut show_captions = None;

                    for (attr, value) in &child_node.attributes {
                        match attr.as_str() {
                            "bw" => black_and_white = Some(value.parse()?),
                            "showCaptions" => show_captions = Some(value.parse()?),
                            "layout" => layout = Some(value.parse()?),
                            "frame" => frame = Some(value.parse()?),
                            _ => (),
                        }
                    }

                    photo_album = Some(PhotoAlbum {
                        black_and_white,
                        frame,
                        layout,
                        show_captions,
                    });
                }
                "custDataLst" => customer_data_list = Some(CustomerDataList::from_xml_element(child_node)?),
                "kinsoku" => {
                    let mut language = None;
                    let mut invalid_start_chars = None;
                    let mut invalid_end_chars = None;

                    for (attr, value) in &child_node.attributes {
                        match attr.as_str() {
                            "lang" => language = Some(value.clone()),
                            "invalStChars" => invalid_start_chars = Some(value.clone()),
                            "invalEndChars" => invalid_end_chars = Some(value.clone()),
                            _ => (),
                        }
                    }

                    let invalid_start_chars = invalid_start_chars.ok_or_else(|| MissingAttributeError::new("invalStChars"))?;
                    let invalid_end_chars = invalid_end_chars.ok_or_else(|| MissingAttributeError::new("invalEndChars"))?;

                    kinsoku = Some(Kinsoku {
                        language,
                        invalid_start_chars,
                        invalid_end_chars,
                    });
                }
                "defaultTextStyle" => default_text_style = Some(::drawingml::TextListStyle::from_xml_element(child_node)?),
                _ => (),
            }
        }

/*
		for (const MXmlNode2 &childNode : xmlNode)
		{
			//else if (childNode.name == mT("modifyVerifier"))
			//	instance->modifyVerifier.reset(ModifyVerifier::FromXmlNode(childNode));
			//else if (childNode.name == mT("extLst"))
			//	instance->extLst.reset(ExtensionList::FromXmlNode(childNode));
		}
                            */


        Ok(Self {
            server_zoom,
            first_slide_num,
            show_special_pls_on_title_slide,
            rtl,
            remove_personal_info_on_save,
            compatibility_mode,
            strict_first_and_last_chars,
            embed_true_type_fonts,
            save_subset_fonts,
            auto_compress_pictures,
            bookmark_id_seed,
            conformance,
            slide_master_id_list,
            notes_master_id_list,
            handout_master_id_list,
            slide_id_list,
            slide_size,
            notes_size,
            smart_tags,
            embedded_font_list,
            custom_show_list,
            photo_album,
            customer_data_list,
            kinsoku,
            default_text_style,
            modify_verifier,
        })
    }
}