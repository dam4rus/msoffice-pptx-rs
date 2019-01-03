use crate::error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError, XmlError};
use crate::relationship::RelationshipId;
use crate::xml::{parse_xml_bool, XmlNode};
use ::std::io::{Read, Seek};
use ::std::str::FromStr;
use ::zip::read::ZipFile;

pub type SlideId = u32; // TODO: 256 <= n <= 2147483648
pub type SlideLayoutId = u32; // TODO: 2147483648 <= n
pub type SlideMasterId = u32; // TODO: 2147483648 <= n
pub type Index = u32;
pub type TLTimeNodeId = u32;
pub type BookmarkIdSeed = u32; // TODO: 1 <= n <= 2147483648
pub type SlideSizeCoordinate = crate::drawingml::PositiveCoordinate32; // TODO: 914400 <= n <= 51206400
pub type Name = String;
pub type TLSubShapeId = crate::drawingml::ShapeId;

pub type Result<T> = ::std::result::Result<T, Box<dyn (::std::error::Error)>>;

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

#[derive(Debug, Clone)]
pub struct IndexRange {
    pub start: Index,
    pub end: Index,
}

impl IndexRange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut start = None;
        let mut end = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "st" => start = Some(value.parse()?),
                "end" => end = Some(value.parse()?),
                _ => ()
            }
        }

        let start = start.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "st"))?;
        let end = end.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "end"))?;

        Ok(Self { start, end })
    }
}

#[derive(Debug, Clone)]
pub struct BackgroundProperties {
    pub shade_to_title: Option<bool>, // false
    pub fill: crate::drawingml::FillProperties,
    pub effect: Option<crate::drawingml::EffectProperties>,
}

impl BackgroundProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let shade_to_title = match xml_node.attribute("shadeToTitle") {
            Some(val) => Some(parse_xml_bool(val)?),
            None => None,
        };

        let mut fill = None;
        let mut effect = None;

        for child_node in &xml_node.child_nodes {
            use crate::drawingml::{FillProperties, EffectProperties};

            if FillProperties::is_choice_member(child_node.local_name()) {
                fill = Some(FillProperties::from_xml_element(child_node)?);
            } else if EffectProperties::is_choice_member(child_node.local_name()) {
                effect = Some(EffectProperties::from_xml_element(child_node)?);
            }
        }

        let fill = fill.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_FillProperties"))?;

        Ok(Self {
            shade_to_title,
            fill,
            effect,
        })
    }
}

#[derive(Debug, Clone)]
pub enum BackgroundGroup {
    Properties(BackgroundProperties),
    Reference(crate::drawingml::StyleMatrixReference),
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
            "bgPr" => Ok(BackgroundGroup::Properties(BackgroundProperties::from_xml_element(
                xml_node,
            )?)),
            "bgRef" => Ok(BackgroundGroup::Reference(
                crate::drawingml::StyleMatrixReference::from_xml_element(xml_node)?,
            )),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_Background").into()),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Background {
    pub black_and_white_mode: Option<crate::drawingml::BlackWhiteMode>, // white
    pub background: BackgroundGroup,
}

impl Background {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let black_and_white_mode = match xml_node.attribute("bwMode") {
            Some(val) => Some(val.parse()?),
            None => None,
        };

        let background_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_Background"))?;
        let background = BackgroundGroup::from_xml_element(background_node)?;

        Ok(Self {
            background,
            black_and_white_mode,
        })
    }
}

#[derive(Default, Debug, Clone)]
pub struct Placeholder {
    pub placeholder_type: Option<PlaceholderType>, // obj
    pub orientation: Option<Direction>,            // horz
    pub size: Option<PlaceholderSize>,             // full
    pub index: Option<u32>,                        // 0
    pub has_custom_prompt: Option<bool>,           // false
}

impl Placeholder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => instance.placeholder_type = Some(value.parse()?),
                "orient" => instance.orientation = Some(value.parse()?),
                "sz" => instance.size = Some(value.parse()?),
                "idx" => instance.index = Some(value.parse()?),
                "hasCustomPrompt" => instance.has_custom_prompt = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Default, Debug, Clone)]
pub struct ApplicationNonVisualDrawingProps {
    pub is_photo: Option<bool>,      // false
    pub is_user_drawn: Option<bool>, // false
    pub placeholder: Option<Placeholder>,
    pub media: Option<crate::drawingml::Media>,
    pub customer_data_list: Option<CustomerDataList>,
}

impl ApplicationNonVisualDrawingProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "isPhoto" => instance.is_photo = Some(parse_xml_bool(value)?),
                "userDrawn" => instance.is_user_drawn = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            use crate::drawingml::Media;

            let local_name = child_node.local_name();
            if Media::is_choice_member(local_name) {
                instance.media = Some(Media::from_xml_element(child_node)?);
            } else {
                match child_node.local_name() {
                    "ph" => instance.placeholder = Some(Placeholder::from_xml_element(child_node)?),
                    "custDataLst" => instance.customer_data_list = Some(
                        CustomerDataList::from_xml_element(child_node)?
                    ),
                    _ => (),
                }
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone)]
pub enum ShapeGroup {
    Shape(Box<Shape>),
    GroupShape(Box<GroupShape>),
    GraphicFrame(Box<GraphicalObjectFrame>),
    Connector(Box<Connector>),
    Picture(Box<Picture>),
    ContentPart(RelationshipId),
}

impl ShapeGroup {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "sp" | "grpSp" | "graphicFrame" | "cxnSp" | "pic" | "contentPart" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "sp" => Ok(ShapeGroup::Shape(Box::new(Shape::from_xml_element(xml_node)?))),
            "grpSp" => Ok(ShapeGroup::GroupShape(Box::new(GroupShape::from_xml_element(
                xml_node,
            )?))),
            "graphicFrame" => Ok(ShapeGroup::GraphicFrame(Box::new(GraphicalObjectFrame::from_xml_element(xml_node)?))),
            "cxnSp" => Ok(ShapeGroup::Connector(Box::new(Connector::from_xml_element(
                xml_node,
            )?))),
            "pic" => Ok(ShapeGroup::Picture(Box::new(Picture::from_xml_element(xml_node)?))),
            "contentPart" => {
                let attr = xml_node
                    .attribute("r:id")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;
                Ok(ShapeGroup::ContentPart(attr.clone()))
            }
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "EG_ShapeGroup"))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Shape {
    pub use_bg_fill: Option<bool>, // false
    pub non_visual_props: Box<ShapeNonVisual>,
    pub shape_props: Box<crate::drawingml::ShapeProperties>,
    pub style: Option<Box<crate::drawingml::ShapeStyle>>,
    pub text_body: Option<crate::drawingml::TextBody>,
}

impl Shape {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let use_bg_fill = match xml_node.attribute("useBgFill") {
            Some(val) => Some(parse_xml_bool(val)?),
            None => None,
        };

        let mut non_visual_props = None;
        let mut shape_props = None;
        let mut style = None;
        let mut text_body = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvSpPr" => non_visual_props = Some(Box::new(ShapeNonVisual::from_xml_element(child_node)?)),
                "spPr" => {
                    shape_props = Some(Box::new(crate::drawingml::ShapeProperties::from_xml_element(
                        child_node,
                    )?))
                }
                "style" => style = Some(Box::new(crate::drawingml::ShapeStyle::from_xml_element(child_node)?)),
                "txBody" => text_body = Some(crate::drawingml::TextBody::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvSpPr"))?;
        let shape_props = shape_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "spPr"))?;

        Ok(Self {
            use_bg_fill,
            non_visual_props,
            shape_props,
            style,
            text_body,
        })
    }
}

#[derive(Debug, Clone)]
pub struct ShapeNonVisual {
    pub drawing_props: Box<crate::drawingml::NonVisualDrawingProps>,
    pub shape_drawing_props: crate::drawingml::NonVisualDrawingShapeProps,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl ShapeNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut shape_drawing_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => {
                    drawing_props = Some(Box::new(crate::drawingml::NonVisualDrawingProps::from_xml_element(
                        child_node,
                    )?))
                }
                "cNvSpPr" => {
                    shape_drawing_props = Some(crate::drawingml::NonVisualDrawingShapeProps::from_xml_element(
                        child_node,
                    )?)
                }
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let shape_drawing_props =
            shape_drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvSpPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;

        Ok(Self {
            drawing_props,
            shape_drawing_props,
            app_props,
        })
    }
}

#[derive(Debug, Clone)]
pub struct GroupShape {
    pub non_visual_props: Box<GroupShapeNonVisual>,
    pub group_shape_props: crate::drawingml::GroupShapeProperties,
    pub shape_array: Vec<ShapeGroup>,
}

impl GroupShape {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_props = None;
        let mut group_shape_props = None;
        let mut shape_array = Vec::new();

        for child_node in &xml_node.child_nodes {
            let child_local_name = child_node.local_name();
            if ShapeGroup::is_choice_member(child_local_name) {
                shape_array.push(ShapeGroup::from_xml_element(child_node)?);
            } else {
                match child_local_name {
                    "nvGrpSpPr" => non_visual_props = Some(Box::new(GroupShapeNonVisual::from_xml_element(child_node)?)),
                    "grpSpPr" => {
                        group_shape_props = Some(crate::drawingml::GroupShapeProperties::from_xml_element(child_node)?)
                    }
                    _ => (),
                }
            }
        }

        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvGrpSpPr"))?;
        let group_shape_props =
            group_shape_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "grpSpPr"))?;

        Ok(Self {
            non_visual_props,
            group_shape_props,
            shape_array,
        })
    }
}

#[derive(Debug, Clone)]
pub struct GroupShapeNonVisual {
    pub drawing_props: Box<crate::drawingml::NonVisualDrawingProps>,
    pub group_drawing_props: crate::drawingml::NonVisualGroupDrawingShapeProps,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl GroupShapeNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut group_drawing_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => {
                    drawing_props = Some(Box::new(crate::drawingml::NonVisualDrawingProps::from_xml_element(
                        child_node,
                    )?))
                }
                "cNvGrpSpPr" => {
                    group_drawing_props = Some(crate::drawingml::NonVisualGroupDrawingShapeProps::from_xml_element(
                        child_node,
                    )?)
                }
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let group_drawing_props =
            group_drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvGrpSpPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;
        Ok(Self {
            drawing_props,
            group_drawing_props,
            app_props,
        })
    }
}

#[derive(Debug, Clone)]
pub struct GraphicalObjectFrame {
    pub black_white_mode: Option<crate::drawingml::BlackWhiteMode>,
    pub non_visual_props: Box<GraphicalObjectFrameNonVisual>,
    pub transform: Box<crate::drawingml::Transform2D>,
    pub graphic: crate::drawingml::GraphicalObject,
}

impl GraphicalObjectFrame {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let black_white_mode = match xml_node.attribute("bwMode") {
            Some(attr) => Some(attr.parse()?),
            None => None,
        };

        let mut non_visual_props = None;
        let mut transform = None;
        let mut graphic = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvGraphicFramePr" => non_visual_props = Some(
                    Box::new(GraphicalObjectFrameNonVisual::from_xml_element(child_node)?)
                ),
                "xfrm" => transform = Some(Box::new(crate::drawingml::Transform2D::from_xml_element(child_node)?)),
                "graphic" => graphic = Some(crate::drawingml::GraphicalObject::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let non_visual_props = 
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvGraphicFramePr"))?;
        let transform = transform.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "xfrm"))?;
        let graphic = graphic.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "graphic"))?;
        
        Ok(Self {
            black_white_mode,
            non_visual_props,
            transform,
            graphic,
        })
    }
}

#[derive(Debug, Clone)]
pub struct GraphicalObjectFrameNonVisual {
    pub drawing_props: Box<crate::drawingml::NonVisualDrawingProps>,
    pub graphic_frame_props: crate::drawingml::NonVisualGraphicFrameProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl GraphicalObjectFrameNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut graphic_frame_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => drawing_props = Some(Box::new(
                    crate::drawingml::NonVisualDrawingProps::from_xml_element(child_node)?
                )),
                "cNvGraphicFramePr" => graphic_frame_props = Some(
                    crate::drawingml::NonVisualGraphicFrameProperties::from_xml_element(child_node)?
                ),
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = 
            drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let graphic_frame_props =
            graphic_frame_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvGraphicFramePr"))?;
        let app_props =
            app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;

        Ok(Self { drawing_props, graphic_frame_props, app_props })
    }
}

#[derive(Debug, Clone)]
pub struct Connector {
    pub non_visual_props: Box<ConnectorNonVisual>,
    pub shape_props: Box<crate::drawingml::ShapeProperties>,
    pub shape_style: Option<Box<crate::drawingml::ShapeStyle>>,
}

impl Connector {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_props = None;
        let mut shape_props = None;
        let mut shape_style = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvCxnSpPr" => non_visual_props = Some(Box::new(ConnectorNonVisual::from_xml_element(child_node)?)),
                "spPr" => {
                    shape_props = Some(Box::new(crate::drawingml::ShapeProperties::from_xml_element(
                        child_node,
                    )?))
                }
                "style" => shape_style = Some(Box::new(crate::drawingml::ShapeStyle::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvCxnSpPr"))?;
        let shape_props = shape_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvCxnSpPr"))?;

        Ok(Self {
            non_visual_props,
            shape_props,
            shape_style,
        })
    }
}

#[derive(Debug, Clone)]
pub struct ConnectorNonVisual {
    pub drawing_props: Box<crate::drawingml::NonVisualDrawingProps>,
    pub connector_props: crate::drawingml::NonVisualConnectorProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl ConnectorNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut connector_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => {
                    drawing_props = Some(Box::new(crate::drawingml::NonVisualDrawingProps::from_xml_element(
                        child_node,
                    )?))
                }
                "cNvCxnSpPr" => {
                    connector_props = Some(crate::drawingml::NonVisualConnectorProperties::from_xml_element(
                        child_node,
                    )?)
                }
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let connector_props =
            connector_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvCxnSpPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;

        Ok(Self {
            drawing_props,
            connector_props,
            app_props,
        })
    }
}

#[derive(Debug, Clone)]
pub struct Picture {
    pub non_visual_props: Box<PictureNonVisual>,
    pub blip_fill: Box<crate::drawingml::BlipFillProperties>,
    pub shape_props: Box<crate::drawingml::ShapeProperties>,
    pub shape_style: Option<Box<crate::drawingml::ShapeStyle>>,
}

impl Picture {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_props = None;
        let mut blip_fill = None;
        let mut shape_props = None;
        let mut shape_style = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvPicPr" => non_visual_props = Some(Box::new(PictureNonVisual::from_xml_element(child_node)?)),
                "blipFill" => {
                    blip_fill = Some(Box::new(crate::drawingml::BlipFillProperties::from_xml_element(
                        child_node,
                    )?))
                }
                "spPr" => {
                    shape_props = Some(Box::new(crate::drawingml::ShapeProperties::from_xml_element(
                        child_node,
                    )?))
                }
                "style" => shape_style = Some(Box::new(crate::drawingml::ShapeStyle::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPicPr"))?;
        let blip_fill = blip_fill.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "blipFill"))?;
        let shape_props = shape_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "spPr"))?;

        Ok(Self {
            non_visual_props,
            blip_fill,
            shape_props,
            shape_style,
        })
    }
}

#[derive(Debug, Clone)]
pub struct PictureNonVisual {
    pub drawing_props: Box<crate::drawingml::NonVisualDrawingProps>,
    pub picture_props: crate::drawingml::NonVisualPictureProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl PictureNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut picture_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => {
                    drawing_props = Some(Box::new(crate::drawingml::NonVisualDrawingProps::from_xml_element(
                        child_node,
                    )?))
                }
                "cNvPicPr" => {
                    picture_props = Some(crate::drawingml::NonVisualPictureProperties::from_xml_element(
                        child_node,
                    )?)
                }
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let picture_props =
            picture_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPicPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;

        Ok(Self {
            drawing_props,
            picture_props,
            app_props,
        })
    }
}

#[derive(Debug, Clone)]
pub struct CommonSlideData {
    pub name: Option<String>,
    pub background: Option<Box<Background>>,
    pub shape_tree: Box<GroupShape>,
    pub customer_data_list: Option<CustomerDataList>,
    pub control_list: Vec<Control>,
}

impl CommonSlideData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name = xml_node.attribute("name").cloned();
        let mut background = None;
        let mut opt_shape_tree = None;
        let mut customer_data_list = None;
        let mut control_list = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "bg" => background = Some(Box::new(Background::from_xml_element(child_node)?)),
                "spTree" => opt_shape_tree = Some(Box::new(GroupShape::from_xml_element(child_node)?)),
                "custDataList" => customer_data_list = Some(CustomerDataList::from_xml_element(child_node)?),
                "controls" => {
                    for control_node in &child_node.child_nodes {
                        control_list.push(Control::from_xml_element(control_node)?);
                    }
                }
                _ => (),
            }
        }

        let shape_tree = opt_shape_tree
            .ok_or_else(|| XmlError::from(MissingChildNodeError::new(xml_node.name.clone(), "spTree")))?;

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
#[derive(Default, Debug, Clone)]
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
                    let id_attr = child_node
                        .attribute("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?;
                    customer_data_list.push(id_attr.clone());
                }
                "tags" => {
                    let id_attr = child_node
                        .attribute("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?;
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

#[derive(Default, Debug, Clone)]
pub struct Control {
    pub picture: Option<Box<Picture>>,
    pub ole_attributes: Box<OleAttributes>,
}

impl Control {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            instance.ole_attributes.try_attribute_parse(attr, value)?;
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "pic" => instance.picture = Some(Box::new(Picture::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Default, Debug, Clone)]
pub struct OleAttributes {
    pub shape_id: Option<crate::drawingml::ShapeId>,
    pub name: Option<String>,       // ""
    pub show_as_icon: Option<bool>, // false
    pub id: Option<RelationshipId>,
    pub image_width: Option<crate::drawingml::PositiveCoordinate32>,
    pub image_height: Option<crate::drawingml::PositiveCoordinate32>,
}

impl OleAttributes {
    pub fn try_attribute_parse<T>(&mut self, attr: T, value: &String) -> Result<()>
    where
        T: AsRef<str>,
    {
        match attr.as_ref() {
            "spid" => self.shape_id = Some(value.parse()?),
            "name" => self.name = Some(value.clone()),
            "showAsIcon" => self.show_as_icon = Some(parse_xml_bool(value)?),
            "r:id" => self.id = Some(value.clone()),
            "imgW" => self.image_width = Some(value.parse()?),
            "imgH" => self.image_height = Some(value.parse()?),
            _ => (),
        }

        Ok(())
    }
}

#[derive(Debug, Clone)]
pub struct SlideSize {
    pub width: SlideSizeCoordinate,
    pub height: SlideSizeCoordinate,
    pub size_type: Option<SlideSizeType>,
}

impl SlideSize {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut width = None;
        let mut height = None;
        let mut size_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "cx" => width = Some(value.parse()?),
                "cy" => height = Some(value.parse()?),
                "type" => size_type = Some(value.parse()?),
                _ => (),
            }
        }

        let width = width.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "cx"))?;
        let height = height.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "cy"))?;

        Ok(Self {
            width,
            height,
            size_type,
        })
    }
}

#[derive(Debug, Clone)]
pub struct SlideIdListEntry {
    pub id: SlideId,
    pub relationship_id: RelationshipId,
}

impl SlideIdListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut relationship_id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => id = Some(value.parse()?),
                "r:id" => relationship_id = Some(value.clone()),
                _ => (),
            }
        }

        let id = id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;
        let relationship_id =
            relationship_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self { id, relationship_id })
    }
}

#[derive(Debug, Clone)]
pub struct SlideLayoutIdListEntry {
    pub id: Option<SlideLayoutId>,
    pub relationship_id: RelationshipId,
}

impl SlideLayoutIdListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut relationship_id = None;
        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => id = Some(value.parse()?),
                "r:id" => relationship_id = Some(value.clone()),
                _ => (),
            }
        }

        let relationship_id =
            relationship_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self { id, relationship_id })
    }
}

#[derive(Debug, Clone)]
pub struct SlideMasterIdListEntry {
    pub id: Option<SlideMasterId>,
    pub relationship_id: RelationshipId,
}

impl SlideMasterIdListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut relationship_id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => id = Some(value.parse()?),
                "r:id" => relationship_id = Some(value.clone()),
                _ => (),
            }
        }

        let relationship_id = relationship_id
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self { id, relationship_id })
    }
}

#[derive(Debug, Clone)]
pub struct NotesMasterIdListEntry {
    pub relationship_id: RelationshipId,
}

impl NotesMasterIdListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let r_id_attr = xml_node
            .attribute("r:id")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self { relationship_id: r_id_attr.clone() })
    }
}

#[derive(Debug, Clone)]
pub struct HandoutMasterIdListEntry {
    pub relationship_id: RelationshipId,
}

impl HandoutMasterIdListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let r_id_attr = xml_node
            .attribute("r:id")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self { relationship_id: r_id_attr.clone() })
    }
}

#[derive(Default, Debug, Clone)]
pub struct SlideMasterTextStyles {
    pub title_styles: Option<Box<crate::drawingml::TextListStyle>>,
    pub body_styles: Option<Box<crate::drawingml::TextListStyle>>,
    pub other_styles: Option<Box<crate::drawingml::TextListStyle>>,
}

impl SlideMasterTextStyles {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "titleStyle" => instance.title_styles = Some(Box::new(
                    crate::drawingml::TextListStyle::from_xml_element(child_node)?
                )),
                "bodyStyle" => instance.body_styles = Some(Box::new(
                    crate::drawingml::TextListStyle::from_xml_element(child_node)?
                )),
                "otherStyle" => instance.other_styles = Some(Box::new(
                    crate::drawingml::TextListStyle::from_xml_element(child_node)?
                )),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Default, Debug, Clone)]
pub struct OrientationTransition {
    pub direction: Option<Direction>, // horz
}

impl OrientationTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let direction = match xml_node.attribute("dir") {
            Some(value) => Some(value.parse()?),
            None => None,
        };

        Ok(Self { direction })
    }
}

#[derive(Default, Debug, Clone)]
pub struct EightDirectionTransition {
    pub direction: Option<TransitionEightDirectionType>, // l
}

impl EightDirectionTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let direction = match xml_node.attribute("dir") {
            Some(value) => Some(value.parse()?),
            None => None,
        };

        Ok(Self { direction })
    }
}

#[derive(Default, Debug, Clone)]
pub struct OptionalBlackTransition {
    pub through_black: Option<bool>, // false
}

impl OptionalBlackTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let through_black = match xml_node.attribute("thruBlk") {
            Some(value) => Some(parse_xml_bool(value)?),
            None => None,
        };

        Ok(Self { through_black })
    }
}

#[derive(Default, Debug, Clone)]
pub struct SideDirectionTransition {
    pub direction: Option<TransitionSideDirectionType>, // l
}

impl SideDirectionTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let direction = match xml_node.attribute("dir") {
            Some(value) => Some(value.parse()?),
            None => None,
        };

        Ok(Self { direction })
    }
}

#[derive(Default, Debug, Clone)]
pub struct SplitTransition {
    pub orientation: Option<Direction>,                  // horz
    pub direction: Option<TransitionInOutDirectionType>, // out
}

impl SplitTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "orient" => instance.orientation = Some(value.parse()?),
                "dir" => instance.direction = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Default, Debug, Clone)]
pub struct CornerDirectionTransition {
    pub direction: Option<TransitionCornerDirectionType>, // lu
}

impl CornerDirectionTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let direction = match xml_node.attribute("dir") {
            Some(value) => Some(value.parse()?),
            None => None,
        };

        Ok(Self { direction })
    }
}

#[derive(Default, Debug, Clone)]
pub struct WheelTransition {
    pub spokes: Option<u32>, // 4
}

impl WheelTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let spokes = match xml_node.attribute("spokes") {
            Some(value) => Some(value.parse()?),
            None => None,
        };

        Ok(Self { spokes })
    }
}

#[derive(Default, Debug, Clone)]
pub struct InOutTransition {
    pub direction: Option<TransitionInOutDirectionType>, // out
}

impl InOutTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let direction = match xml_node.attribute("dir") {
            Some(value) => Some(value.parse()?),
            None => None,
        };

        Ok(Self { direction })
    }
}

#[derive(Debug, Clone)]
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

impl SlideTransitionGroup {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "blinds" | "checker" | "circle" | "dissolve" | "comb" | "cover" | "cut" | "diamond" | "fade" | "newsflash"
            | "plus" | "pull" | "push" | "random" | "randomBar" | "split" | "strips" | "wedge" | "wheel" | "wipe"
            | "zoom" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "blinds" => Ok(SlideTransitionGroup::Blinds(OrientationTransition::from_xml_element(xml_node)?)),
            "checker" => Ok(SlideTransitionGroup::Checker(OrientationTransition::from_xml_element(xml_node)?)),
            "circle" => Ok(SlideTransitionGroup::Circle),
            "dissolve" => Ok(SlideTransitionGroup::Dissolve),
            "comb" => Ok(SlideTransitionGroup::Comb(OrientationTransition::from_xml_element(xml_node)?)),
            "cover" => Ok(SlideTransitionGroup::Cover(EightDirectionTransition::from_xml_element(xml_node)?)),
            "cut" => Ok(SlideTransitionGroup::Cut(OptionalBlackTransition::from_xml_element(xml_node)?)),
            "diamond" => Ok(SlideTransitionGroup::Diamond),
            "fade" => Ok(SlideTransitionGroup::Fade(OptionalBlackTransition::from_xml_element(xml_node)?)),
            "newsflash" => Ok(SlideTransitionGroup::Newsflash),
            "plus" => Ok(SlideTransitionGroup::Plus),
            "pull" => Ok(SlideTransitionGroup::Pull(EightDirectionTransition::from_xml_element(xml_node)?)),
            "push" => Ok(SlideTransitionGroup::Push(SideDirectionTransition::from_xml_element(xml_node)?)),
            "random" => Ok(SlideTransitionGroup::Random),
            "randomBar" => Ok(SlideTransitionGroup::RandomBar(OrientationTransition::from_xml_element(xml_node)?)),
            "split" => Ok(SlideTransitionGroup::Split(SplitTransition::from_xml_element(xml_node)?)),
            "strips" => Ok(SlideTransitionGroup::Strips(CornerDirectionTransition::from_xml_element(xml_node)?)),
            "wedge" => Ok(SlideTransitionGroup::Wedge),
            "wheel" => Ok(SlideTransitionGroup::Wheel(WheelTransition::from_xml_element(xml_node)?)),
            "wipe" => Ok(SlideTransitionGroup::Wipe(SideDirectionTransition::from_xml_element(xml_node)?)),
            "zoom" => Ok(SlideTransitionGroup::Zoom(InOutTransition::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "EG_SlideTransition"))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TransitionStartSoundAction {
    pub is_looping: Option<bool>, // false
    pub sound_file: crate::drawingml::EmbeddedWAVAudioFile,
}

impl TransitionStartSoundAction {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let is_looping = match xml_node.attribute("loop") {
            Some(value) => Some(parse_xml_bool(value)?),
            None => None,
        };

        let sound_file_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "snd"))?;
        let sound_file = crate::drawingml::EmbeddedWAVAudioFile::from_xml_element(sound_file_node)?;

        Ok(Self { is_looping, sound_file })
    }
}

#[derive(Debug, Clone)]
pub enum TransitionSoundAction {
    StartSound(TransitionStartSoundAction),
    EndSound,
}

impl TransitionSoundAction {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "stSnd" | "endSnd" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "stSnd" => Ok(TransitionSoundAction::StartSound(TransitionStartSoundAction::from_xml_element(xml_node)?)),
            "endSnd" => Ok(TransitionSoundAction::EndSound),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "CT_TransitionSoundAction"))),
        }
    }
}

#[derive(Default, Debug, Clone)]
pub struct SlideTransition {
    pub speed: Option<TransitionSpeed>, // fast
    pub advance_on_click: Option<bool>, // true
    pub advance_on_time: Option<u32>,
    pub transition_type: Option<SlideTransitionGroup>,
    pub sound_action: Option<TransitionSoundAction>,
}

impl SlideTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "spd" => instance.speed = Some(value.parse()?),
                "advClick" => instance.advance_on_click = Some(value.parse()?),
                "advTm" => instance.advance_on_time = Some(value.parse()?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            let child_local_name = child_node.local_name();
            if SlideTransitionGroup::is_choice_member(child_local_name) {
                instance.transition_type = Some(SlideTransitionGroup::from_xml_element(child_node)?);
            } else {
                match child_local_name {
                    "sndAc" => instance.sound_action = Some(TransitionSoundAction::from_xml_element(child_node)?),
                    _ => (),
                }
            }
        }

        Ok(instance)
    }
}


#[derive(Default, Debug, Clone)]
pub struct SlideTiming {
    pub time_node_list: Vec<TimeNodeGroup>,
    pub build_list: Vec<Build>,
}

impl SlideTiming {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tnLst" => {
                    for time_node in &child_node.child_nodes {
                        instance.time_node_list.push(TimeNodeGroup::from_xml_element(time_node)?);
                    }
                },
                "bldLst" => {
                    for build_node in &child_node.child_nodes {
                        instance.build_list.push(Build::from_xml_element(build_node)?);
                    }
                },
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone)]
pub enum Build {
    Paragraph(Box<TLBuildParagraph>),
    Diagram(Box<TLBuildDiagram>),
    OleChart(Box<TLOleBuildChart>),
    Graphic(Box<TLGraphicalObjectBuild>),
}

impl Build {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "bldP" | "bldDgm" | "bldOleChart" | "bldGraphic" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bldP" => Ok(Build::Paragraph(Box::new(TLBuildParagraph::from_xml_element(xml_node)?))),
            "bldDgm" => Ok(Build::Diagram(Box::new(TLBuildDiagram::from_xml_element(xml_node)?))),
            "bldOleChart" => Ok(Build::OleChart(Box::new(TLOleBuildChart::from_xml_element(xml_node)?))),
            "bldGraphic" => Ok(Build::Graphic(Box::new(TLGraphicalObjectBuild::from_xml_element(xml_node)?))),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "CT_BuildList"))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLBuildParagraph {
    pub build_common: TLBuildCommonAttributes,
    pub build_type: Option<TLParaBuildType>, // whole
    pub build_level: Option<u32>,            // 1
    pub animate_bg: Option<bool>,            // false
    pub auto_update_anim_bg: Option<bool>,   // true
    pub reverse: Option<bool>,               // false
    pub auto_advance_time: Option<TLTime>,   // indefinite
    pub template_list: Vec<TLTemplate>, // size: 0-9
}

impl TLBuildParagraph {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_id = None;
        let mut group_id = None;
        let mut ui_expand = None;
        let mut build_type = None;
        let mut build_level = None;
        let mut animate_bg = None;
        let mut auto_update_anim_bg = None;
        let mut reverse = None;
        let mut auto_advance_time = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "spid" => shape_id = Some(value.parse()?),
                "grpId" => group_id = Some(value.parse()?),
                "uiExpand" => ui_expand = Some(parse_xml_bool(value)?),
                "build" => build_type = Some(value.parse()?),
                "bldLvl" => build_level = Some(value.parse()?),
                "animBg" => animate_bg = Some(parse_xml_bool(value)?),
                "autoUpdateAnimBg" => auto_update_anim_bg = Some(parse_xml_bool(value)?),
                "rev" => reverse = Some(parse_xml_bool(value)?),
                "advAuto" => auto_advance_time = Some(value.parse()?),
                _ => (),
            }
        }

        let mut template_list = Vec::new();
        if let Some(child_node) = xml_node.child_nodes.get(0) {
            for template_node in &child_node.child_nodes {
                template_list.push(TLTemplate::from_xml_element(template_node)?);
            }
        }

        let shape_id = shape_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
        let group_id = group_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "grpId"))?;

        Ok(Self {
            build_common: TLBuildCommonAttributes {
                shape_id,
                group_id,
                ui_expand
            },
            build_type,
            build_level,
            animate_bg,
            auto_update_anim_bg,
            reverse,
            auto_advance_time,
            template_list,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLPoint {
    pub x: crate::drawingml::Percentage,
    pub y: crate::drawingml::Percentage,
}

impl TLPoint {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut x = None;
        let mut y = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "x" => x = Some(value.parse()?),
                "y" => y = Some(value.parse()?),
                _ => (),
            }
        }

        let x = x.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "x"))?;
        let y = y.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "y"))?;

        Ok(Self { x, y })
    }
}

#[derive(Debug, Clone)]
pub enum TLTime {
    TimePoint(u32),
    Indefinite,
}

impl FromStr for TLTime {
    type Err = crate::error::ParseEnumError;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s {
            "indefinite" => Ok(TLTime::Indefinite),
            _ => Ok(TLTime::TimePoint(s.parse().map_err(|_| Self::Err::new("TLTime"))?)),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLTemplate {
    pub level: Option<u32>, // 0
    pub time_node_list: Vec<TimeNodeGroup>,
}

impl TLTemplate {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let level = match xml_node.attribute("lvl") {
            Some(value) => Some(value.parse()?),
            None => None,
        };

        let time_node_list = match xml_node.child_nodes.get(0) {
            Some(child_node) => {
                let mut vec = Vec::new();
                for time_node in &child_node.child_nodes {
                    vec.push(TimeNodeGroup::from_xml_element(time_node)?);
                }
                vec
            }
            None => return Err(Box::new(MissingChildNodeError::new(xml_node.name.clone(), "tnLst"))),
        };

        Ok(Self { level, time_node_list })
    }
}

#[derive(Debug, Clone)]
pub struct TLBuildCommonAttributes {
    pub shape_id: crate::drawingml::DrawingElementId,
    pub group_id: u32,
    pub ui_expand: Option<bool>, // false
}

#[derive(Debug, Clone)]
pub struct TLBuildDiagram {
    pub build_common: TLBuildCommonAttributes,
    pub build_type: Option<TLDiagramBuildType>, // whole
}

impl TLBuildDiagram {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_id = None;
        let mut group_id = None;
        let mut ui_expand = None;
        let mut build_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "spid" => shape_id = Some(value.parse()?),
                "grpId" => group_id = Some(value.parse()?),
                "uiExpand" => ui_expand = Some(parse_xml_bool(value)?),
                "bld" => build_type = Some(value.parse()?),
                _ => (),
            }
        }

        let shape_id = shape_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
        let group_id = group_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "grpId"))?;

        Ok(Self {
            build_common: TLBuildCommonAttributes {
                shape_id,
                group_id,
                ui_expand,
            },
            build_type,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLOleBuildChart {
    pub build_common: TLBuildCommonAttributes,
    pub build_type: Option<TLOleChartBuildType>, // allAtOnce
    pub animate_bg: Option<bool>,                // true
}

impl TLOleBuildChart {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_id = None;
        let mut group_id = None;
        let mut ui_expand = None;
        let mut build_type = None;
        let mut animate_bg = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "spid" => shape_id = Some(value.parse()?),
                "grpId" => group_id = Some(value.parse()?),
                "uiExpand" => ui_expand = Some(parse_xml_bool(value)?),
                "bld" => build_type = Some(value.parse()?),
                "animBg" => animate_bg = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let shape_id = shape_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
        let group_id = group_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "grpId"))?;

        Ok(Self {
            build_common: TLBuildCommonAttributes {
                shape_id,
                group_id,
                ui_expand,
            },
            build_type,
            animate_bg,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLGraphicalObjectBuild {
    pub build_common: TLBuildCommonAttributes,
    pub build_choice: TLGraphicalObjectBuildChoice,
}

impl TLGraphicalObjectBuild {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_id = None;
        let mut group_id = None;
        let mut ui_expand = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "spid" => shape_id = Some(value.parse()?),
                "grpId" => group_id = Some(value.parse()?),
                "uiExpand" => ui_expand = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let build_choice = match xml_node.child_nodes.get(0) {
            Some(child_node) => TLGraphicalObjectBuildChoice::from_xml_element(child_node)?,
            None => return Err(Box::new(
                MissingChildNodeError::new(xml_node.name.clone(), "TLGraphicalObjectBuildChoice")
            ))
        };

        let shape_id = shape_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
        let group_id = group_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "grpId"))?;

        Ok(Self {
            build_common: TLBuildCommonAttributes {
                shape_id,
                group_id,
                ui_expand,
            },
            build_choice,
        })
    }
}

#[derive(Debug, Clone)]
pub enum TLGraphicalObjectBuildChoice {
    BuildAsOne,
    BuildSubElements(crate::drawingml::AnimationGraphicalObjectBuildProperties),
}

impl TLGraphicalObjectBuildChoice {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "bldAsOne" | "bldSub" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bldAsOne" => Ok(TLGraphicalObjectBuildChoice::BuildAsOne),
            "bldSub" => Ok(TLGraphicalObjectBuildChoice::BuildSubElements(
                crate::drawingml::AnimationGraphicalObjectBuildProperties::from_xml_element(xml_node)?
            )),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "TLGraphicalObjectBuildChoice"))),
        }
    }
}

#[derive(Debug, Clone)]
pub enum TimeNodeGroup {
    Parallel(Box<TLCommonTimeNodeData>),
    Sequence(Box<TLTimeNodeSequence>),
    Exclusive(Box<TLCommonTimeNodeData>),
    Animate(Box<TLAnimateBehavior>),
    AnimateColor(Box<TLAnimateColorBehavior>),
    AnimateEffect(Box<TLAnimateEffectBehavior>),
    AnimateMotion(Box<TLAnimateMotionBehavior>),
    AnimateRotation(Box<TLAnimateRotationBehavior>),
    AnimateScale(Box<TLAnimateScaleBehavior>),
    Command(Box<TLCommandBehavior>),
    Set(Box<TLSetBehavior>),
    Audio(Box<TLMediaNodeAudio>),
    Video(Box<TLMediaNodeVideo>),
}

impl TimeNodeGroup {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "par" | "seq" | "excl" | "anim" | "animClr" | "animEffect" | "animMotion" | "animRot" | "animScale"
            | "cmd" | "set" | "audio" | "video" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "par" => Ok(TimeNodeGroup::Parallel(Box::new(TLCommonTimeNodeData::from_xml_element(xml_node)?))),
            "seq" => Ok(TimeNodeGroup::Sequence(Box::new(TLTimeNodeSequence::from_xml_element(xml_node)?))),
            "excl" => Ok(TimeNodeGroup::Exclusive(Box::new(TLCommonTimeNodeData::from_xml_element(xml_node)?))),
            "anim" => Ok(TimeNodeGroup::Animate(Box::new(TLAnimateBehavior::from_xml_element(xml_node)?))),
            "animClr" => Ok(TimeNodeGroup::AnimateColor(Box::new(TLAnimateColorBehavior::from_xml_element(xml_node)?))),
            "animEffect" => Ok(TimeNodeGroup::AnimateEffect(Box::new(
                TLAnimateEffectBehavior::from_xml_element(xml_node)?
            ))),
            "animMotion" => Ok(TimeNodeGroup::AnimateMotion(Box::new(
                TLAnimateMotionBehavior::from_xml_element(xml_node)?
            ))),
            "animRot" => Ok(TimeNodeGroup::AnimateRotation(Box::new(
                TLAnimateRotationBehavior::from_xml_element(xml_node)?
            ))),
            "animScale" => Ok(TimeNodeGroup::AnimateScale(Box::new(
                TLAnimateScaleBehavior::from_xml_element(xml_node)?
            ))),
            "cmd" => Ok(TimeNodeGroup::Command(Box::new(TLCommandBehavior::from_xml_element(xml_node)?))),
            "set" => Ok(TimeNodeGroup::Set(Box::new(TLSetBehavior::from_xml_element(xml_node)?))),
            "audio" => Ok(TimeNodeGroup::Audio(Box::new(TLMediaNodeAudio::from_xml_element(xml_node)?))),
            "video" => Ok(TimeNodeGroup::Video(Box::new(TLMediaNodeVideo::from_xml_element(xml_node)?))),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "TimeNodeGroup"))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLTimeNodeSequence {
    pub concurrent: Option<bool>,
    pub prev_action_type: Option<TLPreviousActionType>,
    pub next_action_type: Option<TLNextActionType>,
    pub common_time_node_data: Box<TLCommonTimeNodeData>,
    pub prev_condition_list: Vec<TLTimeCondition>,
    pub next_condition_list: Vec<TLTimeCondition>,
}

impl TLTimeNodeSequence {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut concurrent = None;
        let mut prev_action_type = None;
        let mut next_action_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "concurrent" => concurrent = Some(parse_xml_bool(value)?),
                "prevAc" => prev_action_type = Some(value.parse()?),
                "nextAc" => next_action_type = Some(value.parse()?),
                _ => (),
            }
        }

        let mut common_time_node_data = None;
        let mut prev_condition_list = Vec::new();
        let mut next_condition_list = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cTn" => common_time_node_data = Some(Box::new(TLCommonTimeNodeData::from_xml_element(child_node)?)),
                "prevCondLst" => {
                    for condition_node in &child_node.child_nodes {
                        prev_condition_list.push(TLTimeCondition::from_xml_element(condition_node)?);
                    }
                }
                "nextCondLst" => {
                    for condition_node in &child_node.child_nodes {
                        next_condition_list.push(TLTimeCondition::from_xml_element(condition_node)?);
                    }
                }
                _ => (),
            }
        }

        let common_time_node_data =
            common_time_node_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cTn"))?;

        Ok(Self {
            concurrent,
            prev_action_type,
            next_action_type,
            common_time_node_data,
            prev_condition_list,
            next_condition_list,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLAnimateBehavior {
    pub by: Option<String>,
    pub from: Option<String>,
    pub to: Option<String>,
    pub calc_mode: Option<TLAnimateBehaviorCalcMode>,
    pub value_type: Option<TLAnimateBehaviorValueType>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    pub time_animate_value_list: Vec<TLTimeAnimateValue>,
}

impl TLAnimateBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut by = None;
        let mut from = None;
        let mut to = None;
        let mut calc_mode = None;
        let mut value_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "by" => by = Some(value.clone()),
                "from" => from = Some(value.clone()),
                "to" => to = Some(value.clone()),
                "calcmode" => calc_mode = Some(value.parse()?),
                "valueType" => value_type = Some(value.parse()?),
                _ => (),
            }
        }

        let mut common_behavior_data = None;
        let mut time_animate_value_list = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "tavLst" => {
                    for tav_node in &child_node.child_nodes {
                        time_animate_value_list.push(TLTimeAnimateValue::from_xml_element(tav_node)?);
                    }
                }
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            by,
            from,
            to,
            calc_mode,
            value_type,
            common_behavior_data,
            time_animate_value_list,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLAnimateColorBehavior {
    pub color_space: Option<TLAnimateColorSpace>,
    pub direction: Option<TLAnimateColorDirection>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    pub by: Option<TLByAnimateColorTransform>,
    pub from: Option<crate::drawingml::Color>,
    pub to: Option<crate::drawingml::Color>,
}

impl TLAnimateColorBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        use crate::drawingml::Color;

        let mut color_space = None;
        let mut direction = None;
        
        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "clrSpc" => color_space = Some(value.parse()?),
                "dir" => direction = Some(value.parse()?),
                _ => (),
            }
        }

        let mut common_behavior_data = None;
        let mut by = None;
        let mut from = None;
        let mut to = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "by" => {
                    let by_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "TLByAnimateColorTransform"))?;
                    by = Some(TLByAnimateColorTransform::from_xml_element(by_node)?);
                }
                "from" => {
                    let color_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?;
                    from = Some(Color::from_xml_element(color_node)?);
                }
                "to" => {
                    let color_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?;
                    to = Some(Color::from_xml_element(color_node)?);
                }
                _ => (),
            }
        }

        let common_behavior_data = 
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            color_space,
            direction,
            common_behavior_data,
            by,
            from,
            to,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLAnimateEffectBehavior {
    pub transition: Option<TLAnimateEffectTransition>,
    pub filter: Option<String>,
    pub property_list: Option<String>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    pub progress: Option<TLAnimVariant>,
}

impl TLAnimateEffectBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut transition = None;
        let mut filter = None;
        let mut property_list = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "transition" => transition = Some(value.parse()?),
                "filter" => filter = Some(value.clone()),
                "prLst" => property_list = Some(value.clone()),
                _ => (),
            }
        }

        let mut common_behavior_data = None;
        let mut progress = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "progress" => {
                    let progress_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "CT_TLAnimVariant"))?;
                    progress = Some(TLAnimVariant::from_xml_element(progress_node)?);
                },
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            transition,
            filter,
            property_list,
            common_behavior_data,
            progress,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLAnimateMotionBehavior {
    pub origin: Option<TLAnimateMotionBehaviorOrigin>,
    pub path: Option<String>,
    pub path_edit_mode: Option<TLAnimateMotionPathEditMode>,
    pub rotate_angle: Option<crate::drawingml::Angle>,
    pub points_types: Option<String>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    pub by: Option<TLPoint>,
    pub from: Option<TLPoint>,
    pub to: Option<TLPoint>,
    pub rotation_center: Option<TLPoint>,
}

impl TLAnimateMotionBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut origin = None;
        let mut path = None;
        let mut path_edit_mode = None;
        let mut rotate_angle = None;
        let mut points_types = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "origin" => origin = Some(value.parse()?),
                "path" => path = Some(value.clone()),
                "pathEditMode" => path_edit_mode = Some(value.parse()?),
                "rAng" => rotate_angle = Some(value.parse()?),
                "ptsTypes" => points_types = Some(value.clone()),
                _ => (),
            }
        }

        let mut common_behavior_data = None;
        let mut by = None;
        let mut from = None;
        let mut to = None;
        let mut rotation_center = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "by" => by = Some(TLPoint::from_xml_element(child_node)?),
                "from" => from = Some(TLPoint::from_xml_element(child_node)?),
                "to" => to = Some(TLPoint::from_xml_element(child_node)?),
                "rCtr" => rotation_center = Some(TLPoint::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            origin,
            path,
            path_edit_mode,
            rotate_angle,
            points_types,
            common_behavior_data,
            by,
            from,
            to,
            rotation_center,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLAnimateRotationBehavior {
    pub by: Option<crate::drawingml::Angle>,
    pub from: Option<crate::drawingml::Angle>,
    pub to: Option<crate::drawingml::Angle>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
}

impl TLAnimateRotationBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut by = None;
        let mut from = None;
        let mut to = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "by" => by = Some(value.parse()?),
                "from" => from = Some(value.parse()?),
                "to" => to = Some(value.parse()?),
                _ => (),
            }
        }

        let common_behavior_data_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;
        let common_behavior_data = Box::new(TLCommonBehaviorData::from_xml_element(common_behavior_data_node)?);

        Ok(Self {
            by,
            from,
            to,
            common_behavior_data,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLAnimateScaleBehavior {
    pub zoom_contents: Option<bool>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    pub by: Option<TLPoint>,
    pub from: Option<TLPoint>,
    pub to: Option<TLPoint>,
}

impl TLAnimateScaleBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let zoom_contents = match xml_node.attribute("zoomContents") {
            Some(value) => Some(parse_xml_bool(value)?),
            None => None,
        };

        let mut common_behavior_data = None;
        let mut by = None;
        let mut from = None;
        let mut to = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "by" => by = Some(TLPoint::from_xml_element(child_node)?),
                "from" => from = Some(TLPoint::from_xml_element(child_node)?),
                "to" => to = Some(TLPoint::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            zoom_contents,
            common_behavior_data,
            by,
            from,
            to,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLCommandBehavior {
    pub command_type: Option<TLCommandType>,
    pub command: Option<String>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
}

impl TLCommandBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut command_type = None;
        let mut command = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => command_type = Some(value.parse()?),
                "cmd" => command = Some(value.clone()),
                _ => (),
            }
        }

        let common_behavior_data_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;
        let common_behavior_data = Box::new(TLCommonBehaviorData::from_xml_element(common_behavior_data_node)?);

        Ok(Self {
            command_type,
            command,
            common_behavior_data,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLSetBehavior {
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    pub to: Option<TLAnimVariant>,
}

impl TLSetBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut common_behavior_data = None;
        let mut to = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "to" => {
                    let to_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "CT_TLAnimVariant"))?;
                    to = Some(TLAnimVariant::from_xml_element(to_node)?);
                },
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            common_behavior_data,
            to,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLMediaNodeAudio {
    pub is_narration: Option<bool>, // false
    pub common_media_node_data: Box<TLCommonMediaNodeData>,
}

impl TLMediaNodeAudio {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let is_narration = match xml_node.attribute("isNarration") {
            Some(value) => Some(parse_xml_bool(value)?),
            None => None,
        };

        let common_media_node_data_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cMediaNode"))?;
        let common_media_node_data = Box::new(TLCommonMediaNodeData::from_xml_element(common_media_node_data_node)?);

        Ok(Self {
            is_narration,
            common_media_node_data,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLMediaNodeVideo {
    pub fullscreen: Option<bool>, // false
    pub common_media_node_data: Box<TLCommonMediaNodeData>,
}

impl TLMediaNodeVideo {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let fullscreen = match xml_node.attribute("fullScrn") {
            Some(value) => Some(parse_xml_bool(value)?),
            None => None,
        };

        let common_media_node_data_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cMediaNode"))?;
        let common_media_node_data = Box::new(TLCommonMediaNodeData::from_xml_element(common_media_node_data_node)?);

        Ok(Self {
            fullscreen,
            common_media_node_data,
        })
    }
}

#[derive(Default, Debug, Clone)]
pub struct TLTimeAnimateValue {
    pub time: Option<TLTimeAnimateValueTime>, // indefinite
    pub formula: Option<String>,              // ""
    pub value: Option<TLAnimVariant>,
}

impl TLTimeAnimateValue {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "tm" => instance.time = Some(value.parse()?),
                "fmla" => instance.formula = Some(value.clone()),
                _ => (),
            }
        }

        if let Some(child_node) = xml_node.child_nodes.get(0) {
            let val_node = child_node
                .child_nodes
                .get(0)
                .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "CT_TLAnimVariant"))?;
            instance.value = Some(TLAnimVariant::from_xml_element(val_node)?);
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone)]
pub enum TLTimeAnimateValueTime {
    Percentage(crate::drawingml::PositiveFixedPercentage),
    Indefinite,
}

impl FromStr for TLTimeAnimateValueTime {
    type Err = crate::error::ParseEnumError;
    
    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s {
            "indefinite" => Ok(TLTimeAnimateValueTime::Indefinite),
            _ => Ok(TLTimeAnimateValueTime::Percentage(
                s.parse().map_err(|_| Self::Err::new("TLTimeAnimateValueTime"))?)
            ),
        }
    }
}

#[derive(Debug, Clone)]
pub enum TLAnimVariant {
    Bool(bool),
    Int(i32),
    Float(f32),
    String(String),
    Color(crate::drawingml::Color),
}

impl TLAnimVariant {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "boolVal" | "intVal" | "fltVal" | "strVal" | "clrVal" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "boolVal" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLAnimVariant::Bool(parse_xml_bool(val_attr)?))
            }
            "intVal" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLAnimVariant::Int(val_attr.parse()?))
            }
            "fltVal" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLAnimVariant::Float(val_attr.parse()?))
            }
            "strVal" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLAnimVariant::String(val_attr.clone()))
            }
            "clrVal" => {
                let child_node = xml_node
                    .child_nodes
                    .get(0)
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_Color"))?;
                Ok(TLAnimVariant::Color(crate::drawingml::Color::from_xml_element(child_node)?))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TLAnimVariant").into()),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLCommonBehaviorData {
    pub additive: Option<TLBehaviorAdditiveType>,
    pub accumulate: Option<TLBehaviorAccumulateType>,
    pub transform_type: Option<TLBehaviorTransformType>,
    pub from: Option<String>,
    pub to: Option<String>,
    pub by: Option<String>,
    pub runtime_context: Option<String>,
    pub override_type: Option<TLBehaviorOverrideType>,
    pub common_time_node_data: Box<TLCommonTimeNodeData>,
    pub target_element: TLTimeTargetElement,
    pub attr_name_list: Option<Vec<String>>,
}

impl TLCommonBehaviorData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut additive = None;
        let mut accumulate = None;
        let mut transform_type = None;
        let mut from = None;
        let mut to = None;
        let mut by = None;
        let mut runtime_context = None;
        let mut override_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "additive" => additive = Some(value.parse()?),
                "accumulate" => accumulate = Some(value.parse()?),
                "xfrmType" => transform_type = Some(value.parse()?),
                "from" => from = Some(value.clone()),
                "to" => to = Some(value.clone()),
                "by" => by = Some(value.clone()),
                "rctx" => runtime_context = Some(value.clone()),
                "override" => override_type = Some(value.parse()?),
                _ => (),
            }
        }

        let mut common_time_node_data = None;
        let mut target_element = None;
        let mut attr_name_list = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cTn" => common_time_node_data = Some(Box::new(TLCommonTimeNodeData::from_xml_element(child_node)?)),
                "tgtEl" => target_element = Some(TLTimeTargetElement::from_xml_element(child_node)?),
                "attrNameLst" => {
                    let mut vec = Vec::new();
                    for attr_name_node in &child_node.child_nodes {
                        vec.push(match attr_name_node.text {
                            Some(ref text) => text.clone(),
                            None => String::new(), // TODO: maybe it's an error to have an empty node?
                        });
                    }

                    if vec.is_empty() {
                        return Err(Box::new(MissingChildNodeError::new(child_node.name.clone(), "attrName")));
                    }

                    attr_name_list = Some(vec);
                },
                _ => (),
            }
        }

        let common_time_node_data =
            common_time_node_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cTn"))?;
        let target_element =
            target_element.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tgtEl"))?;

        Ok(Self {
            additive,
            accumulate,
            transform_type,
            from,
            to,
            by,
            runtime_context,
            override_type,
            common_time_node_data,
            target_element,
            attr_name_list,
        })
    }
}

#[derive(Debug, Clone)]
pub struct TLCommonMediaNodeData {
    pub volume: Option<crate::drawingml::PositiveFixedPercentage>, // 50000
    pub mute: Option<bool>,                                        // false
    pub number_of_slides: Option<u32>,                             // 1
    pub show_when_stopped: Option<bool>,                           // true
    pub common_time_node_data: Box<TLCommonTimeNodeData>,
    pub target_element: TLTimeTargetElement,
}

impl TLCommonMediaNodeData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut volume = None;
        let mut mute = None;
        let mut number_of_slides = None;
        let mut show_when_stopped = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "vol" => volume = Some(value.parse()?),
                "mute" => mute = Some(parse_xml_bool(value)?),
                "numSld" => number_of_slides = Some(value.parse()?),
                "showWhenStopped" => show_when_stopped = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut common_time_node_data = None;
        let mut target_element = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cTn" => common_time_node_data = Some(Box::new(TLCommonTimeNodeData::from_xml_element(child_node)?)),
                "tgtEl" => target_element = Some(TLTimeTargetElement::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_time_node_data =
            common_time_node_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cTn"))?;
        let target_element =
            target_element.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tgtEl"))?;

        Ok(Self {
            volume,
            mute,
            number_of_slides,
            show_when_stopped,
            common_time_node_data,
            target_element,
        })
    }
}

#[derive(Debug, Clone)]
pub enum TLTimeConditionTriggerGroup {
    TargetElement(TLTimeTargetElement),
    TimeNode(TLTimeNodeId),
    RuntimeNode(TLTriggerRuntimeNode),
}

impl TLTimeConditionTriggerGroup {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "tgtEl" | "tn" | "rtn" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "tgtEl" => {
                let target_element_node = xml_node
                    .child_nodes
                    .get(0)
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "CT_TLTimeTargetElement"))?;
                Ok(TLTimeConditionTriggerGroup::TargetElement(
                    TLTimeTargetElement::from_xml_element(target_element_node)?,
                ))
            },
            "tn" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLTimeConditionTriggerGroup::TimeNode(val_attr.parse()?))
            },
            "rtn" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLTimeConditionTriggerGroup::RuntimeNode(val_attr.parse()?))
            },
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "EG_TLTimeConditionTriggerGroup")))
        }
    }
}

#[derive(Debug, Clone)]
pub enum TLTimeTargetElement {
    SlideTarget,
    SoundTarget(crate::drawingml::EmbeddedWAVAudioFile),
    ShapeTarget(TLShapeTargetElement),
    InkTarget(TLSubShapeId),
}

impl TLTimeTargetElement {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "sldTgt" | "sndTgt" | "spTgt" | "inkTgt" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "sldTgt" => Ok(TLTimeTargetElement::SlideTarget),
            "sndTgt" => Ok(TLTimeTargetElement::SoundTarget(
                crate::drawingml::EmbeddedWAVAudioFile::from_xml_element(xml_node)?
            )),
            "spTgt" => {
                let child_node = xml_node
                    .child_nodes
                    .get(0)
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "CT_TLShapeTargetElement"))?;
                Ok(TLTimeTargetElement::ShapeTarget(TLShapeTargetElement::from_xml_element(child_node)?))
            },
            "inkTgt" => {
                let spid_attr = xml_node
                    .attribute("spid")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
                 Ok(TLTimeTargetElement::InkTarget(spid_attr.parse()?))
            },
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "CT_TLTimeTargetElement"))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLShapeTargetElement {
    pub shape_id: crate::drawingml::DrawingElementId,
    pub target: Option<TLShapeTargetElementGroup>,
}

impl TLShapeTargetElement {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let shape_id_attr = xml_node
            .attribute("spid")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
        let shape_id = shape_id_attr.parse()?;

        let target = match xml_node.child_nodes.get(0) {
            Some(child_node) => Some(TLShapeTargetElementGroup::from_xml_element(child_node)?),
            None => None,
        };

        Ok(Self { shape_id, target })
    }
}

#[derive(Debug, Clone)]
pub enum TLShapeTargetElementGroup {
    Background,
    SubShape(TLSubShapeId),
    OleChartElement(TLOleChartTargetElement),
    TextElement(Option<TLTextTargetElement>),
    GraphicElement(crate::drawingml::AnimationElementChoice),
}

impl TLShapeTargetElementGroup {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "bg" | "subSp" | "oleChartEl" | "txEl" | "graphicEl" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bg" => Ok(TLShapeTargetElementGroup::Background),
            "subSp" => {
                let spid_attr = xml_node
                    .attribute("spid")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
                Ok(TLShapeTargetElementGroup::SubShape(spid_attr.parse()?))
            },
            "oleChartEl" => Ok(TLShapeTargetElementGroup::OleChartElement(
                TLOleChartTargetElement::from_xml_element(xml_node)?
            )),
            "txEl" => Ok(TLShapeTargetElementGroup::TextElement(match xml_node.child_nodes.get(0) {
                    Some(child_node) => Some(TLTextTargetElement::from_xml_element(child_node)?),
                    None => None,
                }
            )),
            "graphicEl" => {
                let child_node = xml_node
                    .child_nodes
                    .get(0)
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "CT_AnimationElementChoice"))?;

                Ok(TLShapeTargetElementGroup::GraphicElement(
                    crate::drawingml::AnimationElementChoice::from_xml_element(child_node)?
                ))
            },
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "TLShapeTargetElementGroup")))
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLOleChartTargetElement {
    pub element_type: TLChartSubelementType,
    pub level: Option<u32>, // 0
}

impl TLOleChartTargetElement {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut element_type = None;
        let mut level = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => element_type = Some(value.parse()?),
                "lvl" => level = Some(value.parse()?),
                _ => (),
            }
        }

        let element_type = element_type.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "type"))?;

        Ok(Self { element_type, level })
    }
}

#[derive(Debug, Clone)]
pub enum TLTextTargetElement {
    CharRange(IndexRange),
    ParagraphRange(IndexRange),
}

impl TLTextTargetElement {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "charRg" | "pRg" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "charRg" => Ok(TLTextTargetElement::CharRange(IndexRange::from_xml_element(xml_node)?)),
            "pRg" => Ok(TLTextTargetElement::ParagraphRange(IndexRange::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "TLTextTargetElement"))),
        }
    }
}

#[derive(Default, Debug, Clone)]
pub struct TLTimeCondition {
    pub trigger_event: Option<TLTriggerEvent>,
    pub delay: Option<TLTime>,
    pub trigger: Option<TLTimeConditionTriggerGroup>,
}

impl TLTimeCondition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "evt" => instance.trigger_event = Some(value.parse()?),
                "delay" => instance.delay = Some(value.parse()?),
                _ => (),
            }
        }

        if let Some(child_node) = xml_node.child_nodes.get(0) {
            instance.trigger = Some(TLTimeConditionTriggerGroup::from_xml_element(child_node)?);
        }

        Ok(instance)
    }
}

#[derive(Default, Debug, Clone)]
pub struct TLCommonTimeNodeData {
    pub id: Option<TLTimeNodeId>,
    pub preset_id: Option<i32>,
    pub preset_class: Option<TLTimeNodePresetClassType>,
    pub preset_subtype: Option<i32>,
    pub duration: Option<TLTime>,
    pub repeat_count: Option<TLTime>, // 1000
    pub repeat_duration: Option<TLTime>,
    pub speed: Option<crate::drawingml::Percentage>, // 100000
    pub acceleration: Option<crate::drawingml::PositiveFixedPercentage>, // 0
    pub deceleration: Option<crate::drawingml::PositiveFixedPercentage>, // 0
    pub auto_reverse: Option<bool>,                  // false
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
    pub start_condition_list: Option<Vec<TLTimeCondition>>,
    pub end_condition_list: Option<Vec<TLTimeCondition>>,
    pub end_sync: Option<TLTimeCondition>,
    pub iterate: Option<TLIterateData>,
    pub child_time_node_list: Option<Vec<TimeNodeGroup>>,
    pub sub_time_node_list: Option<Vec<TimeNodeGroup>>,
}

impl TLCommonTimeNodeData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => instance.id = Some(value.parse()?),
                "presetID" => instance.preset_id = Some(value.parse()?),
                "presetClass" => instance.preset_class = Some(value.parse()?),
                "presetSubtype" => instance.preset_subtype = Some(value.parse()?),
                "dur" => instance.duration = Some(value.parse()?),
                "repeatCount" => instance.repeat_count = Some(value.parse()?),
                "repeatDur" => instance.repeat_duration = Some(value.parse()?),
                "spd" => instance.speed = Some(value.parse()?),
                "accel" => instance.acceleration = Some(value.parse()?),
                "decel" => instance.deceleration = Some(value.parse()?),
                "autoRev" => instance.auto_reverse = Some(parse_xml_bool(value)?),
                "restart" => instance.restart_type = Some(value.parse()?),
                "fill" => instance.fill_type = Some(value.parse()?),
                "syncBehavior" => instance.sync_behavior = Some(value.parse()?),
                "tmFilter" => instance.time_filter = Some(value.clone()),
                "evtFilter" => instance.event_filter = Some(value.clone()),
                "display" => instance.display = Some(parse_xml_bool(value)?),
                "masterRel" => instance.master_relationship = Some(value.parse()?),
                "bldLvl" => instance.build_level = Some(value.parse()?),
                "grpId" => instance.group_id = Some(value.parse()?),
                "afterEffect" => instance.after_effect = Some(parse_xml_bool(value)?),
                "nodeType" => instance.node_type = Some(value.parse()?),
                "nodePh" => instance.node_placeholder = Some(parse_xml_bool(value)?),
                _ => ()
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "stCondLst" => {
                    let mut vec = Vec::new();
                    for cond_node in &child_node.child_nodes {
                        vec.push(TLTimeCondition::from_xml_element(cond_node)?);
                    }

                    if vec.is_empty() {
                        return Err(Box::new(MissingChildNodeError::new(child_node.name.clone(), "cond")));
                    }

                    instance.start_condition_list = Some(vec);
                },
                "endCondLst" => {
                    let mut vec = Vec::new();
                    for cond_node in &child_node.child_nodes {
                        vec.push(TLTimeCondition::from_xml_element(cond_node)?);
                    }

                    if vec.is_empty() {
                        return Err(Box::new(MissingChildNodeError::new(child_node.name.clone(), "cond")));
                    }

                    instance.end_condition_list = Some(vec);
                },
                "endSync" => instance.end_sync = Some(TLTimeCondition::from_xml_element(child_node)?),
                "iterate" => instance.iterate = Some(TLIterateData::from_xml_element(child_node)?),
                "childTnLst" => {
                    let mut vec = Vec::new();
                    for time_node in &child_node.child_nodes {
                        vec.push(TimeNodeGroup::from_xml_element(time_node)?);
                    }

                    if vec.is_empty() {
                        return Err(Box::new(MissingChildNodeError::new(child_node.name.clone(), "TimeNode")));
                    }

                    instance.child_time_node_list = Some(vec);
                },
                "subTnLst" => {
                    let mut vec = Vec::new();
                    for time_node in &child_node.child_nodes {
                        vec.push(TimeNodeGroup::from_xml_element(time_node)?);
                    }

                    if vec.is_empty() {
                        return Err(Box::new(MissingChildNodeError::new(child_node.name.clone(), "TimeNode")));
                    }

                    instance.sub_time_node_list = Some(vec);
                },
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone)]
pub enum TLIterateDataChoice {
    Absolute(TLTime),
    Percent(crate::drawingml::PositivePercentage),
}

impl TLIterateDataChoice {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "tmAbs" | "tmPct" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "tmAbs" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLIterateDataChoice::Absolute(val_attr.parse()?))
            },
            "tmPct" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLIterateDataChoice::Percent(val_attr.parse()?))
            },
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "TLIterateDataChoice"))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLIterateData {
    pub iterate_type: Option<IterateType>, // el
    pub backwards: Option<bool>,           //false
    pub interval: TLIterateDataChoice,
}

impl TLIterateData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut iterate_type = None;
        let mut backwards = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => iterate_type = Some(value.parse()?),
                "backwards" => backwards = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let interval_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "TLIterateDataChoice"))?;
        let interval = TLIterateDataChoice::from_xml_element(interval_node)?;
        
        Ok(Self {
            iterate_type,
            backwards,
            interval,
        })
    }
}

#[derive(Debug, Clone)]
pub enum TLByAnimateColorTransform {
    Rgb(TLByRgbColorTransform),
    Hsl(TLByHslColorTransform),
}

impl TLByAnimateColorTransform {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "rgb" | "hsl" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "rgb" => Ok(TLByAnimateColorTransform::Rgb(TLByRgbColorTransform::from_xml_element(xml_node)?)),
            "hsl" => Ok(TLByAnimateColorTransform::Hsl(TLByHslColorTransform::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "TLByAnimateColorTransform")))
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLByRgbColorTransform {
    pub r: crate::drawingml::FixedPercentage,
    pub g: crate::drawingml::FixedPercentage,
    pub b: crate::drawingml::FixedPercentage,
}

impl TLByRgbColorTransform {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut r = None;
        let mut g = None;
        let mut b = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r" => r = Some(value.parse()?),
                "g" => g = Some(value.parse()?),
                "b" => b = Some(value.parse()?),
                _ => (),
            }
        }

        let r = r.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r"))?;
        let g = g.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "g"))?;
        let b = b.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "b"))?;

        Ok(Self { r, g, b })
    }
}

#[derive(Debug, Clone)]
pub struct TLByHslColorTransform {
    pub h: crate::drawingml::Angle,
    pub s: crate::drawingml::FixedPercentage,
    pub l: crate::drawingml::FixedPercentage,
}

impl TLByHslColorTransform {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut h = None;
        let mut s = None;
        let mut l = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "h" => h = Some(value.parse()?),
                "s" => s = Some(value.parse()?),
                "l" => l = Some(value.parse()?),
                _ => (),
            }
        }

        let h = h.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "h"))?;
        let s = s.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "s"))?;
        let l = l.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "l"))?;

        Ok(Self { h, s, l })
    }
}

#[derive(Debug, Clone)]
pub struct SlideMaster {
    pub preserve: Option<bool>, // false
    pub common_slide_data: Box<CommonSlideData>,
    pub color_mapping: Box<crate::drawingml::ColorMapping>,
    pub slide_layout_id_list: Vec<SlideLayoutIdListEntry>,
    pub transition: Option<Box<SlideTransition>>,
    pub timing: Option<SlideTiming>,
    pub header_footer: Option<HeaderFooter>,
    pub text_styles: Option<SlideMasterTextStyles>,
}

impl SlideMaster {
    pub fn from_zip_file(zip_file: &mut ZipFile<'_>) -> Result<Self> {
        let mut xml_string = String::new();
        zip_file.read_to_string(&mut xml_string)?;
        let xml_node = XmlNode::from_str(xml_string.as_str())?;

        Self::from_xml_element(&xml_node)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let preserve = match xml_node.attribute("preserve") {
            Some(val) => Some(parse_xml_bool(val)?),
            None => None,
        };

        let mut common_slide_data = None;
        let mut color_mapping = None;
        let mut slide_layout_id_list = Vec::new();
        let mut transition = None;
        let mut timing = None;
        let mut header_footer = None;
        let mut text_styles = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cSld" => common_slide_data = Some(Box::new(CommonSlideData::from_xml_element(child_node)?)),
                "clrMap" => {
                    color_mapping = Some(Box::new(crate::drawingml::ColorMapping::from_xml_element(child_node)?))
                }
                "sldLayoutIdLst" => {
                    for slide_layout_id_node in &child_node.child_nodes {
                        slide_layout_id_list.push(SlideLayoutIdListEntry::from_xml_element(slide_layout_id_node)?);
                    }
                }
                "transition" => transition = Some(Box::new(SlideTransition::from_xml_element(child_node)?)),
                "timing" => timing = Some(SlideTiming::from_xml_element(child_node)?),
                "hf" => header_footer = Some(HeaderFooter::from_xml_element(child_node)?),
                "txStyles" => text_styles = Some(SlideMasterTextStyles::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_slide_data = common_slide_data
            .ok_or_else(|| XmlError::from(MissingChildNodeError::new(xml_node.name.clone(), "cSld")))?;
        let color_mapping =
            color_mapping.ok_or_else(|| XmlError::from(MissingChildNodeError::new(xml_node.name.clone(), "clrMap")))?;

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

#[derive(Debug, Clone)]
pub struct SlideLayout {
    pub matching_name: Option<String>,                    // ""
    pub slide_layout_type: Option<SlideLayoutType>,       // cust
    pub preserve: Option<bool>,                           // false
    pub is_user_drawn: Option<bool>,                      // false
    pub show_master_shapes: Option<bool>,                 // true
    pub show_master_placeholder_animations: Option<bool>, // true
    pub common_slide_data: Box<CommonSlideData>,
    pub color_mapping_override: Option<crate::drawingml::ColorMappingOverride>,
    pub transition: Option<Box<SlideTransition>>,
    pub timing: Option<SlideTiming>,
    pub header_footer: Option<HeaderFooter>,
}

impl SlideLayout {
    pub fn from_zip_file(zip_file: &mut ZipFile<'_>) -> Result<Self> {
        let mut xml_string = String::new();
        zip_file.read_to_string(&mut xml_string)?;
        let xml_node = XmlNode::from_str(xml_string.as_str())?;

        Self::from_xml_element(&xml_node)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut matching_name = None;
        let mut slide_layout_type = None;
        let mut preserve = None;
        let mut is_user_drawn = None;
        let mut show_master_shapes = None;
        let mut show_master_placeholder_animations = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "matchingName" => matching_name = Some(value.clone()),
                "type" => slide_layout_type = Some(value.parse()?),
                "preserve" => preserve = Some(parse_xml_bool(value)?),
                "userDrawn" => is_user_drawn = Some(parse_xml_bool(value)?),
                "showMasterSp" => show_master_shapes = Some(parse_xml_bool(value)?),
                "showMasterPhAnim" => show_master_placeholder_animations = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut common_slide_data = None;
        let mut color_mapping_override = None;
        let mut transition = None;
        let mut timing = None;
        let mut header_footer = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cSld" => common_slide_data = Some(Box::new(CommonSlideData::from_xml_element(child_node)?)),
                "clrMapOvr" => {
                    let clr_map_node = child_node.child_nodes.get(0).ok_or_else(|| {
                        MissingChildNodeError::new(child_node.name.clone(), "masterClrMapping|overrideClrMapping")
                    })?;
                    color_mapping_override =
                        Some(crate::drawingml::ColorMappingOverride::from_xml_element(clr_map_node)?);
                }
                "transition" => transition = Some(Box::new(SlideTransition::from_xml_element(child_node)?)),
                "timing" => timing = Some(SlideTiming::from_xml_element(child_node)?),
                "hf" => header_footer = Some(HeaderFooter::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_slide_data =
            common_slide_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cSld"))?;

        Ok(Self {
            matching_name,
            slide_layout_type,
            preserve,
            is_user_drawn,
            show_master_shapes,
            show_master_placeholder_animations,
            common_slide_data,
            color_mapping_override,
            transition,
            timing,
            header_footer,
        })
    }
}

#[derive(Debug, Clone)]
pub struct Slide {
    pub show: Option<bool>,                               // true
    pub show_master_shapes: Option<bool>,                 // true
    pub show_master_placeholder_animations: Option<bool>, // true
    pub common_slide_data: Box<CommonSlideData>,
    pub color_mapping_override: Option<crate::drawingml::ColorMappingOverride>,
    pub transition: Option<Box<SlideTransition>>,
    pub timing: Option<SlideTiming>,
}

impl Slide {
    pub fn from_zip_file(zip_file: &mut ZipFile<'_>) -> Result<Self> {
        let mut xml_string = String::new();
        zip_file.read_to_string(&mut xml_string)?;
        let xml_node = XmlNode::from_str(xml_string.as_str())?;

        Self::from_xml_element(&xml_node)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut show = None;
        let mut show_master_shapes = None;
        let mut show_master_placeholder_animations = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "show" => show = Some(parse_xml_bool(value)?),
                "showMasterSp" => show_master_shapes = Some(parse_xml_bool(value)?),
                "showMasterPhAnim" => show_master_placeholder_animations = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut common_slide_data = None;
        let mut color_mapping_override = None;
        let mut transition = None;
        let mut timing = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cSld" => common_slide_data = Some(Box::new(CommonSlideData::from_xml_element(child_node)?)),
                "clrMapOvr" => {
                    let clr_map_node = child_node.child_nodes.get(0).ok_or_else(|| {
                        MissingChildNodeError::new(child_node.name.clone(), "masterClrMapping|overrideClrMapping")
                    })?;
                    color_mapping_override =
                        Some(crate::drawingml::ColorMappingOverride::from_xml_element(clr_map_node)?);
                }
                "transition" => transition = Some(Box::new(SlideTransition::from_xml_element(child_node)?)),
                "timing" => timing = Some(SlideTiming::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_slide_data =
            common_slide_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cSld"))?;

        Ok(Self {
            show,
            show_master_shapes,
            show_master_placeholder_animations,
            common_slide_data,
            color_mapping_override,
            transition,
            timing,
        })
    }
}

/// EmbeddedFontListEntry
#[derive(Debug, Clone)]
pub struct EmbeddedFontListEntry {
    pub font: crate::drawingml::TextFont,
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
                "font" => font = Some(crate::drawingml::TextFont::from_xml_element(child_node)?),
                "regular" => {
                    let id_attr = child_node
                        .attribute("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?;
                    regular = Some(id_attr.clone());
                }
                "bold" => {
                    let id_attr = child_node
                        .attribute("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?;
                    bold = Some(id_attr.clone());
                }
                "italic" => {
                    let id_attr = child_node
                        .attribute("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?;
                    italic = Some(id_attr.clone());
                }
                "boldItalic" => {
                    let id_attr = child_node
                        .attribute("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?;
                    bold_italic = Some(id_attr.clone());
                }
                _ => (),
            }
        }

        let font = font.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "font"))?;

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
#[derive(Debug, Clone)]
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
            if child_node.local_name() == "sldLst" {
                for slide_node in &child_node.child_nodes {
                    let id_attr = slide_node
                        .attribute("r:id")
                        .ok_or_else(|| MissingAttributeError::new(slide_node.name.clone(), "r:id"))?;
                    slides.push(id_attr.clone());
                }
            }
        }

        let name = name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?;
        let id = id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;

        Ok(Self { name, id, slides })
    }
}

#[derive(Default, Debug, Clone)]
pub struct PhotoAlbum {
    pub black_and_white: Option<bool>,       // false
    pub show_captions: Option<bool>,         // false
    pub layout: Option<PhotoAlbumLayout>,    // PhotoAlbumLayout::FitToSlide
    pub frame: Option<PhotoAlbumFrameShape>, // PhotoAlbumFrameShape::FrameStyle1
}

impl PhotoAlbum {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "bw" => instance.black_and_white = Some(parse_xml_bool(value)?),
                "showCaptions" => instance.show_captions = Some(parse_xml_bool(value)?),
                "layout" => instance.layout = Some(value.parse()?),
                "frame" => instance.frame = Some(value.parse()?),
                _ => ()
            }
        }

        Ok(instance)
    }
}

#[derive(Default, Debug, Clone)]
pub struct HeaderFooter {
    pub slide_number_enabled: Option<bool>, // true
    pub header_enabled: Option<bool>,       // true
    pub footer_enabled: Option<bool>,       // true
    pub date_time_enabled: Option<bool>,    // true
}

impl HeaderFooter {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "sldNum" => instance.slide_number_enabled = Some(parse_xml_bool(value)?),
                "hdr" => instance.header_enabled = Some(parse_xml_bool(value)?),
                "ftr" => instance.footer_enabled = Some(parse_xml_bool(value)?),
                "dt" => instance.date_time_enabled = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone)]
pub struct Kinsoku {
    pub language: Option<String>,
    pub invalid_start_chars: String,
    pub invalid_end_chars: String,
}

impl Kinsoku {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut language = None;
        let mut invalid_start_chars = None;
        let mut invalid_end_chars = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "lang" => language = Some(value.clone()),
                "invalStChars" => invalid_start_chars = Some(value.clone()),
                "invalEndChars" => invalid_end_chars = Some(value.clone()),
                _ => (),
            }
        }

        let invalid_start_chars = invalid_start_chars
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "invalStChars"))?;
        let invalid_end_chars = invalid_end_chars
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "invalEndChars"))?;

        Ok(Self {
            language,
            invalid_start_chars,
            invalid_end_chars,
        })
    }
}

#[derive(Default, Debug, Clone)]
pub struct ModifyVerifier {
    pub algorithm_name: Option<String>,
    pub hash_value: Option<String>,
    pub salt_value: Option<String>,
    pub spin_value: Option<u32>,
}

impl ModifyVerifier {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "algorithmName" => instance.algorithm_name = Some(value.clone()),
                "hashValue" => instance.hash_value = Some(value.clone()),
                "saltValue" => instance.salt_value = Some(value.clone()),
                "spinValue" => instance.spin_value = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

/// Presentation
#[derive(Default, Debug, Clone)]
pub struct Presentation {
    pub server_zoom: Option<crate::drawingml::Percentage>, // 50%
    pub first_slide_num: Option<i32>,                      // 1
    pub show_special_pls_on_title_slide: Option<bool>,     // true
    pub rtl: Option<bool>,                                 // false
    pub remove_personal_info_on_save: Option<bool>,        // false
    pub compatibility_mode: Option<bool>,                  // false
    pub strict_first_and_last_chars: Option<bool>,         // true
    pub embed_true_type_fonts: Option<bool>,               // false
    pub save_subset_fonts: Option<bool>,                   // false
    pub auto_compress_pictures: Option<bool>,              // true
    pub bookmark_id_seed: Option<BookmarkIdSeed>,          // 1
    pub conformance: Option<ConformanceClass>,
    pub slide_master_id_list: Vec<SlideMasterIdListEntry>,
    pub notes_master_id_list: Vec<NotesMasterIdListEntry>, // length = 1
    pub handout_master_id_list: Vec<HandoutMasterIdListEntry>, // length = 1
    pub slide_id_list: Vec<SlideIdListEntry>,
    pub slide_size: Option<SlideSize>,
    pub notes_size: Option<crate::drawingml::PositiveSize2D>,
    pub smart_tags: Option<RelationshipId>,
    pub embedded_font_list: Vec<Box<EmbeddedFontListEntry>>,
    pub custom_show_list: Vec<CustomShow>,
    pub photo_album: Option<PhotoAlbum>,
    pub customer_data_list: Option<CustomerDataList>,
    pub kinsoku: Option<Box<Kinsoku>>,
    pub default_text_style: Option<Box<crate::drawingml::TextListStyle>>,
    pub modify_verifier: Option<Box<ModifyVerifier>>,
}

impl Presentation {
    pub fn from_zip<R>(zipper: &mut zip::ZipArchive<R>) -> Result<Self>
    where
        R: Read + Seek,
    {
        let mut presentation_file = zipper.by_name("ppt/presentation.xml")?;
        let mut xml_string = String::new();
        presentation_file.read_to_string(&mut xml_string)?;

        let root = XmlNode::from_str(xml_string.as_str())?;
        Self::from_xml_element(&root)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "serverZoom" => instance.server_zoom = Some(value.parse()?),
                "firstSlideNum" => instance.first_slide_num = Some(value.parse()?),
                "showSpecialPlsOnTitleSld" => instance.show_special_pls_on_title_slide = Some(parse_xml_bool(value)?),
                "rtl" => instance.rtl = Some(parse_xml_bool(value)?),
                "removePersonalInfoOnSave" => instance.remove_personal_info_on_save = Some(parse_xml_bool(value)?),
                "compatMode" => instance.compatibility_mode = Some(parse_xml_bool(value)?),
                "strictFirstAndLastChars" => instance.strict_first_and_last_chars = Some(parse_xml_bool(value)?),
                "embedTrueTypeFonts" => instance.embed_true_type_fonts = Some(parse_xml_bool(value)?),
                "saveSubsetFonts" => instance.save_subset_fonts = Some(parse_xml_bool(value)?),
                "autoCompressPictures" => instance.auto_compress_pictures = Some(parse_xml_bool(value)?),
                "bookmarkIdSeed" => instance.bookmark_id_seed = Some(value.parse()?),
                "conformance" => instance.conformance = Some(value.parse()?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "sldMasterIdLst" => {
                    for slide_master_id_node in &child_node.child_nodes {
                        instance.slide_master_id_list.push(
                            SlideMasterIdListEntry::from_xml_element(slide_master_id_node)?,
                        );
                    }
                }
                "notesMasterIdLst" => {
                    for notes_master_id_node in &child_node.child_nodes {
                        instance.notes_master_id_list.push(
                            NotesMasterIdListEntry::from_xml_element(notes_master_id_node)?,
                        );
                    }
                }
                "handoutMasterIdLst" => {
                    for handout_master_id_node in &child_node.child_nodes {
                        instance.handout_master_id_list.push(
                            HandoutMasterIdListEntry::from_xml_element(handout_master_id_node)?,
                        );
                    }
                }
                "sldIdLst" => {
                    for slide_id_node in &child_node.child_nodes {
                        instance.slide_id_list.push(SlideIdListEntry::from_xml_element(slide_id_node)?);
                    }
                }
                "sldSz" => instance.slide_size = Some(SlideSize::from_xml_element(child_node)?),
                "notesSz" => instance.notes_size = Some(crate::drawingml::PositiveSize2D::from_xml_element(child_node)?),
                "smartTags" => {
                    let r_id_attr = child_node
                        .attribute("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?;
                    instance.smart_tags = Some(r_id_attr.clone());
                }
                "embeddedFontLst" => {
                    for embedded_font_node in &child_node.child_nodes {
                        instance.embedded_font_list.push(Box::new(
                            EmbeddedFontListEntry::from_xml_element(embedded_font_node)?
                        ));
                    }
                }
                "custShowLst" => {
                    for custom_show_node in &child_node.child_nodes {
                        instance.custom_show_list.push(CustomShow::from_xml_element(custom_show_node)?);
                    }
                }
                "photoAlbum" => instance.photo_album = Some(PhotoAlbum::from_xml_element(child_node)?),
                "custDataLst" => instance.customer_data_list = Some(CustomerDataList::from_xml_element(child_node)?),
                "kinsoku" => instance.kinsoku = Some(Box::new(Kinsoku::from_xml_element(child_node)?)),
                "defaultTextStyle" => instance.default_text_style = Some(Box::new(
                    crate::drawingml::TextListStyle::from_xml_element(child_node)?
                )),
                "modifyVerifier" => instance.modify_verifier = Some(Box::new(
                    ModifyVerifier::from_xml_element(child_node)?
                )),
                _ => (),
            }
        }

        Ok(instance)
    }
}
