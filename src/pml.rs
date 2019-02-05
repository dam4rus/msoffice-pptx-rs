use msoffice_shared::error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError, XmlError, ParseEnumError};
use msoffice_shared::relationship::RelationshipId;
use msoffice_shared::xml::{parse_xml_bool, XmlNode};
use msoffice_shared::decl_simple_type_enum;
use std::io::{Read, Seek};
use std::str::FromStr;
use zip::read::ZipFile;

use enum_from_str::ParseEnumVariantError;
use enum_from_str_derive::FromStr;

pub type SlideId = u32; // TODO: 256 <= n <= 2147483648
pub type SlideLayoutId = u32; // TODO: 2147483648 <= n
pub type SlideMasterId = u32; // TODO: 2147483648 <= n
/// This simple type defines the position of an object in an ordered list.
pub type Index = u32;
pub type TLTimeNodeId = u32;
/// This simple type specifies constraints for value of the Bookmark ID seed.
/// Values represented by this type are restricted to: 1 <= n <= 2147483648
pub type BookmarkIdSeed = u32;
pub type SlideSizeCoordinate = msoffice_shared::drawingml::PositiveCoordinate32; // TODO: 914400 <= n <= 51206400
/// This simple type specifies a name, such as for a comment author or custom show.
pub type Name = String;
pub type TLSubShapeId = msoffice_shared::drawingml::ShapeId;

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

/// This simple type defines a direction of either horizontal or vertical.
#[derive(Debug, Clone, Copy, PartialEq, FromStr)]
pub enum Direction {
    /// Defines a horizontal direction.
    #[from_str="horz"]
    Horizontal, 
    /// Defines a vertical direction.
    #[from_str="vert"]
    Vertical,
}

/// This simple type facilitates the storing of the size of the placeholder. This size is described relative to the body
/// placeholder on the master.
#[derive(Debug, Clone, Copy, PartialEq, FromStr)]
pub enum PlaceholderSize {
    /// Specifies that the placeholder should take the full size of the body placeholder on the master.
    #[from_str="full"]
    Full,
    /// Specifies that the placeholder should take the half size of the body placeholder on the master. Half size
    /// vertically or horizontally? Needs a picture.
    #[from_str="half"]
    Half,
    /// Specifies that the placeholder should take a quarter of the size of the body placeholder on the master. Picture
    /// would be helpful
    #[from_str="quarter"]
    Quarter,
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

/// This simple type specifies the values for photo layouts within a photo album presentation.
/// See Fundamentals And Markup Language Reference for examples
#[derive(Debug, Clone, Copy, PartialEq, FromStr)]
pub enum PhotoAlbumLayout {
    /// Fit Photos to Slide
    #[from_str="fitToSlide"]
    FitToSlide,
    /// 1 Photo per Slide
    #[from_str="pic1"]
    Pic1,
    /// 2 Photo per Slide
    #[from_str="pic2"]
    Pic2,
    /// 4 Photo per Slide
    #[from_str="pic4"]
    Pic4,
    /// 1 Photo per Slide with Titles
    #[from_str="picTitle1"]
    PicTitle1,
    /// 2 Photo per Slide with Titles
    #[from_str="picTitle2"]
    PicTitle2,
    /// 4 Photo per Slide with Titles
    #[from_str = "picTitle4"]
    PicTitle4,
}

/// This simple type specifies the values for photo frame types within a photo album presentation.
/// See Fundamentals And Markup Language Reference for examples
#[derive(Debug, Clone, Copy, PartialEq, FromStr)]
pub enum PhotoAlbumFrameShape {
    /// Rectangle Photo Frame
    #[from_str="frameStyle1"]
    FrameStyle1,
    /// Rounded Rectangle Photo Frame
    #[from_str="frameStyle2"]
    FrameStyle2,
    /// Simple White Photo Frame
    #[from_str="frameStyle3"]
    FrameStyle3,
    /// Simple Black Photo Frame
    #[from_str="frameStyle4"]
    FrameStyle4,
    /// Compound Black Photo Frame
    #[from_str="frameStyle5"]
    FrameStyle5,
    /// Center Shadow Photo Frame
    #[from_str="frameStyle6"]
    FrameStyle6,
    /// Soft Edge Photo Frame
    #[from_str="frameStyle7"]
    FrameStyle7,
}

/// This simple type determines if the Embedded object is re-colored to reflect changes to the color schemes.
#[derive(Debug, Clone, Copy, PartialEq, FromStr)]
pub enum OleObjectFollowColorScheme {
    /// Setting this enumeration causes the Embedded object to not respond to changes in the color scheme in the
    /// presentation.
    #[from_str="none"]
    None,
    /// Setting this enumeration causes the Embedded object to respond to all changes in the color scheme in the
    /// presentation.
    #[from_str="full"]
    Full,
    /// Setting this enumeration causes the Embedded object to respond only to changes in the text and background
    /// colors of the color scheme in the presentation.
    #[from_str="textAndBackground"]
    TextAndBackground,
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

/// This simple type specifies how the animation is applied over subelements of the target element.
#[derive(Debug, Copy, Clone, PartialEq, FromStr)]
pub enum IterateType {
    /// Iterate by element.
    #[from_str="el"]
    Element,
    /// Iterate by Letter.
    #[from_str="wd"]
    Word,
    /// Iterate by Word.
    #[from_str="lt"]
    Letter,
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
    /// This attribute defines the start of the index range.
    pub start: Index,
    /// This attribute defines the end of the index range.
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
                _ => (),
            }
        }

        let start = start.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "st"))?;
        let end = end.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "end"))?;

        Ok(Self { start, end })
    }
}

#[derive(Debug, Clone)]
pub struct BackgroundProperties {
    /// Specifies whether the background of the slide is of a shade to title background type. This
    /// kind of gradient fill is on the slide background and changes based on the placement of
    /// the slide title placeholder.
    /// 
    /// Defaults to false
    pub shade_to_title: Option<bool>,
    pub fill: msoffice_shared::drawingml::FillProperties,
    pub effect: Option<msoffice_shared::drawingml::EffectProperties>,
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
            use msoffice_shared::drawingml::{EffectProperties, FillProperties};

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
    /// This element specifies visual effects used to render the slide background. This includes any fill, image, or effects
    /// that are to make up the background of the slide.
    Properties(BackgroundProperties),
    /// This element specifies the slide background is to use a fill style defined in the style matrix.
    /// The idx attribute refers to the index of a background fill style or fill style within the presentation's
    /// style matrix, defined by the fmtScheme element.
    /// A value of 0 or 1000 indicates no background, values 1-999 refer to the index of a fill style
    /// within the fillStyleLst element, and values 1001 and above refer to the index of a background fill style within
    /// the bgFillStyleLst element. The value 1001 corresponds to the first background fill style, 1002 to the second
    /// background fill style, and so on.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:bgRef idx="2">
    ///   <a:schemeClr val="bg2"/>
    /// </p:bgRef>
    /// ```
    /// 
    /// The above code indicates a slide background with the style's second fill style using the second background color
    /// of the color scheme.
    /// 
    /// ```xml
    /// <p:bgRef idx="1001">
    ///   <a:schemeClr val="bg2"/>
    /// </p:bgRef>
    /// ```
    /// 
    /// The above code indicates a slide background with the style's first background fill style using the second
    /// background color of the color scheme.
    Reference(msoffice_shared::drawingml::StyleMatrixReference),
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
                msoffice_shared::drawingml::StyleMatrixReference::from_xml_element(xml_node)?,
            )),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_Background").into()),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Background {
    /// Specifies that the background should be rendered using only black and white coloring.
    /// That is, the coloring information for the background should be converted to either black
    /// or white when rendering the picture.
    /// 
    /// # Note
    /// 
    /// No gray is to be used in rendering this background, only stark black and stark
    /// white.
    pub black_and_white_mode: Option<msoffice_shared::drawingml::BlackWhiteMode>, // white
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
    /// Specifies what content type a placeholder is intended to contain.
    pub placeholder_type: Option<PlaceholderType>,
    /// Specifies the orientation of a placeholder.
    pub orientation: Option<Direction>,
    /// Specifies the size of a placeholder.
    pub size: Option<PlaceholderSize>,
    /// Specifies the placeholder index. This is used when applying templates or changing
    /// layouts to match a placeholder on one template/master to another.
    pub index: Option<u32>,
    /// Specifies whether the corresponding placeholder should have a custom prompt or not.
    pub has_custom_prompt: Option<bool>,
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

/// This element specifies non-visual properties for objects. These properties include multimedia content associated
/// with an object and properties indicating how the object is to be used or displayed in different contexts.
#[derive(Default, Debug, Clone)]
pub struct ApplicationNonVisualDrawingProps {
    /// Specifies whether the picture belongs to a photo album and should thus be included
    /// when editing a photo album within the generating application.
    pub is_photo: Option<bool>,
    /// Specifies if the corresponding object has been drawn by the user and should thus not be
    /// deleted. This allows for the flagging of slides that contain user drawn data.
    pub is_user_drawn: Option<bool>, // false
    /// This element specifies that the corresponding shape should be represented by the generating application as a
    /// placeholder. When a shape is considered a placeholder by the generating application it can have special
    /// properties to alert the user that they can enter content into the shape. Different placeholder types are allowed
    /// and can be specified by using the placeholder type attribute for this element.
    pub placeholder: Option<Placeholder>,
    pub media: Option<msoffice_shared::drawingml::Media>,
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
            use msoffice_shared::drawingml::Media;

            let local_name = child_node.local_name();
            if Media::is_choice_member(local_name) {
                instance.media = Some(Media::from_xml_element(child_node)?);
            } else {
                match child_node.local_name() {
                    "ph" => instance.placeholder = Some(Placeholder::from_xml_element(child_node)?),
                    "custDataLst" => {
                        instance.customer_data_list = Some(CustomerDataList::from_xml_element(child_node)?)
                    }
                    _ => (),
                }
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone)]
pub enum ShapeGroup {
    /// This element specifies the existence of a single shape. A shape can either be a preset or a custom geometry,
    /// defined using the DrawingML framework. In addition to a geometry each shape can have both visual and non-
    /// visual properties attached. Text and corresponding styling information can also be attached to a shape. This
    /// shape is specified along with all other shapes within either the shape tree or group shape elements.
    Shape(Box<Shape>),
    /// This element specifies a group shape that represents many shapes grouped together. This shape is to be treated
    /// just as if it were a regular shape but instead of being described by a single geometry it is made up of all the
    /// shape geometries encompassed within it. Within a group shape each of the shapes that make up the group are
    /// specified just as they normally would. The idea behind grouping elements however is that a single transform can
    /// apply to many shapes at the same time.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:grpSp>
    ///   <p:nvGrpSpPr>
    ///     <p:cNvPr id="10" name="Group 9"/>
    ///     <p:cNvGrpSpPr/>
    ///     <p:nvPr/>
    ///   </p:nvGrpSpPr>
    ///   <p:grpSpPr>
    ///     <a:xfrm>
    ///       <a:off x="838200" y="990600"/>
    ///       <a:ext cx="2426208" cy="978408"/>
    ///       <a:chOff x="838200" y="990600"/>
    ///       <a:chExt cx="2426208" cy="978408"/>
    ///     </a:xfrm>
    ///   </p:grpSpPr>
    ///   <p:sp>
    ///   ...
    ///   </p:sp>
    ///   <p:sp>
    ///   ...
    ///   </p:sp>
    ///   <p:sp>
    ///   ...
    ///   </p:sp>
    /// </p:grpSp>
    /// ```
    /// 
    /// In the above example we see three shapes specified within a single group. These three shapes have their
    /// position and sizes specified just as they normally would within the shape tree.
    /// The generating application should apply the transformation after the bounding box for the group shape has been
    /// calculated.
    GroupShape(Box<GroupShape>),
    /// This element specifies the existence of a graphics frame. This frame contains a graphic that was generated by an
    /// external source and needs a container in which to be displayed on the slide surface.
    GraphicFrame(Box<GraphicalObjectFrame>),
    /// This element specifies a connection shape that is used to connect two sp elements.
    /// Once a connection is specified using a Connector, it is left to the generating application to determine the
    /// exact path the connector takes.
    /// That is the connector routing algorithm is left up to the generating application as the desired path might be
    /// different depending on the specific needs of the application.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:spTree>
    ///   ...
    ///   <p:sp>
    ///     <p:nvSpPr>
    ///       <p:cNvPr id="1" name="Rectangle 1"/>
    ///       <p:cNvSpPr/>
    ///       <p:nvPr/>
    ///     </p:nvSpPr>
    ///   ...
    ///   </p:sp>
    ///   <p:sp>
    ///     <p:nvSpPr>
    ///       <p:cNvPr id="2" name="Rectangle 2"/>
    ///       <p:cNvSpPr/>
    ///       <p:nvPr/>
    ///     </p:nvSpPr>
    ///     ...
    ///   </p:sp>
    ///   <p:cxnSp>
    ///     <p:nvCxnSpPr>
    ///       <p:cNvPr id="3" name="Elbow Connector 3"/>
    ///       <p:cNvCxnSpPr>
    ///         <a:stCxn id="1" idx="3"/>
    ///         <a:endCxn id="2" idx="1"/>
    ///       </p:cNvCxnSpPr>
    ///       <p:nvPr/>
    ///     </p:nvCxnSpPr>
    ///     ...
    ///   </p:cxnSp>
    /// </p:spTree>
    /// ```
    Connector(Box<Connector>),
    /// This element specifies the existence of a picture object within the document.
    /// 
    /// # Xml example
    /// 
    /// Consider the following PresentationML that specifies the existence of a picture within a document.
    /// This picture can have non-visual properties, a picture fill as well as shape properties attached to it.
    /// 
    /// ```xml
    /// <p:pic>
    ///   <p:nvPicPr>
    ///     <p:cNvPr id="4" name="lake.JPG" descr="Picture of a Lake" />
    ///     <p:cNvPicPr>
    ///       <a:picLocks noChangeAspect="1"/>
    ///     </p:cNvPicPr>
    ///     <p:nvPr/>
    ///   </p:nvPicPr>
    ///   <p:blipFill>
    ///   ...
    ///   </p:blipFill>
    ///   <p:spPr>
    ///   ...
    ///   </p:spPr>
    /// </p:pic>
    /// ```
    Picture(Box<Picture>),
    /// This element specifies a reference to XML content in a format not defined by ECMA-376.
    /// 
    /// The relationship type of the explicit relationship specified by this element shall be
    /// http://purl.oclc.org/ooxml/officeDocument/relationships/customXml and have a TargetMode attribute value of
    /// Internal. If an application cannot process content of the content type specified by the targeted part, then it
    /// should continue to process the file.
    /// If possible, it should also provide some indication that unknown content was not imported.
    /// 
    /// # Note
    /// 
    /// This part allows the native use of other commonly used interchange formats, such as:
    /// * [MathML](http://www.w3.org/TR/MathML2/)
    /// * [SMIL](http://www.w3.org/TR/REC-smil/)
    /// * [SVG](http://www.w3.org/TR/SVG11/)
    /// 
    /// For better interoperability, only standard XML formats should be used.
    ContentPart(RelationshipId),
}

impl ShapeGroup {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
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
            "graphicFrame" => Ok(ShapeGroup::GraphicFrame(Box::new(
                GraphicalObjectFrame::from_xml_element(xml_node)?,
            ))),
            "cxnSp" => Ok(ShapeGroup::Connector(Box::new(Connector::from_xml_element(xml_node)?))),
            "pic" => Ok(ShapeGroup::Picture(Box::new(Picture::from_xml_element(xml_node)?))),
            "contentPart" => {
                let attr = xml_node
                    .attribute("r:id")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;
                Ok(ShapeGroup::ContentPart(attr.clone()))
            }
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "EG_ShapeGroup",
            ))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Shape {
    /// Specifies that the shape fill should be set to that of the slide background surface.
    /// 
    /// # Note
    /// 
    /// This attribute does not set the fill of the shape to be transparent but instead sets it
    /// to be filled with the portion of the slide background that is directly behind it.
    pub use_bg_fill: Option<bool>,
    /// This element specifies all non-visual properties for a shape. This element is a container for the non-visual
    /// identification properties, shape properties and application properties that are to be associated with a shape.
    /// This allows for additional information that does not affect the appearance of the shape to be stored.
    pub non_visual_props: Box<ShapeNonVisual>,
    /// This element specifies the visual shape properties that can be applied to a shape. These properties include the
    /// shape fill, outline, geometry, effects, and 3D orientation.
    pub shape_props: Box<msoffice_shared::drawingml::ShapeProperties>,
    /// This element specifies the style information for a shape. This is used to define a shape's appearance in terms of
    /// the preset styles defined by the style matrix for the theme.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:style>
    ///   <a:lnRef idx="3">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:lnRef>
    ///   <a:fillRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:fillRef>
    ///   <a:effectRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:effectRef>
    ///   <a:fontRef idx="minor">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:fontRef>
    /// </p:style>
    /// ```
    /// 
    /// The parent shape of the above code is to have an outline that uses the third line style defined by the theme, use
    /// the first fill defined by the scheme, and be rendered with the first effect defined by the theme. Text inside the
    /// shape is to use the minor font defined by the theme.
    pub shape_style: Option<Box<msoffice_shared::drawingml::ShapeStyle>>,
    /// This element specifies the existence of text to be contained within the corresponding shape. All visible text and
    /// visible text related properties are contained within this element. There can be multiple paragraphs and within
    /// paragraphs multiple runs of text.
    pub text_body: Option<msoffice_shared::drawingml::TextBody>,
}

impl Shape {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let use_bg_fill = match xml_node.attribute("useBgFill") {
            Some(val) => Some(parse_xml_bool(val)?),
            None => None,
        };

        let mut non_visual_props = None;
        let mut shape_props = None;
        let mut shape_style = None;
        let mut text_body = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvSpPr" => non_visual_props = Some(Box::new(ShapeNonVisual::from_xml_element(child_node)?)),
                "spPr" => {
                    shape_props = Some(Box::new(msoffice_shared::drawingml::ShapeProperties::from_xml_element(
                        child_node,
                    )?))
                }
                "style" => shape_style = Some(Box::new(msoffice_shared::drawingml::ShapeStyle::from_xml_element(child_node)?)),
                "txBody" => text_body = Some(msoffice_shared::drawingml::TextBody::from_xml_element(child_node)?),
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
            shape_style,
            text_body,
        })
    }
}

#[derive(Debug, Clone)]
pub struct ShapeNonVisual {
    pub drawing_props: Box<msoffice_shared::drawingml::NonVisualDrawingProps>,
    /// This element specifies the non-visual drawing properties for a shape. These properties are to be used by the
    /// generating application to determine how the shape should be dealt with.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:sp>
    ///   <p:nvSpPr>
    ///     <p:cNvPr id="2" name="Rectangle 1"/>
    ///     <p:cNvSpPr>
    ///       <a:spLocks noGrp="1"/>
    ///     </p:cNvSpPr>
    ///   </p:nvSpPr>
    ///   ...
    /// </p:sp>
    /// ```
    /// 
    /// This shape lock is stored within the non-visual drawing properties for this shape.
    pub shape_drawing_props: msoffice_shared::drawingml::NonVisualDrawingShapeProps,
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
                    drawing_props = Some(Box::new(msoffice_shared::drawingml::NonVisualDrawingProps::from_xml_element(
                        child_node,
                    )?))
                }
                "cNvSpPr" => {
                    shape_drawing_props = Some(msoffice_shared::drawingml::NonVisualDrawingShapeProps::from_xml_element(
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
    /// This element specifies all non-visual properties for a group shape. This element is a container for the
    /// non-visual identification properties, shape properties and application properties that are to be associated
    /// with a group shape.
    /// This allows for additional information that does not affect the appearance of the group shape to be stored.
    pub non_visual_props: Box<GroupShapeNonVisual>,
    /// This element specifies the properties that are to be common across all of the shapes within the corresponding
    /// group. If there are any conflicting properties within the group shape properties and the individual shape
    /// properties then the individual shape properties should take precedence.
    pub group_shape_props: msoffice_shared::drawingml::GroupShapeProperties,
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
                    "nvGrpSpPr" => {
                        non_visual_props = Some(Box::new(GroupShapeNonVisual::from_xml_element(child_node)?))
                    }
                    "grpSpPr" => {
                        group_shape_props = Some(msoffice_shared::drawingml::GroupShapeProperties::from_xml_element(child_node)?)
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
    pub drawing_props: Box<msoffice_shared::drawingml::NonVisualDrawingProps>,
    /// This element specifies the non-visual drawing properties for a group shape. These non-visual properties are
    /// properties that the generating application would utilize when rendering the slide surface.
    pub group_drawing_props: msoffice_shared::drawingml::NonVisualGroupDrawingShapeProps,
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
                    drawing_props = Some(Box::new(msoffice_shared::drawingml::NonVisualDrawingProps::from_xml_element(
                        child_node,
                    )?))
                }
                "cNvGrpSpPr" => {
                    group_drawing_props = Some(msoffice_shared::drawingml::NonVisualGroupDrawingShapeProps::from_xml_element(
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
    /// Specifies how the graphical object should be rendered, using color, black or white,
    /// or grayscale.
    /// 
    /// # Note
    /// 
    /// This does not mean that the graphical object itself is stored with only black
    /// and white or grayscale information. This attribute instead sets the rendering mode
    /// that the graphical object uses.
    pub black_white_mode: Option<msoffice_shared::drawingml::BlackWhiteMode>,
    /// This element specifies all non-visual properties for a graphic frame. This element is a container for the
    /// non-visual identification properties, shape properties and application properties that are to be associated
    /// with a graphic frame.
    /// This allows for additional information that does not affect the appearance of the graphic frame to be stored.
    pub non_visual_props: Box<GraphicalObjectFrameNonVisual>,
    /// This element specifies the transform to be applied to the corresponding graphic frame. This transformation is
    /// applied to the graphic frame just as it would be for a shape or group shape.
    pub transform: Box<msoffice_shared::drawingml::Transform2D>,
    pub graphic: msoffice_shared::drawingml::GraphicalObject,
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
                "nvGraphicFramePr" => {
                    non_visual_props = Some(Box::new(GraphicalObjectFrameNonVisual::from_xml_element(child_node)?))
                }
                "xfrm" => transform = Some(Box::new(msoffice_shared::drawingml::Transform2D::from_xml_element(child_node)?)),
                "graphic" => graphic = Some(msoffice_shared::drawingml::GraphicalObject::from_xml_element(child_node)?),
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
    pub drawing_props: Box<msoffice_shared::drawingml::NonVisualDrawingProps>,
    /// This element specifies the non-visual drawing properties for a graphic frame. These non-visual properties are
    /// properties that the generating application would utilize when rendering the slide surface.
    pub graphic_frame_props: msoffice_shared::drawingml::NonVisualGraphicFrameProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl GraphicalObjectFrameNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut graphic_frame_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => {
                    drawing_props = Some(Box::new(msoffice_shared::drawingml::NonVisualDrawingProps::from_xml_element(
                        child_node,
                    )?))
                }
                "cNvGraphicFramePr" => {
                    graphic_frame_props = Some(msoffice_shared::drawingml::NonVisualGraphicFrameProperties::from_xml_element(
                        child_node,
                    )?)
                }
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let graphic_frame_props = graphic_frame_props
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvGraphicFramePr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;

        Ok(Self {
            drawing_props,
            graphic_frame_props,
            app_props,
        })
    }
}

#[derive(Debug, Clone)]
pub struct Connector {
    /// This element specifies all non-visual properties for a connection shape. This element is a container for the non-
    /// visual identification properties, shape properties and application properties that are to be associated with a
    /// connection shape. This allows for additional information that does not affect the appearance of the connection
    /// shape to be stored.
    pub non_visual_props: Box<ConnectorNonVisual>,
    /// This element specifies the visual shape properties that can be applied to a shape. These properties include the
    /// shape fill, outline, geometry, effects, and 3D orientation.
    pub shape_props: Box<msoffice_shared::drawingml::ShapeProperties>,
    /// This element specifies the style information for a shape. This is used to define a shape's appearance in terms of
    /// the preset styles defined by the style matrix for the theme.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:style>
    ///   <a:lnRef idx="3">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:lnRef>
    ///   <a:fillRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:fillRef>
    ///   <a:effectRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:effectRef>
    ///   <a:fontRef idx="minor">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:fontRef>
    /// </p:style>
    /// ```
    /// 
    /// The parent shape of the above code is to have an outline that uses the third line style defined by the theme, use
    /// the first fill defined by the scheme, and be rendered with the first effect defined by the theme. Text inside the
    /// shape is to use the minor font defined by the theme.
    pub shape_style: Option<Box<msoffice_shared::drawingml::ShapeStyle>>,
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
                    shape_props = Some(Box::new(msoffice_shared::drawingml::ShapeProperties::from_xml_element(
                        child_node,
                    )?))
                }
                "style" => shape_style = Some(Box::new(msoffice_shared::drawingml::ShapeStyle::from_xml_element(child_node)?)),
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
    pub drawing_props: Box<msoffice_shared::drawingml::NonVisualDrawingProps>,
    /// This element specifies the non-visual drawing properties specific to a connector shape. This includes
    /// information specifying the shapes to which the connector shape is connected.
    pub connector_props: msoffice_shared::drawingml::NonVisualConnectorProperties,
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
                    drawing_props = Some(Box::new(msoffice_shared::drawingml::NonVisualDrawingProps::from_xml_element(
                        child_node,
                    )?))
                }
                "cNvCxnSpPr" => {
                    connector_props = Some(msoffice_shared::drawingml::NonVisualConnectorProperties::from_xml_element(
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
    /// This element specifies all non-visual properties for a picture. This element is a container for the non-visual
    /// identification properties, shape properties and application properties that are to be associated with a picture.
    /// This allows for additional information that does not affect the appearance of the picture to be stored.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:pic>
    ///   ...
    ///   <p:nvPicPr>
    ///   ...
    ///   </p:nvPicPr>
    ///   ...
    /// </p:pic>
    /// ```
    pub non_visual_props: Box<PictureNonVisual>,
    pub blip_fill: Box<msoffice_shared::drawingml::BlipFillProperties>,
    pub shape_props: Box<msoffice_shared::drawingml::ShapeProperties>,
    /// This element specifies the style information for a shape. This is used to define a shape's appearance in terms of
    /// the preset styles defined by the style matrix for the theme.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:style>
    ///   <a:lnRef idx="3">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:lnRef>
    ///   <a:fillRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:fillRef>
    ///   <a:effectRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:effectRef>
    ///   <a:fontRef idx="minor">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:fontRef>
    /// </p:style>
    /// ```
    /// 
    /// The parent shape of the above code is to have an outline that uses the third line style defined by the theme, use
    /// the first fill defined by the scheme, and be rendered with the first effect defined by the theme. Text inside the
    /// shape is to use the minor font defined by the theme.
    pub shape_style: Option<Box<msoffice_shared::drawingml::ShapeStyle>>,
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
                    blip_fill = Some(Box::new(msoffice_shared::drawingml::BlipFillProperties::from_xml_element(
                        child_node,
                    )?))
                }
                "spPr" => {
                    shape_props = Some(Box::new(msoffice_shared::drawingml::ShapeProperties::from_xml_element(
                        child_node,
                    )?))
                }
                "style" => shape_style = Some(Box::new(msoffice_shared::drawingml::ShapeStyle::from_xml_element(child_node)?)),
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
    pub drawing_props: Box<msoffice_shared::drawingml::NonVisualDrawingProps>,
    /// This element specifies the non-visual properties for the picture canvas. These properties are to be used by the
    /// generating application to determine how certain properties are to be changed for the picture object in question.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:pic>
    ///   ...
    ///   <p:nvPicPr>
    ///     <p:cNvPr id="4" name="Lilly_by_Lisher.jpg"/>
    ///     <p:cNvPicPr>
    ///       <a:picLocks noChangeAspect="1"/>
    ///     </p:cNvPicPr>
    ///     <p:nvPr/>
    ///   </p:nvPicPr>
    ///   ...
    /// </p:pic>
    /// ```
    pub picture_props: msoffice_shared::drawingml::NonVisualPictureProperties,
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
                    drawing_props = Some(Box::new(msoffice_shared::drawingml::NonVisualDrawingProps::from_xml_element(
                        child_node,
                    )?))
                }
                "cNvPicPr" => {
                    picture_props = Some(msoffice_shared::drawingml::NonVisualPictureProperties::from_xml_element(
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

/// This element specifies a container for slide information that is relevant to all of the slide types.
/// All slides share a common set of properties that is independent of the slide type; the description of these
/// properties for any particular slide is stored within the slide's common_slide_data container.
/// Slide data specific to the slide type indicated by the parent element is stored elsewhere.
/// 
/// # Note
/// 
/// The actual data in CommonSlideData describe only the particular parent slide; it is only the kind of information
/// stored that is common across all slides.
/// 
/// # Xml example
/// 
/// ```xml
/// <p:sld>
///   <p:cSld>
///     <p:spTree>
///     ...
///     </p:spTree>
///   </p:cSld>
///   ...
/// </p:sld>
/// ```
#[derive(Debug, Clone)]
pub struct CommonSlideData {
    /// Specifies the slide name property that is used to further identify this unique configuration
    /// of common slide data. This might be used to aid in distinguishing different slide layouts or
    /// various other slide types.
    pub name: Option<String>,
    /// This element specifies the background appearance information for a slide. The slide background covers the
    /// entire slide and is visible where no objects exist and as the background for transparent objects.
    pub background: Option<Box<Background>>,
    /// This element specifies all shape-based objects, either grouped or not, that can be referenced on a given slide. As
    /// most objects within a slide are shapes, this represents the majority of content within a slide. Text and effects are
    /// attached to shapes that are contained within the shape_tree element.
    /// 
    /// Each shape-based object within the shape tree, whether grouped or not, shall represent one unique level of z-
    /// ordering on the slide. The z-order for each shape-based object shall be determined by the lexical ordering of
    /// each shape-based object within the shape tree: the first shape-based object shall have the lowest z-order, while
    /// the last shape-based object shall have the highest z-order.
    /// 
    /// The z-ordering of shape-based objects within the shape tree shall also determine the navigation (tab) order of
    /// the shape-based objects: the shape-based object with the lowest z-order (the first shape in lexical order) shall be
    /// first in navigation order, with objects being navigated in ascending z-order.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:sld>
    ///   <p:cSld>
    ///     <p:spTree>
    ///       <p:nvGrpSpPr>
    ///       ...
    ///       </p:nvGrpSpPr>
    ///       <p:grpSpPr>
    ///       ...
    ///       </p:grpSpPr>
    ///       <p:sp>
    ///       ...
    ///       </p:sp>
    ///     </p:spTree>
    ///   </p:cSld>
    ///   ...
    /// </p:sld>
    /// ```
    pub shape_tree: Box<GroupShape>,
    pub customer_data_list: Option<CustomerDataList>,
    /// This element specifies a list of embedded controls for the corresponding slide. Custom embedded controls can
    /// be embedded on slides.
    pub control_list: Option<Vec<Control>>,
}

impl CommonSlideData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name = xml_node.attribute("name").cloned();
        let mut background = None;
        let mut shape_tree = None;
        let mut customer_data_list = None;
        let mut control_list = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "bg" => background = Some(Box::new(Background::from_xml_element(child_node)?)),
                "spTree" => shape_tree = Some(Box::new(GroupShape::from_xml_element(child_node)?)),
                "custDataList" => customer_data_list = Some(CustomerDataList::from_xml_element(child_node)?),
                "controls" => {
                    let mut vec = Vec::new();
                    for control_node in &child_node.child_nodes {
                        vec.push(Control::from_xml_element(control_node)?);
                    }

                    control_list = Some(vec);
                }
                _ => (),
            }
        }

        let shape_tree = shape_tree
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
    /// This element specifies the existence of customer data in the form of tags. This allows for the storage of customer
    /// data within the PresentationML framework. While this is similar to the ext tag in that it can be used store
    /// information, this tag mainly focuses on referencing to other parts of the presentation document. This is
    /// accomplished via the relationship identification attribute that is required for all specified tags.
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
    pub shape_id: Option<msoffice_shared::drawingml::ShapeId>,
    pub name: Option<String>,       // ""
    pub show_as_icon: Option<bool>, // false
    pub id: Option<RelationshipId>,
    pub image_width: Option<msoffice_shared::drawingml::PositiveCoordinate32>,
    pub image_height: Option<msoffice_shared::drawingml::PositiveCoordinate32>,
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

        let relationship_id =
            relationship_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

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

        Ok(Self {
            relationship_id: r_id_attr.clone(),
        })
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

        Ok(Self {
            relationship_id: r_id_attr.clone(),
        })
    }
}

#[derive(Default, Debug, Clone)]
pub struct SlideMasterTextStyles {
    /// This element specifies the text formatting style for the title text within a master slide. This formatting is used on
    /// all title text within related presentation slides. The text formatting is specified by utilizing the DrawingML
    /// framework just as within a regular presentation slide. Within a title style there can be many different style types
    /// defined as there are different kinds of text stored within a slide title.
    pub title_styles: Option<Box<msoffice_shared::drawingml::TextListStyle>>,
    /// This element specifies the text formatting style for all body text within a master slide.
    /// This formatting is used on all body text within presentation slides related to this master.
    /// The text formatting is specified by utilizing the DrawingML framework just as within a regular
    /// presentation slide.
    /// Within the bodyStyle element there can be many different style types defined as there are different kinds of
    /// text stored within the body of a slide.
    pub body_styles: Option<Box<msoffice_shared::drawingml::TextListStyle>>,
    /// This element specifies the text formatting style for the all other text within a master slide. This formatting is
    /// used on all text not covered by the title_styles or body_styles elements within related presentation slides. The text
    /// formatting is specified by utilizing the DrawingML framework just as within a regular presentation slide. Within
    /// the otherStyle element there can be many different style types defined as there are different kinds of text
    /// stored within a slide.
    /// 
    /// # Note
    /// 
    /// The other_styles element is to be used for specifying the text formatting of text within a slide shape but
    /// not within a text box. Text box styling is handled from within the body_styles element.
    pub other_styles: Option<Box<msoffice_shared::drawingml::TextListStyle>>,
}

impl SlideMasterTextStyles {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "titleStyle" => {
                    instance.title_styles =
                        Some(Box::new(msoffice_shared::drawingml::TextListStyle::from_xml_element(child_node)?))
                }
                "bodyStyle" => {
                    instance.body_styles =
                        Some(Box::new(msoffice_shared::drawingml::TextListStyle::from_xml_element(child_node)?))
                }
                "otherStyle" => {
                    instance.other_styles =
                        Some(Box::new(msoffice_shared::drawingml::TextListStyle::from_xml_element(child_node)?))
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Default, Debug, Clone)]
pub struct OrientationTransition {
    /// This attribute specifies a horizontal or vertical transition.
    /// 
    /// Defaults to Direction::Horizontal
    pub direction: Option<Direction>,
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
    /// This attribute specifies if the direction of the transition.
    /// 
    /// Defaults to TransitionEightDirectionType::Left
    pub direction: Option<TransitionEightDirectionType>,
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
    /// This attribute specifies if the transition starts from a black screen (and then transition the
    /// new slide over black).
    /// 
    /// Defaults to false
    pub through_black: Option<bool>,
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
    /// This attribute specifies the direction of the slide transition.
    /// 
    /// Defaults to TransitionSideDirectionType::Left
    pub direction: Option<TransitionSideDirectionType>,
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
    /// This attribute specifies the orientation of a "split" slide transition.
    /// 
    /// Defaults to Direction::Horizontal
    pub orientation: Option<Direction>,
    /// This attribute specifies the direction of a "split" slide transition.
    /// 
    /// Defaults to TransitionInOutDirectionType::Out
    pub direction: Option<TransitionInOutDirectionType>,
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
    /// This attribute specifies if the direction of the transition.
    /// 
    /// Defaults to TransitionCornerDirectionType::LeftUp
    pub direction: Option<TransitionCornerDirectionType>,
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
    /// This attributes specifies the number of spokes ("pie pieces") in the wheel
    /// 
    /// Defaults to 4
    pub spokes: Option<u32>,
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
    /// This element describes the blinds slide transition effect, which uses a set of horizontal or vertical bars and wipes
    /// them either left-to-right or top-to-bottom, respectively, until the new slide is fully shown. The rendering of this
    /// transition depends upon the attributes specified.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:blinds dir="horz"/>
    /// </p:transition>
    /// ```
    Blinds(OrientationTransition),
    /// This element describes the checker slide transition effect, which uses a set of horizontal or vertical
    /// checkerboard squares and wipes them either left-to-right or top-to-bottom, respectively, until the new slide is
    /// fully shown. The rendering of this transition depends upon the attributes specified.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:checker dir="horz"/>
    /// </p:transition>
    /// ```
    Checker(OrientationTransition),
    /// This element describes the circle slide transition effect, which uses a circle pattern centered on the slide that
    /// increases in size until the new slide is fully shown. The rendering of this transition has been shown below.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:circle/>
    /// </p:transition>
    /// ```
    Circle,
    /// This element describes the dissolve slide transition effect, which uses a set of randomly placed squares on the
    /// slide that continue to be added to until the new slide is fully shown. The rendering of this transition has been
    /// shown below.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:dissolve/>
    /// </p:transition>
    /// ```
    Dissolve,
    /// This element describes the comb slide transition effect, which uses a set of horizontal or vertical bars and wipes
    /// them from one end of the slide to the other until the new slide is fully shown. The rendering of this transition
    /// depends upon the attributes specified which have been shown below.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:comb dir="horz"/>
    /// </p:transition>
    /// ```
    Comb(OrientationTransition),
    /// This element describes the cover slide transition effect, which moves the new slide in from an off-screen
    /// location, continually covering more of the previous slide until the new slide is fully shown. The rendering of this
    /// transition depends upon the attributes specified which have been shown below.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:cover dir="d"/>
    /// </p:transition>
    /// ```
    Cover(EightDirectionTransition),
    /// This element describes the cut slide transition effect, which simply replaces the previous slide with the new slide
    /// instantaneously. No animation is used, but an option exists to cut to a black screen before showing the new
    /// slide. The rendering of this transition depends upon the attributes specified which have been shown below.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:cut thruBlk="0"/>
    /// </p:transition>
    /// ```
    Cut(OptionalBlackTransition),
    /// This element describes the diamond slide transition effect, which uses a diamond pattern centered on the slide
    /// that increases in size until the new slide is fully shown. The rendering of this transition has been shown below.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:diamond/>
    /// </p:transition>
    /// ```
    Diamond,
    /// This element describes the fade slide transition effect, which smoothly fades the previous slide either directly to
    /// the new slide or first to a black screen and then to the new slide. The rendering of this transition depends upon
    /// the attributes specified which have been shown below.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:fade thruBlk="0"/>
    /// </p:transition>
    /// ```
    Fade(OptionalBlackTransition),
    /// This element describes the newsflash slide transition effect, which grows and spins the new slide counterclockwise
    /// into place over the previous slide. The rendering of this transition has been shown below.
    /// 
    /// # Xml example
    /// 
    /// Consider the following case in which the “newsflash” slide transition is applied to a slide, along with a
    /// set of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:newsflash/>
    /// </p:transition>
    /// ```
    Newsflash,
    /// This element describes the plus slide transition effect, which uses a plus pattern centered on the slide that
    /// increases in size until the new slide is fully shown.
    /// 
    /// # Xml example
    /// 
    /// Consider the following case in which the “plus” slide transition is applied to a slide, along with a set of
    /// attributes
    /// ```xml
    /// <p:transition>
    ///   <p:plus/>
    /// </p:transition>
    /// ```
    Plus,
    /// This element describes the pull slide transition effect, which moves the previous slide to an off-screen location,
    /// continually revealing more of the new slide until the new slide is fully shown.
    /// 
    /// # Xml example
    /// 
    /// Consider the following cases in which the “pull” slide transition is applied to a slide, along with a set of
    /// attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:pull dir="d"/>
    /// </p:transition>
    /// ```
    Pull(EightDirectionTransition),
    /// This element describes the push slide transition effect, which moves the new slide in from an off-screen
    /// location, continually pushing the previous slide to an opposite off-screen location until the new slide is fully
    /// shown.
    /// 
    /// # Xml example
    /// 
    /// Consider the following cases in which the “push” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:push dir="d"/>
    /// </p:transition>
    /// ```
    Push(SideDirectionTransition),
    /// This element describes the random slide transition effect, which chooses a random transition from the set
    /// available in the rendering application. This transition thus can be different each time it is used.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:random/>
    /// </p:transition>
    /// ```
    Random,
    /// This element describes the randomBar slide transition effect, which uses a set of randomly placed horizontal or
    /// vertical bars on the slide that continue to be added to until the new slide is fully shown.
    /// 
    /// # Xml example
    /// 
    /// Consider the following cases in which the “randomBar” slide transition is applied to a slide, along with
    /// a set of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:randomBar dir="horz"/>
    /// </p:transition>
    /// ```
    RandomBar(OrientationTransition),
    /// This element describes the split slide transition effect, which reveals the new slide directly on top of the
    /// previous one by wiping either horizontal or vertical from the outside in, or from the inside out, until the new
    /// slide is fully shown.
    /// 
    /// # Xml example
    /// 
    /// Consider the following cases in which the “split” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:split orient="horz" dir="in"/>
    /// </p:transition>
    /// ```
    Split(SplitTransition),
    /// This element describes the strips slide transition effect, which uses a set of bars that are arranged in a staggered
    /// fashion and wipes them across the screen until the new slide is fully shown.
    /// 
    /// # Xml example
    /// 
    /// Consider the following cases in which the “strips” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:strips dir="ld"/>
    /// </p:transition>
    /// ```
    Strips(CornerDirectionTransition),
    /// This element describes the wedge slide transition effect, which uses two radial edges that wipe from top to
    /// bottom in opposite directions until the new slide is fully shown.
    /// 
    /// # Xml example
    /// 
    /// Consider the following case in which the “wedge” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:wedge/>
    /// </p:transition>
    /// ```
    Wedge,
    /// This element describes the wheel slide transition effect, which uses a set of radial edges and wipes them in the
    /// clockwise direction until the new slide is fully shown.
    /// 
    /// # Xml example
    /// 
    /// Consider the following cases in which the “wheel” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:wheel spokes="1"/>
    /// </p:transition>
    /// ```
    Wheel(WheelTransition),
    /// This element describes the wipe slide transition effect, which wipes the new slide over the previous slide from
    /// one edge of the screen to the opposite until the new slide is fully shown.
    /// 
    /// # Xml example
    /// 
    /// Consider the following cases in which the “wipe” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:wipe dir="d"/>
    /// </p:transition>
    /// ```
    Wipe(SideDirectionTransition),
    /// This element describes the zoom slide transition effect, which uses a box pattern centered on the slide that
    /// increases in size until the new slide is fully shown.
    /// 
    /// # Xml example
    /// 
    /// Consider the following cases in which the “zoom” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:zoom dir="in"/>
    /// </p:transition>
    /// ```
    Zoom(InOutTransition),
}

impl SlideTransitionGroup {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "blinds" | "checker" | "circle" | "dissolve" | "comb" | "cover" | "cut" | "diamond" | "fade"
            | "newsflash" | "plus" | "pull" | "push" | "random" | "randomBar" | "split" | "strips" | "wedge"
            | "wheel" | "wipe" | "zoom" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "blinds" => Ok(SlideTransitionGroup::Blinds(OrientationTransition::from_xml_element(
                xml_node,
            )?)),
            "checker" => Ok(SlideTransitionGroup::Checker(OrientationTransition::from_xml_element(
                xml_node,
            )?)),
            "circle" => Ok(SlideTransitionGroup::Circle),
            "dissolve" => Ok(SlideTransitionGroup::Dissolve),
            "comb" => Ok(SlideTransitionGroup::Comb(OrientationTransition::from_xml_element(
                xml_node,
            )?)),
            "cover" => Ok(SlideTransitionGroup::Cover(EightDirectionTransition::from_xml_element(
                xml_node,
            )?)),
            "cut" => Ok(SlideTransitionGroup::Cut(OptionalBlackTransition::from_xml_element(
                xml_node,
            )?)),
            "diamond" => Ok(SlideTransitionGroup::Diamond),
            "fade" => Ok(SlideTransitionGroup::Fade(OptionalBlackTransition::from_xml_element(
                xml_node,
            )?)),
            "newsflash" => Ok(SlideTransitionGroup::Newsflash),
            "plus" => Ok(SlideTransitionGroup::Plus),
            "pull" => Ok(SlideTransitionGroup::Pull(EightDirectionTransition::from_xml_element(
                xml_node,
            )?)),
            "push" => Ok(SlideTransitionGroup::Push(SideDirectionTransition::from_xml_element(
                xml_node,
            )?)),
            "random" => Ok(SlideTransitionGroup::Random),
            "randomBar" => Ok(SlideTransitionGroup::RandomBar(
                OrientationTransition::from_xml_element(xml_node)?,
            )),
            "split" => Ok(SlideTransitionGroup::Split(SplitTransition::from_xml_element(
                xml_node,
            )?)),
            "strips" => Ok(SlideTransitionGroup::Strips(
                CornerDirectionTransition::from_xml_element(xml_node)?,
            )),
            "wedge" => Ok(SlideTransitionGroup::Wedge),
            "wheel" => Ok(SlideTransitionGroup::Wheel(WheelTransition::from_xml_element(
                xml_node,
            )?)),
            "wipe" => Ok(SlideTransitionGroup::Wipe(SideDirectionTransition::from_xml_element(
                xml_node,
            )?)),
            "zoom" => Ok(SlideTransitionGroup::Zoom(InOutTransition::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "EG_SlideTransition",
            ))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TransitionStartSoundAction {
    /// This attribute specifies if the sound loops until the next sound event occurs in slideshow.
    /// 
    /// Defaults to false
    pub is_looping: Option<bool>,
    /// This element specifies the audio information to play during a slide transition.
    /// 
    /// # Xml example
    /// 
    /// Consider a slide transition with an audio effect. The <snd> element should be used as follows:
    /// ```xml
    /// <p:transition>
    ///   <p:sndAc>
    ///     <p:stSnd>
    ///       <p:snd r:embed="rId2" />
    ///     </p:stSnd>
    ///   </p:sndAc>
    /// </p:transition>
    /// ```
    pub sound_file: msoffice_shared::drawingml::EmbeddedWAVAudioFile,
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
        let sound_file = msoffice_shared::drawingml::EmbeddedWAVAudioFile::from_xml_element(sound_file_node)?;

        Ok(Self { is_looping, sound_file })
    }
}

#[derive(Debug, Clone)]
pub enum TransitionSoundAction {
    /// This element describes the sound that starts playing during a slide transition.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:sndAc>
    ///     <p:stSnd>
    ///       <p:snd r:embed="rId2"/>
    ///     </p:stSnd>
    ///   </p:sndAc>
    /// </p:transition>
    /// ```
    StartSound(TransitionStartSoundAction),
    /// This element stops all previous sounds during a slide transition.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:transition>
    ///   <p:sndAc>
    ///     <p:endSnd/>
    ///   </p:sndAc>
    /// </p:transition>
    /// ```
    EndSound,
}

impl TransitionSoundAction {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "stSnd" | "endSnd" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "stSnd" => Ok(TransitionSoundAction::StartSound(
                TransitionStartSoundAction::from_xml_element(xml_node)?,
            )),
            "endSnd" => Ok(TransitionSoundAction::EndSound),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "CT_TransitionSoundAction",
            ))),
        }
    }
}

#[derive(Default, Debug, Clone)]
pub struct SlideTransition {
    /// Specifies the transition speed that is to be used when transitioning from the current slide
    /// to the next.
    /// 
    /// Defaults to TransitionSpeed::Fast
    pub speed: Option<TransitionSpeed>,
    /// Specifies whether a mouse click advances the slide or not. If this attribute is not specified
    /// then a value of true is assumed.
    /// 
    /// Defaults to true
    pub advance_on_click: Option<bool>,
    /// Specifies the time, in milliseconds, after which the transition should start. This setting can
    /// be used in conjunction with the advance_on_click attribute. If this attribute is not specified then it
    /// is assumed that no auto-advance occurs.
    pub advance_on_time: Option<u32>,
    pub transition_type: Option<SlideTransitionGroup>,
    /// This element describes a sound action for slide transition. This element specifies that the start of the slide
    /// transition is accompanied by the playback of an audio file; the actual audio file used is specified by the snd
    /// element
    /// 
    /// # Xml example
    /// 
    /// Consider a slide transition with a sound effect. The <sndAc> element should be used as follows:
    /// ```xml
    /// <p:transition>
    ///   <p:sndAc>
    ///     <p:stSnd>
    ///       <p:snd r:embed="rId2"/>
    ///     </p:stSnd>
    ///   </p:sndAc>
    /// </p:transition>
    /// ```
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
    /// This element specifies a list of time node elements used in an animation sequence.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:timing>
    ///   <p:tnLst>
    ///     <p:par> … </p:par>
    ///   </p:tnLst>
    /// </p:timing>
    /// ```
    pub time_node_list: Option<Vec<TimeNodeGroup>>,
    /// This element specifies the list of graphic elements to build. This refers to how the different sub-shapes or sub-
    /// components of a object are displayed. The different objects that can have build properties are text, diagrams,
    /// and charts.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:bldLst>
    ///   <p:bldGraphic spid="1" grpId="0">
    ///     <p:bldSub>
    ///       <a:bldChart bld="category"/>
    ///     </p:bldSub>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    pub build_list: Option<Vec<Build>>,
}

impl SlideTiming {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tnLst" => {
                    let mut vec = Vec::new();
                    for time_node in &child_node.child_nodes {
                        vec.push(TimeNodeGroup::from_xml_element(time_node)?);
                    }

                    if vec.is_empty() {
                        return Err(Box::new(MissingChildNodeError::new(
                            child_node.name.clone(),
                            "tn",
                        )));
                    }

                    instance.time_node_list = Some(vec);
                }
                "bldLst" => {
                    let mut vec = Vec::new();
                    for build_node in &child_node.child_nodes {
                        vec.push(Build::from_xml_element(build_node)?);
                    }

                    if vec.is_empty() {
                        return Err(Box::new(MissingChildNodeError::new(
                            child_node.name.clone(),
                            "bld",
                        )))
                    }
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone)]
pub enum Build {
    /// This element specifies how to build paragraph level properties.
    /// 
    /// # Xml example
    /// 
    /// Consider having animation applied only to 1st level paragraphs. The <bldP> element should be used
    /// as follows:
    /// 
    /// ```xml
    /// <p:bldLst>
    ///   <p:bldP spid="3" grpId="0" build="p"/>
    /// </p:bldLst>
    /// ```
    Paragraph(Box<TLBuildParagraph>),
    /// This element specifies how to build the animation for a diagram.
    /// 
    /// # Xml example
    /// 
    /// Consider the following example where a chart is specified to be animated by category rather than as
    /// one entity. Thus, the bldChart element should be used as follows:
    /// 
    /// ```xml
    /// <p:bdldLst>
    ///   <p:bldGraphic spid="4" grpId="0">
    ///     <p:bldSub>
    ///       <a:bldChart bld="category"/>
    ///     </p:bldSub>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    Diagram(Box<TLBuildDiagram>),
    /// This element describes animation an a embedded Chart.
    /// 
    /// # Xml example
    /// 
    /// Consider displaying animation on a embedded graphical chart. The <bldOleChart>element should be
    /// use as follows:
    /// 
    /// ```xml
    /// <p:bldLst>
    ///   <p:bldOleChart spid="1025" grpId="0"/>
    /// </p:bldLst>
    /// ```
    OleChart(Box<TLOleBuildChart>),
    /// This element specifies how to build a graphical element.
    /// 
    /// # Xml example
    /// 
    /// Consider having a chart graphical element appear as a whole as opposed to by a category. The
    /// <bldGraphic> element should be used as follows:
    /// 
    /// ```xml
    /// <p:bldLdst>
    ///   <p:bldGraphic spid="3" grpId="0">
    ///     <p:bldSub>
    ///       <a:bldChart bld="category"/>
    ///     </p:bldSub>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    Graphic(Box<TLGraphicalObjectBuild>),
}

impl Build {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "bldP" | "bldDgm" | "bldOleChart" | "bldGraphic" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bldP" => Ok(Build::Paragraph(Box::new(TLBuildParagraph::from_xml_element(
                xml_node,
            )?))),
            "bldDgm" => Ok(Build::Diagram(Box::new(TLBuildDiagram::from_xml_element(xml_node)?))),
            "bldOleChart" => Ok(Build::OleChart(Box::new(TLOleBuildChart::from_xml_element(xml_node)?))),
            "bldGraphic" => Ok(Build::Graphic(Box::new(TLGraphicalObjectBuild::from_xml_element(
                xml_node,
            )?))),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "CT_BuildList",
            ))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLBuildParagraph {
    pub build_common: TLBuildCommonAttributes,
    /// This attribute describe the build types.
    /// 
    /// Defaults to TLParaBuildType::Whole
    pub build_type: Option<TLParaBuildType>,
    /// This attribute describes the build level for the paragraph. It is only supported in
    /// paragraph type builds i.e the build attribute shall also be set to "byParagraph" for this
    /// attribute to apply.
    /// 
    /// Defaults to 1
    pub build_level: Option<u32>,
    /// This attribute indicates whether to animate the background of the shape associated with
    /// the text.
    /// 
    /// Defaults to false
    pub animate_bg: Option<bool>,
    /// This attribute indicates whether to automatically update the "animateBg" setting to true
    /// when the shape associated with the text has a fill or line.
    /// 
    /// Defaults to true
    pub auto_update_anim_bg: Option<bool>,
    /// This attribute is only supported in paragraph type builds. This specifies the direction of
    /// the build relative to the order of the elements in the container. When this is set to "true",
    /// the animations for the paragraphs are persisted in reverse order to the order of the
    /// paragraphs themselves such that the last paragraph animates first.
    /// 
    /// Defaults to false
    pub reverse: Option<bool>,               // false
    /// This attribute specifies time after which to automatically advance the build to the next
    /// step.
    /// 
    /// Defaults to TLTime::Indefinite
    pub auto_advance_time: Option<TLTime>,
    pub template_list: Option<Vec<TLTemplate>>,      // size: 0-9
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

        let template_list = match xml_node.child_nodes.get(0) {
            Some(child_node) => {
                let mut vec = Vec::new();
                for template_node in &child_node.child_nodes {
                    vec.push(TLTemplate::from_xml_element(template_node)?);
                }
                Some(vec)
            },
            None => None,
        };

        let shape_id = shape_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
        let group_id = group_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "grpId"))?;

        Ok(Self {
            build_common: TLBuildCommonAttributes {
                shape_id,
                group_id,
                ui_expand,
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
    /// This attribute describes the X coordinate.
    pub x: msoffice_shared::drawingml::Percentage,
    /// This attribute describes the Y coordinate.
    pub y: msoffice_shared::drawingml::Percentage,
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
    type Err = msoffice_shared::error::ParseEnumError;

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
    /// This attribute specifies the shape to which the build applies.
    pub shape_id: msoffice_shared::drawingml::DrawingElementId,
    /// This attribute ties effects persisted in the animation to the build information. The
    /// attribute is used by the editor when changes to the build information are made.
    /// GroupIDs are unique for a given shape. They are not guaranteed to be unique IDs across
    /// all shapes on a slide.
    pub group_id: u32,
    /// This attribute describes the view option indicating if the build should be displayed
    /// expanded.
    /// 
    /// Defaults to false
    pub ui_expand: Option<bool>,
}

#[derive(Debug, Clone)]
pub struct TLBuildDiagram {
    pub build_common: TLBuildCommonAttributes,
    /// This attribute describes how the diagram is built. The animation animates the sub-
    /// elements in the container in the particular order defined by this attribute.
    /// 
    /// Defaults to TLDiagramBuildType::Whole
    pub build_type: Option<TLDiagramBuildType>,
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
    /// This attribute describes how the diagram is built. The animation animates the sub-
    /// elements in the container in the particular order defined by this attribute.
    /// 
    /// Defaults to TLOleChartBuildType::AllAtOnce
    pub build_type: Option<TLOleChartBuildType>,
    /// This attribute describes whether to animate the background of the shape.
    /// 
    /// Defaults to true
    pub animate_bg: Option<bool>,
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
            None => {
                return Err(Box::new(MissingChildNodeError::new(
                    xml_node.name.clone(),
                    "TLGraphicalObjectBuildChoice",
                )))
            }
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
    /// This element specifies in the build list to build the entire graphical object as one entity.
    /// 
    /// # Xml example
    /// 
    /// Consider having a graph appear as on entity as opposed to by category. The <bldAsOne> element
    /// should be used as follows:
    /// 
    /// ```xml
    /// <p:bldLst>
    ///   <p:bldGraphic spid="4" grpId="0">
    ///     <p:bldAsOne/>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    BuildAsOne,
    /// This element specifies the animation properties of a graphical object's sub-elements.
    /// 
    /// # Xml example
    /// 
    /// Consider applying animation to a graphical element consisting of a diagram. The <bldSub> element
    /// should be used as follows:
    /// ```xml
    /// <p:bldLst>
    ///   <p:bldGraphic spid="5" grpId="0">
    ///     <p:bldSub>
    ///       <a:bldDgm bld="one"/>
    ///     </p:bldSub>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    BuildSubElements(msoffice_shared::drawingml::AnimationGraphicalObjectBuildProperties),
}

impl TLGraphicalObjectBuildChoice {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
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
                msoffice_shared::drawingml::AnimationGraphicalObjectBuildProperties::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TLGraphicalObjectBuildChoice",
            ))),
        }
    }
}

#[derive(Debug, Clone)]
pub enum TimeNodeGroup {
    /// This element describes the Parallel time node which can be activated along with other parallel time node
    /// containers.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:timing>
    ///   <p:tnLst>
    ///     <p:par>
    ///       <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
    ///         <p:childTnLst>
    ///           <p:seq concurrent="1" nextAc="seek">
    ///             …
    ///           </p:seq>
    ///         </p:childTnLst>
    ///       </p:cTn>
    ///     </p:par>
    ///   </p:tnLst>
    /// </p:timing>
    /// ```
    Parallel(Box<TLCommonTimeNodeData>),
    /// This element describes the Sequence time node and it can only be activated when the one before it finishes.
    /// 
    /// # Xml example
    /// 
    /// For example, suppose we have a simple animation with a blind entrance.
    /// ```xml
    /// <p:timing>
    ///   <p:tnLst>
    ///     <p:par>
    ///       <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
    ///         <p:childTnLst>
    ///           <p:seq concurrent="1" nextAc="seek">
    ///           …
    ///           </p:seq>
    ///         </p:childTnLst>
    ///       </p:cTn>
    ///     </p:par>
    ///   </p:tnLst>
    /// </p:timing>
    /// ```
    Sequence(Box<TLTimeNodeSequence>),
    /// This element describes the Exclusive time node. This time node is used to pause all other timelines when it is
    /// activated.
    Exclusive(Box<TLCommonTimeNodeData>),
    /// This element is a generic animation element that requires little or no semantic understanding of the attribute
    /// being animated. It can animate text within a shape or even the shape itself.
    /// 
    /// # Xml example
    /// 
    /// Consider trying to emphasize text within a shape by changing the size of its font by 150%. The
    /// <anim> element should be used as follows:
    /// 
    /// ```xml
    /// <p:anim to="1.5" calcmode="lin" valueType="num">
    ///   <p:cBhvr override="childStyle">
    ///     <p:cTn id="1" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="1">
    ///         <p:txEl>
    ///           <p:charRg st="1" end="4"/>
    ///         </p:txEl>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>style.fontSize</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    /// </p:anim>
    /// ```
    Animate(Box<TLAnimateBehavior>),
    /// This animation element is responsible for animating the color of an object.
    /// 
    /// # Xml example
    /// 
    /// Consider trying to emphasize a shape by changing its fill color to scheme color accent2. The
    /// <animClr> element should be used as follows:
    /// ```xml
    /// <p:animClr clrSpc="rgb">
    ///   <p:cBhvr>
    ///     <p:cTn id="1" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="1"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>fillcolor</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:to>
    ///     <a:schemeClr val="accent2"/>
    ///   </p:to>
    /// </p:animClr>
    /// ```
    AnimateColor(Box<TLAnimateColorBehavior>),
    /// This animation behavior provides the ability to do image transform/filter effects on elements. Some visual
    /// effects are dynamic in nature and have a progress that animates from 0 to 1 over a period of time to do visual
    /// transitions between hidden and visible states. Other filters are static and apply a effects like a blur or drop-
    /// shadow which aren't inherently time-based.
    /// 
    /// # Xml example
    /// 
    /// Consider trying to emphasize a shape by creating an entrance animation using a "blinds" motion.
    /// ```xml
    /// <p:animEffect transition="in" filter="blinds(horizontal)">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    AnimateEffect(Box<TLAnimateEffectBehavior>),
    /// Animate motion provides an abstracted way to move positioned elements. It provides the ability to specify
    /// from/to/by motion as well as to use more detailed path descriptions for motion over polylines or bezier curves.
    /// 
    /// # Xml example
    /// 
    /// Consider animating a shape from its original position to the right.. The <animMotion> element should
    /// be used as follows:
    /// 
    /// ```xml
    /// <p:animMotion origin="layout" path="M 0 0 L 0.25 0 E" pathEditMode="relative">
    ///   <p:cBhvr>
    ///     <p:cTn id="1" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="1"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>ppt_x</p:attrName>
    ///       <p:attrName>ppt_y</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    /// </p:animMotion>
    /// ```
    AnimateMotion(Box<TLAnimateMotionBehavior>),
    /// This animation element is responsible for animating the rotation of an object. Rotation values set in the "by",
    /// "to, and "from" attributes are specified in degrees measured to a 60,000th, i.e 1 degree is 60,000. Rotation
    /// values can be larger than 360°.
    /// 
    /// The sign of the rotation angle specifies the direction for rotation. A negative rotation specifies that the rotation
    /// should appear in the host to go counter-clockwise".
    /// 
    /// # Xml example
    /// 
    /// Consider trying to emphasize a shape by rotating it 360 degrees clockwise. The <animRot> element
    /// should be used as follows:
    /// 
    /// ```xml
    /// <p:animRot by="21600000">
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="5"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>r</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    /// </p:animRot>
    /// ```
    AnimateRotation(Box<TLAnimateRotationBehavior>),
    /// This animation element is responsible for animating the scale of an object. When animating the scale, the
    /// element shall scale around the reference point of the element and the positioning system used should be
    /// consistent with the one used for motion paths. When animating the width and height of an element, all of the
    /// width/height animation values are calculated first then the scale animations are applied on top of that. So for
    /// example, an animation from 0 to 100 of the width with a concurrent scale from 100% to 200% would result in
    /// the element appearing to scale from 0 to 200.
    /// 
    /// # Xml example
    /// 
    /// Consider trying to emphasize a shape by scaling it larger by 150%. The <animScale> element should
    /// be used as follows:
    /// 
    /// ```xml
    /// <p:childTnLst>
    ///   <p:animScale>
    ///     <p:cBhvr>
    ///       <p:cTn id="6" dur="2000" fill="hold"/>
    ///       <p:tgtEl>
    ///         <p:spTgt spid="5"/>
    ///       </p:tgtEl>
    ///     </p:cBhvr>
    ///     <p:by x="150000" y="150000"/>
    ///   </p:animScale>
    /// </p:childTnLst>
    /// ```
    AnimateScale(Box<TLAnimateScaleBehavior>),
    /// This element describes the several non-durational commands that can be executed within a timeline. This can
    /// be used to send events, call functions on elements, and send verbs to embedded objects. For example “Object
    /// Action” effects for Embedded objects and Media commands for sounds/movies such as "PlayFrom(0.0)" and
    /// "togglePause".
    Command(Box<TLCommandBehavior>),
    /// This element allows the setting of a particular property value to a fixed value while the behavior is active and
    /// restores the value when the behavior is reset or turned off.
    /// 
    /// # Xml example
    /// 
    /// For example, suppose we want to set certain properties during an animation effect. The <set>
    /// element should be used as follows:
    /// ```xml
    /// <p:childTnLst>
    ///   <p:set>
    ///     <p:cBhvr>
    ///       <p:cTn id="6" dur="1" fill="hold"> … </p:cTn>
    ///       <p:tgtEl>
    ///         <p:spTgt spid="4"/>
    ///       </p:tgtEl>
    ///       <p:attrNameLst>
    ///         <p:attrName>style.visibility</p:attrName>
    ///       </p:attrNameLst>
    ///     </p:cBhvr>
    ///     <p:to>
    ///       <p:strVal val="visible"/>
    ///     </p:to>
    ///   </p:set>
    ///   <p:animEffect transition="in" filter="blinds(horizontal)">
    ///     …
    ///   </p:animEffect>
    /// </p:childTnLst>
    /// ```
    Set(Box<TLSetBehavior>),
    /// This element is used to include audio during an animation. This element specifies that this node within the
    /// animation tree triggers the playback of an audio file; the actual audio file used is specified by the sndTgt
    /// element (§19.5.70).
    /// 
    /// # Xml example
    /// 
    /// Consider adding applause sound to an animation sequence. The audio element is used as follows:
    /// 
    /// ```xml
    /// <p:cTn ...>
    ///   <p:stCondLst>...</p:stCondLst>
    ///   <p:childTnLst>...</p:childTnLst>
    ///   <p:subTnLst>
    ///     <p:audio>
    ///       <p:cMediaNode vol="50%">...
    ///         <p:tgtEl>
    ///           <p:sndTgt r:embed="rId2" />
    ///         </p:tgtEl>
    ///       </p:cMediaNode>
    ///     </p:audio>
    ///   </p:subTnLst>
    /// </p:cTn>
    /// ```
    /// 
    /// The audio element specifies the location of the audio playback within the animation; its child sndTgt element
    /// specifies that the audio to be played is the target of the relationship with ID rId2.
    Audio(Box<TLMediaNodeAudio>),
    /// This element specifies video information in an animation sequence. This element specifies that this node within
    /// the animation tree triggers the playback of a video file; the actual video file used is specified by the videoFile
    /// element
    /// 
    /// # Xml example
    /// 
    /// Consider a slide with an animated video content. The <video> element is used as follows:
    /// ```xml
    /// <p:cSld>
    ///   <p:spTree>
    ///     <p:pic>
    ///       <p:nvPicPr>
    ///       <p:cNvPr id="4"/>
    ///       …
    ///       <p:nvPr>
    ///         <a:videoFile r:link="rId1" contentType="video/ogg"/>
    ///       </p:nvPr>
    ///     </p:nvPicPr>
    ///     …
    ///     </p:pic>
    ///   </p:spTree>
    /// </p:cSld>
    /// …
    /// <p:childTnLst>
    ///   <p:seq concurrent="1" nextAc="seek">
    ///     …
    ///   </p:seq>
    ///   <p:video>
    ///     <p:cMediaNode>
    ///       …
    ///       <p:tgtEl>
    ///         <p:spTgt spid="4"/>
    ///       </p:tgtEl>
    ///     </p:cMediaNode>
    ///   </p:video>
    /// </p:childTnLst>
    /// ```
    /// 
    /// The video element specifies the location of the video playback within the animation sequence; its child spTgt
    /// element specifies that the shape which contains the video to be played has a shape ID of 4. If we look at the
    /// shape with that ID value, its child videoFile element references an external video file of content type video/ogg
    /// located at the target of the relationship with ID rId1
    Video(Box<TLMediaNodeVideo>),
}

impl TimeNodeGroup {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
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
                TLAnimateEffectBehavior::from_xml_element(xml_node)?,
            ))),
            "animMotion" => Ok(TimeNodeGroup::AnimateMotion(Box::new(
                TLAnimateMotionBehavior::from_xml_element(xml_node)?,
            ))),
            "animRot" => Ok(TimeNodeGroup::AnimateRotation(Box::new(
                TLAnimateRotationBehavior::from_xml_element(xml_node)?,
            ))),
            "animScale" => Ok(TimeNodeGroup::AnimateScale(Box::new(
                TLAnimateScaleBehavior::from_xml_element(xml_node)?,
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
    /// This attribute specifies if concurrency is enabled or disabled. By default this attribute has
    /// a value of "disabled". When the value is set to "enabled", the previous element is left
    /// enabled when advancing to the next element in a sequence instead of being ended. This
    /// is only relevant for advancing via the next condition element being triggered. The only
    /// other way to advance to the next element would be to have the current element end,
    /// which implies it is no longer concurrent.
    pub concurrent: Option<bool>,
    /// This attribute specifies what to do when going backwards in a sequence. By default it is
    /// set to TLPreviousActionType::None and nothing special is done. When the value is TLPreviousActionType::SkipTimed,
    /// the sequence continues to go backwards until it reaches a sequence element that was defined to begin
    /// only on the next condition element.
    pub prev_action_type: Option<TLPreviousActionType>,
    /// This attribute specifies what to do when going forward in sequence. By default this
    /// attribute has a value of TLNextActionType::None. When this is set to seek it seeks the element to a natural
    /// end time (not necessarily the actual end time).
    /// 
    /// The natural end position is defined as the latest non-infinite end time of the children. If a
    /// child loops forever, the end of its first loop is used as its "end time" for the purposes of
    /// this calculation.
    /// 
    /// Some container elements can have infinite durations due to an infinite-duration child
    /// element. The engine needs to recurse down through all infinite duration containers to
    /// calculate their natural duration in case a child might have non-infinite duration within it
    /// that needs to be taken into account.
    pub next_action_type: Option<TLNextActionType>,
    pub common_time_node_data: Box<TLCommonTimeNodeData>,
    /// This element describes a list of conditions that shall be met in order to go backwards in an animation sequence.
    /// 
    /// # Xml example
    /// 
    /// Consider trying to emphasize text within a shape by changing the size of its font.
    /// ```xml
    /// <p:seq concurrent="1" nextAc="seek">
    ///   <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
    ///   </p:cTn>
    ///   <p:prevCondLst>
    ///     <p:cond evt="onPrev" delay="0">
    ///       <p:tgtEl>
    ///         <p:sldTgt/>
    ///       </p:tgtEl>
    ///     </p:cond>
    ///   </p:prevCondLst>
    ///   <p:nextCondLst>
    ///   </p:nextCondLst>
    /// </p:seq>
    /// ```
    pub prev_condition_list: Vec<TLTimeCondition>,
    /// This element describes a list of conditions that shall be met to advance to the next animation sequence.
    /// 
    /// # Xml example
    /// 
    /// Consider a shape with a text emphasis changing the size of its font.
    /// ```xml
    /// <p:seq concurrent="1" nextAc="seek">
    ///   <p:cTn id="2" dur="indefinite" nodeType="mainSeq"> … </p:cTn>
    ///   <p:prevCondLst> … </p:prevCondLst>
    ///   <p:nextCondLst>
    ///     <p:cond evt="onNext" delay="0">
    ///       <p:tgtEl>
    ///         <p:sldTgt/>
    ///       </p:tgtEl>
    ///     </p:cond>
    ///   </p:nextCondLst>
    /// </p:seq>
    /// ```
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
    /// This attribute specifies a relative offset value for the animation with respect to its
    /// position before the start of the animation.
    pub by: Option<String>,
    /// This attribute specifies the starting value of the animation.
    pub from: Option<String>,
    /// This attribute specifies the ending value for the animation as a percentage.
    pub to: Option<String>,
    /// This attribute specifies the interpolation mode for the animation.
    pub calc_mode: Option<TLAnimateBehaviorCalcMode>,
    /// This attribute specifies the type of property value.
    pub value_type: Option<TLAnimateBehaviorValueType>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// This element specifies a list of time animated value elements.
    /// 
    /// ```xml
    /// <p:anim calcmode="lin" valueType="num">
    ///   <p:cBhvr additive="base"> … </p:cBhvr>
    ///   <p:tavLst>
    ///     <p:tav tm="0%">
    ///       <p:val>
    ///         <p:strVal val="1+#ppt_h/2"/>
    ///       </p:val>
    ///     </p:tav>
    ///     <p:tav tm="100000">
    ///       <p:val>
    ///         <p:strVal val="#ppt_y"/>
    ///       </p:val>
    ///     </p:tav>
    ///   </p:tavLst>
    /// </p:anim>
    /// ```
    pub time_animate_value_list: Option<Vec<TLTimeAnimateValue>>,
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
        let mut time_animate_value_list = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "tavLst" => {
                    let mut vec = Vec::new();
                    for tav_node in &child_node.child_nodes {
                        vec.push(TLTimeAnimateValue::from_xml_element(tav_node)?);
                    }

                    time_animate_value_list = Some(vec);
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
    /// This attribute specifies the color space in which to interpolate the animation. Values for
    /// example can be HSL & RGB.
    /// 
    /// The values for from/to/by/etc. can still be specified in any supported color format
    /// without affecting the color space within which the animation happens.
    /// 
    /// The RGB color space is best used for doing animations between two different colors since
    /// it doesn't require going through any other hues between the two colors specified. The
    /// HSL space is useful for animating through a rainbow of colors or for modifying just the
    /// saturation by 30% for example.
    pub color_space: Option<TLAnimateColorSpace>,
    /// This attribute specifies which direction to cycle the hue around the color wheel. Values
    /// are clockwise or counter clockwise. Default is clockwise.
    pub direction: Option<TLAnimateColorDirection>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// This element describes the relative offset value for the color animation.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animClr clrSpc="hsl">
    ///   <p:cBhvr>
    ///     <p:cTn id="8" dur="500" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>stroke.color</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:by>
    ///     <p:hsl h="0" s="0" l="0"/>
    ///   </p:by>
    /// </p:animClr>
    /// ```
    pub by: Option<TLByAnimateColorTransform>,
    /// This element is used to specify the starting color of the target element.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animClr clrSpc="rgb" dir="cw">
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl> … </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>fillcolor</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:from>
    ///     <a:schemeClr val="accent3"/>
    ///   </p:from>
    ///   <p:to>
    ///     <a:schemeClr val="accent2"/>
    ///   </p:to>
    /// </p:animClr>
    /// ```
    pub from: Option<msoffice_shared::drawingml::Color>,
    /// This element specifies the resulting color for the animation color change.
    /// 
    /// # Xml example
    /// 
    /// Consider emphasize a shape by changing its fill color from blue to red. The <to> element should be
    /// used as follows:
    /// ```xml
    /// <p:childTnLst>
    ///   <p:animClr clrSpc="rgb">
    ///     <p:cBhvr> … </p:cBhvr>
    ///     <p:to>
    ///       <a:schemeClr val="accent2"/>
    ///     </p:to>
    ///   </p:animClr>
    /// </p:childTnLst>
    /// ```
    pub to: Option<msoffice_shared::drawingml::Color>,
}

impl TLAnimateColorBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        use msoffice_shared::drawingml::Color;

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
                    let by_node = child_node.child_nodes.get(0).ok_or_else(|| {
                        MissingChildNodeError::new(child_node.name.clone(), "TLByAnimateColorTransform")
                    })?;
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
    /// This attribute specifies whether to transition the element in or out or treat it as a static
    /// filter. The values are "None", "In" and "Out", and the default value is "In".
    /// 
    /// When a value of "In" is specified, the element is not visible at the start of the animation
    /// and is completely visible be the end of the duration. When "Out" is specified, the element
    /// is visible at the start and not visible at the end of the effect. This visibility is in addition to
    /// the effect of setting CSS visibility or display attributes.
    pub transition: Option<TLAnimateEffectTransition>,
    /// This attribute specifies the animation types and subtypes to be used for the effect.
    /// Multiple animations are allowed to be listed so that in the event that a superseding
    /// animation (leftmost) cannot be rendered, a fallback animation is available. That is, the
    /// rendering application parses the list from left to right until a supported animation is
    /// found.
    /// 
    /// The syntax used for the filter attribute value is as follows: "type(subtype);type(subtype)".
    /// Subtype can be a string value such as "fromLeft" or a numerical value depending on the
    /// type specified.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animEffect transition="in" filter="blinds(horizontal);blinds(vertical)">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:spTgtspid="5"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    /// 
    /// There are two animation filters shown in this example. The first is the blinds (horizontal),
    /// which the rendering application is to use as the primary animation effect. If, however,
    /// the rendering application does not support this animation, the blinds (vertical) animation
    /// is used. In this example there are only two animation filters listed, a primary and a
    /// fallback, but it is possible to list multiple fallback filters using the syntax defined above.
    pub filter: Option<String>,
    /// This attribute specifies a list of properties that coincide with the effect specified.
    /// Although there are many animation types allowed, this attribute allows the setting of
    /// specific property settings in order to describe an even wider variety of animation types.
    /// 
    /// The syntax used for the prLst attribute value is as follows: “name:value;name:value”.
    /// When multiple animation types are listed in the filter attribute, the rendering application
    /// attempts to apply each property value even though some might not apply to it.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animEffect filter="image" prLst="opacity: 0.5">
    ///   <p:cBhvr rctx="IE">
    ///     <p:cTn id="7" dur="indefinite"/>
    ///     <p:tgtEl>
    ///       <p:spTgtspid="3"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    /// 
    /// The animation filter specified is an image filter type that has a specific property called
    /// opacity set to a value of 0.5.
    pub property_list: Option<String>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// This element defines the progression of an animation. The default for the way animation progress happens
    /// through an animEffect is a linear ramp from 0 to 1, starting at the effect’s begin time & ending at the effect’s
    /// end time. When you specify a value for the progress attribute, you are overriding this default behaviour. The
    /// value between 0 and 1 represents a percentage through the effect, where 0 is 0% and 1 is 100%.
    /// 
    /// Each animEffect is in fact an object-based transition. These transitions can be specified as “In” (where the object
    /// is not visible at 0% and becomes completely visible at 100%) or “Out” (where the object is visible at 0% and
    /// becomes completely invisible at 100%). You would set the progress attribute if you want to use the animEffect
    /// as a “static” effect, where the transition properties do not actually change over time. As an alternative to using
    /// the progress attribute, you can use the tmFilter (time filter), which is a base attribute of any effect/timenode, to
    /// specify the way that progress through an effect should be performed dynamically.
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
                }
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
    /// Specifies what the origin of the motion path is relative to such as the layout of the slide,
    /// or the parent.
    pub origin: Option<TLAnimateMotionBehaviorOrigin>,
    /// Specifies the path primitive followed by coordinates for the animation motion. The
    /// allowed values that are understood within a path are as follows:
    /// 
    /// M = move to, L = line to, C = curve to, Z=close loop, E=end
    /// UPPERCASE = absolute coords, lowercase = relative coords
    /// Thus total allowed set = {M,L,C,Z,E,m,l,c,z,e)
    /// 
    /// # Example
    /// 
    /// The following string is a sample path.
    /// path: “M 0 0 L 1 1 c 1 2 3 4 4 4 Z”
    pub path: Option<String>,
    /// This attribute specifies how the motion path moves when the target element is moved.
    pub path_edit_mode: Option<TLAnimateMotionPathEditMode>,
    /// The attribute describes the relative angle of the motion path.
    pub rotate_angle: Option<msoffice_shared::drawingml::Angle>,
    /// This attribute describes the point type of the points in the path attribute. The allowed
    /// values that are understood for the ptsTypes attribute are as follows:
    /// 
    /// A = Auto, F = Corner, T = Straight, S = Smooth
    /// UPPERCASE = Straight Line follows point, lowercase = curve follows point.
    /// Thus, the total allowed set = {A,F,T,S,a,f,t,s}
    pub points_types: Option<String>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// This element describes the relative offset value for the animation.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animScale>
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    ///   <p:by x="150.000%" y="150.000%"/>
    /// </p:animScale>
    /// ```
    pub by: Option<TLPoint>,
    pub from: Option<TLPoint>,
    /// This element specifies the target location for an animation motion or animation scale effect
    /// 
    /// # Xml example
    /// 
    /// Consider an animation with a "light speed" entrance effect.
    /// ```xml
    /// <p:animScale>
    ///   <p:cBhvr>
    ///     <p:cTn id="9" dur="200" decel="10.5%" autoRev="1" fill="hold">
    ///       <p:stCondLst>
    ///         <p:cond delay="600"/>
    ///       </p:stCondLst>
    ///     </p:cTn>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    ///   <p:from x="100%" y="100%"/>
    ///   <p:to x="80%" y="100%"/>
    /// </p:animScale>
    /// ```
    pub to: Option<TLPoint>,
    /// This element describes the center of the rotation used to rotate a motion path by X angle.
    /// 
    /// # Xml example
    /// 
    /// For example, suppose we have a simple animation with a checkerbox text entrance.
    /// ```xml
    /// <p:animMotion origin="layout" path="M 0 0 L 0.25 0.33333 E" pathEditMode="relative" rAng="0" ptsTypes="">
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>ppt_x</p:attrName>
    ///       <p:attrName>ppt_y</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:rCtr x="56.7%" y="83.4%"/>
    /// </p:animMotion>
    /// ```
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
    /// This attribute describes the relative offset value for the animation.
    pub by: Option<msoffice_shared::drawingml::Angle>,
    /// This attribute describes the starting value for the animation.
    pub from: Option<msoffice_shared::drawingml::Angle>,
    /// This attribute describes the ending value for the animation.
    pub to: Option<msoffice_shared::drawingml::Angle>,
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
    /// This attribute specifies whether to zoom the contents of an object when doing a scaling
    /// animation.
    pub zoom_contents: Option<bool>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// This element describes the relative offset value for the animation.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animScale>
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    ///   <p:by x="150.000%" y="150.000%"/>
    /// </p:animScale>
    /// ```
    pub by: Option<TLPoint>,
    /// This element specifies an x/y co-ordinate to start the animation from.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animScale>
    ///   <p:cBhvr>
    ///     <p:cTn> … </p:cTn>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    ///   <p:from x="100%" y="100%"/>
    ///   <p:to x="80%" y="100%"/>
    /// </p:animScale>
    /// ```
    pub from: Option<TLPoint>,
    /// This element specifies the target location for an animation motion or animation scale effect
    /// 
    /// # Xml example
    /// 
    /// Consider an animation with a "light speed" entrance effect.
    /// ```xml
    /// <p:animScale>
    ///   <p:cBhvr>
    ///     <p:cTn id="9" dur="200" decel="10.5%" autoRev="1" fill="hold">
    ///       <p:stCondLst>
    ///         <p:cond delay="600"/>
    ///       </p:stCondLst>
    ///     </p:cTn>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    ///   <p:from x="100%" y="100%"/>
    ///   <p:to x="80%" y="100%"/>
    /// </p:animScale>
    /// ```
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
    /// This attribute specifies the kind of command that is issued by the rendering application to
    /// the appropriate target application or object.
    /// 
    /// There are three possible values, call, evt, and verb. A call command type is used to
    /// specify the class of commands that can then be issued.
    /// 
    /// * Call commands: This command type is used to call methods on the object specified (play(), pause(), etc.)
    /// 
    /// * Event commands: This command type is used to set an event for the object at this point in the timeline
    ///                   (onstopaudio, etc.)
    /// 
    /// * Verb Commands: This command type is used to set verbs for the object to occur at this point in the timeline
    ///                  (0, 1, etc.)
    pub command_type: Option<TLCommandType>,
    /// This attribute defines the actual command to be issued. Depending on the command
    /// specified, the actual command can be made to invoke a wide range of actions on the
    /// linked or embedded object
    /// 
    /// Reserved Values (when command_type == TLCommandType::Call):
    /// * play: play corresponding media
    /// * playFrom(s): play corresponding media starting from s, where s is the number of
    ///                seconds from the beginning of the clip
    /// * pause: pause corresponding media
    /// * resume: resume play of corresponding media
    /// * stop: stop play of corresponding media
    /// * togglePause: play corresponding media if media is already paused, pause
    ///                corresponding media if media is already playing. If the corresponding
    ///                media is not active, this command restarts the media and plays from
    ///                its beginning.
    /// 
    /// Reserved Values (when command_type == TLCommandType::Event):
    /// * onstopaudio: stop play of all audio
    /// 
    /// Reserved Values (when command_type == TLCommandType::Verb):
    /// * 0: Open the object for editing
    /// * 1: Open the object for viewing
    /// 
    /// The value of the cmd attribute shall be the string representation of an integer that
    /// represents the embedded object verb number. This verb number determines the action
    /// that the rendering application should take corresponding to this object when this point in
    /// the animation is reached.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:cmd type="evt" cmd="onstopaudio">
    ///   <p:cBhvr>
    ///     <p:cTn display="0" masterRel="sameClick">
    ///       <p:stCondLst>
    ///         <p:cond evt="begin" delay="0">
    ///           <p:tn val="5"/>
    ///         </p:cond>
    ///       </p:stCondLst>
    ///     </p:cTn>
    ///     <p:tgtEl>
    ///       <p:sldTgt/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:cmd>
    /// ```
    /// 
    /// In the above example, the event of onstopaudio stops all audio from playing once this
    /// particular animation is reached in the timeline.
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
    /// The element specifies the certain attribute of a time node after an animation effect.
    /// 
    /// # Xml example
    /// 
    /// Consider an animation effect that leaves a string value visible afterwards. The <to> element should
    /// be used as follows:
    /// ```xml
    /// <p:childTnLst>
    ///   <p:set>
    ///     <p:cBhvr> … </p:cBhvr>
    ///     <p:to>
    ///       <p:strVal val="visible"/>
    ///     </p:to>
    ///   </p:set>
    ///   <p:anim calcmode="lin" valueType="num"> … </p:anim> …
    /// </p:childTnLst>
    /// ```
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
                }
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
    /// This attribute indicates whether the audio is a narration for the slide.
    /// 
    /// Defaults to false
    pub is_narration: Option<bool>,
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
    /// This attribute specifies if the video is displayed in full-screen.
    /// 
    /// Defaults to false
    pub fullscreen: Option<bool>,
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

/// This element defines a "keypoint" in animation interpolation.
/// 
/// # Xml example
/// 
/// Consider a shape with a "fly-in" animation. The <tav> element should be used as follows:
/// ```xml
/// <p:anim calcmode="lin" valueType="num">
///   <p:cBhvr additive="base"> … </p:cBhvr>
///   <p:tavLst>
///     <p:tav tm="0%">
///       <p:val>
///         <p:strVal val="1+#ppt_h/2"/>
///       </p:val>
///     </p:tav>
///     <p:tav tm="100%">
///       <p:val>
///         <p:strVal val="#ppt_y"/>
///       </p:val>
///     </p:tav>
///   </p:tavLst>
/// </p:anim>
/// ```
#[derive(Default, Debug, Clone)]
pub struct TLTimeAnimateValue {
    /// This attribute specifies the time at which the attribute being animated takes on the value.
    /// 
    /// Defaults to TLTimeAnimateValueTime::Indefinite
    pub time: Option<TLTimeAnimateValueTime>,
    /// This attribute allows for the specification of a formula to be used for describing a
    /// complex motion for an animated object. The formula manipulates the motion of the
    /// object by modifying a property of the object over a specified period of time. Each formula
    /// has zero or more inputs specified by the ($) symbol, zero or more variables specified by
    /// the (#) symbol pre-pended to the variable name and a target variable which is specified
    /// by the previously specified attrName element. The formula can contain one or more of
    /// any of the constants, operators or functions listed below. In addition to this, the formula
    /// can also contain floating point numbers and parentheses.
    /// 
    /// Mathematical operations have the following order of precedence, listed from lowest to
    /// highest. Operators listed on the same line have equal precedence.
    /// 
    /// * “+”, “-“
    /// * “*”, “/”, “%”
    /// * “^”
    /// * Unary minus, Unary plus (e.g. -2, meaning 3*-2 is the same as 3*(-2))
    /// * Variables, Constants (including numbers) and Functions (as listed previously)
    /// 
    /// # Language Description
    /// 
    /// Digit       = '0' | '1' | ‘2’ | ‘3’ | ‘4’ | ‘5’ | ‘6’ | ‘7’ | ‘8’ | '9' ;
    /// 
    /// number      = digit , { digit } ;
    /// 
    /// exponent    = [ '-' ] , ( 'e' | 'E' ) , number ;
    /// 
    /// value       = number , [ '.' number ] , [ exponent ] ;
    /// 
    /// variable    = '$' | 'ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h' ;
    /// 
    /// constant    = value | 'pi' | 'e' ;
    /// 
    /// ident       = 'abs' | ‘acos’ | ‘asin’ | ‘atan’ | ‘ceil’
    ///               | ‘cos’ | ‘cosh’ | ‘deg’ | ‘exp’ | ‘floor’ | ‘ln’
    ///               | ‘max’ | ‘min’ | ‘rad’ | ‘rand’ | ‘sin’ | ‘sinh’
    ///               | ‘sqrt’ | ‘tan’ | 'tanh' ;
    /// 
    /// function    = ident , '(' , formula [ ',' , formula ] , ')' ;
    /// 
    /// formula     = term , { [ '+' | '-' ] , term } ;
    /// 
    /// term        = power , { [ '*' | '/' | '%' ] , power } ;
    /// 
    /// power       = unary [ '^' , unary ] ;
    /// 
    /// unary       = [ '+' | '-' ] , factor ;
    /// 
    /// factor      = variable | constant | function | parens ;
    /// 
    /// parens      = '(' , formula , ')' ;
    /// 
    /// ## Note
    /// 
    /// Formulas can only support a calcMode (Calculation Mode) of linear or discrete. If
    /// another calcMode is specified or no calcMode is specified then a calcMode of linear is
    /// assumed.
    /// 
    /// Any additional characters in the formula string that are not contained within the
    /// set described are considered invalid.
    /// 
    /// # Variables
    /// 
    /// |Name       |Description                                        |
    /// |-----------|---------------------------------------------------|
    /// |$          |Formula input                                      |
    /// |ppt_x      |Pre-animation x position of the object on the slide|
    /// |ppt_y      |Pre-animation y position of the object on the slide|
    /// |ppt_w      |Pre-animation width of the object                  |
    /// |ppt_h      |Pre-animation height of the object                 |
    /// 
    /// # Constants
    /// 
    /// |Name       |Description                                        |
    /// |-----------|---------------------------------------------------|
    /// |pi         |Mathematical constant pi                           |
    /// |e          |Mathematical constant e                            |
    /// 
    /// # Operators
    /// 
    /// |Name       |Description        |Usage                                  |
    /// |-----------|-------------------|---------------------------------------|
    /// |+          |Addition           |“x+y”, adds x to the value y           |
    /// |-          |Subtraction        |“x-y”, subtracts y from the value x    |
    /// |*          |Multiplication     |“x*y”, multiplies x by the value y     |
    /// |/          |Division           |“x/y”, divides x by the value y        |
    /// |%          |Modulus            |“x%y”, the remainder of x/y            |
    /// |^          |Power              |“x^y”, x raised to the power y         |
    /// 
    /// # Functions
    /// 
    /// |Name       |Description                |Usage                                                              |
    /// |-----------|---------------------------|-------------------------------------------------------------------|
    /// |abs        |Absolute value             |“abs(x)”, absolute value of x                                      |
    /// |acos       |Arc Cosine                 |“acos(x)”, arc cosine of the value x                               |
    /// |asin       |Arc Sine                   |“asin(x)”, arc sine of the value x                                 |
    /// |atan       |Arc Tangent                |“atan(x)”, arc tangent of the value x                              |
    /// |ceil       |Ceil value                 |“ceil(x)”, value of x rounded up                                   |
    /// |cos        |Cosine                     |“cos(x)”, cosine of the value of x                                 |
    /// |cosh       |Hyperbolic Cosine          |“cosh(x)", hyperbolic cosine of the value x                        |
    /// |deg        |Radiant to Degree convert  |“deg(x)”, the degree value of radiant value x                      |
    /// |exp        |Exponent                   |“exp(x)”, value of constant e raised to the power of x             |
    /// |floor      |Floor value                |“floor(x)”, value of x rounded down                                |
    /// |ln         |Natural logarithm          |“ln(x)”, natural logarithm of x                                    |
    /// |max        |Maximum of two values      |“max(x,y)”, returns x if (x > y) or returns y if (y > x)           |
    /// |min        |Minimum of two values      |“min(x,y)", returns x if (x < y) or returns y if (y < x)           |
    /// |rad        |Degree to Radiant convert  |“rad(x)”, the radiant value of degree value x                      |
    /// |rand       |Random value               |“rand(x)”, returns a random floating point value between 0 and x   |
    /// |sin        |Sine                       |“sin(x)”, sine of the value x                                      |
    /// |sinh       |Hyperbolic Sine            |"sinh(x)”, hyperbolic sine of the value x                          |
    /// |sqrt       |Square root                |“sqrt(x)”, square root of the value x                              |
    /// |tan        |Tangent                    |“tan(x)”, tangent of the value x                                   |
    /// |tanh       |Hyperbolic Tangent         |“tanh(x)", hyperbolic tangent of the value x                       |
    /// 
    /// # Xml example
    /// 
    /// <p:animcalcmode="lin" valueType="num">
    ///   <p:cBhvr>
    ///     <p:cTn id="9" dur="664" tmFilter="0.0,0.0; 0.25,0.07;0.50,0.2; 0.75,0.467; 1.0,1.0">
    ///       <p:stCondLst>
    ///         <p:cond delay="0"/>
    ///       </p:stCondLst>
    ///     </p:cTn>
    ///     <p:tgtEl>
    ///       <p:spTgtspid="4"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>ppt_y</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:tavLst>
    ///     <p:tav tm="0%" fmla="#ppt_y-sin(pi*$)/3">
    ///       <p:val>
    ///         <p:fltValval="0.5"/>
    ///       </p:val>
    ///     </p:tav>
    ///     <p:tav tm="100%">
    ///       <p:val>
    ///         <p:fltValval="1"/>
    ///       </p:val>
    ///     </p:tav>
    ///   </p:tavLst>
    /// </p:anim>
    /// 
    /// The animation example above modifies the ppt_y variable of the object by subtracting
    /// sin(pi*$)/3 from the non-animated value of ppt_y. The start value is 0.5 and the end
    /// value is 1 specified in each of the val elements. The total time for this animation is
    /// specified within the dur attribute and the filtered time graph is specified by the tmFilter
    /// attribute. The end result is that the object moves from a point above its non-animated
    /// position back to its non-animated position. With the specification of the tmFilter it has a
    /// modified time graph such that it also appears to accelerate as it reaches its final position.
    /// 
    /// ## Note
    /// 
    /// For this example, the non-animated value of ppt_y is the value of this variable if
    /// the object were to be statically rendered on the slide without animation properties.
    pub formula: Option<String>,
    /// The element specifies a value for a time animate.
    /// 
    /// # Xml example
    /// 
    /// Consider a shape with a fade in animation effect. The <val> element should be used as follows:
    /// ```xml
    /// <p:anim calcmode="lin" valueType="num">
    ///   <p:cBhvr additive="base"> … </p:cBhvr>
    ///   <p:tavLst>
    ///     <p:tav tm="0%">
    ///       <p:val>
    ///         <p:strVal val="0-#ppt_w/2"/>
    ///       </p:val>
    ///     </p:tav>
    ///     <p:tav tm="100%">
    ///       <p:val>
    ///         <p:strVal val="#ppt_x"/>
    ///       </p:val>
    ///     </p:tav>
    ///   </p:tavLst>
    /// </p:anim>
    /// ```
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
    Percentage(msoffice_shared::drawingml::PositiveFixedPercentage),
    Indefinite,
}

impl FromStr for TLTimeAnimateValueTime {
    type Err = msoffice_shared::error::ParseEnumError;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s {
            "indefinite" => Ok(TLTimeAnimateValueTime::Indefinite),
            _ => Ok(TLTimeAnimateValueTime::Percentage(
                s.parse().map_err(|_| Self::Err::new("TLTimeAnimateValueTime"))?,
            )),
        }
    }
}

#[derive(Debug, Clone)]
pub enum TLAnimVariant {
    /// This element specifies a boolean value to be used for evaluation by a parent element. The exact meaning of the
    /// value contained within this element is not defined here but is dependent on the usage of this element in
    /// conjunction with one of the listed parent elements.
    Bool(bool),
    /// This element specifies an integer value to be used for evaluation by a parent element. The exact meaning of the
    /// value contained within this element is not defined here but is dependent on the usage of this element in
    /// conjunction with one of the listed parent elements.
    Int(i32),
    /// This element specifies a floating point value to be used for evaluation by a parent element. The exact meaning
    /// of the value contained within this element is not defined here but is dependent on the usage of this element in
    /// conjunction with one of the listed parent elements.
    Float(f32),
    /// This element specifies a string value to be used for evaluation by a parent element. The exact meaning of the
    /// value contained within this element is not defined here but is dependent on the usage of this element in
    /// conjunction with one of the listed parent elements.
    String(String),
    /// This element describes the color variant. This is used to specify a color that is to be used for animating the color
    /// property of an object.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:set>
    ///   <p:cBhvr override="childStyle">
    ///     …
    ///   </p:cBhvr>
    ///   <p:to>
    ///     <p:clrVal>
    ///       <a:schemeClr val="accent2"/>
    ///     </p:clrVal>
    ///   </p:to>
    /// </p:set>
    /// ```
    Color(msoffice_shared::drawingml::Color),
}

impl TLAnimVariant {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
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
                Ok(TLAnimVariant::Color(msoffice_shared::drawingml::Color::from_xml_element(
                    child_node,
                )?))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TLAnimVariant").into()),
        }
    }
}

/// This element describes the common behaviors of animations.
/// 
/// # Xml example
/// 
/// ```xml
/// <p:anim to="1.5" calcmode="lin" valueType="num">
///   <p:cBhvr override="childStyle">
///     <p:cTn id="6" dur="2000" fill="hold"/>
///     <p:tgtEl>
///       <p:spTgt spid="3">
///         <p:txEl>
///           <p:charRg st="4294967295" end="4294967295"/>
///         </p:txEl>
///       </p:spTgt>
///     </p:tgtEl>
///     <p:attrNameLst>
///       <p:attrName>style.fontSize</p:attrName>
///     </p:attrNameLst>
///   </p:cBhvr>
/// </p:anim>
/// ```
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
    /// This element specifies the target children elements which have the animation effects applied to.
    /// 
    /// # Xml example
    /// 
    /// Consider a shape with ID 3 with a fade effect animation applied to it. The <tgtEl> element should be
    /// used as follows:
    /// ```xml
    /// <p:animEffect transition="in" filter="fade">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="2000"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    pub target_element: TLTimeTargetElement,
    /// This element is used to describe a list of attributes in which to apply an animation to.
    /// 
    /// The elements of this list is used to contain an attribute value for an Attribute Name List. This value defines the specific
    /// attribute that an animation should be applied to, such as fill, style, and shadow, etc. A specific property is
    /// defined by using a "property.sub-property" format which is often extended to multiple sub properties as seen in
    /// the allowed values below.
    /// 
    /// Allowed property values:
    /// * style.opacity
    /// * style.rotation
    /// * style.visibility
    /// * style.color
    /// * style.fontSize
    /// * style.fontWeight
    /// * style.fontStyle
    /// * style.fontFamily
    /// * style.textEffectEmboss
    /// * style.textShadow
    /// * style.textTransform
    /// * style.textDecorationUnderline
    /// * style.textEffectOutline
    /// * style.textDecorationLineThrough
    /// * style.sRotation
    /// * imageData.cropTop
    /// * imageData.cropBottom
    /// * imageData.cropLeft
    /// * imageData.cropRight
    /// * imageData.gain
    /// * imageData.blacklevel
    /// * imageData.gamma
    /// * imageData.grayscale
    /// * imageData.chromakey
    /// * fill.on
    /// * fill.type
    /// * fill.color
    /// * fill.opacity
    /// * fill.color2
    /// * fill.method
    /// * fill.opacity2
    /// * fill.angle
    /// * fill.focus
    /// * fill.focusposition.x
    /// * fill.focusposition.y
    /// * fill.focussize.x
    /// * fill.focussize.y
    /// * stroke.on
    /// * stroke.color
    /// * stroke.weight
    /// * stroke.opacity
    /// * stroke.linestyle
    /// * stroke.dashstyle
    /// * stroke.filltype
    /// * stroke.src
    /// * stroke.color2
    /// * stroke.imagesize.x
    /// * stroke.imagesize.y
    /// * stroke.startArrow
    /// * stroke.endArrow
    /// * stroke.startArrowWidth
    /// * stroke.startArrowLength
    /// * stroke.endArrowWidth
    /// * stroke.endArrowLength
    /// * shadow.on
    /// * shadow.type
    /// * shadow.color
    /// * shadow.color2
    /// * shadow.opacity
    /// * shadow.offset.x
    /// * shadow.offset.y
    /// * shadow.offset2.x
    /// * shadow.offset2.y
    /// * shadow.origin.x
    /// * shadow.origin.y
    /// * shadow.matrix.xtox
    /// * shadow.matrix.ytox
    /// * shadow.matrix.xtoy
    /// * shadow.matrix.ytoy
    /// * shadow.matrix.perspectiveX
    /// * shadow.matrix.perspectiveY
    /// * skew.on
    /// * skew.offset.x
    /// * skew.offset.y
    /// * skew.origin.x
    /// * skew.origin.y
    /// * skew.matrix.xtox
    /// * skew.matrix.ytox
    /// * skew.matrix.xtoy
    /// * skew.matrix.ytoy
    /// * skew.matrix.perspectiveX
    /// * skew.matrix.perspectiveY
    /// * extrusion.on
    /// * extrusion.type
    /// * extrusion.render
    /// * extrusion.viewpointorigin.x
    /// * extrusion.viewpointorigin.y
    /// * extrusion.viewpoint.x
    /// * extrusion.viewpoint.y
    /// * extrusion.viewpoint.z
    /// * extrusion.plane
    /// * extrusion.skewangle
    /// * extrusion.skewamt
    /// * extrusion.backdepth,
    /// * extrusion.foredepth
    /// * extrusion.orientation.x
    /// * extrusion.orientation.y
    /// * extrusion.orientation.zand
    /// * extrusion.orientationangle
    /// * extrusion.color,
    /// * extrusion.rotationangle.x
    /// * extrusion.rotationangle.y
    /// * extrusion.lockrotationcenter
    /// * extrusion.autorotationcenter
    /// * extrusion.rotationcenter.x
    /// * extrusion.rotationcenter.y
    /// * extrusion.rotationcenter.z
    /// * extrusion.colormode.
    /// 
    /// # Xml example
    /// 
    /// Consider trying to emphasize the txt font size within the body of a shape. The attribute would be
    /// 'style.fontSize' and this can be done by doing the following:
    /// 
    /// ```xml
    /// <p:anim to="1.5" calcmode="lin" valueType="num">
    ///   <p:cBhvr override="childStyle">
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3">
    ///         <p:txEl>
    ///           <p:charRg st="4294967295" end="4294967295"/>
    ///         </p:txEl>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>style.fontSize</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    /// </p:anim>
    /// ```
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
                        return Err(Box::new(MissingChildNodeError::new(
                            child_node.name.clone(),
                            "attrName",
                        )));
                    }

                    attr_name_list = Some(vec);
                }
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

/// This element is used to describe behavior of media elements, such as sound or movies, in an animation.
/// 
/// # Xml example
/// 
/// ```xml
/// <p:audio>
///   <p:cMediaNode mute="1">
///     <p:cTn display="0" masterRel="sameClick">
///       <p:stCondLst> … </p:stCondLst>
///       <p:endCondLst> … </p:endCondLst>
///     </p:cTn>
///     <p:tgtEl> … </p:tgtEl>
///   </p:cMediaNode>
/// </p:audio>
/// ```
#[derive(Debug, Clone)]
pub struct TLCommonMediaNodeData {
    /// This attribute describes the volume of the media element.
    /// 
    /// Defaults to 50000
    pub volume: Option<msoffice_shared::drawingml::PositiveFixedPercentage>,
    /// This attribute describes whether the media should be mute.
    /// 
    /// Defaults to false
    pub mute: Option<bool>,
    /// This attribute describes the numbers of slides across which the media should play.
    /// 
    /// Defaults to 1
    pub number_of_slides: Option<u32>,
    /// This attribute describes whether the media should be displayed when it is stopped.
    /// 
    /// Defaults to true
    pub show_when_stopped: Option<bool>,
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
    /// This element describes the time node trigger choice.
    /// 
    /// # Xml example
    /// 
    /// Consider a time node with an event condition. The <tn> element should be used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="5">
    ///     <p:stCondLst>
    ///       <p:cond delay="0"/>
    ///     </p:stCondLst>
    ///     <p:endCondLst>
    ///       <p:cond evt="begin" delay="0">
    ///         <p:tn val="5"/>
    ///       </p:cond>
    ///     </p:endCondLst>
    ///     <p:childTnLst> … </p:childTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    TimeNode(TLTimeNodeId),
    /// This element specifies the child time node that triggers a time condition. References a child time node or all
    /// child nodes. Order is based on the child's end time.
    /// 
    /// # Xml example
    /// 
    /// Consider an animation which ends the synchronization of all parallel time nodes when all the child
    /// nodes have ended their animation. The <rtn> element should be used as follows:
    /// ```xml
    /// <p:cTn>
    ///   <p:stCondLst> … </p:stCondLst>
    ///   <p:endSync evt="end" delay="0">
    ///     <p:rtn val="all"/>
    ///   </p:endSync>
    ///   <p:childTnLst> … </p:childTnLst>
    /// </p:cTn>
    /// ```
    RuntimeNode(TLTriggerRuntimeNode),
}

impl TLTimeConditionTriggerGroup {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
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
            }
            "tn" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLTimeConditionTriggerGroup::TimeNode(val_attr.parse()?))
            }
            "rtn" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLTimeConditionTriggerGroup::RuntimeNode(val_attr.parse()?))
            }
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "EG_TLTimeConditionTriggerGroup",
            ))),
        }
    }
}

#[derive(Debug, Clone)]
pub enum TLTimeTargetElement {
    /// This element specifies the slide as the target element.
    /// 
    /// # Xml example
    /// 
    /// For example, suppose we have a simple animation with a blind entrance.
    /// ```xml
    /// <p:seq concurrent="1" nextAc="seek">
    ///   <p:cTn id="2" dur="indefinite" nodeType="mainSeq"> … </p:cTn>
    ///   <p:prevCondLst> … </p:prevCondLst>
    ///   <p:nextCondLst>
    ///     <p:cond evt="onNext" delay="0">
    ///       <p:tgtEl>
    ///         <p:sldTgt/>
    ///       </p:tgtEl>
    ///     </p:cond>
    ///   </p:nextCondLst>
    /// </p:seq>
    /// ```
    SlideTarget,
    /// This element describes the sound information for a target object.
    /// 
    /// # Xml example
    /// 
    /// Consider a shape with a sound effect animation. The <sndTgt> element should be used as follows:
    /// ```xml
    /// <p:subTnLst>
    ///   <p:audio>
    ///     <p:cMediaNode>
    ///       <p:cTn display="0" masterRel="sameClick"> … </p:cTn>
    ///       <p:tgtEl>
    ///         <p:sndTgt r:embed="rId2" r:link="rId3"/>
    ///       </p:tgtEl>
    ///     </p:cMediaNode>
    ///   </p:audio>
    /// </p:subTnLst>
    /// ```
    SoundTarget(msoffice_shared::drawingml::EmbeddedWAVAudioFile),
    /// The element specifies the shape in which to apply a certain animation to.Err
    /// 
    /// # Xml example
    /// 
    /// Consider a shape whose id is 3 in which we want to apply a fade animation effect. The <spTgt> should
    /// be used as follows:
    /// ```xml
    /// <p:animEffect transition="in" filter="fade">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="2000"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    ShapeTarget(TLShapeTargetElement),
    /// This element specifies an animation target element that is represented by a sub-shape in a legacy graphical
    /// object.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animEffect transition="in" filter="blinds(horizontal)">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:inkTgt spid="_x0000_s2057"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    InkTarget(TLSubShapeId),
}

impl TLTimeTargetElement {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
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
                msoffice_shared::drawingml::EmbeddedWAVAudioFile::from_xml_element(xml_node)?,
            )),
            "spTgt" => {
                let child_node = xml_node
                    .child_nodes
                    .get(0)
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "CT_TLShapeTargetElement"))?;
                
                Ok(TLTimeTargetElement::ShapeTarget(
                    TLShapeTargetElement::from_xml_element(child_node)?,
                ))
            }
            "inkTgt" => {
                let spid_attr = xml_node
                    .attribute("spid")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
                
                Ok(TLTimeTargetElement::InkTarget(spid_attr.parse()?))
            }
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "CT_TLTimeTargetElement",
            ))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLShapeTargetElement {
    /// This attribute specifies the shape identifier.
    pub shape_id: msoffice_shared::drawingml::DrawingElementId,
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
    /// This element is used to specify animating the background of an object.
    /// 
    /// # Xml example
    /// 
    /// Consider adding animation to the background of Shape Id 3. The <bg> tag can be used as follows:
    /// 
    /// ```xml
    /// <p:tgtEl>
    ///   <p:spTgt spid="3">
    ///     <p:bg/>
    ///   </p:spTgt>
    /// </p:tgtEl>
    /// ```
    Background,
    /// This element specifies the subshape of a legacy graphical object to animate.
    /// 
    /// # Xml example
    /// 
    /// Consider adding animation to a legacy diagram. The <subSp> element should be used as follows:
    /// ```xml
    /// <p:animEffect transition="in" filter="blinds(horizontal)">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="2053">
    ///         <p:subSp spid="_x0000_s70664"/>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    SubShape(TLSubShapeId),
    /// This element specifies the subelement of an embedded chart to animate.
    /// 
    /// # Xml example
    /// 
    /// Consider an embedded Chart with a entrance animation effect applied to each of the graph's
    /// categories. The <oldChartEl> element should be used as follows:
    /// ```xml
    /// <p:animEffect transition="in" filter="blinds(horizontal)">
    ///   <p:cBhvr>
    ///     <p:cTn id="12" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="19460">
    ///         <p:oleChartEl type="category" lvl="1"/>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    OleChartElement(TLOleChartTargetElement),
    /// This element specifies a text element to animate.
    /// 
    /// # Xml example
    /// 
    /// Consider a shape containing text to be animated. The <txEl> should be used as follows:
    /// ```xml
    /// <p:tgtEl>
    ///   <p:spTgt spid="5">
    ///     <p:txEl>
    ///       <p:pRg st="1" end="1"/>
    ///     </p:txEl>
    ///   </p:spTgt>
    /// </p:tgtEl>
    /// ```
    TextElement(Option<TLTextTargetElement>),
    /// This element specifies a graphical element which to animate
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:set>
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="1" fill="hold"> … </p:cTn>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4">
    ///         <p:graphicEl>
    ///           <a:dgm id="{87C2C707-C3F4-4E81-A967-A8B8AE13E575}"/>
    ///         </p:graphicEl>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///     <p:attrNameLst> … </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:to> … </p:to>
    /// </p:set>
    /// ```
    GraphicElement(msoffice_shared::drawingml::AnimationElementChoice),
}

impl TLShapeTargetElementGroup {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
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
            }
            "oleChartEl" => Ok(TLShapeTargetElementGroup::OleChartElement(
                TLOleChartTargetElement::from_xml_element(xml_node)?,
            )),
            "txEl" => Ok(TLShapeTargetElementGroup::TextElement(
                match xml_node.child_nodes.get(0) {
                    Some(child_node) => Some(TLTextTargetElement::from_xml_element(child_node)?),
                    None => None,
                },
            )),
            "graphicEl" => {
                let child_node = xml_node
                    .child_nodes
                    .get(0)
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "CT_AnimationElementChoice"))?;

                Ok(TLShapeTargetElementGroup::GraphicElement(
                    msoffice_shared::drawingml::AnimationElementChoice::from_xml_element(child_node)?,
                ))
            }
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TLShapeTargetElementGroup",
            ))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLOleChartTargetElement {
    /// This attribute specifies how to chart should be built during its animation.
    pub element_type: TLChartSubelementType,
    /// This attribute describes the element levels to animate.
    /// 
    /// Defaults to 0
    pub level: Option<u32>,
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
    /// This element specifies animation on a character range defined by a start and end character position.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animMotion>
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3">
    ///         <p:txEl>
    ///           <p:charRg st="0" end="9"/>
    ///         </p:txEl>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>ppt_x</p:attrName>
    ///       <p:attrName>ppt_y</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    /// </p:animMotion>
    /// ```
    CharRange(IndexRange),
    /// This element specifies a text range to animate based on starting and ending paragraph number.
    /// 
    /// # Xml example
    /// 
    /// Consider an animation entrance of the first 3 text paragraphs. The <pRg> element should be used as
    /// follows:
    /// ```xml
    /// <p:animEffect transition="in" filter="checkerboard(across)">
    ///   <p:cBhvr>
    ///     <p:cTn id="12" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3">
    ///         <p:txEl>
    ///           <p:pRg st="0" end="2"/>
    ///         </p:txEl>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    ParagraphRange(IndexRange),
}

impl TLTextTargetElement {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "charRg" | "pRg" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "charRg" => Ok(TLTextTargetElement::CharRange(IndexRange::from_xml_element(xml_node)?)),
            "pRg" => Ok(TLTextTargetElement::ParagraphRange(IndexRange::from_xml_element(
                xml_node,
            )?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TLTextTargetElement",
            ))),
        }
    }
}

/// This element specifies conditions on time nodes in a timeline. It is used within a list of start condition or list of
/// end condition elements.
/// 
/// # Xml example
/// 
/// ```xml
/// <p:cTn>
///   <p:stCondLst>
///     <p:cond delay="2000"/>
///   </p:stCondLst>
///   <p:childTnLst>
///     <p:set> … </p:set>
///     <p:animEffect transition="in" filter="blinds(horizontal)">
///       <p:cBhvr>
///         <p:cTn id="7" dur="1000"/>
///         <p:tgtEl>
///           <p:spTgt spid="4"/>
///         </p:tgtEl>
///       </p:cBhvr>
///     </p:animEffect>
///   </p:childTnLst>
/// </p:cTn>
/// ```
#[derive(Default, Debug, Clone)]
pub struct TLTimeCondition {
    /// This attribute describes the event that triggers an animation.
    pub trigger_event: Option<TLTriggerEvent>,
    /// This attribute describes the delay after an animation is triggered.
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

/// This element describes the properties that are common for time nodes.
#[derive(Default, Debug, Clone)]
pub struct TLCommonTimeNodeData {
    /// This attribute specifies the identifier for the timenode.
    pub id: Option<TLTimeNodeId>,
    /// This attribute describes the preset identifier for the time node.
    pub preset_id: Option<i32>,
    /// This attribute descries the class of effect in which it belongs.
    pub preset_class: Option<TLTimeNodePresetClassType>,
    /// This attribute is a bitflag that specifies a direction or some other attribute of the effect.
    /// For example it can be set to specify a “From Bottom” for the Fly In effect, or “Bold” for
    /// the Change Font Style effect.
    pub preset_subtype: Option<i32>,
    /// This attribute describes the duration of the time node, expressed as unit time.
    pub duration: Option<TLTime>,
    /// This attribute describes the number of times the element should repeat, in units of
    /// thousandths.
    /// 
    /// Defaults to 100_0
    pub repeat_count: Option<TLTime>,
    /// This attribute describes the amount of time over which the element should repeat. If
    /// absent, the attribute is taken to be the same as the specified duration.
    pub repeat_duration: Option<TLTime>,
    /// This attribute specifies the percentage by which to speed up (or slow down) the timing. If
    /// negative, the timing is reversed.
    /// 
    /// Defaults to 100_000
    /// 
    /// # Example
    /// 
    /// If speed is 200% (200_000) and the specified duration is 10 seconds, the actual duration is 5 seconds.
    pub speed: Option<msoffice_shared::drawingml::Percentage>,
    /// This attribute describes the percentage of specified duration over which the element's
    /// time takes to accelerate from 0 up to the "run rate."
    /// 
    /// Defaults to 0
    pub acceleration: Option<msoffice_shared::drawingml::PositiveFixedPercentage>,
    /// This attribute describes the percentage of specified duration over which the element's
    /// time takes to decelerate from the "run rate" down to 0.
    /// 
    /// Defaults to 0
    pub deceleration: Option<msoffice_shared::drawingml::PositiveFixedPercentage>,
    /// This attribute describes whether to automatically play the animation in reverse after
    /// playing it in the forward direction.
    /// 
    /// Defaults to false
    pub auto_reverse: Option<bool>,
    /// This attribute specifies if a node is to restart when it completes its action
    pub restart_type: Option<TLTimeNodeRestartType>,
    /// This attribute describes the fill type for the time node.
    pub fill_type: Option<TLTimeNodeFillType>,
    /// This attribute specifies how the time node synchronizes to its group.
    pub sync_behavior: Option<TLTimeNodeSyncType>,
    /// This attribute specifies the time filter for the time node.
    pub time_filter: Option<String>,
    /// This attribute describes the event filter for this time node.
    pub event_filter: Option<String>,
    /// This attribute describes whether the state of the time node is visible or hidden.
    pub display: Option<bool>,
    /// This attribute specifies how the time node plays back relative to its master time node.
    pub master_relationship: Option<TLTimeNodeMasterRelation>,
    /// This attribute describes the build level of the animation.
    pub build_level: Option<i32>,
    /// This attribute describes the Group ID of the time node.
    pub group_id: Option<u32>,
    /// This attribute specifies whether there is an after effect applied to the time node.
    pub after_effect: Option<bool>,
    /// This attribute specifies the type of time node.
    pub node_type: Option<TLTimeNodeType>,
    /// This attribute describes whether this node is a placeholder.
    pub node_placeholder: Option<bool>,
    /// This element contains a list conditions that shall be met for a time node to be activated.
    /// 
    /// # Xml example
    /// 
    /// example, suppose we have a shape with an entrance appearance after 5 seconds. The
    /// <stCondLst> element should be used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="5" nodeType="clickEffect">
    ///     <p:stCondLst>
    ///       <p:cond delay="5000"/>
    ///     </p:stCondLst>
    ///     <p:childTnLst> … </p:childTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    pub start_condition_list: Option<Vec<TLTimeCondition>>,
    /// This element describes a list of the end conditions that shall be met in order to stop the time node.
    /// 
    /// # Xml example
    /// 
    /// Consider a shape a shape with an audio attached to the animation. The <endCondList> element
    /// should be used as follows to specifies when the sound is done:
    /// ```xml
    /// <p:audio>
    ///   <p:cMediaNode>
    ///     <p:cTn display="0" masterRel="sameClick">
    ///       <p:stCondLst> … </p:stCondLst>
    ///       <p:endCondLst>
    ///         <p:cond evt="onStopAudio" delay="0">
    ///           <p:tgtEl>
    ///             <p:sldTgt/>
    ///           </p:tgtEl>
    ///         </p:cond>
    ///       </p:endCondLst>
    ///     </p:cTn>
    ///     <p:tgtEl> … </p:tgtEl>
    ///   </p:cMediaNode>
    /// </p:audio>
    /// ```
    pub end_condition_list: Option<Vec<TLTimeCondition>>,
    /// This element is used to synchronizes the stopping of parallel elements in the timing tree. It is used on interactive
    /// timeline sequences to specify that the interactive sequence’s duration ends when all of the child timenodes
    /// have ended. It is also used to make interactive sequences restart-able (so that the entire interactive sequence
    /// can be repeated if the trigger object is clicked on repeatedly).
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:seq concurrent="1" nextAc="seek">
    ///   <p:cTn>
    ///     <p:stCondLst/>
    ///     <p:endSync evt="end" delay="0">
    ///       <p:rtn val="all"/>
    ///     </p:endSync>
    ///     <p:childTnLst/>
    ///   </p:cTn>
    ///   <p:nextCondLst/>
    /// </p:seq>
    /// ```
    pub end_sync: Option<TLTimeCondition>,
    /// This element specifies how the animation should be successively applied to sub elements of the target element
    /// for a repeated effect. It can be applied to contained timing and animation structures over the letters, words, or
    /// shapes within a target element.
    /// 
    /// # Xml example
    /// 
    /// Consider a text animation where the words appear letter by letter. The <iterate> element should be
    /// used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="1" >
    ///     <p:stCondLst> … </p:stCondLst>
    ///     <p:iterate type="lt">
    ///       <p:tmPct val="10000"/>
    ///     </p:iterate>
    ///     <p:childTnLst> … </p:childTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    pub iterate: Option<TLIterateData>,
    /// This element describes the list of time nodes that have a fixed location in the timing tree based on their parent
    /// time node. The children's start time is defined relative to their parent time node’s start.
    pub child_time_node_list: Option<Vec<TimeNodeGroup>>,
    /// This element describes time nodes that have a start time which is not based on the containing timenode. It is
    /// instead based on their master relationship (masterRel). At runtime, they are inserted dynamically into the
    /// timing tree as child timenodes for playback, based on the logic defined by the master relationship. These
    /// elements are used for animations such as "dim after" and "play sound effects"
    /// 
    /// # Xml example
    /// 
    /// Consider an animation with a "Fly In" effect on paragraphs so that each paragraph flies in on a
    /// separate click. Then the "Dim After" effect for paragraph 1 doe not happen until paragraph 2 flies in. The
    /// <subTnLst> element should be used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="5" grpId="0" nodeType="clickEffect">
    ///     <p:stCondLst> … </p:stCondLst>
    ///     <p:childTnLst> … </p:childTnLst>
    ///     <p:subTnLst>
    ///       <p:set>
    ///         <p:cBhvr override="childStyle">
    ///           <p:cTn fill="hold" masterRel="nextClick" afterEffect="1"/>
    ///           <p:tgtEl> … </p:tgtEl>
    ///           <p:attrNameLst> … </p:attrNameLst>
    ///         </p:cBhvr>
    ///         <p:to> … </p:to>
    ///         </p:set>
    ///     </p:subTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
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
                _ => (),
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
                }
                "endCondLst" => {
                    let mut vec = Vec::new();
                    for cond_node in &child_node.child_nodes {
                        vec.push(TLTimeCondition::from_xml_element(cond_node)?);
                    }

                    if vec.is_empty() {
                        return Err(Box::new(MissingChildNodeError::new(child_node.name.clone(), "cond")));
                    }

                    instance.end_condition_list = Some(vec);
                }
                "endSync" => instance.end_sync = Some(TLTimeCondition::from_xml_element(child_node)?),
                "iterate" => instance.iterate = Some(TLIterateData::from_xml_element(child_node)?),
                "childTnLst" => {
                    let mut vec = Vec::new();
                    for time_node in &child_node.child_nodes {
                        vec.push(TimeNodeGroup::from_xml_element(time_node)?);
                    }

                    if vec.is_empty() {
                        return Err(Box::new(MissingChildNodeError::new(
                            child_node.name.clone(),
                            "TimeNode",
                        )));
                    }

                    instance.child_time_node_list = Some(vec);
                }
                "subTnLst" => {
                    let mut vec = Vec::new();
                    for time_node in &child_node.child_nodes {
                        vec.push(TimeNodeGroup::from_xml_element(time_node)?);
                    }

                    if vec.is_empty() {
                        return Err(Box::new(MissingChildNodeError::new(
                            child_node.name.clone(),
                            "TimeNode",
                        )));
                    }

                    instance.sub_time_node_list = Some(vec);
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone)]
pub enum TLIterateDataChoice {
    /// This element describes the duration of the iteration interval in absolute time.
    /// 
    /// # Xml example
    /// 
    /// Consider a text animation where the words appear letter by letter every 10 seconds. The <tmAbs>
    /// element should be used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="5" >
    ///     <p:stCondLst> … </p:stCondLst>
    ///     <p:iterate type="lt">
    ///       <p:tmAbs val="10000"/>
    ///     </p:iterate>
    ///     <p:childTnLst> … </p:childTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    Absolute(TLTime),
    /// This element describes the duration of the iteration interval in a percentage of time.
    /// 
    /// # Xml example
    /// 
    /// Consider a text animation where the words appear letter by letter every 10th of the animation
    /// duration. The <tmPct> element should be used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="5">
    ///     <p:stCondLst> … </p:stCondLst>
    ///     <p:iterate type="lt">
    ///       <p:tmPct val="10%"/>
    ///     </p:iterate>
    ///     <p:childTnLst> … </p:childTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    Percent(msoffice_shared::drawingml::PositivePercentage),
}

impl TLIterateDataChoice {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
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
            }
            "tmPct" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TLIterateDataChoice::Percent(val_attr.parse()?))
            }
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TLIterateDataChoice",
            ))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLIterateData {
    /// This attribute specifies the iteration behavior and applies it to each letter, word or shape
    /// within a container element.
    /// 
    /// Values are by word, by letter, or by element. If there is no text or block elements such as
    /// shapes within the container or a single word, letter, or shape (depending on iterate type)
    /// then no iteration happens and the behavior is applied to the element itself instead.
    /// 
    /// Defaults to IterateType::Element
    pub iterate_type: Option<IterateType>,
    /// This attribute specifies whether to go backwards in the timeline to the previous node.
    /// 
    /// Defaults to false
    pub backwards: Option<bool>,
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
    /// The element specifies an incremental RGB value to add to the color property
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animClr clrSpc="rgb">
    ///   <p:cBhvr>
    ///     <p:cTn id="8" dur="500" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>stroke.color</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:by>
    ///     <p:rgb r="10" g="20" b="30"/>
    ///   </p:by>
    /// </p:animClr>
    /// ```
    Rgb(TLByRgbColorTransform),
    /// This element specifies an incremental HSL (Hue, Saturation, Lightness) value to add to a color animation.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:animClr clrSpc="hsl">
    ///   <p:cBhvr>
    ///     <p:cTn id="8" dur="500" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>stroke.color</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:by>
    ///     <p:hsl h="0" s="0" l="0"/>
    ///   </p:by>
    /// </p:animClr>
    /// ```
    Hsl(TLByHslColorTransform),
}

impl TLByAnimateColorTransform {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "rgb" | "hsl" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "rgb" => Ok(TLByAnimateColorTransform::Rgb(TLByRgbColorTransform::from_xml_element(
                xml_node,
            )?)),
            "hsl" => Ok(TLByAnimateColorTransform::Hsl(TLByHslColorTransform::from_xml_element(
                xml_node,
            )?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TLByAnimateColorTransform",
            ))),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TLByRgbColorTransform {
    /// This attribute specifies a red component luminance as a percentage. Values are in the range [-100%, 100%].
    pub r: msoffice_shared::drawingml::FixedPercentage,
    /// This attribute specifies a green component luminance as a percentage. Values are in the range [-100%, 100%].
    pub g: msoffice_shared::drawingml::FixedPercentage,
    /// This attribute specifies a blue component luminance as a percentage. Values are in the range [-100%, 100%].
    pub b: msoffice_shared::drawingml::FixedPercentage,
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
    /// Specifies hue as an angle. The values range from [0, 360] degrees
    pub h: msoffice_shared::drawingml::Angle,
    /// Specifies a lightness as a percentage. The values are in the range [-100%, 100%].
    pub s: msoffice_shared::drawingml::FixedPercentage,
    /// Specifies a saturation as a percentage. The values are in the range [-100%, 100%].
    pub l: msoffice_shared::drawingml::FixedPercentage,
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

/// This element specifies an instance of a slide master slide. Within a slide master slide are contained all elements
/// that describe the objects and their corresponding formatting for within a presentation slide. Within a slide
/// master slide are two main elements. The common_slide_data element specifies the common slide elements such as shapes and
/// their attached text bodies. Then the text_styles element specifies the formatting for the text within each of these
/// shapes. The other properties within a slide master slide specify other properties for within a presentation slide
/// such as color information, headers and footers, as well as timing and transition information for all corresponding
/// presentation slides.
#[derive(Debug, Clone)]
pub struct SlideMaster {
    /// Specifies whether the corresponding slide layout is deleted when all the slides that follow
    /// that layout are deleted. If this attribute is not specified then a value of false should be
    /// assumed by the generating application. This would mean that the slide would in fact be
    /// deleted if no slides within the presentation were related to it.
    /// 
    /// Defaults to false
    pub preserve: Option<bool>,
    pub common_slide_data: Box<CommonSlideData>,
    /// This element specifies the mapping layer that transforms one color scheme definition to another. Each attribute
    /// represents a color name that can be referenced in this master, and the value is the corresponding color in the
    /// theme.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1"
    /// accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5"
    /// accent6="accent6" hlink="hlink" folHlink="folHlink"/>
    /// ```
    pub color_mapping: Box<msoffice_shared::drawingml::ColorMapping>,
    /// This element specifies the existence of the slide layout identification list. This list is contained within the slide
    /// master and is used to determine which layouts are being used within the slide master file. Each layout within the
    /// list of slide layouts has its own identification number and relationship identifier that uniquely identifies it within
    /// both the presentation document and the particular master slide within which it is used.
    /// 
    /// The SlideLayoutIdListEntry specifies the relationship information for each slide layout that is used within the slide master.
    /// The slide master has relationship identifiers that it uses internally for determining the slide layouts that should be
    /// used. Then, to resolve what these slide layouts should be the sldLayoutId elements in the sldLayoutIdLst are
    /// utilized.
    pub slide_layout_id_list: Option<Vec<SlideLayoutIdListEntry>>,
    /// This element specifies the kind of slide transition that should be used to transition to the current slide from the
    /// previous slide. That is, the transition information is stored on the slide that appears after the transition is
    /// complete.
    pub transition: Option<Box<SlideTransition>>,
    /// This element specifies the timing information for handling all animations and timed events within the
    /// corresponding slide. This information is tracked via time nodes within the timing element. More information on
    /// the specifics of these time nodes and how they are to be defined can be found within the Animation section of
    /// the PresentationML framework.
    pub timing: Option<SlideTiming>,
    /// This element specifies the header and footer information for a slide. Headers and footers consist of
    /// placeholders for text that should be consistent across all slides and slide types, such as a date and time, slide
    /// numbering, and custom header and footer text.
    pub header_footer: Option<HeaderFooter>,
    /// This element specifies the text styles within a slide master. Within this element is the styling information for title
    /// text, the body text and other slide text as well. This element is only for use within the Slide Master and thus sets
    /// the text styles for the corresponding presentation slides.
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
        let mut slide_layout_id_list = None;
        let mut transition = None;
        let mut timing = None;
        let mut header_footer = None;
        let mut text_styles = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cSld" => common_slide_data = Some(Box::new(CommonSlideData::from_xml_element(child_node)?)),
                "clrMap" => {
                    color_mapping = Some(
                        Box::new(msoffice_shared::drawingml::ColorMapping::from_xml_element(child_node)?)
                    )
                }
                "sldLayoutIdLst" => {
                    let mut vec = Vec::new();
                    for slide_layout_id_node in &child_node.child_nodes {
                        vec.push(SlideLayoutIdListEntry::from_xml_element(slide_layout_id_node)?);
                    }
                    slide_layout_id_list = Some(vec);
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

/// This element specifies an instance of a slide layout. The slide layout contains in essence a template slide design
/// that can be applied to any existing slide. When applied to an existing slide all corresponding content should be
/// mapped to the new slide layout.
/// 
/// 
#[derive(Debug, Clone)]
pub struct SlideLayout {
    /// Specifies a name to be used in place of the name attribute within the cSld element. This
    /// is used for layout matching in response to layout changes and template applications.
    /// 
    /// Defaults to ""
    pub matching_name: Option<String>,
    /// Specifies the slide layout type that is used by this slide.
    /// 
    /// Defaults to SlideLayoutType::Custom
    pub slide_layout_type: Option<SlideLayoutType>,
    /// Specifies whether the corresponding slide layout is deleted when all the slides that follow
    /// that layout are deleted. If this attribute is not specified then a value of false should be
    /// assumed by the generating application. This would mean that the slide would in fact be
    /// deleted if no slides within the presentation were related to it.
    /// 
    /// Defaults to false
    pub preserve: Option<bool>,
    /// Specifies if the corresponding object has been drawn by the user and should thus not be
    /// deleted. This allows for the flagging of slides that contain user drawn data.
    pub is_user_drawn: Option<bool>,
    /// Specifies if shapes on the master slide should be shown on slides or not.
    /// 
    /// Defaults to true
    pub show_master_shapes: Option<bool>,
    /// Specifies whether or not to display animations on placeholders from the master slide.
    /// 
    /// Defaults to true
    pub show_master_placeholder_animations: Option<bool>,
    pub common_slide_data: Box<CommonSlideData>,
    pub color_mapping_override: Option<msoffice_shared::drawingml::ColorMappingOverride>,
    /// This element specifies the kind of slide transition that should be used to transition to the current slide from the
    /// previous slide. That is, the transition information is stored on the slide that appears after the transition is
    /// complete.
    pub transition: Option<Box<SlideTransition>>,
    /// This element specifies the timing information for handling all animations and timed events within the
    /// corresponding slide. This information is tracked via time nodes within the timing element. More information on
    /// the specifics of these time nodes and how they are to be defined can be found within the Animation section of
    /// the PresentationML framework.
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
                        Some(msoffice_shared::drawingml::ColorMappingOverride::from_xml_element(clr_map_node)?);
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

/// This element specifies a slide within a slide list. The slide list is used to specify an ordering of slides.
/// 
/// # Xml example
/// 
/// ```xml
/// <p:custShowLst>
///   <p:custShow name="Custom Show 1" id="0">
///     <p:sldLst>
///       <p:sld r:id="rId4"/>
///       <p:sld r:id="rId3"/>
///       <p:sld r:id="rId2"/>
///       <p:sld r:id="rId5"/>
///     </p:sldLst>
///   </p:custShow>
/// </p:custShowLst>
/// ```
/// In the above example the order specified to present the slides is slide 4, then 3, 2 and finally 5.
#[derive(Debug, Clone)]
pub struct Slide {
    /// Specifies that the current slide should be shown in slide show. If this attribute is omitted
    /// then a value of true is assumed.
    /// 
    /// Defaults to true
    pub show: Option<bool>,
    /// Specifies if shapes on the master slide should be shown on slides or not.
    /// 
    /// Defaults to true
    pub show_master_shapes: Option<bool>,
    /// Specifies whether or not to display animations on placeholders from the master slide.
    /// 
    /// Defaults to true
    pub show_master_placeholder_animations: Option<bool>,
    pub common_slide_data: Box<CommonSlideData>,
    /// This element provides a mechanism with which to override the color schemes listed within the
    /// SlideMaster::color_mapping element.
    /// If the ColorMappingOverride::UseMaster element is present, the color scheme defined by the master is used.
    /// If the ColorMappingOverride::Override element is present, it defines a new color scheme specific to the
    /// parent notes slide, presentation slide, or slide layout.
    pub color_mapping_override: Option<msoffice_shared::drawingml::ColorMappingOverride>,
    /// This element specifies the kind of slide transition that should be used to transition to the current slide from the
    /// previous slide. That is, the transition information is stored on the slide that appears after the transition is
    /// complete.
    pub transition: Option<Box<SlideTransition>>,
    /// This element specifies the timing information for handling all animations and timed events within the
    /// corresponding slide. This information is tracked via time nodes within the timing element. More information on
    /// the specifics of these time nodes and how they are to be defined can be found within the Animation section of
    /// the PresentationML framework.
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
                        Some(msoffice_shared::drawingml::ColorMappingOverride::from_xml_element(clr_map_node)?);
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

#[derive(Debug, Clone)]
pub struct EmbeddedFontListEntry {
    /// This element specifies specific properties describing an embedded font. Once specified, this font is available
    /// for use within the presentation.
    /// Within a font specification there can be regular, bold, italic and boldItalic versions of the font specified.
    /// The actual font data for each of these is referenced using a relationships file that contains links to all
    /// available fonts.
    /// This font data contains font information for each of the characters to be made available in each version of 
    /// the font.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:regular r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    /// 
    /// # Font Substitution Logic
    /// 
    /// If the specified font is not available on a system being used for rendering, then the attributes of this
    /// element are to be utilized in selecting an alternate font.
    /// 
    /// # Note
    /// 
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine which
    /// characters are to be stored in the corresponding font data files.
    pub font: msoffice_shared::drawingml::TextFont,
    /// This element specifies a regular embedded font that is linked to a parent typeface. Once specified, this regular
    /// version of the given typeface name is available for use within the presentation. The actual font data is
    /// referenced using a relationships file that contains links to all fonts available. This font data contains font
    /// information for each of the characters to be made available.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:regular r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    /// 
    /// # Note
    /// 
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine which
    /// characters are to be stored in the corresponding font data files.
    pub regular: Option<RelationshipId>,
    /// This element specifies a bold embedded font that is linked to a parent typeface. Once specified, this bold 
    /// version of the given typeface name is available for use within the presentation. The actual font data is 
    /// referenced using a relationships file that contains links to all fonts available. This font data contains font 
    /// information for each of the characters to be made available.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:bold r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    /// 
    /// # Note
    /// 
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine
    /// which characters are to be stored in the corresponding font data files.
    pub bold: Option<RelationshipId>,
    /// This element specifies an italic embedded font that is linked to a parent typeface. Once specified, this italic
    /// version of the given typeface name is available for use within the presentation. The actual font data is
    /// referenced using a relationships file that contains links to all fonts available. This font data contains font
    /// information for each of the characters to be made available.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:italic r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    /// 
    /// # Note
    /// 
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine which
    /// characters are to be stored in the corresponding font data files.
    pub italic: Option<RelationshipId>,
    /// This element specifies a bold italic embedded font that is linked to a parent typeface. Once specified, this
    /// bold italic version of the given typeface name is available for use within the presentation. The actual font
    /// data is referenced using a relationships file that contains links to all fonts available. This font data 
    /// contains font information for each of the characters to be made available.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:boldItalic r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    /// 
    /// # Note
    /// 
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine
    /// which characters are to be stored in the corresponding font data files.
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
                "font" => font = Some(msoffice_shared::drawingml::TextFont::from_xml_element(child_node)?),
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
    /// Specifies a name for the custom show.
    pub name: Name,
    /// Specifies the identification number for this custom show. This should be unique among
    /// all the custom shows within the corresponding presentation.
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
    /// Specifies whether all pictures in the photo album are to be displayed as black and white.
    /// 
    /// Defaults to false
    pub black_and_white: Option<bool>,
    /// Specifies whether to show captions for pictures in the photo album. Captions are text
    /// boxes grouped with each image, with the group set to not allow ungrouping.
    /// 
    /// Defaults to false
    pub show_captions: Option<bool>,
    /// Specifies the layout that is to be used to arrange the pictures in the photo album on
    /// individual slides.
    /// 
    /// Defaults to PhotoAlbumLayout::FitToSlide
    pub layout: Option<PhotoAlbumLayout>,
    /// Specifies the frame type that is to be used on all the pictures in the photo album.
    /// 
    /// Defaults to PhotoAlbumFrameShape::FrameStyle1
    pub frame: Option<PhotoAlbumFrameShape>,
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
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Default, Debug, Clone)]
pub struct HeaderFooter {
    /// Specifies whether the slide number placeholder is enabled. If this attribute is not
    /// specified, a value of true should be assumed by the generating application.
    pub slide_number_enabled: Option<bool>, // true
    /// Specifies whether the Header placeholder is enabled for this master. If this attribute is
    /// not specified, a value of true should be assumed by the generating application.
    pub header_enabled: Option<bool>,
    /// Specifies whether the Footer placeholder is enabled for this master. If this attribute is not
    /// specified, a value of true should be assumed by the generating application.
    pub footer_enabled: Option<bool>,
    /// Specifies whether the Date/Time placeholder is enabled for this master. If this attribute is
    /// not specified, a value of true should be assumed by the generating application.
    pub date_time_enabled: Option<bool>,
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
    /// Specifies the corresponding East Asian language that these settings apply to.
    pub language: Option<String>,
    /// Specifies the characters that cannot start a line of text.
    pub invalid_start_chars: String,
    /// Specifies the characters that cannot end a line of text.
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

        let invalid_start_chars =
            invalid_start_chars.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "invalStChars"))?;
        let invalid_end_chars =
            invalid_end_chars.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "invalEndChars"))?;

        Ok(Self {
            language,
            invalid_start_chars,
            invalid_end_chars,
        })
    }
}

#[derive(Default, Debug, Clone)]
pub struct ModifyVerifier {
    /// Specifies the specific cryptographic hashing algorithm which shall be used along with the
    /// salt attribute and input password in order to compute the hash value.
    /// 
    /// The following values are reserved:
    /// * MD2: Specifies that the MD2 algorithm, as defined by RFC 1319, shall be used.
    /// __It is recommended that applications should avoid using this algorithm to store new hash values, due to
    /// publically known breaks.__
    /// 
    /// * MD4: Specifies that the MD4 algorithm, as defined by RFC 1320, shall be used.
    /// __It is recommended that applications should avoid using this algorithm to store new hash values, due to
    /// publically known breaks.__
    /// 
    /// * MD5: Specifies that the MD5 algorithm, as defined by RFC 1321, shall be used.
    /// __It is recommended that applications should avoid using this algorithm to store new hash values, due to
    /// publically known breaks.__
    /// 
    /// * RIPEMD-128: Specifies that the RIPEMD-128 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// __It is recommended that applications should avoid using this algorithm to store new hash values, due to
    /// publically known breaks.__
    /// 
    /// * RIPEMD-160: Specifies that the RIPEMD-160 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// * SHA-1: Specifies that the SHA-1 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// * SHA-256: Specifies that the SHA-256 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// * SHA-384: Specifies that the SHA-384 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// * SHA-512: Specifies that the SHA-512 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// * WHIRLPOOL: Specifies that the WHIRLPOOL algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// 
    /// # Xml example
    /// 
    /// Consider an Office Open XML document with the following information stored in one of its protection elements:
    /// ```xml
    /// < ... algorithmName="SHA-1" hashValue="9oN7nWkCAyEZib1RomSJTjmPpCY=" />
    /// ```
    /// The algorithm_name attribute value of “SHA-1” specifies that the SHA-1 hashing algorithm must be used to
    /// generate a hash from the user-defined password.
    pub algorithm_name: Option<String>,
    /// Specifies the hash value for the password required to edit this chartsheet. This value shall
    /// be compared with the resulting hash value after hashing the user-supplied password
    /// using the algorithm specified by the preceding attributes and parent XML element, and if
    /// the two values match, the protection shall no longer be enforced.
    /// 
    /// If this value is omitted, then the reservationPassword attribute shall contain the
    /// password hash for the workbook.
    /// 
    /// # Xml example
    /// 
    /// Consider an Office Open XML document with the following information stored in one of its protection elements:
    /// ```xml
    /// <... algorithmName="SHA-1" hashValue="9oN7nWkCAyEZib1RomSJTjmPpCY=" />
    /// ```
    /// The hashValue attribute value of 9oN7nWkCAyEZib1RomSJTjmPpCY= specifies that the
    /// user-supplied password must be hashed using the pre-processing defined by the parent
    /// element (if any) followed by the SHA-1 algorithm (specified via the algorithmName
    /// attribute value of SHA-1 ) and that the resulting has value must be
    /// 9oN7nWkCAyEZib1RomSJTjmPpCY= for the protection to be disabled.
    pub hash_value: Option<String>,
    /// Specifies the salt which was prepended to the user-supplied password before it was
    /// hashed using the hashing algorithm defined by the preceding attribute values to generate
    /// the hashValue attribute, and which shall also be prepended to the user-supplied
    /// password before attempting to generate a hash value for comparison. A salt is a random
    /// string which is added to a user-supplied password before it is hashed in order to prevent
    /// a malicious party from pre-calculating all possible password/hash combinations and
    /// simply using those pre-calculated values (often referred to as a "dictionary attack").
    /// 
    /// If this attribute is omitted, then no salt shall be prepended to the user-supplied password
    /// before it is hashed for comparison with the stored hash value.
    /// 
    /// # Xml example
    /// 
    /// Consider an Office Open XML document with the following information stored in one of its protection elements:
    /// ```xml
    /// <... saltValue="ZUdHa+D8F/OAKP3I7ssUnQ==" hashValue="9oN7nWkCAyEZib1RomSJTjmPpCY=" />
    /// ```
    /// 
    /// The saltValue attribute value of ZUdHa+D8F/OAKP3I7ssUnQ== specifies that the user-
    /// supplied password must have this value prepended before it is run through the specified
    /// hashing algorithm to generate a resulting hash value for comparison.
    pub salt_value: Option<String>,
    /// Specifies the number of times the hashing function shall be iteratively run (runs using
    /// each iteration's result plus a 4 byte value (0-based, little endian) containing the number
    /// of the iteration as the input for the next iteration) when attempting to compare a user-
    /// supplied password with the value stored in the hashValue attribute.
    /// 
    /// # Rationale
    /// 
    /// Running the algorithm many times increases the cost of exhaustive search
    /// attacks correspondingly. Storing this value allows for the number of iterations to be
    /// increased over time to accommodate faster hardware (and hence the ability to run more
    /// iterations in less time).
    /// 
    /// # Xml example
    /// 
    /// Consider an Office Open XML document with the following information stored in one of its protection elements:
    /// ```xml
    /// <... spinCount="100000" hashValue="9oN7nWkCAyEZib1RomSJTjmPpCY=" />
    /// ```
    /// The spinCount attribute value of 100000 specifies that the hashing function must be run
    /// one hundred thousand times to generate a hash value for comparison with the
    /// hashValue attribute.
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

/// This element specifies within it fundamental presentation-wide properties.
/// 
/// # Xml example
/// 
/// Consider the following presentation with a single slide master and two slides. In addition to these
/// commonly used elements there can also be the specification of other properties such as slide size, notes size and
/// default text styles.
/// ```xml
/// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="...">
///  <p:sldMasterIdLst>
///    <p:sldMasterId id="2147483648" r:id="rId1"/>
///  </p:sldMasterIdLst>
///  <p:sldIdLst>
///    <p:sldId id="256" r:id="rId3"/>
///    <p:sldId id="257" r:id="rId4"/>
///  </p:sldIdLst>
///  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
///  <p:notesSz cx="6858000" cy="9144000"/>
///  <p:defaultTextStyle>
///  ...
///  </p:defaultTextStyle>
/// </p:presentation>
/// ```
#[derive(Default, Debug, Clone)]
pub struct Presentation {
    pub server_zoom: Option<msoffice_shared::drawingml::Percentage>, // 50%
    pub first_slide_num: Option<i32>,                      // 1
    pub show_special_pls_on_title_slide: Option<bool>,     // true
    pub rtl: Option<bool>,                                 // false
    pub remove_personal_info_on_save: Option<bool>,        // false
    /// Specifies whether the generating application is to be in a compatibility mode which
    /// serves to inform the user of any loss of content or functionality when working with older
    /// formats.
    /// 
    /// Defaults to false
    pub compatibility_mode: Option<bool>,
    pub strict_first_and_last_chars: Option<bool>,         // true
    pub embed_true_type_fonts: Option<bool>,               // false
    pub save_subset_fonts: Option<bool>,                   // false
    /// Specifies whether the generating application should automatically compress all pictures
    /// for this presentation.
    /// 
    /// Defaults to true
    pub auto_compress_pictures: Option<bool>,
    /// Specifies a seed for generating bookmark IDs to ensure IDs remain unique across the
    /// document. This value specifies the number to be used as the ID for the next new
    /// bookmark created.
    /// 
    /// Defaults to 1
    pub bookmark_id_seed: Option<BookmarkIdSeed>,
    /// Specifies the conformance class (§2.1) to which the PresentationML document conforms.
    /// 
    /// Defaults to ConformanceClass::Transitional
    pub conformance: Option<ConformanceClass>,
    /// This element specifies a list of identification information for the slide master slides that are available 
    /// within the corresponding presentation.
    /// A slide master is a slide that is specifically designed to be a template for all related child layout slides.
    /// 
    /// The SlideMasterIdListEntry specifies a slide master that is available within the corresponding presentation.
    /// A slide master is a slide that is specifically designed to be a template for all related child layout slides.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:sldMasterIdLst>
    ///     <p:sldMasterId id="2147483648" r:id="rId1"/>
    ///   </p:sldMasterIdLst>
    ///   ...
    /// </p:presentation>
    pub slide_master_id_list: Vec<SlideMasterIdListEntry>,
    /// The specifies a list of identification information for the notes master slides that are available within the
    /// corresponding presentation. A notes master is a slide that is specifically designed for the printing of the slide
    /// along with any attached notes.
    /// 
    /// The NotesMasterIdListEntry specifies a notes master that is available within the corresponding presentation.
    /// A notes master is a slide that is specifically designed for the printing of the slide along with any 
    /// attached notes.
    /// 
    /// # Xml example
    /// 
    /// Consider the following specification of a notes master within a presentation
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:notesMasterIdLst>
    ///     <p:notesMasterId r:id="rId8"/>
    ///   </p:notesMasterIdLst>
    ///   ...
    /// </p:presentation>
    /// ```xml
    /// 
    /// # Note
    /// 
    /// Even though the reference documentation states that this element is a list, the Xml schema states that it has
    /// only a single element.
    pub notes_master_id: Option<NotesMasterIdListEntry>,
    /// This element specifies a list of identification information for the handout master slides that are available
    /// within the corresponding presentation.
    /// A handout master is a slide that is specifically designed for printing as a handout.
    /// 
    /// The HandoutMasterIdListEntry specifies a handout master that is available within the corresponding presentation. A handout
    /// master is a slide that is specifically designed for printing as a handout.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:handoutMasterIdLst>
    ///     <p:handoutMasterId r:id="rId8"/>
    ///   </p:handoutMasterIdLst>
    ///   ...
    /// </p:presentation>
    /// ```
    /// 
    /// # Note
    /// 
    /// Even though the reference documentation states that this element is a list, the Xml schema states that it has
    /// only a single element.
    pub handout_master_id: Option<HandoutMasterIdListEntry>,
    /// This element specifies a list of identification information for the slides that are available within the
    /// corresponding presentation. A slide contains the information that is specific to a single slide such as slide-
    /// specific shape and text information.
    /// 
    /// The SlideIdListEntry specifies a list of presentation slides. A presentation slide contains the information
    /// that is specific to a single slide such as slide-specific shape and text information.
    pub slide_id_list: Vec<SlideIdListEntry>,
    /// This element specifies the size of the presentation slide surface. Objects within a presentation slide can be
    /// specified outside these extents, but this is the size of background surface that is shown when the slide is
    /// presented or printed.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
    ///   ...
    /// </p:presentation>
    pub slide_size: Option<SlideSize>,
    /// This element specifies the size of slide surface used for notes slides and handout slides. Objects within a
    /// notes slide can be specified outside these extents, but the notes slide has a background surface of the 
    /// specified size when presented or printed.
    /// This element is intended to specify the region to which content is fitted in any special format of printout 
    /// the application might choose to generate, such as an outline handout.
    /// 
    /// # Xml example
    /// 
    /// Consider the following specifying of the size of a notes slide.
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:notesSz cx="9144000" cy="6858000"/>
    ///   ...
    /// </p:presentation>
    /// ```
    pub notes_size: Option<msoffice_shared::drawingml::PositiveSize2D>,
    /// This element specifies that references to smart tags exist within this document.
    /// To denote the location of smart tags on individual runs of text, there smart tag identifier attributes are 
    /// specified for each run to which a smart tag applies.
    /// These are further specified in the run property attributes within DrawingML.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:presentation>
    ///   ...
    ///   <p:smartTags r:id="rId1"/>
    /// </p:presentation>
    /// ```
    /// 
    /// The presence of the smartTags element specifies that there is smart tag information within the PresentationML
    /// package. Individual runs are then inspected for the value of the smtId attribute to determine where smart tags
    /// might apply, for example:
    /// 
    /// ```xml
    /// <p:txBody>
    ///  <a:bodyPr/>
    ///  <a:lstStyle/>
    ///  <a:p>
    ///    <a:r>
    ///      <a:rPr lang="en-US" dirty="0" smtId="1"/>
    ///      <a:t>CNTS</a:t>
    ///    </a:r>
    ///    <a:endParaRPr lang="en-US" dirty="0"/>
    ///  </a:p>
    /// </p:txBody>
    /// ```
    /// 
    /// In the sample above there is a smart tag identifier of 1 specified for this run of text to denote that the text
    /// should be inspected for smart tag information.
    /// 
    /// # Note
    /// 
    /// For a complete definition of smart tags, which are semantically identical throughout Office Open XML,
    /// see §17.5.1.
    pub smart_tags: Option<RelationshipId>,
    /// This element specifies a list of fonts that are embedded within the corresponding presentation. The font data
    /// for these fonts is stored alongside the other document parts within the document container. The actual font 
    /// data is referenced within the EmbeddedFontListEntry element.
    /// 
    /// The EmbeddedFontListEntry element specifies an embedded font. Once specified, this font is available for use within the 
    /// presentation.
    /// Within a font specification there can be regular, bold, italic and boldItalic versions of the font specified.
    /// The actual font data for each of these is referenced using a relationships file that contains links to all 
    /// available fonts.
    /// This font data contains font information for each of the characters to be made available in each version of the
    /// font.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:regular r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    /// 
    /// # Note
    /// 
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine
    /// which characters are to be stored in the corresponding font data files.
    pub embedded_font_list: Vec<Box<EmbeddedFontListEntry>>,
    /// This element specifies a list of all custom shows that are available within the corresponding presentation.
    /// A custom show is a defined slide sequence that allows for the displaying of the slides with the presentation in
    /// any arbitrary order.
    pub custom_show_list: Vec<CustomShow>,
    /// This element specifies that the corresponding presentation contains a photo album. A photo album specifies a
    /// list of images within the presentation that spread across one or more slides, all of which share a consistent
    /// layout. Each image in the album is formatted with a consistent style. This functionality enables the application
    /// to manage all of the images together and modify their ordering, layout, and formatting as a set.
    /// 
    /// This element does not enforce the specified properties on individual photo album images; rather, it specifies
    /// common settings that should be applied by default to all photo album images and their containing slides.
    /// Images that are part of the photo album are identified by the presence of the isPhoto element in the definition
    /// of the picture.
    /// 
    /// # Xml example
    /// 
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:photoAlbum bw="1" layout="2pic"/>
    ///   ...
    /// </p:presentation>
    pub photo_album: Option<PhotoAlbum>,
    /// This element allows for the specifying of customer defined data within the PresentationML framework.
    /// References to custom data or tags can be defined within this list.
    /// 
    /// The elements of this list specifies customer data which allows for the specifying and persistence of customer
    /// specific data within the presentation.
    pub customer_data_list: Option<CustomerDataList>,
    /// This element specifies the presentation-wide kinsoku settings that define the line breaking behaviour of East
    /// Asian text within the corresponding presentation.
    pub kinsoku: Option<Box<Kinsoku>>,
    /// This element specifies the default text styles that are to be used within the presentation.
    /// The text style defined here can be referenced when inserting a new slide if that slide is not associated with a
    /// master slide or if no styling information has been otherwise specified for the text within the
    /// presentation slide.
    pub default_text_style: Option<Box<msoffice_shared::drawingml::TextListStyle>>,
    /// This element specifies the write protection settings which have been applied to a PresentationML document.
    /// Write protection refers to a mode in which the document's contents should not be modified, and the document
    /// should not be resaved using the same file name.
    /// 
    /// When present, the application shall require a password to enable modifications to the document. If the
    /// supplied password does not match the hash value in this attribute, then write protection shall be enabled.
    /// If this element is omitted, then no write protection shall be applied to the current document.
    /// Since this protection does not encrypt the document, malicious applications might circumvent its use.
    /// 
    /// The password supplied to the algorithm is to be a UTF-16LE encoded string; strings longer than 510 octets are
    /// truncated to 510 octets. If there is a leading BOM character (U+FEFF) in the encoded password it is removed
    /// before hash calculation. The attributes of this element specify the algorithm to be used to verify the password
    /// provided by the user.
    /// 
    /// # Xml example
    /// 
    /// Consider a PresentationML document that can only be opened in a write protected state unless a
    /// password is provided, in which case the file would be opened in an editable state. This requirement would be
    /// specified using the following PresentationML:
    /// ```xml
    /// <p:modifyVerifier p:algorithmName="SHA-512" ...
    /// p:hashValue="9oN7nWkCAyEZib1RomSJTjmPpCY=" ... />
    /// ```
    /// ...In order for the hosting application to enable edits to the document, the hosting application would have to
    /// be provided with a password that the hosting application would then hash using the algorithm specified by the
    /// algorithm attributes and compare to the value of the hashValue attribute ( 9oN7nWkCAyEZib1RomSJTjmPpCY= ).
    /// If the two values matched, the file would be opened in an editable state.
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
                        instance
                            .slide_master_id_list
                            .push(SlideMasterIdListEntry::from_xml_element(slide_master_id_node)?);
                    }
                }
                "notesMasterIdLst" => {
                    if let Some(notes_master_id_node) = child_node.child_nodes.get(0) {
                        instance.notes_master_id =
                            Some(NotesMasterIdListEntry::from_xml_element(notes_master_id_node)?);
                    }
                }
                "handoutMasterIdLst" => {
                    if let Some(handout_master_id_node) = child_node.child_nodes.get(0) {
                        instance.handout_master_id =
                            Some(HandoutMasterIdListEntry::from_xml_element(handout_master_id_node)?);
                    }
                }
                "sldIdLst" => {
                    for slide_id_node in &child_node.child_nodes {
                        instance
                            .slide_id_list
                            .push(SlideIdListEntry::from_xml_element(slide_id_node)?);
                    }
                }
                "sldSz" => instance.slide_size = Some(SlideSize::from_xml_element(child_node)?),
                "notesSz" => {
                    instance.notes_size = Some(msoffice_shared::drawingml::PositiveSize2D::from_xml_element(child_node)?)
                }
                "smartTags" => {
                    let r_id_attr = child_node
                        .attribute("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?;
                    instance.smart_tags = Some(r_id_attr.clone());
                }
                "embeddedFontLst" => {
                    for embedded_font_node in &child_node.child_nodes {
                        instance
                            .embedded_font_list
                            .push(Box::new(EmbeddedFontListEntry::from_xml_element(embedded_font_node)?));
                    }
                }
                "custShowLst" => {
                    for custom_show_node in &child_node.child_nodes {
                        instance
                            .custom_show_list
                            .push(CustomShow::from_xml_element(custom_show_node)?);
                    }
                }
                "photoAlbum" => instance.photo_album = Some(PhotoAlbum::from_xml_element(child_node)?),
                "custDataLst" => instance.customer_data_list = Some(CustomerDataList::from_xml_element(child_node)?),
                "kinsoku" => instance.kinsoku = Some(Box::new(Kinsoku::from_xml_element(child_node)?)),
                "defaultTextStyle" => {
                    instance.default_text_style =
                        Some(Box::new(msoffice_shared::drawingml::TextListStyle::from_xml_element(child_node)?))
                }
                "modifyVerifier" => {
                    instance.modify_verifier = Some(Box::new(ModifyVerifier::from_xml_element(child_node)?))
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}
