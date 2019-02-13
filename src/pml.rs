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
