use std::io::{ Read, Seek };
use std::str;

use drawingml;
use relationship;

use zip;
use quick_xml;

pub type SlideId = u32; // TODO: 256 <= n <= 2147483648
pub type SlideLayoutId = u32; // TODO: 2147483648 <= n
pub type SlideMasterId = u32; // TODO: 2147483648 <= n
pub type Index = u32;
pub type TLTimeNodeId = u32;
pub type BookmarkIdSeed = u32; // TODO: 1 <= n <= 2147483648
pub type SlideSizeCoordinate = drawingml::PositiveCoordinate32; // TODO: 914400 <= n <= 51206400
pub type Name = String;
pub type TLSubShapeId = drawingml::ShapeId;

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
    pub fill: drawingml::FillProperties,
    pub effect: Option<drawingml::EffectProperties>,
    pub shade_to_title: Option<bool>, // false
}

pub enum BackgroundGroup {
    Properties(BackgroundProperties),
    Reference(drawingml::StyleMatrixReference),
}

pub struct Background {
    pub background: BackgroundGroup,
    pub black_and_white_mode: Option<drawingml::BlackWhiteMode>, // white
}

pub struct Placeholder {
    pub placeholder_type: Option<PlaceholderType>, // obj
    pub orientation: Option<Direction>, // horz
    pub size: Option<PlaceholderSize>, // full
    pub index: Option<u32>, // 0
    pub has_custom_prompt: Option<bool>, // false
}

pub struct ApplicationNonVisualDrawingProps {
    pub is_photo: Option<bool>, // false
    pub is_user_drawn: Option<bool>, // false
    pub placeholder: Option<Placeholder>,
    pub media: Option<drawingml::Media>,
    //pub customer_data_list: Option<CustomerDataList>,

}

pub enum ShapeGroup {
    Shape(Shape),
    GroupShape(GroupShape),
    GraphicFrame(GraphicalObjectFrame),
    Connector(Connector),
    Picture(Picture),
    ContentPart(relationship::RelationshipId),
}

pub struct Shape {
    pub non_visual_props: ShapeNonVisual,
    pub shape_props: drawingml::ShapeProperties,
    pub style: Option<drawingml::ShapeStyle>,
    pub text_body: Option<drawingml::TextBody>,
    pub use_bg_fill: Option<bool>, // false
}

pub struct ShapeNonVisual {
    pub drawing_props: drawingml::NonVisualDrawingProps,
    pub shape_drawing_props: drawingml::NonVisualDrawingShapeProps,
    pub app_props: ApplicationNonVisualDrawingProps,
}

pub struct GroupShape {
    pub non_visual_props: GroupShapeNonVisual,
    pub group_shape_props: drawingml::GroupShapeProperties,
    pub shape_array: Vec<ShapeGroup>,
}

pub struct GroupShapeNonVisual {
    pub drawing_props: drawingml::NonVisualDrawingProps,
    pub group_drawing_props: drawingml::NonVisualGroupDrawingShapeProps,
    pub app_props: ApplicationNonVisualDrawingProps,
}

pub struct GraphicalObjectFrame {
    pub non_visual_props: GraphicalObjectFrameNonVisual,
    pub transform: drawingml::Transform2D,
    pub graphic: drawingml::GraphicalObject,
    pub black_white_mode: Option<drawingml::BlackWhiteMode>,
}

pub struct GraphicalObjectFrameNonVisual {
    pub drawing_props: drawingml::NonVisualDrawingProps,
    pub graphic_frame_props: drawingml::NonVisualGraphicFrameProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

pub struct Connector {
    pub non_visual_props: ConnectorNonVisual,
    pub shape_props: drawingml::ShapeProperties,
    pub shape_style: Option<drawingml::ShapeStyle>,
}

pub struct ConnectorNonVisual {
    pub drawing_props: drawingml::NonVisualDrawingProps,
    pub connector_props: drawingml::NonVisualConnectorProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

pub struct Picture {
    pub non_visual_props: PictureNonVisual,
    pub blip_fill: drawingml::BlipFillProperties,
    pub shape_propers: drawingml::ShapeProperties,
    pub shape_style: Option<drawingml::ShapeStyle>,
}

pub struct PictureNonVisual {
    pub drawing_props: drawingml::NonVisualDrawingProps,
    pub picture_props: drawingml::NonVisualPictureProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

pub struct CommonSlideData {
    pub background: Option<Background>,
    pub shape_tree: GroupShape,
    pub customer_data_list: Option<CustomerDataList>,
    pub control_list: Vec<Control>,
    pub name: Option<String>,
}

pub struct CustomerDataList {
    pub customer_data_list: Vec<relationship::RelationshipId>,
    pub tags: Option<relationship::RelationshipId>,
}

pub struct Control {
    pub picture: Option<Picture>,
    pub ole_attributes: OleAttributes,
}

pub struct OleAttributes {
    pub shape_id: Option<drawingml::ShapeId>,
    pub name: Option<String>, // ""
    pub show_as_icon: Option<bool>, // false
    pub id: Option<relationship::RelationshipId>,
    pub image_width: Option<drawingml::PositiveCoordinate32>,
    pub image_height: Option<drawingml::PositiveCoordinate32>,
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

pub struct SlideLayoutIdListEntry {
    pub id: Option<SlideLayoutId>,
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

pub struct SlideMasterTextStyles {
    pub title_styles: Option<drawingml::TextListStyle>,
    pub body_styles: Option<drawingml::TextListStyle>,
    pub other_styles: Option<drawingml::TextListStyle>,
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
    pub sound_file: drawingml::EmbeddedWAVAudioFile,
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
    pub x: drawingml::Percentage,
    pub y: drawingml::Percentage,
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
    pub shape_id: drawingml::DrawingElementId,
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
    BuildSubElements(drawingml::AnimationGraphicalObjectBuildProperties),
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
    pub from: Option<drawingml::Color>,
    pub to: Option<drawingml::Color>,
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
    pub rotate_angle: Option<drawingml::Angle>,
    pub points_types: Option<String>,
}

pub struct TLAnimateRotationBehavior {
    pub common_behavior_data: TLCommonBehaviorData,
    pub by: Option<drawingml::Angle>,
    pub from: Option<drawingml::Angle>,
    pub to: Option<drawingml::Angle>,
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
    Percentage(drawingml::PositiveFixedPercentage),
    Indefinite,
}

pub enum TLAnimVariant {
    Bool(bool),
    Int(i32),
    Float(f32),
    String(String),
    Color(drawingml::Color),
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
    pub volume: Option<drawingml::PositiveFixedPercentage>, // 50000
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
    SoundTarget(drawingml::EmbeddedWAVAudioFile),
    ShapeTarget(TLShapeTargetElement),
    InkTarget(TLSubShapeId),
}

pub struct TLShapeTargetElement {
    pub target: Option<TLShapeTargetElementGroup>,
    pub shape_id: drawingml::DrawingElementId,
}

pub enum TLShapeTargetElementGroup {
    Background,
    SubShape(TLSubShapeId),
    OleChartElement(TLOleChartTargetElement),
    TextElement(Option<TLTextTargetElement>),
    GraphicElement(drawingml::AnimationElementChoice),
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
    pub speed: Option<drawingml::Percentage>, // 100000
    pub acceleration: Option<drawingml::PositiveFixedPercentage>, // 0
    pub deceleration: Option<drawingml::PositiveFixedPercentage>, // 0
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
    Percent(drawingml::PositivePercentage),
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
    pub r: drawingml::FixedPercentage,
    pub g: drawingml::FixedPercentage,
    pub b: drawingml::FixedPercentage,
}

pub struct TLByHslColorTransform {
    pub h: drawingml::Angle,
    pub s: drawingml::FixedPercentage,
    pub l: drawingml::FixedPercentage,
}

pub struct TopLevelSlideData {
    pub color_mapping: drawingml::ColorMapping,
}

pub struct ChildSlideData {
    pub color_mapping_override: Option<drawingml::ColorMappingOverride>,
    pub show_master_shapes: Option<bool>, // true
    pub show_master_placholder_animations: Option<bool>, // true
}

pub struct SlideMaster {
    pub common_slide_data: CommonSlideData,
    pub top_level_slide_data: TopLevelSlideData,
    pub slide_layout_id_list: Vec<SlideLayoutIdListEntry>,
    pub transition: Option<SlideTransition>,
    pub timing: Option<SlideTiming>,
    pub header_footer: Option<HeaderFooter>,
    pub text_styles: Option<SlideMasterTextStyles>,
    pub preserve: Option<bool>, // false
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
    pub conformance: Option<ConformanceClass>,
    pub slide_master_id_list: Vec<SlideMasterIdListEntry>,
    pub notes_master_id_list: Vec<NotesMasterIdListEntry>, // length = 1
    pub handout_master_id_list: Vec<HandoutMasterIdListEntry>, // length = 1
    pub slide_id_list: Vec<SlideIdListEntry>,
    pub slide_size: Option<SlideSize>,
    pub notes_size: Option<drawingml::PositiveSize2D>,
    pub smart_tags: Option<relationship::RelationshipId>,
    pub embedded_font_list: Vec<EmbeddedFontListEntry>,
    pub custom_show_list: Vec<CustomShow>,
    pub photo_album: Option<PhotoAlbum>,
    pub customer_data_list: Option<CustomerDataList>,
    pub kinsoku: Option<Kinsoku>,
    pub default_text_style: Option<drawingml::TextListStyle>,
    pub modify_verifier: Option<ModifyVerifier>,
}

impl Presentation {
    fn new() -> Presentation {
        Presentation {
            server_zoom: None,
            first_slide_num: None,
            show_special_pls_on_title_slide: None,
            rtl: None,
            remove_personal_info_on_save: None,
            compatibility_mode: None,
            strict_first_and_last_chars: None,
            embed_true_type_fonts: None,
            save_subset_fonts: None,
            auto_compress_pictures: None,
            bookmark_id_seed: None,
            conformance: None,
            slide_master_id_list: Vec::new(),
            notes_master_id_list: Vec::new(),
            handout_master_id_list: Vec::new(),
            slide_id_list: Vec::new(),
            slide_size: None,
            notes_size: None,
            smart_tags: None,
            embedded_font_list: Vec::new(),
            custom_show_list: Vec::new(),
            photo_album: None,
            customer_data_list: None,
            kinsoku: None,
            default_text_style: None,
            modify_verifier: None,
        }
    }

    pub fn from_xml<R>(zipper: &mut zip::ZipArchive<R>) -> Option<Presentation> where R: Read + Seek {
        let mut presentation_file = match zipper.by_name("ppt/presentation.xml") {
            Ok(f) => f,
            Err(_) => return None,
        };
        
        let mut presentation = Presentation::new();
        let mut xml_string = String::new();

        match presentation_file.read_to_string(&mut xml_string) {
            Ok(_) => {
                let mut buffer = Vec::new();
                let mut xml_reader = quick_xml::Reader::from_str(xml_string.as_str());
                loop {
                    match xml_reader.read_event(&mut buffer) {
                        Ok(quick_xml::events::Event::Start(ref element)) => {
                            /*
                            Presentation *instance = new Presentation();
		for (const std::unique_ptr<MXmlAttribute2> &attr : xmlNode.attributes)
		{
			if (attr->name == mT("serverZoom"))
				instance->serverZoom.SetSome(attr->value.ToInt(50000));
			else if (attr->name == mT("firstSlideNum"))
				instance->firstSlideNum.SetSome(attr->value.ToInt(1));
			else if (attr->name == mT("showSpecialPlsOnTitleSld"))
				instance->showSpecialPlsOnTitleSld.SetSome(attr->value.ToBool());
			else if (attr->name == mT("rtl"))
				instance->rtl.SetSome(attr->value.ToBool());
			else if (attr->name == mT("removePersonalInfoOnSave"))
				instance->removePersonalInfoOnSave.SetSome(attr->value.ToBool());
			else if (attr->name == mT("compatMode"))
				instance->compatMode.SetSome(attr->value.ToBool());
			else if (attr->name == mT("strictFirstAndLastChars"))
				instance->strictFirstAndLastChars.SetSome(attr->value.ToBool());
			else if (attr->name == mT("embedTrueTypeFonts"))
				instance->embedTrueTypeFonts.SetSome(attr->value.ToBool());
			else if (attr->name == mT("saveSubsetFonts"))
				instance->saveSubsetFonts.SetSome(attr->value.ToBool());
			else if (attr->name == mT("autoCompressPictures"))
				instance->autoCompressPictures.SetSome(attr->value.ToBool());
			else if (attr->name == mT("bookmarkIdSeed"))
				instance->bookmarkIdSeed.SetSome(attr->value.ToUInt(1));
			else if (attr->name == mT("conformance"))
				instance->conformance = attr->value;
		}

		for (const MXmlNode2 &childNode : xmlNode)
		{
			if (childNode.name == mT("sldMasterIdLst"))
			{
				for (const MXmlNode2 &sldMasterIdNode : childNode)
					instance->sldMasterIdLst.Emplace(SlideMasterIdListEntry::FromXmlNode(sldMasterIdNode));
			}
			else if (childNode.name == mT("notesMasterIdLst"))
			{
				for (const MXmlNode2 &notesMasterIdNode : childNode)
					instance->notesMasterIdLst.Emplace(NotesMasterIdListEntry::FromXmlNode(notesMasterIdNode));
			}
			else if (childNode.name == mT("handoutMasterIdLst"))
			{
				for (const MXmlNode2 &handoutMasterIdNode : childNode)
					instance->handoutMasterIdLst.Emplace(HandoutMasterIdListEntry::FromXmlNode(handoutMasterIdNode));
			}
			else if (childNode.name == mT("sldIdLst"))
			{
				for (const MXmlNode2 &sldIdNode : childNode)
					instance->sldIdLst.Emplace(SlideIdListEntry::FromXmlNode(sldIdNode));
			}
			else if (childNode.name == mT("sldSz"))
				instance->sldSz.reset(SlideSize::FromXmlNode(childNode));
			else if (childNode.name == mT("notesSz"))
				instance->notesSz.reset(DrawingML::PositiveSize2D::FromXmlNode(childNode));
			//else if (childNode.name == mT("smartTags"))
			//	instance->smartTags.reset(SmartTags::FromXmlNode(childNode));
			else if (childNode.name == mT("embeddedFontLst"))
			{
				for (const MXmlNode2 &embeddedFontNode : childNode)
					instance->embeddedFontLst.Emplace(EmbeddedFontListEntry::FromXmlNode(embeddedFontNode));
			}
			else if (childNode.name == mT("custShowLst"))
			{
				for (const MXmlNode2 &custShowNode : childNode)
					instance->custShowLst.Emplace(CustomShow::FromXmlNode(custShowNode));
			}
			else if (childNode.name == mT("photoAlbum"))
				instance->photoAlbum.reset(PhotoAlbum::FromXmlNode(childNode));
			//else if (childNode.name == mT("custDataLst"))
			//	instance->custDataLst.reset(CustomerDataList::FromXmlNode(childNode));
			//else if (childNode.name == mT("kinsoku"))
			//	instance->kinsoku.reset(Kinsoku::FromXmlNode(childNode));
			else if (childNode.name == mT("defaultTextStyle"))
				instance->defaultTextStyle.reset(DrawingML::TextListStyle::FromXmlNode(childNode));
			//else if (childNode.name == mT("modifyVerifier"))
			//	instance->modifyVerifier.reset(ModifyVerifier::FromXmlNode(childNode));
			//else if (childNode.name == mT("extLst"))
			//	instance->extLst.reset(ExtensionList::FromXmlNode(childNode));
		}
                            */


                            match element.local_name() {
                                b"presentation" => {
                                    for attr in element.attributes() {
                                        if let Ok(a) = attr {
                                            match a.key {
                                                b"serverZoom" => presentation.server_zoom = Some(str::from_utf8(&a.value).unwrap().parse::<i32>().unwrap() as f32 / 100_000.0),
                                                _ => (),

                                            }
                                        }

                                            /*
                                            			if (attr->name == mT("serverZoom"))
				instance->serverZoom.SetSome(attr->value.ToInt(50000));
			else if (attr->name == mT("firstSlideNum"))
				instance->firstSlideNum.SetSome(attr->value.ToInt(1));
			else if (attr->name == mT("showSpecialPlsOnTitleSld"))
				instance->showSpecialPlsOnTitleSld.SetSome(attr->value.ToBool());
			else if (attr->name == mT("rtl"))
				instance->rtl.SetSome(attr->value.ToBool());
			else if (attr->name == mT("removePersonalInfoOnSave"))
				instance->removePersonalInfoOnSave.SetSome(attr->value.ToBool());
			else if (attr->name == mT("compatMode"))
				instance->compatMode.SetSome(attr->value.ToBool());
			else if (attr->name == mT("strictFirstAndLastChars"))
				instance->strictFirstAndLastChars.SetSome(attr->value.ToBool());
			else if (attr->name == mT("embedTrueTypeFonts"))
				instance->embedTrueTypeFonts.SetSome(attr->value.ToBool());
			else if (attr->name == mT("saveSubsetFonts"))
				instance->saveSubsetFonts.SetSome(attr->value.ToBool());
			else if (attr->name == mT("autoCompressPictures"))
				instance->autoCompressPictures.SetSome(attr->value.ToBool());
			else if (attr->name == mT("bookmarkIdSeed"))
				instance->bookmarkIdSeed.SetSome(attr->value.ToUInt(1));
			else if (attr->name == mT("conformance"))
				instance->conformance = attr->value;
                                            */
                                    }
                                }
                            }
                        },
                        Ok(quick_xml::events::Event::Eof) => break,
                        _ => (),
                    }

                    buffer.clear();
                }
            }
            Err(_) => return None,
        }

        Some(presentation)
    }
}