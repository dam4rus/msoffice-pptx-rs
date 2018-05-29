// TODO: This module defines shared types between different OOX file formats. It should be refactored into a different crate, if these types are needed.

type Percentage = f32;
type PositivePercentage = f32; // TODO: 0 <= n < inf
type PositiveFixedPercentage = f32; // TODO: 0 <= n <= 100000
type FixedPercentage = f32; // TODO: -100000 <= n <= 100000
type Color = i32;


/*
typedef MString Guid;

    typedef int Percentage;
    typedef int PositivePercentage;
    typedef int PositiveFixedPercentage;
    typedef int FixedPercentage;
    typedef int PositiveFixedAngle;

    typedef MString HexColorRGB;

    typedef MInt64 Coordinate;
    typedef MInt64 PositiveCoordinate;

    typedef int Coordinate32; 
    typedef Coordinate32 LineWidth;
    typedef Coordinate32 PositiveCoordinate32;

    typedef MUInt DrawingElementId;

    typedef int Angle;
    typedef Angle FixedAngle;
    typedef Angle PositiveFixedAngle;

    typedef MString GeomGuideName;
    typedef MString GeomGuideFormula;

    typedef MUInt StyleMatrixColumnIndex;

    typedef int TextColumnCount;
    typedef Percentage TextFontScalePercent;
    typedef Percentage TextSpacingPercent;
    typedef int TextSpacingPoint;
    typedef Coordinate32 TextMargin;
    typedef Coordinate32 TextIndent;
    typedef int TextIndentLevelType;
    typedef Percentage TextBulletSizePercent;
    typedef int TextFontSize;
    typedef MString TextTypeface;
    typedef MString Panose;
    typedef int TextBulletStartAtNum;
    typedef MString TextLanguageID;
    typedef int TextNonNegativePoint;
    typedef int TextPoint;

    typedef MString ShapeID;

    enum class TileFlipMode;
    enum class RectAlignment;
    enum class BlackWhiteMode;
    enum class TextVerticalType;
    enum class TextAnchoringType;
    enum class TextHorzOverflowType;

    class RelativeRect;
    class Point2D;
    class PositiveSize2D;
    class Transform2D;

    class AlphaModulateFixedEffect;
    class LuminanceEffect;
    class DuotoneEffect;
    class EffectChoice;
    class EffectPropertiesChoice;
    class EffectStyleItem;
    class Blip;

    typedef MArray<std::unique_ptr<EffectStyleItem>> EffectStyleList;

    class GradientStop;
    class LinearShadeProperties;
    class PathShadeProperties;
    class ShadePropertiesChoice;

    class TileInfoProperties;
    class StretchInfoProperties;
    class FillModeProperties;

    class SolidColorFillProperties;
    class GradientFillProperties;
    class BlipFillProperties;
    class PatternFillProperties;
    class FillPropertiesChoice;
    class FillProperties;
    class LineEndProperties;
    class LineProperties;

    typedef MArray<std::unique_ptr<FillPropertiesChoice>> FillStyleList;
    typedef MArray<std::unique_ptr<FillPropertiesChoice>> BackgroundFillStyleList;
    typedef MArray<std::unique_ptr<LineProperties>> LineStyleList;

    class ThemeableFillStyle;

    class NonVisualDrawingProps;
    class NonVisualDrawingShapeProps;
    class NonVisualConnectorProperties;
    class NonVisualGraphicFrameProperties;
    class NonVisualGroupDrawingShapeProps;
    class NonVisualPictureProperties;

    class Connection;

    class ShapeProperties;
    class ShapeStyle;
    class GroupShapeProperties;
    class GraphicalObject;

    class StyleMatrixReference;

    class EmbeddedWAVAudioFile;
    class MediaChoice;

    class FontCollection;
    class FontReference;
    class TextFont;
    class TextBody;
    class TextBodyProperties;
    class TextListStyle;
    class TextSpacingChoice;
    class TextBulletChoice;
    class TextParagraphProperties;
    class TextCharacterProperties;
    class TextParagraph;
    class Hyperlink;

    class CustomGeometry2D;
    class PresetGeometry2D;
    class AdjCoordinate;
    class AdjAngle;

    class Picture;

    class ColorMappingOverride;
    class ColorMapping;
    class ColorSchemeAndMapping;
    class CustomColor;
    class OfficeStyleSheet;

    class Table;
    class TableStyle;

    typedef MArray<std::unique_ptr<ColorSchemeAndMapping>> ColorSchemeList;
    typedef MArray<std::unique_ptr<CustomColor>> CustomColorList;

    class AnimationGraphicalObjectBuildPropertiesChoice;
    class AnimationElementChoice;

    class DashStop;
    typedef MArray<std::unique_ptr<DashStop>> DashStopList;

    class Path2D;
    typedef MArray<std::unique_ptr<Path2D>> Path2DList;

    class ConnectionSite;
    typedef MArray<std::unique_ptr<ConnectionSite>> ConnectionSiteList;

    class GeomGuide;
    typedef MArray<std::unique_ptr<GeomGuide>> GeomGuideList;

    class TextTabStop;
    typedef MArray<std::unique_ptr<TextTabStop>> TextTabStopList;
*/

decl_oox_enum! {
    pub enum TileFlipMode {
        None = "none",
        X = "x",
        Y = "y",
        XY = "xy",
    }
}

decl_oox_enum! {
    pub enum RectAlignment {
        L = "l",
        T = "t",
        R = "r",
        B = "b",
        Tl = "tl",
        Tr = "tr",
        Bl = "bl",
        Br = "br",
        Ctr = "ctr",
    }
}

decl_oox_enum! {
    pub enum PathFillMode {
        None = "none",
        Norm = "norm",
        Lighten = "lighten",
        LightenLess = "lightenLess",
        Darken = "darken",
        DarkenLess = "darkenLess",
    }
}

decl_oox_enum! {
    pub enum ShapeType {
        Line = "line",
        LineInv = "lineInv",
        Triangle = "triangle",
        RtTriangle = "rtTriangle",
        Rect = "rect",
        Diamond = "diamond",
        Parallelogram = "parallelogram",
        Trapezoid = "trapezoid",
        NonIsoscelesTrapezoid = "nonIsoscelesTrapezoid",
        Pentagon = "pentagon",
        Hexagon = "hexagon",
        Heptagon = "heptagon",
        Octagon = "octagon",
        Decagon = "decagon",
        Dodecagon = "dodecagon",
        Star4 = "star4",
        Star5 = "star5",
        Star6 = "star6",
        Star7 = "star7",
        Star8 = "star8",
        Star10 = "star10",
        Star12 = "star12",
        Star16 = "star16",
        Star24 = "star24",
        Star32 = "star32",
        RoundRect = "roundRect",
        Round1Rect = "round1Rect",
        Round2SameRect = "round2SameRect",
        Round2DiagRect = "round2DiagRect",
        SnipRoundRect = "snipRoundRect",
        Snip1Rect = "snip1Rect",
        Snip2SameRect = "snip2SameRect",
        Snip2DiagRect = "snip2DiagRect",
        Plaque = "plaque",
        Ellipse = "ellipse",
        Teardrop = "teardrop",
        HomePlate = "homePlate",
        Chevron = "chevron",
        PieWedge = "pieWedge",
        Pie = "pie",
        BlockArc = "blockArc",
        Donut = "donut",
        NoSmoking = "noSmoking",
        RightArrow = "rightArrow",
        LeftArrow = "leftArrow",
        UpArrow = "upArrow",
        DownArrow = "downArrow",
        StripedRightArrow = "stripedRightArrow",
        NotchedRightArrow = "notchedRightArrow",
        BentUpArrow = "bentUpArrow",
        LeftRightArrow = "leftRightArrow",
        UpDownArrow = "upDownArrow",
        LeftUpArrow = "leftUpArrow",
        LeftRightUpArrow = "leftRightUpArrow",
        QuadArrow = "quadArrow",
        LeftArrowCallout = "leftArrowCallout",
        RightArrowCallout = "rightArrowCallout",
        UpArrowCallout = "upArrowCallout",
        DownArrowCallout = "downArrowCallout",
        LeftRightArrowCallout = "leftRightArrowCallout",
        UpDownArrowCallout = "upDownArrowCallout",
        QuadArrowCallout = "quadArrowCallout",
        BentArrow = "bentArrow",
        UturnArrow = "uturnArrow",
        CircularArrow = "circularArrow",
        LeftCircularArrow = "leftCircularArrow",
        LeftRightCircularArrow = "leftRightCircularArrow",
        CurvedRightArrow = "curvedRightArrow",
        CurvedLeftArrow = "curvedLeftArrow",
        CurvedUpArrow = "curvedUpArrow",
        CurvedDownArrow = "curvedDownArrow",
        SwooshArrow = "swooshArrow",
        Cube = "cube",
        Can = "can",
        LightningBolt = "lightningBolt",
        Heart = "heart",
        Sun = "sun",
        Moon = "moon",
        SmileyFace = "smileyFace",
        IrregularSeal1 = "irregularSeal1",
        IrregularSeal2 = "irregularSeal2",
        FoldedCorner = "foldedCorner",
        Bevel = "bevel",
        Frame = "frame",
        HalfFrame = "halfFrame",
        Corner = "corner",
        DiagStripe = "diagStripe",
        Chord = "chord",
        Arc = "arc",
        LeftBracket = "leftBracket",
        RightBracket = "rightBracket",
        LeftBrace = "leftBrace",
        RightBrace = "rightBrace",
        BracketPair = "bracketPair",
        BracePair = "bracePair",
        StraightConnector1 = "straightConnector1",
        BentConnector2 = "bentConnector2",
        BentConnector3 = "bentConnector3",
        BentConnector4 = "bentConnector4",
        BentConnector5 = "bentConnector5",
        CurvedConnector2 = "curvedConnector2",
        CurvedConnector3 = "curvedConnector3",
        CurvedConnector4 = "curvedConnector4",
        CurvedConnector5 = "curvedConnector5",
        Callout1 = "callout1",
        Callout2 = "callout2",
        Callout3 = "callout3",
        AccentCallout1 = "accentCallout1",
        AccentCallout2 = "accentCallout2",
        AccentCallout3 = "accentCallout3",
        BorderCallout1 = "borderCallout1",
        BorderCallout2 = "borderCallout2",
        BorderCallout3 = "borderCallout3",
        AccentBorderCallout1 = "accentBorderCallout1",
        AccentBorderCallout2 = "accentBorderCallout2",
        AccentBorderCallout3 = "accentBorderCallout3",
        WedgeRectCallout = "wedgeRectCallout",
        WedgeRoundRectCallout = "wedgeRoundRectCallout",
        WedgeEllipseCallout = "wedgeEllipseCallout",
        CloudCallout = "cloudCallout",
        Cloud = "cloud",
        Ribbon = "ribbon",
        Ribbon2 = "ribbon2",
        EllipseRibbon = "ellipseRibbon",
        EllipseRibbon2 = "ellipseRibbon2",
        LeftRightRibbon = "leftRightRibbon",
        VerticalScroll = "verticalScroll",
        HorizontalScroll = "horizontalScroll",
        Wave = "wave",
        DoubleWave = "doubleWave",
        Plus = "plus",
        FlowChartProcess = "flowChartProcess",
        FlowChartDecision = "flowChartDecision",
        FlowChartInputOutput = "flowChartInputOutput",
        FlowChartPredefinedProcess = "flowChartPredefinedProcess",
        FlowChartInternalStorage = "flowChartInternalStorage",
        FlowChartDocument = "flowChartDocument",
        FlowChartMultidocument = "flowChartMultidocument",
        FlowChartTerminator = "flowChartTerminator",
        FlowChartPreparation = "flowChartPreparation",
        FlowChartManualInput = "flowChartManualInput",
        FlowChartManualOperation = "flowChartOperation",
        FlowChartConnector = "flowChartConnector",
        FlowChartPunchedCard = "flowChartPunchedCard",
        FlowChartPunchedTape = "flowChartPunchedTape",
        FlowChartSummingJunction = "flowChartSummingJunction",
        FlowChartOr = "flowChartOr",
        FlowChartCollate = "flowChartCollate",
        FlowChartSort = "flowChartSort",
        FlowChartExtract = "flowChartExtract",
        FlowChartMerge = "flowChartMerge",
        FlowChartOfflineStorage = "flowChartOfflineStorage",
        FlowChartOnlineStorage = "flowChartOnlineStorage",
        FlowChartMagneticTape = "flowChartMagneticTape",
        FlowChartMagneticDisk = "flowChartMagneticDisk",
        FlowChartMagneticDrum = "flowChartMagneticDrum",
        FlowChartDisplay = "flowChartDisplay",
        FlowChartDelay = "flowChartDelay",
        FlowChartAlternateProcess = "flowChartAlternateProcess",
        FlowChartOffpageConnector = "flowChartOffpageConnector",
        ActionButtonBlank = "actionButtonBlank",
        ActionButtonHome = "actionButtonHome",
        ActionButtonHelp = "actionButtonHelp",
        ActionButtonInformation = "actionButtonInformation",
        ActionButtonForwardNext = "actionButtonForwardNext",
        ActionButtonBackPrevious = "actionButtonBackPrevious",
        ActionButtonEnd = "actionButtonEnd",
        ActionButtonBeginning = "actionButtonBeginning",
        ActionButtonReturn = "actionButtonReturn",
        ActionButtonDocument = "actionButtonDocument",
        ActionButtonSound = "actionButtonSound",
        ActionButtonMovie = "actionButtonMovie",
        Gear6 = "gear6",
        Gear9 = "gear9",
        Funnel = "funnel",
        MathPlus = "mathPlus",
        MathMinus = "mathMinus",
        MathMultiply = "mathMultiply",
        MathDivide = "mathDivide",
        MathEqual = "mathEqual",
        MathNotEqual = "mathNotEqual",
        CornerTabs = "cornerTabs",
        SquareTabs = "squareTabs",
        PlaqueTabs = "plaqueTabs",
        ChartX = "chartX",
        ChartStar = "chartStar",
        ChartPlus = "chartPlus",
    }
}

decl_oox_enum! {
    pub enum LineCap {
        Rnd = "rnd",
        Sq = "sq",
        Flat = "flat",
    }
}

decl_oox_enum! {
    pub enum CompoundLine {
        Sng = "sng",
        Dbl = "dbl",
        ThickThin = "thickThin",
        ThinThick = "thinThick",
        Tri = "tri",
    }
}

decl_oox_enum! {
    pub enum PenAlignment {
        Ctr = "ctr",
        In = "in",
    }
}

decl_oox_enum! {
    pub enum PresetLineDashVal {
        Solid = "solid",
        Dot = "dot",
        Dash = "dash",
        LgDash = "lgDash",
        DashDot = "dashDot",
        LgDashDot = "lgDashDot",
        LgDashDotDot = "ldDashDotDot",
        SysDash = "sysDash",
        SysDot = "sysDot",
        SysDashDot = "sysDashDot",
        SysDashDotDot = "sysDashDotDot",
    }
}

decl_oox_enum! {
    pub enum LineEndType {
        None = "none",
        Triangle = "triangle",
        Stealth = "stealth",
        Diamond = "diamond",
        Oval = "oval",
        Arrow = "arrow",
    }
}

decl_oox_enum! {
    pub enum LineEndWidth {
        Sm = "sm",
        Med = "med",
        Lg = "lg",
    }
}

decl_oox_enum! {
    pub enum LineEndLength {
        Sm = "sm",
        Med = "med",
        Lg = "lg",
    }
}

decl_oox_enum! {
    pub enum BlendMode {
        Over = "over",
        Mult = "mult",
        Screen = "screen",
        Darken = "darken",
        Lighten = "lighten",
    }
}

decl_oox_enum! {
    pub enum PresetShadowVal {
        Shdw1 = "shdw1",
        Shdw2 = "shdw2",
        Shdw3 = "shdw3",
        Shdw4 = "shdw4",
        Shdw5 = "shdw5",
        Shdw6 = "shdw6",
        Shdw7 = "shdw7",
        Shdw8 = "shdw8",
        Shdw9 = "shdw9",
        Shdw10 = "shdw10",
        Shdw11 = "shdw11",
        Shdw12 = "shdw12",
        Shdw13 = "shdw13",
        Shdw14 = "shdw14",
        Shdw15 = "shdw15",
        Shdw16 = "shdw16",
        Shdw17 = "shdw17",
        Shdw18 = "shdw18",
        Shdw19 = "shdw19",
        Shdw20 = "shdw20",
    }
}

decl_oox_enum! {
    pub enum EffectContainerType {
        Sib = "sib",
        Tree = "tree",
    }
}

decl_oox_enum! {
    pub enum FontCollectionIndex {
        Major = "major",
        Minor = "minor",
        None = "none",
    }
}

decl_oox_enum! {
    pub enum AnimationBuildType {
        AllAtOnce = "allAtOnce",
    }
}

decl_oox_enum! {
    pub enum AnimationDgmOnlyBuildType {
        One = "one",
        LvlOne = "lvlOne",
        LvlAtOnce = "lvlAtOnce",
    }
}

decl_oox_enum! {
    pub enum AnimationChartOnlyBuildType {
        Series = "series",
        Category = "category",
        SeriesEl = "seriesEl",
        CategoryEl = "categoryEl",
    }
}

decl_oox_enum! {
    pub enum DgmBuildStep {
        Sp = "sp",
        Bg = "bg",
    }
}

decl_oox_enum! {
    pub enum ChartBuildStep {
        Category = "category",
        PtInCategory = "ptInCategory",
        Series = "series",
        PtInSeries = "ptInSeries",
        AllPts = "allPts",
        GridLegend = "gridLegend",
    }
}

decl_oox_enum! {
    pub enum OnOffStyleType {
        On = "on",
        Off = "off",
        Def = "def",
    }
}


pub struct RelativeRect {
    pub left: Option<Percentage>,
    pub top: Option<Percentage>,
    pub right: Option<Percentage>,
    pub bottom: Option<Percentage>,
}


pub struct AlphaModulateFixedEffect {
    pub amount: Option<PositivePercentage>
}


pub struct LuminanceEffect {
    pub bright: Option<FixedPercentage>,
    pub contrast: Option<FixedPercentage>,
}


pub struct DuotoneEffect {
    pub colors: [Color; 2],
}


pub enum EffectGroup {
    AlphaModFixed(AlphaModulateFixedEffect),
    Luminance(LuminanceEffect),
    Duotone(DuotoneEffect),
}
/*
*/