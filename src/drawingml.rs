// TODO: This module defines shared types between different OOX file formats. It should be refactored into a different crate, if these types are needed.
use crate::error::{
    AdjustParseError, Limit, LimitViolationError, MissingAttributeError, MissingChildNodeError, NotGroupMemberError,
};
use crate::relationship::RelationshipId;
use crate::xml::{parse_xml_bool, XmlNode};
use std::io::Read;
use std::str::FromStr;
use zip::read::ZipFile;

pub type Result<T> = ::std::result::Result<T, Box<dyn (::std::error::Error)>>;

pub type Guid = String; // TODO: move to shared common types. pattern="\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}"
pub type Percentage = f32;
pub type PositivePercentage = f32; // TODO: 0 <= n < inf
pub type PositiveFixedPercentage = f32; // TODO: 0 <= n <= 100000
pub type FixedPercentage = f32; // TODO: -100000 <= n <= 100000
pub type HexColorRGB = String;
pub type Coordinate = i64;
pub type PositiveCoordinate = u64;
pub type Coordinate32 = i32;
pub type PositiveCoordinate32 = u32;
pub type LineWidth = Coordinate32;
pub type DrawingElementId = u32;
pub type Angle = i32;
pub type FixedAngle = Angle; // TODO: -5400000 <= n <= 5400000
pub type PositiveFixedAngle = Angle; // TODO: 0 <= n <= 21600000
pub type GeomGuideName = String;
pub type GeomGuideFormula = String;
pub type StyleMatrixColumnIndex = u32;
pub type TextColumnCount = i32; // TODO: 1 <= n <= 16
pub type TextFontScalePercent = Percentage; // TODO: 1000 <= n <= 100000
pub type TextSpacingPercent = Percentage; // TODO: 0 <= n <= 13200000
pub type TextSpacingPoint = i32; // TODO: 0 <= n <= 158400
pub type TextMargin = Coordinate32; // TODO: 0 <= n <= 51206400
pub type TextIndent = Coordinate32; // TODO: -51206400 <= n <= 51206400
pub type TextIndentLevelType = i32; // TODO; 0 <= n <= 8
pub type TextBulletSizePercent = Percentage; // TODO: 0.25 <= n <= 4.0
pub type TextFontSize = i32; // TODO: 100 <= n <= 400000
pub type TextTypeFace = String;
pub type TextLanguageID = String;
pub type Panose = String; // TODO: hex, length=10
pub type TextBulletStartAtNum = i32; // TODO: 1 <= n <= 32767
pub type Lang = String;
pub type TextNonNegativePoint = i32; // TODO: 0 <= n <= 400000
pub type TextPoint = i32; // TODO: -400000 <= n <= 400000
pub type ShapeId = String;

decl_simple_type_enum! {
    pub enum TileFlipMode {
        None = "none",
        X = "x",
        Y = "y",
        XY = "xy",
    }
}

decl_simple_type_enum! {
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

decl_simple_type_enum! {
    pub enum PathFillMode {
        None = "none",
        Norm = "norm",
        Lighten = "lighten",
        LightenLess = "lightenLess",
        Darken = "darken",
        DarkenLess = "darkenLess",
    }
}

decl_simple_type_enum! {
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

decl_simple_type_enum! {
    pub enum LineCap {
        Round = "rnd",
        Square = "sq",
        Flat = "flat",
    }
}

decl_simple_type_enum! {
    pub enum CompoundLine {
        Single = "sng",
        Double = "dbl",
        ThickThin = "thickThin",
        ThinThick = "thinThick",
        Triple = "tri",
    }
}

decl_simple_type_enum! {
    pub enum PenAlignment {
        Center = "ctr",
        Inset = "in",
    }
}

decl_simple_type_enum! {
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

decl_simple_type_enum! {
    pub enum LineEndType {
        None = "none",
        Triangle = "triangle",
        Stealth = "stealth",
        Diamond = "diamond",
        Oval = "oval",
        Arrow = "arrow",
    }
}

decl_simple_type_enum! {
    pub enum LineEndWidth {
        Small = "sm",
        Medium = "med",
        Large = "lg",
    }
}

decl_simple_type_enum! {
    pub enum LineEndLength {
        Small = "sm",
        Medium = "med",
        Large = "lg",
    }
}

decl_simple_type_enum! {
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

decl_simple_type_enum! {
    pub enum EffectContainerType {
        Sib = "sib",
        Tree = "tree",
    }
}

decl_simple_type_enum! {
    pub enum FontCollectionIndex {
        Major = "major",
        Minor = "minor",
        None = "none",
    }
}

decl_simple_type_enum! {
    pub enum DgmBuildStep {
        Sp = "sp",
        Bg = "bg",
    }
}

decl_simple_type_enum! {
    pub enum ChartBuildStep {
        Category = "category",
        PtInCategory = "ptInCategory",
        Series = "series",
        PtInSeries = "ptInSeries",
        AllPts = "allPts",
        GridLegend = "gridLegend",
    }
}

decl_simple_type_enum! {
    pub enum OnOffStyleType {
        On = "on",
        Off = "off",
        Def = "def",
    }
}

decl_simple_type_enum! {
    pub enum SystemColorVal {
        ScrollBar = "scrollBar",
        Background = "background",
        ActiveCaption = "activeCaption",
        InactiveCaption = "inactiveCaption",
        Menu = "menu",
        Window = "window",
        WindowFrame = "windowFrame",
        MenuText = "menuText",
        WindowText = "windowText",
        CaptionText = "captionText",
        ActiveBorder = "activeBorder",
        InactiveBorder = "inactiveBorder",
        AppWorkspace = "appWorkspace",
        Highlight = "highlight",
        HighlightText = "highlightText",
        BtnFace = "btnFace",
        BtnShadow = "btnShadow",
        GrayText = "grayText",
        BtnText = "btnText",
        InactiveCaptionText = "inactiveCaptionText",
        BtnHighlight = "btnHighlight",
        DkShadow3d = "3dDkShadow",
        Light3d = "3dLight",
        InfoText = "infoText",
        InfoBk = "infoBk",
        HotLight = "hotLight",
        GradientActiveCaption = "gradientActiveCaption",
        GradientInactiveCaption = "gradientInactiveCaption",
        MenuHighlight = "menuHighlight",
        MenuBar = "menubar",
    }
}

decl_simple_type_enum! {
    pub enum PresetColorVal {
        AliceBlue = "aliceBlue",
        AntiqueWhite = "antiqueWhite",
        Aqua = "aqua",
        Aquamarine = "aquamarine",
        Azure = "azure",
        Beige = "beige",
        Bisque = "bisque",
        Black = "black",
        BlanchedAlmond = "blanchedAlmond",
        Blue = "blue",
        BlueViolet = "blueViolet",
        Brown = "brown",
        BurlyWood = "burlyWood",
        CadetBlue = "cadetBlue",
        Chartreuse = "chartreuse",
        Chocolate = "chocolate",
        Coral = "coral",
        CornflowerBlue = "cornflowerBlue",
        Cornsilk = "cornsilk",
        Crimson = "crimson",
        Cyan = "cyan",
        DarkBlue = "darkBlue",
        DarkCyan = "darkCyan",
        DarkGoldenrod = "darkGoldenrod",
        DarkGray = "darkGray",
        DarkGrey = "darkGrey",
        DarkGreen = "darkGreen",
        DarkKhaki = "darkKhaki",
        DarkMagenta = "darkMagenta",
        DarkOliveGreen = "darkOliveGreen",
        DarkOrange = "darkOrange",
        DarkOrchid = "darkOrchid",
        DarkRed = "darkRed",
        DarkSalmon = "darkSalmon",
        DarkSeaGreen = "darkSeaGreen",
        DarkSlateBlue = "darkSlateBlue",
        DarkSlateGray = "darkSlateGray",
        DarkSlateGrey = "darkSlateGrey",
        DarkTurqoise = "darkTurquoise",
        DarkViolet = "darkViolet",
        DkBlue = "dkBlue",
        DkCyan = "dkCyan",
        DkGoldenrod = "dkGoldenrod",
        DkGray = "dkGray",
        DkGrey = "dkGrey",
        DkGreen = "dkGreen",
        DkKhaki = "dkKhaki",
        DkMagenta = "dkMagenta",
        DkOliveGreen = "dkOliveGreen",
        DkOrange = "dkOrange",
        DkOrchid = "dkOrchid",
        DkRed = "dkRed",
        DkSalmon = "dkSalmon",
        DkSeaGreen = "dkSeaGreen",
        DkSlateBlue = "dkSlateBlue",
        DkSlateGray = "dkSlateGray",
        DkSlateGrey = "dkSlateGrey",
        DkTurquoise = "dkTurquoise",
        DkViolet = "dkViolet",
        DeepPink = "deepPink",
        DeepSkyBlue = "deepSkyBlue",
        DimGray = "dimGray",
        DimGrey = "dimGrey",
        DodgerBluet = "dodgerBlue",
        Firebrick = "firebrick",
        FloralWhite = "floralWhite",
        ForestGreen = "forestGreen",
        Fuchsia = "fuchsia",
        Gainsboro = "gainsboro",
        GhostWhite = "ghostWhite",
        Gold = "gold",
        Goldenrod = "goldenrod",
        Gray = "gray",
        Grey = "grey",
        Green = "green",
        GreenYellow = "greenYellow",
        Honeydew = "honeydew",
        HotPink = "hotPink",
        IndianRed = "indianRed",
        Indigo = "indigo",
        Ivory = "ivory",
        Khaki = "khaki",
        Lavender = "lavender",
        LavenderBlush = "lavenderBlush",
        LawnGreen = "lawnGreen",
        LemonChiffon = "lemonChiffon",
        LightBlue = "lightBlue",
        LightCoral = "lightCoral",
        LightCyan = "lightCyan",
        LightGoldenrodYellow = "lightGoldenrodYellow",
        LightGray = "lightGray",
        LightGrey = "lightGrey",
        LightGreen = "lightGreen",
        LightPink = "lightPink",
        LightSalmon = "lightSalmon",
        LightSeaGreen = "lightSeaGreen",
        LightSkyBlue = "lightSkyBlue",
        LightSlateGray = "lightSlateGray",
        LightSlateGrey = "lightSlateGrey",
        LightSteelBlue = "lightSteelBlue",
        LightYellow = "lightYellow",
        LtBlue = "ltBlue",
        LtCoral = "ltCoral",
        LtCyan = "ltCyan",
        LtGoldenrodYellow = "ltGoldenrodYellow",
        LtGray = "ltGray",
        LtGrey = "ltGrey",
        LtGreen = "ltGreen",
        LtPink = "ltPink",
        LtSalmon = "ltSalmon",
        LtSeaGreen = "ltSeaGreen",
        LtSkyBlue = "ltSkyBlue",
        LtSlateGray = "ltSlateGray",
        LtSlateGrey = "ltSlateGrey",
        LtSteelBlue = "ltSteelBlue",
        LtYellow = "ltYellow",
        Lime = "lime",
        LimeGreen = "limeGreen",
        Linen = "linen",
        Magenta = "magenta",
        Maroon = "maroon",
        MedAquamarine = "medAquamarine",
        MedBlue = "medBlue",
        MedOrchid = "medOrchid",
        MedPurple = "medPurple",
        MedSeaGreen = "medSeaGreen",
        MedSlateBlue = "medSlateBlue",
        MedSpringGreen = "medSpringGreen",
        MedTurquoise = "medTurquoise",
        MedVioletRed = "medVioletRed",
        MediumAquamarine = "mediumAquamarine",
        MediumBlue = "mediumBlue",
        MediumOrchid = "mediumOrchid",
        MediumPurple = "mediumPurple",
        MediumSeaGreen = "mediumSeaGreen",
        MediumSlateBlue = "mediumSlateBlue",
        MediumSpringGreen = "mediumSpringGreen",
        MediumTurquoise = "mediumTurquoise",
        MediumVioletRed = "mediumVioletRed",
        MidnightBlue = "midnightBlue",
        MintCream = "mintCream",
        MistyRose = "mistyRose",
        Moccasin = "moccasin",
        NavajoWhite = "navajoWhite",
        Navy = "navy",
        OldLace = "oldLace",
        Olive = "olive",
        OliveDrab = "oliveDrab",
        Orange = "orange",
        OrangeRed = "orangeRed",
        Orchid = "orchid",
        PaleGoldenrod = "paleGoldenrod",
        PaleGreen = "paleGreen",
        PaleTurquoise = "paleTurquoise",
        PaleVioletRed = "paleVioletRed",
        PapayaWhip = "papayaWhip",
        PeachPuff = "peachPuff",
        Peru = "peru",
        Pink = "pink",
        Plum = "plum",
        PowderBlue = "powderBlue",
        Purple = "purple",
        Red = "red",
        RosyBrown = "rosyBrown",
        RoyalBlue = "royalBlue",
        SaddleBrown = "saddleBrown",
        Salmon = "salmon",
        SandyBrown = "sandyBrown",
        SeaGreen = "seaGreen",
        SeaShell = "seaShell",
        Sienna = "sienna",
        Silver = "silver",
        SkyBlue = "skyBlue",
        SlateBlue = "slateBlue",
        SlateGray = "slateGray",
        SlateGrey = "slateGrey",
        Snow = "snow",
        SpringGreen = "springGreen",
        SteelBlue = "steelBlue",
        Tan = "tan",
        Teal = "teal",
        Thistle = "thistle",
        Tomato = "tomato",
        Turquoise = "turquoise",
        Violet = "violet",
        Wheat = "wheat",
        White = "white",
        WhiteSmoke = "whiteSmoke",
        Yellow = "yellow",
        YellowGreen = "yellowGreen",
    }
}

decl_simple_type_enum! {
    pub enum SchemeColorVal {
        Background1 = "bg1",
        Text1 = "tx1",
        Background2 = "bg2",
        Text2 = "tx2",
        Accent1 = "accent1",
        Accent2 = "accent2",
        Accent3 = "accent3",
        Accent4 = "accent4",
        Accent5 = "accent5",
        Hypelinglink = "hlink",
        FollowedHyperlink = "folHlink",
        PlaceholderColor = "phClr",
        Dark1 = "dk1",
        Light1 = "lt1",
        Dark2 = "dk2",
        Light2 = "lt2",
    }
}

decl_simple_type_enum! {
    pub enum ColorSchemeIndex {
        Dark1 = "dk1",
        Light1 = "lt1",
        Dark2 = "dk2",
        Light2 = "lt2",
        Accent1 = "accent1",
        Accent2 = "accent2",
        Accent3 = "accent3",
        Accent4 = "accent4",
        Accent5 = "accent5",
        Accent6 = "accent6",
        Hyperlink = "hlink",
        FollowedHyperlink = "folHlink",
    }
}

decl_simple_type_enum! {
    pub enum TextAlignType {
        Left = "l",
        Center = "ctr",
        Right = "r",
        Justified = "just",
        JustifiedLow = "justLow",
        Distritbuted = "dist",
        ThaiDistributed = "thaiDist",
    }
}

decl_simple_type_enum! {
    pub enum TextFontAlignType {
        Auto = "auto",
        Top = "t",
        Center = "ctr",
        Baseline = "base",
        Bottom = "b",
    }
}

decl_simple_type_enum! {
    pub enum TextAutonumberScheme {
        AlphaLcParenBoth = "alphaLcParenBoth",
        AlphaUcParenBoth = "alphaUcParenBoth",
        AlphaLcParenR = "alphaLcParenR",
        AlphaUcParenR = "alphaUcParenR",
        AlphaLcPeriod = "alphaLcPeriod",
        AlphaUcPeriod = "alphaUcPeriod",
        ArabicParenBoth = "arabicParenBoth",
        ArabicParenR = "arabicParenR",
        ArabicPeriod = "arabicPeriod",
        ArabicPlain = "arabicPlain",
        RomanLcParenBoth = "romanLcParenBoth",
        RomanUcParenBoth = "romanUcParenBoth",
        RomanLcParenR = "romanLcParenR",
        RomanUcParenR = "romanUcParenR",
        RomanLcPeriod = "romanLcPeriod",
        RomanUcPeriod = "romanUcPeriod",
        CircleNumDbPlain = "circleNumDbPlain",
        CircleNumWdBlackPlain = "circleNumWdBlackPlain",
        CircleNumWdWhitePlain = "circleNumWdWhitePlain",
        ArabicDbPeriod = "arabicDbPeriod",
        ArabicDbPlain = "arabicDbPlain",
        Ea1ChsPeriod = "ea1ChsPeriod",
        Ea1ChsPlain = "ea1ChsPlain",
        Ea1ChtPeriod = "ea1ChtPeriod",
        Ea1ChtPlain = "ea1ChtPlain",
        Ea1JpnChsDbPeriod = "ea1JpnChsDbPeriod",
        Ea1JpnKorPlain = "ea1JpnKorPlain",
        Ea1JpnKorPeriod = "ea1JpnKorPeriod",
        Arabic1Minus = "arabic1Minus",
        Arabic2Minus = "arabic2Minus",
        Hebrew2Minus = "hebrew2Minus",
        ThaiAlphaPeriod = "thaiAlphaPeriod",
        ThaiAlphaParenR = "thaiAlphaParenR",
        ThaiAlphaParenBoth = "thaiAlphaParenBoth",
        ThaiNumPeriod = "thaiNumPeriod",
        ThaiNumParenR = "thaiNumParenR",
        ThaiNumParenBoth = "thaiNumParenBoth",
        HindiAlphaPeriod = "hindiAlphaPeriod",
        HindiNumPeriod = "hindiNumPeriod",
        HindiNumParenR = "hindiNumParenR",
        HindiAlpha1Period = "hindiAlpha1Period",
    }
}

decl_simple_type_enum! {
    pub enum PathShadeType {
        Shape = "shape",
        Circle = "circle",
        Rect = "rect",
    }
}

decl_simple_type_enum! {
    pub enum PresetPatternVal {
        Percent5 = "pct5",
        Percent10 = "pct10",
        Percent20 = "pct20",
        Percent25 = "pct25",
        Percent30 = "pct30",
        Percent40 = "pct40",
        Percent50 = "pct50",
        Percent60 = "pct60",
        Percent70 = "pct70",
        Percent75 = "pct75",
        Percent80 = "pct80",
        Percent90 = "pct90",
        Horizontal = "horz",
        Vertical = "vert",
        LightHorizontal = "ltHorz",
        LightVertical = "ltVert",
        DarkHorizontal = "dkHorz",
        DarkVertical = "dkVert",
        NarrowHorizontal = "narHorz",
        NarrowVertical = "narVert",
        DashedHorizontal = "dashHorz",
        DashedVertical = "dashVert",
        Cross = "cross",
        DownwardDiagonal = "dnDiag",
        UpwardDiagonal = "upDiag",
        LightDownwardDiagonal = "ltDnDiag",
        LightUpwardDiagonal = "ltUpDiag",
        DarkDownwardDiagonal = "dkDnDiag",
        DarkUpwardDiagonal = "dkUpDiag",
        WideDownwardDiagonal = "wdDnDiag",
        WideUpwardDiagonal = "wdUpDiag",
        DashedDownwardDiagonal = "dashDnDiag",
        DashedUpwardDiagonal = "dashUpDiag",
        DiagonalCross = "diagCross",
        SmallCheckerBoard = "smCheck",
        LargeCheckerBoard = "lgCheck",
        SmallGrid = "smGrid",
        LargeGrid = "lgGrid",
        DottedGrid = "dotGrid",
        SmallConfetti = "smConfetti",
        LargeConfetti = "lgConfetti",
        HorizontalBrick = "horzBrick",
        DiagonalBrick = "diagBrick",
        SolidDiamond = "solidDmnd",
        OpenDiamond = "openDmnd",
        DottedDiamond = "dotDmnd",
        Plaid = "plaid",
        Sphere = "sphere",
        Weave = "weave",
        Divot = "divot",
        Shingle = "shingle",
        Wave = "wave",
        Trellis = "trellis",
        ZigZag = "zigzag",
    }
}

decl_simple_type_enum! {
    pub enum BlendMode {
        Overlay = "over",
        Multiply = "mult",
        Screen = "screen",
        Lighten = "lighten",
        Darken = "darken",
    }
}

decl_simple_type_enum! {
    pub enum TextTabAlignType {
        Left = "l",
        Center = "ctr",
        Right = "r",
        Decimal = "dec",
    }
}

decl_simple_type_enum! {
    pub enum TextUnderlineType {
        None = "none",
        Words = "words",
        Single = "sng",
        Double = "dbl",
        Heavy = "heavy",
        Dotted = "dotted",
        DottedHeavy = "dottedHeavy",
        Dash = "dash",
        DashHeavy = "dashHeavy",
        DashLong = "dashLong",
        DashLongHeavy = "dashLongHeavy",
        DotDash = "dotDash",
        DotDashHeavy = "dotDashHeavy",
        DotDotDash = "dotDotDash",
        DotDotDashHeavy = "dotDotDashHeavy",
        Wavy = "wavy",
        WavyHeavy = "wavyHeavy",
        WavyDouble = "wavyDbl",
    }
}

decl_simple_type_enum! {
    pub enum TextStrikeType {
        NoStrike = "noStrike",
        SingleStrike = "sngStrike",
        DoubleStrike = "dblStrike",
    }
}

decl_simple_type_enum! {
    pub enum TextCapsType {
        None = "none",
        Small = "small",
        All = "all",
    }
}

decl_simple_type_enum! {
    pub enum TextShapeType {
        NoShape = "textNoShape",
        Plain = "textPlain",
        Stop = "textStop",
        Triangle = "textTriangle",
        TriangleInverted = "textTriangleInverted",
        Chevron = "textChevron",
        ChevronInverted = "textChevronInverted",
        RingInside = "textRingInside",
        RingOutside = "textRingOutside",
        ArchUp = "textArchUp",
        ArchDown = "textArchDown",
        Circle = "textCircle",
        Button = "textButton",
        ArchUpPour = "textArchUpPour",
        ArchDownPour = "textArchDownPour",
        CirclePour = "textCirclePour",
        ButtonPour = "textButtonPour",
        CurveUp = "textCurveUp",
        CurveDown = "textCurveDown",
        CanUp = "textCanUp",
        CanDown = "textCanDown",
        Wave1 = "textWave1",
        Wave2 = "textWave2",
        Wave4 = "textWave4",
        DoubleWave1 = "textDoubleWave1",
        Inflate = "textInflate",
        Deflate = "textDeflate",
        InflateBottom = "textInflateBottom",
        DeflateBottom = "textDeflateBottom",
        InflateTop = "textInflateTop",
        DeflateTop = "textDeflateTop",
        DeflateInflate = "textDeflateInflate",
        DeflateInflateDeflate = "textDeflateInflateDeflate",
        FadeLeft = "textFadeLeft",
        FadeUp = "textFadeUp",
        FadeRight = "textFadeRight",
        FadeDown = "textFadeDown",
        SlantUp = "textSlantUp",
        SlantDown = "textSlantDown",
        CascadeUp = "textCascadeUp",
        CascadeDown = "textCascadeDown",
    }
}

decl_simple_type_enum! {
    pub enum TextVertOverflowType {
        Overflow = "overflow",
        Ellipsis = "ellipsis",
        Clip = "clip",
    }
}

decl_simple_type_enum! {
    pub enum TextHorzOverflowType {
        Overflow = "overflow",
        Clip = "clip",
    }
}

decl_simple_type_enum! {
    pub enum TextVerticalType {
        Horizontal = "horz",
        Vertical = "vert",
        Vertical270 = "vert270",
        WordArtVertical = "wordArtVert",
        EastAsianVertical = "eaVert",
        MongolianVertical = "mongolianVert",
        WordArtVerticalRtl = "wordArtVertRtl",
    }
}

decl_simple_type_enum! {
    pub enum TextWrappingType {
        None = "none",
        Square = "square",
    }
}

decl_simple_type_enum! {
    pub enum TextAnchoringType {
        Top = "t",
        Center = "ctr",
        Bottom = "b",
        Justified = "just",
        Distributed = "dist",
    }
}

decl_simple_type_enum! {
    pub enum BlackWhiteMode {
        Color = "clr",
        Auto = "auto",
        Gray = "gray",
        LightGray = "ltGray",
        InverseGray = "invGray",
        GrayWhite = "grayWhite",
        BlackGray = "blackGray",
        BlackWhite = "blackWhite",
        Black = "black",
        White = "white",
        Hidden = "hidden",
    }
}

decl_simple_type_enum! {
    pub enum AnimationBuildType {
        AllAtOnce = "allAtOnce",
    }
}

decl_simple_type_enum! {
    pub enum AnimationDgmOnlyBuildType {
        One = "one",
        LvlOne = "lvlOne",
        LvlAtOnce = "lvlAtOnce",
    }
}

decl_simple_type_enum! {
    pub enum AnimationDgmBuildType {
        AllAtOnce = "allAtOnce",
        One = "one",
        LvlOne = "lvlOne",
        LvlAtOnce = "lvlAtOnce",
    }
}

decl_simple_type_enum! {
    pub enum AnimationChartOnlyBuildType {
        Series = "series",
        Category = "category",
        SeriesElement = "seriesElement",
        CategoryElement = "categoryElement",
    }
}

decl_simple_type_enum! {
    pub enum AnimationChartBuildType {
        AllAtOnce = "allAtOnce",
        Series = "series",
        Category = "category",
        SeriesElement = "seriesElement",
        CategoryElement = "categoryElement",
    }
}

decl_simple_type_enum! {
    pub enum BlipCompression {
        Email = "email",
        Screen = "screen",
        Print = "print",
        HqPrint = "hqprint",
        None = "none",
    }
}

/// ColorTransform
pub enum ColorTransform {
    Tint(PositiveFixedPercentage),
    Shade(PositiveFixedPercentage),
    Complement,
    Inverse,
    Grayscale,
    Alpha(PositiveFixedPercentage),
    AlphaOffset(FixedPercentage),
    AlphaModulate(PositivePercentage),
    Hue(PositiveFixedAngle),
    HueOffset(Angle),
    HueModulate(PositivePercentage),
    Saturation(Percentage),
    SaturationOffset(Percentage),
    SaturationModulate(Percentage),
    Luminance(Percentage),
    LuminanceOffset(Percentage),
    LuminanceModulate(Percentage),
    Red(Percentage),
    RedOffset(Percentage),
    RedModulate(Percentage),
    Green(Percentage),
    GreenOffset(Percentage),
    GreenModulate(Percentage),
    Blue(Percentage),
    BlueOffset(Percentage),
    BlueModulate(Percentage),
    Gamma,
    InverseGamma,
}

impl ColorTransform {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "tint" | "shade" | "comp" | "inv" | "gray" | "alpha" | "alphaOff" | "alphaMod" | "hue" | "hueOff"
            | "hueMod" | "sat" | "satOff" | "satMod" | "lum" | "lumOff" | "lumMod" | "red" | "redOff" | "redMod"
            | "green" | "greenOff" | "greenMod" | "blue" | "blueOff" | "blueMod" | "gamma" | "invGamma" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<ColorTransform> {
        match xml_node.local_name() {
            "tint" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::Tint(value.parse::<PositiveFixedPercentage>()?))
            }
            "shade" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::Shade(value.parse::<PositiveFixedPercentage>()?))
            }
            "comp" => Ok(ColorTransform::Complement),
            "inv" => Ok(ColorTransform::Inverse),
            "gray" => Ok(ColorTransform::Grayscale),
            "alpha" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::Alpha(value.parse::<PositiveFixedPercentage>()?))
            }
            "alphaOff" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::AlphaOffset(value.parse::<FixedPercentage>()?))
            }
            "alphaMod" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::AlphaModulate(value.parse::<FixedPercentage>()?))
            }
            "hue" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::Hue(value.parse::<PositiveFixedAngle>()?))
            }
            "hueOff" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::HueOffset(value.parse::<Angle>()?))
            }
            "hueMod" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::HueModulate(value.parse::<PositivePercentage>()?))
            }
            "sat" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::Saturation(value.parse::<Percentage>()?))
            }
            "satOff" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::SaturationOffset(value.parse::<Percentage>()?))
            }
            "satMod" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::SaturationModulate(value.parse::<Percentage>()?))
            }
            "lum" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::Luminance(value.parse::<Percentage>()?))
            }
            "lumOff" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::LuminanceOffset(value.parse::<Percentage>()?))
            }
            "lumMod" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::LuminanceModulate(value.parse::<Percentage>()?))
            }
            "red" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::Red(value.parse::<Percentage>()?))
            }
            "redOff" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::RedOffset(value.parse::<Percentage>()?))
            }
            "redMod" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::RedModulate(value.parse::<Percentage>()?))
            }
            "green" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::Green(value.parse::<Percentage>()?))
            }
            "greenOff" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::GreenOffset(value.parse::<Percentage>()?))
            }
            "greenMod" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::GreenModulate(value.parse::<Percentage>()?))
            }
            "blue" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::Blue(value.parse::<Percentage>()?))
            }
            "blueOff" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::BlueOffset(value.parse::<Percentage>()?))
            }
            "blueMod" => {
                let value = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(ColorTransform::BlueModulate(value.parse::<Percentage>()?))
            }
            "gamma" => Ok(ColorTransform::Gamma),
            "invGamma" => Ok(ColorTransform::InverseGamma),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_ColorTransform").into()),
        }
    }
}

/// ScRgbColor
pub struct ScRgbColor {
    pub r: Percentage,
    pub g: Percentage,
    pub b: Percentage,
    pub color_transforms: Vec<ColorTransform>,
}

impl ScRgbColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<ScRgbColor> {
        let mut opt_r = None;
        let mut opt_g = None;
        let mut opt_b = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r" => opt_r = Some(value.parse::<Percentage>()?),
                "g" => opt_g = Some(value.parse::<Percentage>()?),
                "b" => opt_b = Some(value.parse::<Percentage>()?),
                _ => (),
            }
        }

        let r = opt_r.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r"))?;
        let g = opt_g.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "g"))?;
        let b = opt_b.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "b"))?;

        let mut color_transforms = Vec::new();

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                color_transforms.push(ColorTransform::from_xml_element(child_node)?);
            }
        }

        Ok(Self {
            r,
            g,
            b,
            color_transforms,
        })
    }
}

/// SRgbColor
pub struct SRgbColor {
    pub value: u32,
    pub color_transforms: Vec<ColorTransform>,
}

impl SRgbColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<SRgbColor> {
        let val_attr = xml_node
            .attribute("val")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
        let value = u32::from_str_radix(val_attr, 16)?;

        let mut color_transforms = Vec::new();

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                color_transforms.push(ColorTransform::from_xml_element(child_node)?);
            }
        }

        Ok(Self {
            value,
            color_transforms,
        })
    }
}

/// HslColor
pub struct HslColor {
    pub hue: PositiveFixedAngle,
    pub saturation: Percentage,
    pub luminance: Percentage,
    pub color_transforms: Vec<ColorTransform>,
}

impl HslColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<HslColor> {
        let mut opt_h = None;
        let mut opt_s = None;
        let mut opt_l = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "hue" => opt_h = Some(value.parse::<PositiveFixedAngle>()?),
                "sat" => opt_s = Some(value.parse::<Percentage>()?),
                "lum" => opt_l = Some(value.parse::<Percentage>()?),
                _ => (),
            }
        }

        let hue = opt_h.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "hue"))?;
        let saturation = opt_s.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "sat"))?;
        let luminance = opt_l.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "lum"))?;

        let mut color_transforms = Vec::new();

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                color_transforms.push(ColorTransform::from_xml_element(child_node)?);
            }
        }

        Ok(Self {
            hue,
            saturation,
            luminance,
            color_transforms,
        })
    }
}

/// SystemColor
pub struct SystemColor {
    pub value: SystemColorVal,
    pub last_color: Option<HexColorRGB>,
    pub color_transforms: Vec<ColorTransform>,
}

impl SystemColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<SystemColor> {
        let mut opt_val = None;
        let mut last_color = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "val" => opt_val = Some(value.parse::<SystemColorVal>()?),
                "lastClr" => last_color = Some(value.clone()),
                _ => (),
            }
        }

        let value = opt_val.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;

        let mut color_transforms = Vec::new();

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                color_transforms.push(ColorTransform::from_xml_element(child_node)?);
            }
        }

        Ok(Self {
            value,
            last_color,
            color_transforms,
        })
    }
}

/// PresetColor
pub struct PresetColor {
    pub value: PresetColorVal,
    pub color_transforms: Vec<ColorTransform>,
}

impl PresetColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<PresetColor> {
        let attr_val = xml_node
            .attribute("val")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
        let value = attr_val.parse::<PresetColorVal>()?;

        let mut color_transforms = Vec::new();

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                color_transforms.push(ColorTransform::from_xml_element(child_node)?);
            }
        }

        Ok(Self {
            value,
            color_transforms,
        })
    }
}

/// SchemeColor
pub struct SchemeColor {
    pub value: SchemeColorVal,
    pub color_transforms: Vec<ColorTransform>,
}

impl SchemeColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<SchemeColor> {
        let attr_val = xml_node
            .attribute("val")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
        let value = attr_val.parse::<SchemeColorVal>()?;

        let mut color_transforms = Vec::new();

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                color_transforms.push(ColorTransform::from_xml_element(child_node)?);
            }
        }

        Ok(Self {
            value,
            color_transforms,
        })
    }
}

/// Color
pub enum Color {
    ScRgbColor(Box<ScRgbColor>),
    SRgbColor(Box<SRgbColor>),
    HslColor(Box<HslColor>),
    SystemColor(Box<SystemColor>),
    SchemeColor(Box<SchemeColor>),
    PresetColor(Box<PresetColor>),
}

impl Color {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "scrgbClr" | "srgbClr" | "hslClr" | "sysClr" | "schemeClr" | "prstClr" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Color> {
        match xml_node.local_name() {
            "scrgbClr" => Ok(Color::ScRgbColor(Box::new(ScRgbColor::from_xml_element(xml_node)?))),
            "srgbClr" => Ok(Color::SRgbColor(Box::new(SRgbColor::from_xml_element(xml_node)?))),
            "hslClr" => Ok(Color::HslColor(Box::new(HslColor::from_xml_element(xml_node)?))),
            "sysClr" => Ok(Color::SystemColor(Box::new(SystemColor::from_xml_element(xml_node)?))),
            "schemeClr" => Ok(Color::SchemeColor(Box::new(SchemeColor::from_xml_element(xml_node)?))),
            "prstClr" => Ok(Color::PresetColor(Box::new(PresetColor::from_xml_element(xml_node)?))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_ColorChoice").into()),
        }
    }
}

pub struct CustomColor {
    pub color: Color,
    pub name: Option<String>,
}

impl CustomColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name = xml_node.attribute("name").cloned();
        let color_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_ColorChoice"))?;
        let color = Color::from_xml_element(color_node)?;

        Ok(Self { name, color })
    }
}

pub struct ColorMapping {
    pub background1: ColorSchemeIndex,
    pub text1: ColorSchemeIndex,
    pub background2: ColorSchemeIndex,
    pub text2: ColorSchemeIndex,
    pub accent1: ColorSchemeIndex,
    pub accent2: ColorSchemeIndex,
    pub accent3: ColorSchemeIndex,
    pub accent4: ColorSchemeIndex,
    pub accent5: ColorSchemeIndex,
    pub accent6: ColorSchemeIndex,
    pub hyperlink: ColorSchemeIndex,
    pub followed_hyperlink: ColorSchemeIndex,
}

impl ColorMapping {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut background1 = None;
        let mut text1 = None;
        let mut background2 = None;
        let mut text2 = None;
        let mut accent1 = None;
        let mut accent2 = None;
        let mut accent3 = None;
        let mut accent4 = None;
        let mut accent5 = None;
        let mut accent6 = None;
        let mut hyperlink = None;
        let mut followed_hyperlink = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "bg1" => background1 = Some(value.parse()?),
                "tx1" => text1 = Some(value.parse()?),
                "bg2" => background2 = Some(value.parse()?),
                "tx2" => text2 = Some(value.parse()?),
                "accent1" => accent1 = Some(value.parse()?),
                "accent2" => accent2 = Some(value.parse()?),
                "accent3" => accent3 = Some(value.parse()?),
                "accent4" => accent4 = Some(value.parse()?),
                "accent5" => accent5 = Some(value.parse()?),
                "accent6" => accent6 = Some(value.parse()?),
                "hlink" => hyperlink = Some(value.parse()?),
                "folHlink" => followed_hyperlink = Some(value.parse()?),
                _ => (),
            }
        }

        let background1 = background1.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "bg1"))?;
        let text1 = text1.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "tx1"))?;
        let background2 = background2.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "bg2"))?;
        let text2 = text2.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "tx2"))?;
        let accent1 = accent1.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent1"))?;
        let accent2 = accent2.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent2"))?;
        let accent3 = accent3.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent3"))?;
        let accent4 = accent4.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent4"))?;
        let accent5 = accent5.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent5"))?;
        let accent6 = accent6.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent6"))?;
        let hyperlink = hyperlink.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "hlink"))?;
        let followed_hyperlink =
            followed_hyperlink.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "folHlink"))?;

        Ok(Self {
            background1,
            text1,
            background2,
            text2,
            accent1,
            accent2,
            accent3,
            accent4,
            accent5,
            accent6,
            hyperlink,
            followed_hyperlink,
        })
    }
}

pub struct ColorScheme {
    pub name: String,
    pub dark1: Color,
    pub light1: Color,
    pub dark2: Color,
    pub light2: Color,
    pub accent1: Color,
    pub accent2: Color,
    pub accent3: Color,
    pub accent4: Color,
    pub accent5: Color,
    pub accent6: Color,
    pub hyperlink: Color,
    pub followed_hyperlink: Color,
}

impl ColorScheme {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name_attr = xml_node
            .attribute("name")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?;
        let name = name_attr.clone();

        let mut dk1 = None;
        let mut lt1 = None;
        let mut dk2 = None;
        let mut lt2 = None;
        let mut accent1 = None;
        let mut accent2 = None;
        let mut accent3 = None;
        let mut accent4 = None;
        let mut accent5 = None;
        let mut accent6 = None;
        let mut hyperlink = None;
        let mut follow_hyperlink = None;

        for child_node in &xml_node.child_nodes {
            let scheme_node = child_node
                .child_nodes
                .get(0)
                .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "scheme color value"))?;

            match child_node.local_name() {
                "dk1" => dk1 = Some(Color::from_xml_element(&scheme_node)?),
                "lt1" => lt1 = Some(Color::from_xml_element(&scheme_node)?),
                "dk2" => dk2 = Some(Color::from_xml_element(&scheme_node)?),
                "lt2" => lt2 = Some(Color::from_xml_element(&scheme_node)?),
                "accent1" => accent1 = Some(Color::from_xml_element(&scheme_node)?),
                "accent2" => accent2 = Some(Color::from_xml_element(&scheme_node)?),
                "accent3" => accent3 = Some(Color::from_xml_element(&scheme_node)?),
                "accent4" => accent4 = Some(Color::from_xml_element(&scheme_node)?),
                "accent5" => accent5 = Some(Color::from_xml_element(&scheme_node)?),
                "accent6" => accent6 = Some(Color::from_xml_element(&scheme_node)?),
                "hlink" => hyperlink = Some(Color::from_xml_element(&scheme_node)?),
                "folHlink" => follow_hyperlink = Some(Color::from_xml_element(&scheme_node)?),
                _ => (),
            }
        }

        let dark1 = dk1.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "dk1"))?;
        let light1 = lt1.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lt1"))?;
        let dark2 = dk2.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "dk2"))?;
        let light2 = lt2.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lt2"))?;
        let accent1 = accent1.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent1"))?;
        let accent2 = accent2.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent2"))?;
        let accent3 = accent3.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent3"))?;
        let accent4 = accent4.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent4"))?;
        let accent5 = accent5.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent5"))?;
        let accent6 = accent6.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent6"))?;
        let hyperlink = hyperlink.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "hlink"))?;
        let followed_hyperlink =
            follow_hyperlink.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "folHlink"))?;

        Ok(Self {
            name,
            dark1,
            light1,
            dark2,
            light2,
            accent1,
            accent2,
            accent3,
            accent4,
            accent5,
            accent6,
            hyperlink,
            followed_hyperlink,
        })
    }
}

pub enum ColorMappingOverride {
    UseMaster,
    Override(Box<ColorMapping>),
}

impl ColorMappingOverride {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "masterClrMapping" | "overrideClrMapping" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "masterClrMapping" => Ok(ColorMappingOverride::UseMaster),
            "overrideClrMapping" => Ok(ColorMappingOverride::Override(Box::new(
                ColorMapping::from_xml_element(xml_node)?,
            ))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "CT_ColorMappingOverride").into()),
        }
    }
}

pub struct ColorSchemeAndMapping {
    pub color_scheme: Box<ColorScheme>,
    pub color_mapping: Option<Box<ColorMapping>>,
}

impl ColorSchemeAndMapping {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut color_scheme = None;
        let mut color_mapping = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "clrScheme" => color_scheme = Some(Box::new(ColorScheme::from_xml_element(child_node)?)),
                "clrMap" => color_mapping = Some(Box::new(ColorMapping::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let color_scheme = 
            color_scheme.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "clrScheme"))?;

        Ok(Self { color_scheme, color_mapping })
    }
}

/// GradientStop
pub struct GradientStop {
    pub position: PositiveFixedPercentage,
    pub color: Color,
}

impl GradientStop {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let pos_attr = xml_node
            .attribute("pos")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "pos"))?;
        let position = pos_attr.parse::<PositiveFixedPercentage>()?;

        let child_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "color"))?;
        if !Color::is_choice_member(child_node.local_name()) {
            return Err(NotGroupMemberError::new(child_node.name.clone(), "EG_Color").into());
        }

        let color = Color::from_xml_element(child_node)?;
        Ok(Self { position, color })
    }
}

pub struct LinearShadeProperties {
    pub angle: Option<PositiveFixedAngle>,
    pub scaled: Option<bool>,
}

impl LinearShadeProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut angle = None;
        let mut scaled = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "ang" => angle = Some(value.parse::<PositiveFixedAngle>()?),
                "scaled" => scaled = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(Self { angle, scaled })
    }
}

pub struct PathShadeProperties {
    pub path: Option<PathShadeType>,
    pub fill_to_rect: Option<RelativeRect>,
}

impl PathShadeProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let path = match xml_node.attribute("path") {
            Some(val) => Some(val.parse()?),
            None => None,
        };

        let fill_to_rect = match xml_node.child_nodes.get(0) {
            Some(node) => Some(RelativeRect::from_xml_element(node)?),
            None => None,
        };

        Ok(Self { path, fill_to_rect })
    }
}

pub enum ShadeProperties {
    Linear(LinearShadeProperties),
    Path(PathShadeProperties),
}

impl ShadeProperties {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "lin" | "path" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "lin" => Ok(ShadeProperties::Linear(LinearShadeProperties::from_xml_element(
                xml_node,
            )?)),
            "path" => Ok(ShadeProperties::Path(PathShadeProperties::from_xml_element(xml_node)?)),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_ShadeProperties").into()),
        }
    }
}

/// GradientFillProperties
pub struct GradientFillProperties {
    pub flip: Option<TileFlipMode>,
    pub rotate_with_shape: Option<bool>,
    pub gradient_stop_list: Vec<GradientStop>, // length: 2 <= n <= inf
    pub shade_properties: Option<ShadeProperties>,
    pub tile_rect: Option<RelativeRect>,
}

impl GradientFillProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut flip = None;
        let mut rotate_with_shape = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "flip" => flip = Some(value.parse::<TileFlipMode>()?),
                "rotWithShape" => rotate_with_shape = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut gradient_stop_list = Vec::new();
        let mut shade_properties = None;
        let mut tile_rect = None;

        for child_node in &xml_node.child_nodes {
            let local_name = child_node.local_name();
            if ShadeProperties::is_choice_member(local_name) {
                shade_properties = Some(ShadeProperties::from_xml_element(child_node)?);
            } else {
                match child_node.local_name() {
                    "gsLst" => {
                        for gs_node in &child_node.child_nodes {
                            gradient_stop_list.push(GradientStop::from_xml_element(gs_node)?);
                        }
                    }
                    "tileRect" => tile_rect = Some(RelativeRect::from_xml_element(child_node)?),
                    _ => (),
                }
            }
        }

        Ok(Self {
            flip,
            rotate_with_shape,
            gradient_stop_list,
            shade_properties,
            tile_rect,
        })
    }
}

pub struct TileInfoProperties {
    pub translate_x: Option<Coordinate>,
    pub translate_y: Option<Coordinate>,
    pub scale_x: Option<Percentage>,
    pub scale_y: Option<Percentage>,
    pub flip_mode: Option<TileFlipMode>,
    pub alignment: Option<RectAlignment>,
}

impl TileInfoProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut translate_x = None;
        let mut translate_y = None;
        let mut scale_x = None;
        let mut scale_y = None;
        let mut flip_mode = None;
        let mut alignment = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "tx" => translate_x = Some(value.parse()?),
                "ty" => translate_y = Some(value.parse()?),
                "sx" => scale_x = Some(value.parse()?),
                "sy" => scale_y = Some(value.parse()?),
                "flip" => flip_mode = Some(value.parse()?),
                "algn" => alignment = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            translate_x,
            translate_y,
            scale_x,
            scale_y,
            flip_mode,
            alignment,
        })
    }
}

pub struct StretchInfoProperties {
    pub fill_rect: Option<RelativeRect>,
}

impl StretchInfoProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let fill_rect = match xml_node.child_nodes.get(0) {
            Some(rect_node) => Some(RelativeRect::from_xml_element(rect_node)?),
            None => None,
        };

        Ok(Self { fill_rect })
    }
}

pub enum FillModeProperties {
    Tile(Box<TileInfoProperties>),
    Stretch(Box<StretchInfoProperties>),
}

impl FillModeProperties {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "tile" | "stretch" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "tile" => Ok(FillModeProperties::Tile(Box::new(
                TileInfoProperties::from_xml_element(xml_node)?,
            ))),
            "stretch" => Ok(FillModeProperties::Stretch(Box::new(
                StretchInfoProperties::from_xml_element(xml_node)?,
            ))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_FillModeProperties").into()),
        }
    }
}

pub struct BlipFillProperties {
    pub dpi: Option<u32>,
    pub rotate_with_shape: Option<bool>,
    pub blip: Option<Box<Blip>>,
    pub source_rect: Option<RelativeRect>,
    pub fill_mode_properties: Option<FillModeProperties>,
}

impl BlipFillProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut dpi = None;
        let mut rotate_with_shape = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "dpi" => dpi = Some(value.parse()?),
                "rotWithShape" => rotate_with_shape = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut blip = None;
        let mut source_rect = None;
        let mut fill_mode_properties = None;

        for child_node in &xml_node.child_nodes {
            let child_local_name = child_node.local_name();

            if FillModeProperties::is_choice_member(child_local_name) {
                fill_mode_properties = Some(FillModeProperties::from_xml_element(child_node)?);
            } else {
                match child_local_name {
                    "blip" => blip = Some(Box::new(Blip::from_xml_element(child_node)?)),
                    "srcRect" => source_rect = Some(RelativeRect::from_xml_element(child_node)?),
                    _ => (),
                }
            }
        }

        Ok(Self {
            dpi,
            rotate_with_shape,
            blip,
            source_rect,
            fill_mode_properties,
        })
    }
}

pub struct PatternFillProperties {
    pub fg_color: Option<Color>,
    pub bg_color: Option<Color>,
    pub preset: Option<PresetPatternVal>,
}

impl PatternFillProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let preset = match xml_node.attribute("prst") {
            Some(val) => Some(val.parse()?),
            None => None,
        };

        let mut fg_color = None;
        let mut bg_color = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "fgClr" => {
                    let fg_color_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?;
                    fg_color = Some(Color::from_xml_element(fg_color_node)?);
                }
                "bgClr" => {
                    let bg_color_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?;
                    bg_color = Some(Color::from_xml_element(bg_color_node)?);
                }
                _ => (),
            }
        }

        Ok(Self { preset, fg_color, bg_color })
    }
}

pub enum FillProperties {
    NoFill,
    SolidFill(Color),
    GradientFill(Box<GradientFillProperties>),
    BlipFill(Box<BlipFillProperties>),
    PatternFill(Box<PatternFillProperties>),
    GroupFill,
}

impl FillProperties {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "noFill" | "solidFill" | "gradFill" | "blipFill" | "pattFill" | "grpFill" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "noFill" => Ok(FillProperties::NoFill),
            "solidFill" => {
                let child_node = xml_node
                    .child_nodes
                    .get(0)
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "color"))?;
                Ok(FillProperties::SolidFill(Color::from_xml_element(&child_node)?))
            }
            "gradFill" => Ok(FillProperties::GradientFill(Box::new(
                GradientFillProperties::from_xml_element(xml_node)?,
            ))),
            "blipFill" => Ok(FillProperties::BlipFill(Box::new(
                BlipFillProperties::from_xml_element(xml_node)?,
            ))),
            "pattFill" => Ok(FillProperties::PatternFill(Box::new(
                PatternFillProperties::from_xml_element(xml_node)?,
            ))),
            "grpFill" => Ok(FillProperties::GroupFill),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_FillProperties").into()),
        }
    }
}

/// LineFillProperties
pub enum LineFillProperties {
    NoFill,
    SolidFill(Color),
    GradientFill(Box<GradientFillProperties>),
    PatternFill(Box<PatternFillProperties>),
}

impl LineFillProperties {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "noFill" | "solidFill" | "gradFill" | "pattFill" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<LineFillProperties> {
        match xml_node.local_name() {
            "noFill" => Ok(LineFillProperties::NoFill),
            "solidFill" => {
                let child_node = xml_node
                    .child_nodes
                    .get(0)
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "color"))?;

                if !Color::is_choice_member(child_node.local_name()) {
                    return Err(NotGroupMemberError::new(child_node.name.clone(), "EG_Color").into());
                }

                Ok(LineFillProperties::SolidFill(Color::from_xml_element(child_node)?))
            }
            "gradFill" => Ok(LineFillProperties::GradientFill(Box::new(
                GradientFillProperties::from_xml_element(xml_node)?,
            ))),
            "pattFill" => Ok(LineFillProperties::PatternFill(Box::new(
                PatternFillProperties::from_xml_element(xml_node)?,
            ))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_LineFillProperties").into()),
        }
    }
}

/// DashStop
pub struct DashStop {
    pub dash_length: PositivePercentage,
    pub space_length: PositivePercentage,
}

impl DashStop {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<DashStop> {
        let mut opt_dash_length = None;
        let mut opt_space_length = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "d" => opt_dash_length = Some(value.parse::<PositivePercentage>()?),
                "sp" => opt_space_length = Some(value.parse::<PositivePercentage>()?),
                _ => (),
            }
        }

        let dash_length = opt_dash_length.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "d"))?;
        let space_length = opt_space_length.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "sp"))?;

        Ok(Self {
            dash_length,
            space_length,
        })
    }
}

/// LineDashProperties
pub enum LineDashProperties {
    PresetDash(PresetLineDashVal),
    CustomDash(Vec<DashStop>),
}

impl LineDashProperties {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "prstDash" | "custDash" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<LineDashProperties> {
        match xml_node.local_name() {
            "prstDash" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(LineDashProperties::PresetDash(val_attr.parse::<PresetLineDashVal>()?))
            }
            "custDash" => {
                let mut dash_vec = Vec::new();
                for child_node in &xml_node.child_nodes {
                    if child_node.local_name() == "ds" {
                        match DashStop::from_xml_element(child_node) {
                            Ok(val) => dash_vec.push(val),
                            Err(err) => println!("Failed to parse 'ds' element: {}", err),
                        }
                    }
                }

                Ok(LineDashProperties::CustomDash(dash_vec))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_LineDashProperties").into()),
        }
    }
}

/// LineJoinProperties
pub enum LineJoinProperties {
    Round,
    Bevel,
    Miter(Option<PositivePercentage>),
}

impl LineJoinProperties {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "round" | "bevel" | "miter" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<LineJoinProperties> {
        match xml_node.local_name() {
            "round" => Ok(LineJoinProperties::Round),
            "bevel" => Ok(LineJoinProperties::Bevel),
            "miter" => {
                let lim = match xml_node.attribute("lim") {
                    Some(ref attr) => Some(attr.parse::<PositivePercentage>()?),
                    None => None,
                };
                Ok(LineJoinProperties::Miter(lim))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_LineJoinProperties").into()),
        }
    }
}

/// LineEndProperties
pub struct LineEndProperties {
    pub end_type: Option<LineEndType>,
    pub width: Option<LineEndWidth>,
    pub length: Option<LineEndLength>,
}

impl LineEndProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<LineEndProperties> {
        let mut end_type = None;
        let mut width = None;
        let mut length = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => end_type = Some(value.parse::<LineEndType>()?),
                "width" => width = Some(value.parse::<LineEndWidth>()?),
                "length" => length = Some(value.parse::<LineEndLength>()?),
                _ => (),
            }
        }

        Ok(Self {
            end_type,
            width,
            length,
        })
    }
}

/// LineProperties
pub struct LineProperties {
    pub width: Option<LineWidth>,
    pub cap: Option<LineCap>,
    pub compound: Option<CompoundLine>,
    pub pen_alignment: Option<PenAlignment>,
    pub fill_properties: Option<LineFillProperties>,
    pub dash_properties: Option<LineDashProperties>,
    pub join_properties: Option<LineJoinProperties>,
    pub head_end: Option<LineEndProperties>,
    pub tail_end: Option<LineEndProperties>,
}

impl LineProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<LineProperties> {
        let mut width = None;
        let mut cap = None;
        let mut compound = None;
        let mut pen_alignment = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "w" => width = Some(value.parse::<LineWidth>()?),
                "cap" => cap = Some(value.parse::<LineCap>()?),
                "cmpd" => compound = Some(value.parse::<CompoundLine>()?),
                "algn" => pen_alignment = Some(value.parse::<PenAlignment>()?),
                _ => (),
            }
        }

        let mut fill_properties = None;
        let mut dash_properties = None;
        let mut join_properties = None;
        let mut head_end = None;
        let mut tail_end = None;

        for child_node in &xml_node.child_nodes {
            if LineFillProperties::is_choice_member(child_node.local_name()) {
                fill_properties = Some(LineFillProperties::from_xml_element(child_node)?);
            } else if LineDashProperties::is_choice_member(child_node.local_name()) {
                dash_properties = Some(LineDashProperties::from_xml_element(child_node)?);
            } else if LineJoinProperties::is_choice_member(child_node.local_name()) {
                join_properties = Some(LineJoinProperties::from_xml_element(child_node)?);
            } else {
                match child_node.local_name() {
                    "headEnd" => head_end = Some(LineEndProperties::from_xml_element(child_node)?),
                    "tailEnd" => tail_end = Some(LineEndProperties::from_xml_element(child_node)?),
                    _ => (),
                }
            }
        }

        Ok(Self {
            width,
            cap,
            compound,
            pen_alignment,
            fill_properties,
            dash_properties,
            join_properties,
            head_end,
            tail_end,
        })
    }
}

/// RelativeRect
pub struct RelativeRect {
    pub left: Option<Percentage>,
    pub top: Option<Percentage>,
    pub right: Option<Percentage>,
    pub bottom: Option<Percentage>,
}

impl RelativeRect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<RelativeRect> {
        let mut left = None;
        let mut top = None;
        let mut right = None;
        let mut bottom = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "l" => left = Some(value.parse::<Percentage>()?),
                "t" => top = Some(value.parse::<Percentage>()?),
                "r" => right = Some(value.parse::<Percentage>()?),
                "b" => bottom = Some(value.parse::<Percentage>()?),
                _ => (),
            }
        }

        Ok(Self {
            left,
            top,
            right,
            bottom,
        })
    }
}

pub struct Point2D {
    pub x: Coordinate,
    pub y: Coordinate,
}

impl Point2D {
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

/// PositiveSize2D
pub struct PositiveSize2D {
    pub width: PositiveCoordinate,
    pub height: PositiveCoordinate,
}

impl PositiveSize2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut opt_width = None;
        let mut opt_height = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "cx" => opt_width = Some(value.parse::<PositiveCoordinate>()?),
                "cy" => opt_height = Some(value.parse::<PositiveCoordinate>()?),
                _ => (),
            }
        }

        let width = opt_width.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "cx"))?;
        let height = opt_height.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "cy"))?;

        Ok(Self { width, height })
    }
}

pub struct StyleMatrixReference {
    pub index: StyleMatrixColumnIndex,
    pub color: Option<Color>,
}

impl StyleMatrixReference {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let idx_attr = xml_node
            .attribute("idx")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "idx"))?;
        let index = idx_attr.parse()?;

        let color = match xml_node.child_nodes.get(0) {
            Some(node) => Some(Color::from_xml_element(node)?),
            None => None,
        };

        Ok(Self { index, color })
    }
}

/// EffectContainer
pub struct EffectContainer {
    pub container_type: Option<EffectContainerType>,
    pub name: Option<String>,
    pub effects: Vec<Effect>,
}

impl EffectContainer {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<EffectContainer> {
        let mut container_type = None;
        let mut name = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => container_type = Some(value.parse::<EffectContainerType>()?),
                "name" => name = Some(value.clone()),
                _ => (),
            }
        }

        let mut effects = Vec::new();
        for child_node in &xml_node.child_nodes {
            if Effect::is_choice_member(child_node.local_name()) {
                effects.push(Effect::from_xml_element(child_node)?);
            }
        }

        Ok(Self {
            container_type,
            name,
            effects,
        })
    }
}

/// AlphaBiLevelEffect
pub struct AlphaBiLevelEffect {
    pub threshold: PositiveFixedPercentage,
}

impl AlphaBiLevelEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<AlphaBiLevelEffect> {
        let thresh_attr = xml_node
            .attribute("thresh")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "thresh"))?;
        let threshold = thresh_attr.parse::<PositiveFixedPercentage>()?;
        Ok(Self { threshold })
    }
}

/// AlphaInverseEffect
pub struct AlphaInverseEffect {
    pub color: Option<Color>,
}

impl AlphaInverseEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<AlphaInverseEffect> {
        let color = match xml_node.child_nodes.get(0) {
            Some(child_node) => Some(Color::from_xml_element(child_node)?),
            None => None,
        };

        Ok(Self { color })
    }
}

/// AlphaModulateEffect
pub struct AlphaModulateEffect {
    pub container: EffectContainer,
}

impl AlphaModulateEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<AlphaModulateEffect> {
        let child_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "container"))?;

        let container = EffectContainer::from_xml_element(child_node)?;

        Ok(Self { container })
    }
}

/// AlphaModulateFixedEffect
pub struct AlphaModulateFixedEffect {
    pub amount: Option<PositivePercentage>, // 1.0
}

impl AlphaModulateFixedEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let amount = match xml_node.attribute("amt") {
            Some(attr) => Some(attr.parse()?),
            None => None,
        };

        Ok(Self { amount })
    }
}

pub struct AlphaOutsetEffect {
    pub radius: Option<Coordinate>,
}

impl AlphaOutsetEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let radius = match xml_node.attribute("rad") {
            Some(attr) => Some(attr.parse()?),
            None => None,
        };

        Ok(Self { radius })
    }
}

pub struct AlphaReplaceEffect {
    pub alpha: PositiveFixedPercentage,
}

impl AlphaReplaceEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let a_attr = xml_node
            .attribute("a")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "a"))?;
        let alpha = a_attr.parse()?;

        Ok(Self { alpha })
    }
}

pub struct BiLevelEffect {
    pub threshold: PositiveFixedPercentage,
}

impl BiLevelEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let thresh_attr = xml_node
            .attribute("thresh")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "thresh"))?;
        let threshold = thresh_attr.parse()?;

        Ok(Self { threshold })
    }
}

pub struct BlendEffect {
    pub blend: BlendMode,
    pub container: EffectContainer,
}

impl BlendEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let blend_attr = xml_node
            .attribute("blend")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "blend"))?;
        let blend = blend_attr.parse()?;

        let container_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cont"))?;
        let container = EffectContainer::from_xml_element(container_node)?;

        Ok(Self { blend, container })
    }
}

pub struct BlurEffect {
    pub radius: Option<PositiveCoordinate>, // 0
    pub grow: Option<bool>,                 // true
}

impl BlurEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut radius = None;
        let mut grow = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "rad" => radius = Some(value.parse()?),
                "grow" => grow = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(Self { radius, grow })
    }
}

pub struct ColorChangeEffect {
    pub use_alpha: Option<bool>, // true
    pub color_from: Color,
    pub color_to: Color,
}

impl ColorChangeEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let use_alpha = match xml_node.attribute("useA") {
            Some(attr) => Some(parse_xml_bool(attr)?),
            None => None,
        };

        let mut color_from = None;
        let mut color_to = None;
        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "clrFrom" => {
                    let color_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?;
                    color_from = Some(Color::from_xml_element(color_node)?);
                }
                "clrTo" => {
                    let color_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?;
                    color_to = Some(Color::from_xml_element(color_node)?);
                }
                _ => (),
            }
        }

        let color_from = color_from.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "clrFrom"))?;
        let color_to = color_to.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "clrTo"))?;

        Ok(Self { use_alpha, color_from, color_to })
    }
}

pub struct ColorReplaceEffect {
    pub color: Color,
}

impl ColorReplaceEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let color_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_Color"))?;
        let color = Color::from_xml_element(color_node)?;

        Ok(Self { color })
    }
}

pub struct LuminanceEffect {
    pub brightness: Option<FixedPercentage>,
    pub contrast: Option<FixedPercentage>,
}

impl LuminanceEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut brightness = None;
        let mut contrast = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "bright" => brightness = Some(value.parse()?),
                "contrast" => contrast = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self { brightness, contrast })
    }
}

pub struct DuotoneEffect {
    pub colors: [Color; 2],
}

impl DuotoneEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let color_1_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_Color"))?;
        let color_2_node = xml_node
            .child_nodes
            .get(1)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_Color"))?;

        let color_1 = Color::from_xml_element(color_1_node)?;
        let color_2 = Color::from_xml_element(color_2_node)?;

        Ok(Self { colors: [color_1, color_2] })
    }
}

pub struct FillEffect {
    pub fill_properties: FillProperties,
}

impl FillEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let fill_properties_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_FillProperties"))?;
        let fill_properties = FillProperties::from_xml_element(fill_properties_node)?;

        Ok(Self { fill_properties })
    }
}

pub struct FillOverlayEffect {
    pub blend_mode: BlendMode,
    pub fill: FillProperties,
}

impl FillOverlayEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let blend_mode_attr = xml_node
            .attribute("blend")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "blend"))?;
        let blend_mode = blend_mode_attr.parse()?;

        let fill_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_FillProperties"))?;
        let fill = FillProperties::from_xml_element(fill_node)?;
        
        Ok(Self { blend_mode, fill })
    }
}

pub struct GlowEffect {
    pub radius: Option<PositiveCoordinate>, // 0
    pub color: Color,
}

impl GlowEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let radius = match xml_node.attribute("rad") {
            Some(attr) => Some(attr.parse()?),
            None => None,
        };

        let color_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_ColorChoice"))?;
        let color = Color::from_xml_element(color_node)?;

        Ok(Self { radius, color })
    }
}

pub struct HslEffect {
    pub hue: Option<PositiveFixedAngle>,     // 0
    pub saturation: Option<FixedPercentage>, // 0%
    pub luminance: Option<FixedPercentage>,  // 0%
}

impl HslEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut hue = None;
        let mut saturation = None;
        let mut luminance = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "hue" => hue = Some(value.parse()?),
                "sat" => saturation = Some(value.parse()?),
                "lum" => luminance = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self { hue, saturation, luminance })
    }
}

pub struct InnerShadowEffect {
    pub blur_radius: Option<PositiveCoordinate>, // 0
    pub distance: Option<PositiveCoordinate>,    // 0
    pub direction: Option<PositiveFixedAngle>,   // 0
    pub color: Color,
}

impl InnerShadowEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut blur_radius = None;
        let mut distance = None;
        let mut direction = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "blurRad" => blur_radius = Some(value.parse()?),
                "dist" => distance = Some(value.parse()?),
                "dir" => direction = Some(value.parse()?),
                _ => (),
            }
        }

        let color_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_ColorChoice"))?;
        let color = Color::from_xml_element(color_node)?;

        Ok(Self {
            blur_radius,
            distance,
            direction,
            color,
        })
    }
}

pub struct OuterShadowEffect {
    pub blur_radius: Option<PositiveCoordinate>, // 0
    pub distance: Option<PositiveCoordinate>,    // 0
    pub direction: Option<PositiveFixedAngle>,   // 0
    pub scale_x: Option<Percentage>,             // 100000
    pub scale_y: Option<Percentage>,             // 100000
    pub skew_x: Option<FixedAngle>,              // 0
    pub skew_y: Option<FixedAngle>,              // 0
    pub alignment: Option<RectAlignment>,        // b
    pub rotate_with_shape: Option<bool>,         // true
    pub color: Color,
}

impl OuterShadowEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut blur_radius = None;
        let mut distance = None;
        let mut direction = None;
        let mut scale_x = None;
        let mut scale_y = None;
        let mut skew_x = None;
        let mut skew_y = None;
        let mut alignment = None;
        let mut rotate_with_shape = None;
        
        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "blurRad" => blur_radius = Some(value.parse()?),
                "dist" => distance = Some(value.parse()?),
                "dir" => direction = Some(value.parse()?),
                "sx" => scale_x = Some(value.parse()?),
                "sy" => scale_y = Some(value.parse()?),
                "kx" => skew_x = Some(value.parse()?),
                "ky" => skew_y = Some(value.parse()?),
                "algn" => alignment = Some(value.parse()?),
                "rotWithShape" => rotate_with_shape = Some(value.parse()?),
                _ => ()
            }
        }

        let color_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_ColorChoice"))?;
        let color = Color::from_xml_element(color_node)?;

        Ok(Self {
            blur_radius,
            distance,
            direction,
            scale_x,
            scale_y,
            skew_x,
            skew_y,
            alignment,
            rotate_with_shape,
            color,
        })
    }
}

pub struct PresetShadowEffect {
    pub preset: PresetShadowVal,
    pub distance: Option<PositiveCoordinate>,  // 0
    pub direction: Option<PositiveFixedAngle>, // 0
    pub color: Color,
}

impl PresetShadowEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut preset = None;
        let mut distance = None;
        let mut direction = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "prst" => preset = Some(value.parse()?),
                "dist" => distance = Some(value.parse()?),
                "dir" => direction = Some(value.parse()?),
                _ => (),
            }
        }

        let color_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_ColorChoice"))?;
        let color = Color::from_xml_element(color_node)?;

        let preset = preset.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "prst"))?;

        Ok(Self { preset, distance, direction, color })
    }
}

pub struct ReflectionEffect {
    pub blur_radius: Option<PositiveCoordinate>,         // 0
    pub start_opacity: Option<PositiveFixedPercentage>,  // 100000
    pub start_position: Option<PositiveFixedPercentage>, // 0
    pub end_opacity: Option<PositiveFixedPercentage>,    // 0
    pub end_position: Option<PositiveFixedPercentage>,   // 100000
    pub distance: Option<PositiveCoordinate>,            // 0
    pub direction: Option<PositiveFixedAngle>,           // 0
    pub fade_direction: Option<PositiveFixedAngle>,      // 5400000
    pub scale_x: Option<Percentage>,                     // 100000
    pub scale_y: Option<Percentage>,                     // 100000
    pub skew_x: Option<FixedAngle>,                      // 0
    pub skew_y: Option<FixedAngle>,                      // 0
    pub alignment: Option<RectAlignment>,                // b
    pub rotate_with_shape: Option<bool>,                 // true
}

impl ReflectionEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut blur_radius = None;
        let mut start_opacity = None;
        let mut start_position = None;
        let mut end_opacity = None;
        let mut end_position = None;
        let mut distance = None;
        let mut direction = None;
        let mut fade_direction = None;
        let mut scale_x = None;
        let mut scale_y = None;
        let mut skew_x = None;
        let mut skew_y = None;
        let mut alignment = None;
        let mut rotate_with_shape = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "blurRad" => blur_radius = Some(value.parse()?),
                "stA" => start_opacity = Some(value.parse()?),
                "stPos" => start_position = Some(value.parse()?),
                "endA" => end_opacity = Some(value.parse()?),
                "endPos" => end_position = Some(value.parse()?),
                "dist" => distance = Some(value.parse()?),
                "dir" => direction = Some(value.parse()?),
                "fadeDir" => fade_direction = Some(value.parse()?),
                "sx" => scale_x = Some(value.parse()?),
                "sy" => scale_y = Some(value.parse()?),
                "kx" => skew_x = Some(value.parse()?),
                "ky" => skew_y = Some(value.parse()?),
                "algn" => alignment = Some(value.parse()?),
                "rotWithShape" => rotate_with_shape = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            blur_radius,
            start_opacity,
            start_position,
            end_opacity,
            end_position,
            distance,
            direction,
            fade_direction,
            scale_x,
            scale_y,
            skew_x,
            skew_y,
            alignment,
            rotate_with_shape,
        })
    }
}

pub struct RelativeOffsetEffect {
    pub translate_x: Option<Percentage>, // 0
    pub translate_y: Option<Percentage>, // 0
}

impl RelativeOffsetEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut translate_x = None;
        let mut translate_y = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "tx" => translate_x = Some(value.parse()?),
                "ty" => translate_y = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self { translate_x, translate_y })
    }
}

pub struct SoftEdgesEffect {
    pub radius: PositiveCoordinate,
}

impl SoftEdgesEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let radius_attr = xml_node
            .attribute("rad")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "rad"))?;
        
        let radius = radius_attr.parse()?;

        Ok(Self { radius })
    }
}

pub struct TintEffect {
    pub hue: Option<PositiveFixedAngle>, // 0
    pub amount: Option<FixedPercentage>, // 0
}

impl TintEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut hue = None;
        let mut amount = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "hue" => hue = Some(value.parse()?),
                "amt" => amount = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self { hue, amount })
    }
}

pub struct TransformEffect {
    pub scale_x: Option<Percentage>,     // 100000
    pub scale_y: Option<Percentage>,     // 100000
    pub translate_x: Option<Coordinate>, // 0
    pub translate_y: Option<Coordinate>, // 0
    pub skew_x: Option<FixedAngle>,      // 0
    pub skew_y: Option<FixedAngle>,      // 0
}

impl TransformEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut scale_x = None;
        let mut scale_y = None;
        let mut translate_x = None;
        let mut translate_y = None;
        let mut skew_x = None;
        let mut skew_y = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "sx" => scale_x = Some(value.parse()?),
                "sy" => scale_y = Some(value.parse()?),
                "kx" => skew_x = Some(value.parse()?),
                "ky" => skew_y = Some(value.parse()?),
                "tx" => translate_x = Some(value.parse()?),
                "ty" => translate_y = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            scale_x,
            scale_y,
            translate_x,
            translate_y,
            skew_x,
            skew_y,
        })
    }
}

// TODO: maybe Box ReflectionEffect variant (sizeof==120)
pub enum Effect {
    Container(EffectContainer),
    EffectReference(String),
    AlphaBiLevel(AlphaBiLevelEffect),
    AlphaCeiling,
    AlphaFloor,
    AlphaInverse(AlphaInverseEffect),
    AlphaModulate(AlphaModulateEffect),
    AlphaModulateFixed(AlphaModulateFixedEffect),
    AlphaOutset(AlphaOutsetEffect),
    AlphaReplace(AlphaReplaceEffect),
    BiLevel(BiLevelEffect),
    Blend(BlendEffect),
    Blur(BlurEffect),
    ColorChange(ColorChangeEffect),
    ColorReplace(ColorReplaceEffect),
    Duotone(DuotoneEffect),
    Fill(FillEffect),
    FillOverlay(FillOverlayEffect),
    Glow(GlowEffect),
    Grayscale,
    Hsl(HslEffect),
    InnerShadow(InnerShadowEffect),
    Luminance(LuminanceEffect),
    OuterShadow(OuterShadowEffect),
    PresetShadow(PresetShadowEffect),
    Reflection(ReflectionEffect),
    RelativeOffset(RelativeOffsetEffect),
    SoftEdges(SoftEdgesEffect),
    Tint(TintEffect),
    Transform(TransformEffect),
}

impl Effect {
    pub fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>
    {
        match name.as_ref() {
            "cont" | "effect" | "alphaBiLevel" | "alphaCeiling" | "alphaFloor" | "alphaInv" | "alphaMod"
            | "alphaModFix" | "alphaOutset" | "alphaRepl" | "biLevel" | "blend" | "blur" | "clrChange" | "clrRepl"
            | "duotone" | "fill" | "fillOverlay" | "glow" | "grayscl" | "hsl" | "innerShdw" | "lum" | "outerShdw"
            | "prstShdw" | "reflection" | "relOff" | "softEdge" | "tint" | "xfrm" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "cont" => Ok(Effect::Container(EffectContainer::from_xml_element(xml_node)?)),
            "effect" => {
                let ref_attr = xml_node
                    .attribute("ref")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "ref"))?;
                Ok(Effect::EffectReference(ref_attr.clone()))
            }
            "alphaBiLevel" => Ok(Effect::AlphaBiLevel(AlphaBiLevelEffect::from_xml_element(xml_node)?)),
            "alphaCeiling" => Ok(Effect::AlphaCeiling),
            "alphaFloor" => Ok(Effect::AlphaFloor),
            "alphaInv" => Ok(Effect::AlphaInverse(AlphaInverseEffect::from_xml_element(xml_node)?)),
            "alphaMod" => Ok(Effect::AlphaModulate(AlphaModulateEffect::from_xml_element(xml_node)?)),
            "alphaModFix" => Ok(Effect::AlphaModulateFixed(AlphaModulateFixedEffect::from_xml_element(xml_node)?)),
            "alphaOutset" => Ok(Effect::AlphaOutset(AlphaOutsetEffect::from_xml_element(xml_node)?)),
            "alphaRepl" => Ok(Effect::AlphaReplace(AlphaReplaceEffect::from_xml_element(xml_node)?)),
            "biLevel" => Ok(Effect::BiLevel(BiLevelEffect::from_xml_element(xml_node)?)),
            "blend" => Ok(Effect::Blend(BlendEffect::from_xml_element(xml_node)?)),
            "blur" => Ok(Effect::Blur(BlurEffect::from_xml_element(xml_node)?)),
            "clrChange" => Ok(Effect::ColorChange(ColorChangeEffect::from_xml_element(xml_node)?)),
            "clrRepl" => Ok(Effect::ColorReplace(ColorReplaceEffect::from_xml_element(xml_node)?)),
            "duotone" => Ok(Effect::Duotone(DuotoneEffect::from_xml_element(xml_node)?)),
            "fill" => Ok(Effect::Fill(FillEffect::from_xml_element(xml_node)?)),
            "fillOverlay" => Ok(Effect::FillOverlay(FillOverlayEffect::from_xml_element(xml_node)?)),
            "glow" => Ok(Effect::Glow(GlowEffect::from_xml_element(xml_node)?)),
            "grayscl" => Ok(Effect::Grayscale),
            "hsl" => Ok(Effect::Hsl(HslEffect::from_xml_element(xml_node)?)),
            "innerShdw" => Ok(Effect::InnerShadow(InnerShadowEffect::from_xml_element(xml_node)?)),
            "lum" => Ok(Effect::Luminance(LuminanceEffect::from_xml_element(xml_node)?)),
            "outerShdw" => Ok(Effect::OuterShadow(OuterShadowEffect::from_xml_element(xml_node)?)),
            "prstShdw" => Ok(Effect::PresetShadow(PresetShadowEffect::from_xml_element(xml_node)?)),
            "reflection" => Ok(Effect::Reflection(ReflectionEffect::from_xml_element(xml_node)?)),
            "relOff" => Ok(Effect::RelativeOffset(RelativeOffsetEffect::from_xml_element(xml_node)?)),
            "softEdge" => Ok(Effect::SoftEdges(SoftEdgesEffect::from_xml_element(xml_node)?)),
            "tint" => Ok(Effect::Tint(TintEffect::from_xml_element(xml_node)?)),
            "xfrm" => Ok(Effect::Transform(TransformEffect::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "EG_Effect"))),
        }
    }
}

pub struct EffectList {
    pub blur: Option<BlurEffect>,
    pub fill_overlay: Option<FillOverlayEffect>,
    pub glow: Option<GlowEffect>,
    pub inner_shadow: Option<InnerShadowEffect>,
    pub outer_shadow: Option<OuterShadowEffect>,
    pub preset_shadow: Option<PresetShadowEffect>,
    pub reflection: Option<ReflectionEffect>,
    pub soft_edges: Option<SoftEdgesEffect>,
}

pub enum EffectProperties {
    EffectList(Box<EffectList>),
    EffectContainer(Box<EffectContainer>),
}

pub struct EffectStyleItem {
    pub effect_props: EffectProperties,
    //pub scene_3d: Option<Scene3D>,
    //pub shape_3d: Option<Shape3D>,
}

/// BlipEffect
pub enum BlipEffect {
    AlphaBiLevel(AlphaBiLevelEffect),
    AlphaCeiling,
    AlphaFloor,
    AlphaInv(AlphaInverseEffect),
    AlphaMod(AlphaModulateEffect),
    AlphaModFix(AlphaModulateFixedEffect),
    AlphaRepl(AlphaReplaceEffect),
    BiLevel(BiLevelEffect),
    Blur(BlurEffect),
    ClrChange(ColorChangeEffect),
    ClrRepl(ColorReplaceEffect),
    Duotone(DuotoneEffect),
    FillOverlay(FillOverlayEffect),
    Grayscl,
    Hsl(HslEffect),
    Lum(LuminanceEffect),
    Tint(TintEffect),
}

impl BlipEffect {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "alphaBiLevel" | "alphaCeiling" | "alphaFloor" | "alphaInv" | "alphaMod" | "alphaModFixed" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<BlipEffect> {
        match xml_node.local_name() {
            "alphaBiLevel" => Ok(BlipEffect::AlphaBiLevel(AlphaBiLevelEffect::from_xml_element(
                xml_node,
            )?)),
            "alphaCeiling" => Ok(BlipEffect::AlphaCeiling),
            "alphaFloor" => Ok(BlipEffect::AlphaFloor),
            "alphaInv" => Ok(BlipEffect::AlphaInv(AlphaInverseEffect::from_xml_element(xml_node)?)),
            "alphaMod" => Ok(BlipEffect::AlphaMod(AlphaModulateEffect::from_xml_element(xml_node)?)),
            "alphaModFixed" => Ok(BlipEffect::AlphaModFix(AlphaModulateFixedEffect::from_xml_element(
                xml_node,
            )?)),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_BlipEffect").into()),
        }
    }
}

/// Blip
pub struct Blip {
    pub embed_rel_id: Option<RelationshipId>,
    pub linked_rel_id: Option<RelationshipId>,
    pub compression: Option<BlipCompression>,
    pub effects: Vec<BlipEffect>,
}

impl Blip {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut embed_rel_id = None;
        let mut linked_rel_id = None;
        let mut compression = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r:embed" => embed_rel_id = Some(value.clone()),
                "r:link" => linked_rel_id = Some(value.clone()),
                "cstate" => compression = Some(value.parse::<BlipCompression>()?),
                _ => (),
            }
        }

        let mut effects = Vec::new();

        for child_node in &xml_node.child_nodes {
            if BlipEffect::is_choice_member(child_node.local_name()) {
                effects.push(BlipEffect::from_xml_element(child_node)?);
            }
        }

        Ok(Self {
            embed_rel_id,
            linked_rel_id,
            compression,
            effects,
        })
    }
}

/// TextFont
pub struct TextFont {
    pub typeface: TextTypeFace,
    pub panose: Option<Panose>,
    pub pitch_family: Option<i32>, // 0
    pub charset: Option<i32>,      // 1
}

impl TextFont {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextFont> {
        let mut typeface = None;
        let mut panose = None;
        let mut pitch_family = None;
        let mut charset = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "typeface" => typeface = Some(value.clone()),
                "panose" => panose = Some(value.clone()),
                "pitchFamily" => pitch_family = Some(value.parse::<i32>()?),
                "charset" => charset = Some(value.parse::<i32>()?),
                _ => (),
            }
        }

        let typeface = typeface.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "typeface"))?;

        Ok(Self {
            typeface,
            panose,
            pitch_family,
            charset,
        })
    }
}

pub struct SupplementalFont {
    pub script: String,
    pub typeface: TextTypeFace,
}

impl SupplementalFont {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut script = None;
        let mut typeface = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "script" => script = Some(value.clone()),
                "typeface" => typeface = Some(value.clone()),
                _ => (),
            }
        }

        let script = script.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "script"))?;
        let typeface = typeface.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "typeface"))?;

        Ok(Self { script, typeface })
    }
}

/// TextSpacing
pub enum TextSpacing {
    Percent(TextSpacingPercent),
    Point(TextSpacingPoint),
}

impl TextSpacing {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextSpacing> {
        match xml_node.local_name() {
            "spcPct" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TextSpacing::Percent(val_attr.parse::<TextSpacingPercent>()?))
            }
            "spcPts" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TextSpacing::Point(val_attr.parse::<TextSpacingPoint>()?))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextSpacing").into()),
        }
    }
}

/// TextBulletColor
pub enum TextBulletColor {
    FollowText,
    Color(Color),
}

impl TextBulletColor {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "buClrTx" | "buClr" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextBulletColor> {
        match xml_node.local_name() {
            "buClrTx" => Ok(TextBulletColor::FollowText),
            "buClr" => {
                let child_node = xml_node
                    .child_nodes
                    .get(0)
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "color"))?;
                Ok(TextBulletColor::Color(Color::from_xml_element(child_node)?))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBulletColor").into()),
        }
    }
}

/// TextBulletSize
pub enum TextBulletSize {
    FollowText,
    Percent(TextBulletSizePercent),
    Point(TextFontSize),
}

impl TextBulletSize {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "buSzTx" | "buSzPct" | "buSzPts" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextBulletSize> {
        match xml_node.local_name() {
            "buSzTx" => Ok(TextBulletSize::FollowText),
            "buSzPct" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TextBulletSize::Percent(val_attr.parse::<TextBulletSizePercent>()?))
            }
            "buSzPts" => {
                let val_attr = xml_node
                    .attribute("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(TextBulletSize::Point(val_attr.parse::<TextFontSize>()?))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBulletSize").into()),
        }
    }
}

/// TextBulletTypeface
pub enum TextBulletTypeface {
    FollowText,
    Font(TextFont),
}

impl TextBulletTypeface {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "buFontTx" | "buFont" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextBulletTypeface> {
        match xml_node.local_name() {
            "buFontTx" => Ok(TextBulletTypeface::FollowText),
            "buFont" => Ok(TextBulletTypeface::Font(TextFont::from_xml_element(xml_node)?)),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBulletTypeface").into()),
        }
    }
}

/// TextBullet
pub enum TextBullet {
    None,
    AutoNumbered(TextAutonumberedBullet),
    Character(String),
    Picture(Box<Blip>),
}

impl TextBullet {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "buNone" | "buAutoNum" | "buChar" | "buBlip" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextBullet> {
        match xml_node.local_name() {
            "buNone" => Ok(TextBullet::None),
            "buAutoNum" => Ok(TextBullet::AutoNumbered(TextAutonumberedBullet::from_xml_element(
                xml_node,
            )?)),
            "buChar" => {
                let char_attr = xml_node
                    .attribute("char")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "char"))?;
                Ok(TextBullet::Character(char_attr.clone()))
            }
            "buBlip" => match xml_node.child_nodes.get(0) {
                Some(child_node) => Ok(TextBullet::Picture(Box::new(Blip::from_xml_element(child_node)?))),
                None => Err(MissingChildNodeError::new(xml_node.name.clone(), "EG_TextBullet").into()),
            },
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBullet").into()),
        }
    }
}

/// TextAutonumberedBullet
pub struct TextAutonumberedBullet {
    pub scheme: TextAutonumberScheme,
    pub start_at: Option<TextBulletStartAtNum>,
}

impl TextAutonumberedBullet {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextAutonumberedBullet> {
        let mut scheme = None;
        let mut start_at = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => scheme = Some(value.parse::<TextAutonumberScheme>()?),
                "startAt" => start_at = Some(value.parse::<TextBulletStartAtNum>()?),
                _ => (),
            }
        }

        let scheme = scheme.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "type"))?;

        Ok(Self { scheme, start_at })
    }
}

/// TextTabStop
pub struct TextTabStop {
    pub position: Option<Coordinate32>,
    pub alignment: Option<TextTabAlignType>,
}

impl TextTabStop {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextTabStop> {
        let mut position = None;
        let mut alignment = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "pos" => position = Some(value.parse::<Coordinate32>()?),
                "algn" => alignment = Some(value.parse::<TextTabAlignType>()?),
                _ => (),
            }
        }

        Ok(Self { position, alignment })
    }
}

pub enum TextUnderlineLine {
    FollowText,
    Line(Option<Box<LineProperties>>),
}

impl TextUnderlineLine {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "uLnTx" | "uLn" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "uLnTx" => Ok(TextUnderlineLine::FollowText),
            "uLn" => Ok(TextUnderlineLine::Line(match xml_node.child_nodes.get(0) {
                Some(node) => Some(Box::new(LineProperties::from_xml_element(node)?)),
                None => None,
            })),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextUnderlineLine").into()),
        }
    }
}

pub enum TextUnderlineFill {
    FollowText,
    Fill(FillProperties),
}

impl TextUnderlineFill {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "uFillTx" | "uFill" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "uFillTx" => Ok(TextUnderlineFill::FollowText),
            "uFill" => {
                let fill_node = xml_node
                    .child_nodes
                    .get(0)
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_FillProperties"))?;
                Ok(TextUnderlineFill::Fill(FillProperties::from_xml_element(fill_node)?))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextUnderlineFill").into()),
        }
    }
}

pub struct Hyperlink {
    pub relationship_id: Option<RelationshipId>,
    pub invalid_url: Option<String>,
    pub action: Option<String>,
    pub target_frame: Option<String>,
    pub tooltip: Option<String>,
    pub history: Option<bool>,         // true
    pub highlight_click: Option<bool>, // false
    pub end_sound: Option<bool>,       // false
    pub sound: Option<EmbeddedWAVAudioFile>,
}

impl Hyperlink {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut relationship_id = None;
        let mut invalid_url = None;
        let mut action = None;
        let mut target_frame = None;
        let mut tooltip = None;
        let mut history = None;
        let mut highlight_click = None;
        let mut end_sound = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r:id" => relationship_id = Some(value.clone()),
                "invalidUrl" => invalid_url = Some(value.clone()),
                "action" => action = Some(value.clone()),
                "tgtFrame" => target_frame = Some(value.clone()),
                "tooltip" => tooltip = Some(value.clone()),
                "history" => history = Some(parse_xml_bool(value)?),
                "highlightClick" => highlight_click = Some(parse_xml_bool(value)?),
                "endSnd" => end_sound = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let sound = None;
        // TODO implement
        // let sound = match xml_node.child_nodes.get(0) {
        //     Some(node) => Some(EmbeddedWAVAudioFile::from_xml_element(node)?),
        //     None => None,
        // };

        Ok(Self {
            relationship_id,
            invalid_url,
            action,
            target_frame,
            tooltip,
            history,
            highlight_click,
            end_sound,
            sound,
        })
    }
}

/// TextCharacterProperties
pub struct TextCharacterProperties {
    pub kumimoji: Option<bool>,
    pub language: Option<TextLanguageID>,
    pub alternative_language: Option<TextLanguageID>,
    pub font_size: Option<TextFontSize>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<TextUnderlineType>,
    pub strikethrough: Option<TextStrikeType>,
    pub kerning: Option<TextNonNegativePoint>,
    pub caps_type: Option<TextCapsType>,
    pub spacing: Option<TextPoint>,
    pub normalize_heights: Option<bool>,
    pub baseline: Option<Percentage>,
    pub no_proofing: Option<bool>,
    pub dirty: Option<bool>,          // true
    pub spelling_error: Option<bool>, // false
    pub smarttag_clean: Option<bool>, // true
    pub smarttag_id: Option<u32>,     // 0
    pub bookmark_link_target: Option<String>,
    pub line_properties: Option<Box<LineProperties>>,
    pub fill_properties: Option<FillProperties>,
    pub effect_properties: Option<EffectProperties>,
    pub highlight_color: Option<Color>,
    pub text_underline_line: Option<TextUnderlineLine>,
    pub text_underline_fill: Option<TextUnderlineFill>,
    pub latin_font: Option<TextFont>,
    pub east_asian_font: Option<TextFont>,
    pub complex_script_font: Option<TextFont>,
    pub symbol_font: Option<TextFont>,
    pub hyperlink_click: Option<Box<Hyperlink>>,
    pub hyperlink_mouse_over: Option<Box<Hyperlink>>,
    pub rtl: Option<bool>,
}

impl TextCharacterProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextCharacterProperties> {
        let mut kumimoji = None;
        let mut language = None;
        let mut alternative_language = None;
        let mut font_size = None;
        let mut bold = None;
        let mut italic = None;
        let mut underline = None;
        let mut strikethrough = None;
        let mut kerning = None;
        let mut caps_type = None;
        let mut spacing = None;
        let mut normalize_heights = None;
        let mut baseline = None;
        let mut no_proofing = None;
        let mut dirty = None;
        let mut spelling_error = None;
        let mut smarttag_clean = None;
        let mut smarttag_id = None;
        let mut bookmark_link_target = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "kumimoji" => kumimoji = Some(parse_xml_bool(value)?),
                "lang" => language = Some(value.clone()),
                "altLang" => alternative_language = Some(value.clone()),
                "sz" => font_size = Some(value.parse::<TextFontSize>()?),
                "b" => bold = Some(parse_xml_bool(value)?),
                "i" => italic = Some(parse_xml_bool(value)?),
                "u" => underline = Some(value.parse::<TextUnderlineType>()?),
                "strike" => strikethrough = Some(value.parse::<TextStrikeType>()?),
                "kern" => kerning = Some(value.parse::<TextNonNegativePoint>()?),
                "cap" => caps_type = Some(value.parse::<TextCapsType>()?),
                "spc" => spacing = Some(value.parse::<TextPoint>()?),
                "normalizeH" => normalize_heights = Some(parse_xml_bool(value)?),
                "baseline" => baseline = Some(value.parse::<Percentage>()?),
                "noProof" => no_proofing = Some(parse_xml_bool(value)?),
                "dirty" => dirty = Some(parse_xml_bool(value)?),
                "err" => spelling_error = Some(parse_xml_bool(value)?),
                "smtClean" => smarttag_clean = Some(parse_xml_bool(value)?),
                "smtId" => smarttag_id = Some(value.parse::<u32>()?),
                "bmk" => bookmark_link_target = Some(value.clone()),
                _ => (),
            }
        }

        let mut line_properties = None;
        let mut fill_properties = None;
        let effect_properties = None;
        let mut highlight_color = None;
        let mut text_underline_line = None;
        let mut text_underline_fill = None;
        let mut latin_font = None;
        let mut east_asian_font = None;
        let mut complex_script_font = None;
        let symbol_font = None;
        let mut hyperlink_click = None;
        let mut hyperlink_mouse_over = None;
        let mut rtl = None;

        for child_node in &xml_node.child_nodes {
            let child_local_name = child_node.local_name();
            if FillProperties::is_choice_member(child_local_name) {
                fill_properties = Some(FillProperties::from_xml_element(child_node)?);
            // TODO implement
            //} else if EffectProperties::is_choice_member(child_local_name) {
            //    effect_properties = Some(EffectProperties::from_xml_element(child_node)?);
            } else if TextUnderlineLine::is_choice_member(child_local_name) {
                text_underline_line = Some(TextUnderlineLine::from_xml_element(child_node)?);
            } else if TextUnderlineFill::is_choice_member(child_local_name) {
                text_underline_fill = Some(TextUnderlineFill::from_xml_element(child_node)?);
            } else {
                match child_local_name {
                    "ln" => line_properties = Some(Box::new(LineProperties::from_xml_element(child_node)?)),
                    "highlight" => {
                        let color_node = child_node
                            .child_nodes
                            .get(0)
                            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "CT_Color"))?;
                        highlight_color = Some(Color::from_xml_element(color_node)?);
                    }
                    "latin" => latin_font = Some(TextFont::from_xml_element(child_node)?),
                    "ea" => east_asian_font = Some(TextFont::from_xml_element(child_node)?),
                    "cs" => complex_script_font = Some(TextFont::from_xml_element(child_node)?),
                    "hlinkClick" => hyperlink_click = Some(Box::new(Hyperlink::from_xml_element(child_node)?)),
                    "hlinkMouseOver" => hyperlink_mouse_over = Some(Box::new(Hyperlink::from_xml_element(child_node)?)),
                    "rtl" => {
                        rtl = match child_node.text {
                            Some(ref s) => Some(parse_xml_bool(s)?),
                            None => None,
                        }
                    }
                    _ => (),
                }
            }
        }

        Ok(Self {
            kumimoji,
            language,
            alternative_language,
            font_size,
            bold,
            italic,
            underline,
            strikethrough,
            kerning,
            caps_type,
            spacing,
            normalize_heights,
            baseline,
            no_proofing,
            dirty,
            spelling_error,
            smarttag_clean,
            smarttag_id,
            bookmark_link_target,
            line_properties,
            fill_properties,
            effect_properties,
            highlight_color,
            text_underline_line,
            text_underline_fill,
            latin_font,
            east_asian_font,
            complex_script_font,
            symbol_font,
            hyperlink_click,
            hyperlink_mouse_over,
            rtl,
        })
    }
}

/// TextParagraphProperties
pub struct TextParagraphProperties {
    pub margin_left: Option<TextMargin>,
    pub margin_right: Option<TextMargin>,
    pub level: Option<TextIndentLevelType>,
    pub indent: Option<TextIndent>,
    pub align: Option<TextAlignType>,
    pub default_tab_size: Option<Coordinate32>,
    pub rtl: Option<bool>,
    pub east_asian_line_break: Option<bool>,
    pub font_align: Option<TextFontAlignType>,
    pub latin_line_break: Option<bool>,
    pub hanging_punctuations: Option<bool>,
    pub line_spacing: Option<TextSpacing>,
    pub space_before: Option<TextSpacing>,
    pub space_after: Option<TextSpacing>,
    pub bullet_color: Option<TextBulletColor>,
    pub bullet_size: Option<TextBulletSize>,
    pub bullet_typeface: Option<TextBulletTypeface>,
    pub bullet: Option<TextBullet>,
    pub tab_stop_list: Vec<TextTabStop>,
    pub default_run_properties: Option<Box<TextCharacterProperties>>,
}

impl TextParagraphProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextParagraphProperties> {
        let mut margin_left = None;
        let mut margin_right = None;
        let mut level = None;
        let mut indent = None;
        let mut align = None;
        let mut default_tab_size = None;
        let mut rtl = None;
        let mut east_asian_line_break = None;
        let mut font_align = None;
        let mut latin_line_break = None;
        let mut hanging_punctuations = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "marL" => margin_left = Some(value.parse::<TextMargin>()?),
                "marR" => margin_right = Some(value.parse::<TextMargin>()?),
                "lvl" => level = Some(value.parse::<TextIndentLevelType>()?),
                "indent" => indent = Some(value.parse::<TextIndent>()?),
                "algn" => align = Some(value.parse::<TextAlignType>()?),
                "defTabSz" => default_tab_size = Some(value.parse::<Coordinate32>()?),
                "rtl" => rtl = Some(parse_xml_bool(value)?),
                "eaLnBrk" => east_asian_line_break = Some(parse_xml_bool(value)?),
                "fontAlgn" => font_align = Some(value.parse::<TextFontAlignType>()?),
                "latinLnBrk" => latin_line_break = Some(parse_xml_bool(value)?),
                "hangingPunct" => hanging_punctuations = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut line_spacing = None;
        let mut space_before = None;
        let mut space_after = None;
        let mut bullet_color = None;
        let mut bullet_size = None;
        let mut bullet_typeface = None;
        let mut bullet = None;
        let mut tab_stop_list = Vec::new();
        let mut default_run_properties = None;

        for child_node in &xml_node.child_nodes {
            if TextBulletColor::is_choice_member(child_node.local_name()) {
                bullet_color = Some(TextBulletColor::from_xml_element(child_node)?);
            } else if TextBulletColor::is_choice_member(child_node.local_name()) {
                bullet_size = Some(TextBulletSize::from_xml_element(child_node)?);
            } else if TextBulletTypeface::is_choice_member(child_node.local_name()) {
                bullet_typeface = Some(TextBulletTypeface::from_xml_element(child_node)?);
            } else if TextBullet::is_choice_member(child_node.local_name()) {
                bullet = Some(TextBullet::from_xml_element(child_node)?);
            } else {
                match child_node.local_name() {
                    "lnSpc" => {
                        let line_spacing_node = child_node
                            .child_nodes
                            .get(0)
                            .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "lnSpc child"))?;
                        line_spacing = Some(TextSpacing::from_xml_element(line_spacing_node)?);
                    }
                    "spcBef" => {
                        let space_before_node = child_node
                            .child_nodes
                            .get(0)
                            .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "spcBef child"))?;
                        space_before = Some(TextSpacing::from_xml_element(space_before_node)?);
                    }
                    "spcAft" => {
                        let space_after_node = child_node
                            .child_nodes
                            .get(0)
                            .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "spcAft child"))?;
                        space_after = Some(TextSpacing::from_xml_element(space_after_node)?);
                    }
                    "tabLst" => tab_stop_list.push(TextTabStop::from_xml_element(child_node)?),
                    "defRPr" => {
                        default_run_properties = Some(Box::new(TextCharacterProperties::from_xml_element(child_node)?))
                    }
                    _ => (),
                }
            }
        }

        Ok(Self {
            margin_left,
            margin_right,
            level,
            indent,
            align,
            default_tab_size,
            rtl,
            east_asian_line_break,
            font_align,
            latin_line_break,
            hanging_punctuations,
            line_spacing,
            space_before,
            space_after,
            bullet_color,
            bullet_size,
            bullet_typeface,
            bullet,
            tab_stop_list,
            default_run_properties,
        })
    }
}

pub struct TextParagraph {
    pub properties: Option<Box<TextParagraphProperties>>,
    pub text_run_list: Vec<TextRun>,
    pub end_paragraph_char_properties: Option<Box<TextCharacterProperties>>,
}

impl TextParagraph {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut properties = None;
        let mut text_run_list = Vec::new();
        let mut end_paragraph_char_properties = None;

        for child_node in &xml_node.child_nodes {
            let local_name = child_node.local_name();
            if TextRun::is_choice_member(local_name) {
                text_run_list.push(TextRun::from_xml_element(child_node)?);
            } else {
                match child_node.local_name() {
                    "pPr" => properties = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?)),
                    "endParaRPr" => {
                        end_paragraph_char_properties =
                            Some(Box::new(TextCharacterProperties::from_xml_element(child_node)?))
                    }
                    _ => (),
                }
            }
        }

        Ok(Self {
            properties,
            text_run_list,
            end_paragraph_char_properties,
        })
    }
}

pub enum TextRun {
    RegularTextRun(Box<RegularTextRun>),
    LineBreak(Box<TextLineBreak>),
    TextField(Box<TextField>),
}

impl TextRun {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "r" | "br" | "fld" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "r" => Ok(TextRun::RegularTextRun(Box::new(RegularTextRun::from_xml_element(
                xml_node,
            )?))),
            "br" => Ok(TextRun::LineBreak(Box::new(TextLineBreak::from_xml_element(xml_node)?))),
            "fld" => Ok(TextRun::TextField(Box::new(TextField::from_xml_element(xml_node)?))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextRun").into()),
        }
    }
}

pub struct RegularTextRun {
    pub char_properties: Option<Box<TextCharacterProperties>>,
    pub text: String,
}

impl RegularTextRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut char_properties = None;
        let mut text = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rPr" => char_properties = Some(Box::new(TextCharacterProperties::from_xml_element(child_node)?)),
                "t" => text = child_node.text.clone(),
                _ => (),
            }
        }

        let text = text.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "t"))?;
        Ok(Self { char_properties, text })
    }
}

pub struct TextLineBreak {
    pub char_properties: Option<Box<TextCharacterProperties>>,
}

impl TextLineBreak {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let char_properties = match xml_node.child_nodes.get(0) {
            Some(node) => Some(Box::new(TextCharacterProperties::from_xml_element(node)?)),
            None => None,
        };

        Ok(Self { char_properties })
    }
}

pub struct TextField {
    pub id: Guid,
    pub field_type: Option<String>,
    pub char_properties: Option<Box<TextCharacterProperties>>,
    pub paragraph_properties: Option<Box<TextParagraph>>,
    pub text: Option<String>,
}

impl TextField {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut field_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => id = Some(value.clone()),
                "type" => field_type = Some(value.clone()),
                _ => (),
            }
        }

        let id = id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;

        let mut char_properties = None;
        let mut paragraph_properties = None;
        let mut text = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rPr" => char_properties = Some(Box::new(TextCharacterProperties::from_xml_element(child_node)?)),
                "pPr" => paragraph_properties = Some(Box::new(TextParagraph::from_xml_element(child_node)?)),
                "t" => text = child_node.text.clone(),
                _ => (),
            }
        }

        Ok(Self {
            id,
            field_type,
            char_properties,
            paragraph_properties,
            text,
        })
    }
}

/// TextListStyle
pub struct TextListStyle {
    pub def_paragraph_props: Option<Box<TextParagraphProperties>>,
    pub lvl1_paragraph_props: Option<Box<TextParagraphProperties>>,
    pub lvl2_paragraph_props: Option<Box<TextParagraphProperties>>,
    pub lvl3_paragraph_props: Option<Box<TextParagraphProperties>>,
    pub lvl4_paragraph_props: Option<Box<TextParagraphProperties>>,
    pub lvl5_paragraph_props: Option<Box<TextParagraphProperties>>,
    pub lvl6_paragraph_props: Option<Box<TextParagraphProperties>>,
    pub lvl7_paragraph_props: Option<Box<TextParagraphProperties>>,
    pub lvl8_paragraph_props: Option<Box<TextParagraphProperties>>,
    pub lvl9_paragraph_props: Option<Box<TextParagraphProperties>>,
}

impl TextListStyle {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut def_paragraph_props = None;
        let mut lvl1_paragraph_props = None;
        let mut lvl2_paragraph_props = None;
        let mut lvl3_paragraph_props = None;
        let mut lvl4_paragraph_props = None;
        let mut lvl5_paragraph_props = None;
        let mut lvl6_paragraph_props = None;
        let mut lvl7_paragraph_props = None;
        let mut lvl8_paragraph_props = None;
        let mut lvl9_paragraph_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "defPPr" => {
                    def_paragraph_props = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                }
                "lvl1pPr" => {
                    lvl1_paragraph_props = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                }
                "lvl2pPr" => {
                    lvl2_paragraph_props = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                }
                "lvl3pPr" => {
                    lvl3_paragraph_props = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                }
                "lvl4pPr" => {
                    lvl4_paragraph_props = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                }
                "lvl5pPr" => {
                    lvl5_paragraph_props = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                }
                "lvl6pPr" => {
                    lvl6_paragraph_props = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                }
                "lvl7pPr" => {
                    lvl7_paragraph_props = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                }
                "lvl8pPr" => {
                    lvl8_paragraph_props = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                }
                "lvl9pPr" => {
                    lvl9_paragraph_props = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                }
                _ => (),
            }
        }

        Ok(Self {
            def_paragraph_props,
            lvl1_paragraph_props,
            lvl2_paragraph_props,
            lvl3_paragraph_props,
            lvl4_paragraph_props,
            lvl5_paragraph_props,
            lvl6_paragraph_props,
            lvl7_paragraph_props,
            lvl8_paragraph_props,
            lvl9_paragraph_props,
        })
    }
}

pub struct TextBody {
    pub body_properties: Box<TextBodyProperties>,
    pub list_style: Option<Box<TextListStyle>>,
    pub paragraph_array: Vec<Box<TextParagraph>>,
}

impl TextBody {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut body_properties = None;
        let mut list_style = None;
        let mut paragraph_array = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "bodyPr" => body_properties = Some(Box::new(TextBodyProperties::from_xml_element(child_node)?)),
                "lstStyle" => list_style = Some(Box::new(TextListStyle::from_xml_element(child_node)?)),
                "p" => paragraph_array.push(Box::new(TextParagraph::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let body_properties =
            body_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "bodyPr"))?;

        Ok(Self {
            body_properties,
            list_style,
            paragraph_array,
        })
    }
}

pub struct TextBodyProperties {
    pub rotate_angle: Option<Angle>,
    pub paragraph_spacing: Option<bool>,
    pub vertical_overflow: Option<TextVertOverflowType>,
    pub horizontal_overflow: Option<TextHorzOverflowType>,
    pub vertical_type: Option<TextVerticalType>,
    pub wrap_type: Option<TextWrappingType>,
    pub left_inset: Option<Coordinate32>,
    pub top_inset: Option<Coordinate32>,
    pub right_inset: Option<Coordinate32>,
    pub bottom_inset: Option<Coordinate32>,
    pub column_count: Option<TextColumnCount>,
    pub space_between_columns: Option<PositiveCoordinate32>,
    pub rtl_columns: Option<bool>,
    pub is_from_word_art: Option<bool>,
    pub anchor: Option<TextAnchoringType>,
    pub anchor_center: Option<bool>,
    pub force_antialias: Option<bool>,
    pub upright: Option<bool>,
    pub compatible_line_spacing: Option<bool>,
    pub preset_text_warp: Option<Box<PresetTextShape>>,
    pub auto_fit_type: Option<TextAutoFit>,
    //pub scene_3d: Option<Scene3D>,
    //pub text_3d: Option<Text3D>,
}

impl TextBodyProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut rotate_angle = None;
        let mut paragraph_spacing = None;
        let mut vertical_overflow = None;
        let mut horizontal_overflow = None;
        let mut vertical_type = None;
        let mut wrap_type = None;
        let mut left_inset = None;
        let mut top_inset = None;
        let mut right_inset = None;
        let mut bottom_inset = None;
        let mut column_count = None;
        let mut space_between_columns = None;
        let mut rtl_columns = None;
        let mut is_from_word_art = None;
        let mut anchor = None;
        let mut anchor_center = None;
        let mut force_antialias = None;
        let mut upright = None;
        let mut compatible_line_spacing = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "rot" => rotate_angle = Some(value.parse::<Angle>()?),
                "spcFirstLastPara" => paragraph_spacing = Some(parse_xml_bool(value)?),
                "vertOverflow" => vertical_overflow = Some(value.parse::<TextVertOverflowType>()?),
                "horzOverflow" => horizontal_overflow = Some(value.parse::<TextHorzOverflowType>()?),
                "vert" => vertical_type = Some(value.parse::<TextVerticalType>()?),
                "wrap" => wrap_type = Some(value.parse::<TextWrappingType>()?),
                "lIns" => left_inset = Some(value.parse::<Coordinate32>()?),
                "tIns" => top_inset = Some(value.parse::<Coordinate32>()?),
                "rIns" => right_inset = Some(value.parse::<Coordinate32>()?),
                "bIns" => bottom_inset = Some(value.parse::<Coordinate32>()?),
                "numCol" => column_count = Some(value.parse::<TextColumnCount>()?),
                "spcCol" => space_between_columns = Some(value.parse::<PositiveCoordinate32>()?),
                "rtlCol" => rtl_columns = Some(parse_xml_bool(value)?),
                "fromWordArt" => is_from_word_art = Some(parse_xml_bool(value)?),
                "anchor" => anchor = Some(value.parse::<TextAnchoringType>()?),
                "anchorCtr" => anchor_center = Some(parse_xml_bool(value)?),
                "forceAA" => force_antialias = Some(parse_xml_bool(value)?),
                "upright" => upright = Some(parse_xml_bool(value)?),
                "compatLnSpc" => compatible_line_spacing = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut preset_text_warp = None;
        let mut auto_fit_type = None;

        for child_node in &xml_node.child_nodes {
            let child_local_name = child_node.local_name();
            if TextAutoFit::is_choice_member(child_local_name) {
                auto_fit_type = Some(TextAutoFit::from_xml_element(child_node)?);
            } else if child_local_name == "prstTxWarp" {
                preset_text_warp = Some(Box::new(PresetTextShape::from_xml_element(child_node)?));
            }
        }

        Ok(Self {
            rotate_angle,
            paragraph_spacing,
            vertical_overflow,
            horizontal_overflow,
            vertical_type,
            wrap_type,
            left_inset,
            top_inset,
            right_inset,
            bottom_inset,
            column_count,
            space_between_columns,
            rtl_columns,
            is_from_word_art,
            anchor,
            anchor_center,
            force_antialias,
            upright,
            compatible_line_spacing,
            auto_fit_type,
            preset_text_warp,
        })
    }
}

pub enum TextAutoFit {
    NoAutoFit,
    NormalAutoFit(TextNormalAutoFit),
    ShapeAutoFit,
}

impl TextAutoFit {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "noAutofit" | "normAutofit" | "spAutoFit" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "noAutofit" => Ok(TextAutoFit::NoAutoFit),
            "normAutofit" => Ok(TextAutoFit::NormalAutoFit(TextNormalAutoFit::from_xml_element(
                xml_node,
            )?)),
            "spAutoFit" => Ok(TextAutoFit::ShapeAutoFit),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextAutofit").into()),
        }
    }
}

pub struct TextNormalAutoFit {
    pub font_scale: Option<TextFontScalePercent>,           // 100000
    pub line_spacing_reduction: Option<TextSpacingPercent>, // 0
}

impl TextNormalAutoFit {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut font_scale = None;
        let mut line_spacing_reduction = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "fontScale" => font_scale = Some(value.parse::<TextFontScalePercent>()?),
                "lnSpcReduction" => line_spacing_reduction = Some(value.parse::<TextSpacingPercent>()?),
                _ => (),
            }
        }

        Ok(Self {
            font_scale,
            line_spacing_reduction,
        })
    }
}

pub struct PresetTextShape {
    pub preset: TextShapeType,
    pub adjust_value_list: Vec<GeomGuide>,
}

impl PresetTextShape {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let preset_attr = xml_node
            .attribute("prst")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "prst"))?;
        let preset = preset_attr.parse()?;

        let mut adjust_value_list = Vec::new();
        if let Some(node) = xml_node.child_nodes.get(0) {
            for av_node in &node.child_nodes {
                adjust_value_list.push(GeomGuide::from_xml_element(av_node)?);
            }
        }

        Ok(Self {
            preset,
            adjust_value_list,
        })
    }
}

pub struct FontScheme {
    pub name: String,
    pub major_font: Box<FontCollection>,
    pub minor_font: Box<FontCollection>,
}

impl FontScheme {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name_attr = xml_node
            .attribute("name")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?;
        let name = name_attr.clone();
        let mut major_font = None;
        let mut minor_font = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "majorFont" => major_font = Some(Box::new(FontCollection::from_xml_element(child_node)?)),
                "minorFont" => minor_font = Some(Box::new(FontCollection::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let major_font = major_font.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "majorFont"))?;
        let minor_font = minor_font.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "minorFont"))?;

        Ok(Self {
            name,
            major_font,
            minor_font,
        })
    }
}

pub struct FontCollection {
    pub latin: TextFont,
    pub east_asian: TextFont,
    pub complex_script: TextFont,
    pub supplemental_font_list: Vec<SupplementalFont>,
}

impl FontCollection {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut opt_latin = None;
        let mut opt_ea = None;
        let mut opt_cs = None;
        let mut supplemental_font_list = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "latin" => opt_latin = Some(TextFont::from_xml_element(child_node)?),
                "ea" => opt_ea = Some(TextFont::from_xml_element(child_node)?),
                "cs" => opt_cs = Some(TextFont::from_xml_element(child_node)?),
                "font" => supplemental_font_list.push(SupplementalFont::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let latin = opt_latin.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "latin"))?;
        let east_asian = opt_ea.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "ea"))?;
        let complex_script = opt_cs.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cs"))?;

        Ok(Self {
            latin,
            east_asian,
            complex_script,
            supplemental_font_list,
        })
    }
}

pub struct NonVisualDrawingProps {
    pub id: DrawingElementId,
    pub name: String,
    pub description: Option<String>,
    pub hidden: Option<bool>, // false
    pub title: Option<String>,
    pub hyperlink_click: Option<Box<Hyperlink>>,
    pub hyperlink_hover: Option<Box<Hyperlink>>,
}

impl NonVisualDrawingProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut opt_id = None;
        let mut opt_name = None;
        let mut description = None;
        let mut hidden = None;
        let mut title = None;
        let mut hyperlink_click = None;
        let mut hyperlink_hover = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => opt_id = Some(value.parse::<DrawingElementId>()?),
                "name" => opt_name = Some(value.clone()),
                "descr" => description = Some(value.clone()),
                "hidden" => hidden = Some(parse_xml_bool(value)?),
                "title" => title = Some(value.clone()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "hlinkClick" => hyperlink_click = Some(Box::new(Hyperlink::from_xml_element(child_node)?)),
                "hlinkHover" => hyperlink_hover = Some(Box::new(Hyperlink::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let id = opt_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;
        let name = opt_name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?;

        Ok(Self {
            id,
            name,
            description,
            hidden,
            title,
            hyperlink_click,
            hyperlink_hover,
        })
    }
}

#[derive(Default, Debug, Copy, Clone)]
pub struct Locking {
    pub no_grouping: Option<bool>,            // false
    pub no_select: Option<bool>,              // false
    pub no_rotate: Option<bool>,              // false
    pub no_change_aspect_ratio: Option<bool>, // false
    pub no_move: Option<bool>,                // false
    pub no_resize: Option<bool>,              // false
    pub no_edit_points: Option<bool>,         // false
    pub no_adjust_handles: Option<bool>,      // false
    pub no_change_arrowheads: Option<bool>,   // false
    pub no_change_shape_type: Option<bool>,   // false
}

impl Locking {
    pub fn try_attribute_parse(&mut self, attr: &str, value: &str) -> Result<()> {
        match attr {
            "noGrp" => self.no_grouping = Some(parse_xml_bool(value)?),
            "noSelect" => self.no_select = Some(parse_xml_bool(value)?),
            "noRot" => self.no_rotate = Some(parse_xml_bool(value)?),
            "noChangeAspect" => self.no_change_aspect_ratio = Some(parse_xml_bool(value)?),
            "noMove" => self.no_move = Some(parse_xml_bool(value)?),
            "noResize" => self.no_resize = Some(parse_xml_bool(value)?),
            "noEditPoints" => self.no_edit_points = Some(parse_xml_bool(value)?),
            "noAdjustHandles" => self.no_adjust_handles = Some(parse_xml_bool(value)?),
            "noChangeArrowheads" => self.no_change_arrowheads = Some(parse_xml_bool(value)?),
            "noChangeShapeType" => self.no_change_shape_type = Some(parse_xml_bool(value)?),
            _ => (),
        }

        Ok(())
    }
}

pub struct ShapeLocking {
    pub locking: Locking,
    pub no_text_edit: Option<bool>, // false
}

impl ShapeLocking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut locking: Locking = Default::default();
        let mut no_text_edit = None;

        for (attr, value) in &xml_node.attributes {
            if attr.as_str() == "noTextEdit" {
                no_text_edit = Some(parse_xml_bool(value)?);
            } else {
                locking.try_attribute_parse(attr, value)?;
            }
        }

        Ok(Self { locking, no_text_edit })
    }
}

pub struct GroupLocking {
    pub no_grouping: Option<bool>,            // false
    pub no_ungrouping: Option<bool>,          // false
    pub no_select: Option<bool>,              // false
    pub no_rotate: Option<bool>,              // false
    pub no_change_aspect_ratio: Option<bool>, // false
    pub no_move: Option<bool>,                // false
    pub no_resize: Option<bool>,              // false
}

impl GroupLocking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut no_grouping = None;
        let mut no_ungrouping = None;
        let mut no_select = None;
        let mut no_rotate = None;
        let mut no_change_aspect_ratio = None;
        let mut no_move = None;
        let mut no_resize = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "noGrp" => no_grouping = Some(parse_xml_bool(value)?),
                "noUngrp" => no_ungrouping = Some(parse_xml_bool(value)?),
                "noSelect" => no_select = Some(parse_xml_bool(value)?),
                "noRot" => no_rotate = Some(parse_xml_bool(value)?),
                "noChangeAspect" => no_change_aspect_ratio = Some(parse_xml_bool(value)?),
                "noMove" => no_move = Some(parse_xml_bool(value)?),
                "noResize" => no_resize = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(Self {
            no_grouping,
            no_ungrouping,
            no_select,
            no_rotate,
            no_change_aspect_ratio,
            no_move,
            no_resize,
        })
    }
}

pub struct GraphicalObjectFrameLocking {
    pub no_grouping: Option<bool>,      // false
    pub no_drilldown: Option<bool>,     // false
    pub no_select: Option<bool>,        // false
    pub no_change_aspect: Option<bool>, // false
    pub no_move: Option<bool>,          // false
    pub no_resize: Option<bool>,        // false
}

pub struct ConnectorLocking {
    pub locking: Locking,
}

impl ConnectorLocking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut locking: Locking = Default::default();

        for (attr, value) in &xml_node.attributes {
            locking.try_attribute_parse(attr, value)?;
        }

        Ok(Self { locking })
    }
}

pub struct PictureLocking {
    pub locking: Locking,
    pub no_crop: Option<bool>, // false
}

impl PictureLocking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut locking: Locking = Default::default();
        let mut no_crop = None;
        for (attr, value) in &xml_node.attributes {
            if attr.as_str() == "noCrop" {
                no_crop = Some(parse_xml_bool(value)?);
            } else {
                locking.try_attribute_parse(attr, value)?;
            }
        }

        Ok(Self { locking, no_crop })
    }
}

pub struct NonVisualDrawingShapeProps {
    pub is_text_box: Option<bool>, // false
    pub shape_locks: Option<ShapeLocking>,
}

impl NonVisualDrawingShapeProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let is_text_box = match xml_node.attribute("txBox") {
            Some(attr) => Some(parse_xml_bool(attr)?),
            None => None,
        };

        let shape_locks = match xml_node.child_nodes.get(0) {
            Some(sp_lock_node) => Some(ShapeLocking::from_xml_element(sp_lock_node)?),
            None => None,
        };

        Ok(Self {
            is_text_box,
            shape_locks,
        })
    }
}

pub struct NonVisualGroupDrawingShapeProps {
    pub locks: Option<GroupLocking>,
}

impl NonVisualGroupDrawingShapeProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let locks = match xml_node.child_nodes.get(0) {
            Some(node) => Some(GroupLocking::from_xml_element(node)?),
            None => None,
        };

        Ok(Self { locks })
    }
}

pub struct NonVisualGraphicFrameProperties {
    pub graphic_frame_locks: Option<GraphicalObjectFrameLocking>,
}

pub struct NonVisualConnectorProperties {
    pub connector_locks: Option<ConnectorLocking>,
    pub start_connection: Option<Connection>,
    pub end_connection: Option<Connection>,
}

impl NonVisualConnectorProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut connector_locks = None;
        let mut start_connection = None;
        let mut end_connection = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cxnSpLocks" => connector_locks = Some(ConnectorLocking::from_xml_element(child_node)?),
                "stCxn" => start_connection = Some(Connection::from_xml_element(child_node)?),
                "endCxn" => end_connection = Some(Connection::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(Self {
            connector_locks,
            start_connection,
            end_connection,
        })
    }
}

pub struct NonVisualPictureProperties {
    pub prefer_relative_resize: Option<bool>, // true
    pub picture_locks: Option<PictureLocking>,
}

impl NonVisualPictureProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let prefer_relative_resize = match xml_node.attribute("preferRelativeResize") {
            Some(attr) => Some(parse_xml_bool(attr)?),
            None => None,
        };

        let picture_locks = match xml_node.child_nodes.get(0) {
            Some(node) => Some(PictureLocking::from_xml_element(node)?),
            None => None,
        };

        Ok(Self {
            prefer_relative_resize,
            picture_locks,
        })
    }
}

pub struct Connection {
    pub id: DrawingElementId,
    pub shape_index: u32,
}

impl Connection {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut shape_index = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => id = Some(value.parse()?),
                "idx" => shape_index = Some(value.parse()?),
                _ => (),
            }
        }

        let id = id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;
        let shape_index = shape_index.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "idx"))?;

        Ok(Self { id, shape_index })
    }
}

pub struct EmbeddedWAVAudioFile {
    pub embed_rel_id: RelationshipId,
    pub name: Option<String>,
    pub built_in: Option<bool>, // false
}

pub struct AudioCDTime {
    pub track: u8,
    pub time: Option<u32>, // 0
}

pub struct AudioCD {
    pub start_time: AudioCDTime,
    pub end_time: AudioCDTime,
}

pub struct AudioFile {
    pub link: RelationshipId,
    pub content_type: Option<String>,
}

pub struct VideoFile {
    pub link: RelationshipId,
    pub content_type: Option<String>,
}

pub struct QuickTimeFile {
    pub link: RelationshipId,
}

pub enum Media {
    AudioCd(AudioCD),
    WavAudioFile(EmbeddedWAVAudioFile),
    AudioFile(AudioFile),
    VideoFile(VideoFile),
    QuickTimeFile(QuickTimeFile),
}

pub struct Transform2D {
    pub rotate_angle: Option<Angle>,   // 0
    pub flip_horizontal: Option<bool>, // false
    pub flip_vertical: Option<bool>,   // false
    pub offset: Option<Point2D>,
    pub extents: Option<PositiveSize2D>,
}

impl Transform2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut rotate_angle = None;
        let mut flip_horizontal = None;
        let mut flip_vertical = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "rot" => rotate_angle = Some(value.parse()?),
                "flipH" => flip_horizontal = Some(parse_xml_bool(value)?),
                "flipV" => flip_vertical = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut offset = None;
        let mut extents = None;
        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "off" => offset = Some(Point2D::from_xml_element(child_node)?),
                "ext" => extents = Some(PositiveSize2D::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(Self {
            rotate_angle,
            flip_horizontal,
            flip_vertical,
            offset,
            extents,
        })
    }
}

pub struct GroupTransform2D {
    pub rotate_angle: Option<Angle>,   // 0
    pub flip_horizontal: Option<bool>, // false
    pub flip_vertical: Option<bool>,   // false
    pub offset: Option<Point2D>,
    pub extents: Option<PositiveSize2D>,
    pub child_offset: Option<Point2D>,
    pub child_extents: Option<PositiveSize2D>,
}

impl GroupTransform2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut rotate_angle = None;
        let mut flip_horizontal = None;
        let mut flip_vertical = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "rot" => rotate_angle = Some(value.parse()?),
                "flipH" => flip_horizontal = Some(parse_xml_bool(value)?),
                "flipV" => flip_vertical = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut offset = None;
        let mut extents = None;
        let mut child_offset = None;
        let mut child_extents = None;
        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "off" => offset = Some(Point2D::from_xml_element(child_node)?),
                "ext" => extents = Some(PositiveSize2D::from_xml_element(child_node)?),
                "chOff" => child_offset = Some(Point2D::from_xml_element(child_node)?),
                "chExt" => child_extents = Some(PositiveSize2D::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(Self {
            rotate_angle,
            flip_horizontal,
            flip_vertical,
            offset,
            extents,
            child_offset,
            child_extents,
        })
    }
}

pub struct GroupShapeProperties {
    pub black_and_white_mode: Option<BlackWhiteMode>,
    pub transform: Option<Box<GroupTransform2D>>,
    pub fill_properties: Option<FillProperties>,
    pub effect_properties: Option<EffectProperties>,
    //pub scene_3d: Option<Scene3D>,
}

impl GroupShapeProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let black_and_white_mode = match xml_node.attribute("bwMode") {
            Some(attr) => Some(attr.parse()?),
            None => None,
        };

        let mut transform = None;
        let mut fill_properties = None;
        let effect_properties = None;

        for child_node in &xml_node.child_nodes {
            let child_local_name = child_node.local_name();
            if child_local_name == "xfrm" {
                transform = Some(Box::new(GroupTransform2D::from_xml_element(child_node)?));
            } else if FillProperties::is_choice_member(child_local_name) {
                fill_properties = Some(FillProperties::from_xml_element(child_node)?);
                // } else if EffectProperties::is_choice_member(child_local_name) {
                //     effect_properties = Some(EffectProperties::from_xml_element(child_node)?);
            }
        }

        Ok(Self {
            black_and_white_mode,
            transform,
            fill_properties,
            effect_properties,
        })
    }
}

pub enum Geometry {
    Custom(Box<CustomGeometry2D>),
    Preset(Box<PresetGeometry2D>),
}

impl Geometry {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "custGeom" | "prstGeom" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "custGeom" => Ok(Geometry::Custom(Box::new(CustomGeometry2D::from_xml_element(
                xml_node,
            )?))),
            "prstGeom" => Ok(Geometry::Preset(Box::new(PresetGeometry2D::from_xml_element(
                xml_node,
            )?))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_Geometry").into()),
        }
    }
}

pub struct GeomGuide {
    pub name: GeomGuideName,
    pub formula: GeomGuideFormula,
}

impl GeomGuide {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut name = None;
        let mut formula = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "name" => name = Some(value.clone()),
                "fmla" => formula = Some(value.clone()),
                _ => (),
            }
        }

        let name = name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?;
        let formula = formula.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "fmla"))?;
        Ok(Self { name, formula })
    }
}

pub enum AdjustHandle {
    XY(Box<XYAdjustHandle>),
    Polar(Box<PolarAdjustHandle>),
}

impl AdjustHandle {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "ahXY" | "ahPolar" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "ahXY" => Ok(AdjustHandle::XY(Box::new(XYAdjustHandle::from_xml_element(xml_node)?))),
            "ahPolar" => Ok(AdjustHandle::Polar(Box::new(PolarAdjustHandle::from_xml_element(
                xml_node,
            )?))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "AdjustHandle").into()),
        }
    }
}

pub enum AdjCoordinate {
    Coordinate(Coordinate),
    GeomGuideName(GeomGuideName),
}

impl FromStr for AdjCoordinate {
    type Err = AdjustParseError;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s.parse::<Coordinate>() {
            Ok(coord) => Ok(AdjCoordinate::Coordinate(coord)),
            Err(_) => Ok(AdjCoordinate::GeomGuideName(GeomGuideName::from(s))),
        }
    }
}

pub enum AdjAngle {
    Angle(Angle),
    GeomGuideName(GeomGuideName),
}

impl FromStr for AdjAngle {
    type Err = AdjustParseError;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s.parse::<Angle>() {
            Ok(angle) => Ok(AdjAngle::Angle(angle)),
            Err(_) => Ok(AdjAngle::GeomGuideName(GeomGuideName::from(s))),
        }
    }
}

pub struct AdjPoint2D {
    pub x: AdjCoordinate,
    pub y: AdjCoordinate,
}

impl AdjPoint2D {
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

pub struct GeomRect {
    pub left: AdjCoordinate,
    pub top: AdjCoordinate,
    pub right: AdjCoordinate,
    pub bottom: AdjCoordinate,
}

impl GeomRect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut left = None;
        let mut top = None;
        let mut right = None;
        let mut bottom = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "l" => left = Some(value.parse()?),
                "t" => top = Some(value.parse()?),
                "r" => right = Some(value.parse()?),
                "b" => bottom = Some(value.parse()?),
                _ => (),
            }
        }

        let left = left.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "l"))?;
        let top = top.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "t"))?;
        let right = right.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r"))?;
        let bottom = bottom.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "b"))?;

        Ok(Self {
            left,
            top,
            right,
            bottom,
        })
    }
}

pub struct XYAdjustHandle {
    pub guide_reference_x: Option<GeomGuideName>,
    pub guide_reference_y: Option<GeomGuideName>,
    pub min_x: Option<AdjCoordinate>,
    pub max_x: Option<AdjCoordinate>,
    pub min_y: Option<AdjCoordinate>,
    pub max_y: Option<AdjCoordinate>,
    pub position: AdjPoint2D,
}

impl XYAdjustHandle {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut guide_reference_x = None;
        let mut guide_reference_y = None;
        let mut min_x = None;
        let mut max_x = None;
        let mut min_y = None;
        let mut max_y = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "gdRefX" => guide_reference_x = Some(value.clone()),
                "gdRefY" => guide_reference_y = Some(value.clone()),
                "minX" => min_x = Some(value.parse()?),
                "maxX" => max_x = Some(value.parse()?),
                "minY" => min_y = Some(value.parse()?),
                "maxY" => max_y = Some(value.parse()?),
                _ => (),
            }
        }

        let pos_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "pos"))?;
        let position = AdjPoint2D::from_xml_element(pos_node)?;

        Ok(Self {
            guide_reference_x,
            guide_reference_y,
            min_x,
            max_x,
            min_y,
            max_y,
            position,
        })
    }
}

pub struct PolarAdjustHandle {
    pub guide_reference_radial: Option<GeomGuideName>,
    pub guide_reference_angle: Option<GeomGuideName>,
    pub min_radial: Option<AdjCoordinate>,
    pub max_radial: Option<AdjCoordinate>,
    pub min_angle: Option<AdjAngle>,
    pub max_angle: Option<AdjAngle>,
    pub position: AdjPoint2D,
}

impl PolarAdjustHandle {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut guide_reference_radial = None;
        let mut guide_reference_angle = None;
        let mut min_radial = None;
        let mut max_radial = None;
        let mut min_angle = None;
        let mut max_angle = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "gdRefR" => guide_reference_radial = Some(value.clone()),
                "gdRefAng" => guide_reference_angle = Some(value.clone()),
                "minR" => min_radial = Some(value.parse()?),
                "maxR" => max_radial = Some(value.parse()?),
                "minAng" => min_angle = Some(value.parse()?),
                "maxAng" => max_angle = Some(value.parse()?),
                _ => (),
            }
        }

        let pos_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "pos"))?;
        let position = AdjPoint2D::from_xml_element(pos_node)?;

        Ok(Self {
            guide_reference_radial,
            guide_reference_angle,
            min_radial,
            max_radial,
            min_angle,
            max_angle,
            position,
        })
    }
}

pub struct ConnectionSite {
    pub angle: AdjAngle,
    pub position: AdjPoint2D,
}

impl ConnectionSite {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let angle_attr = xml_node
            .attribute("ang")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "ang"))?;
        let angle = angle_attr.parse()?;

        let pos_node = xml_node
            .child_nodes
            .get(0)
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "pos"))?;
        let position = AdjPoint2D::from_xml_element(pos_node)?;

        Ok(Self { angle, position })
    }
}

pub enum Path2DCommand {
    Close,
    MoveTo(AdjPoint2D),
    LineTo(AdjPoint2D),
    ArcTo(Path2DArcTo),
    QuadBezierTo(AdjPoint2D, AdjPoint2D),
    CubicBezTo(AdjPoint2D, AdjPoint2D, AdjPoint2D),
}

pub struct Path2DArcTo {
    pub width_radius: AdjCoordinate,
    pub height_radius: AdjCoordinate,
    pub start_angle: AdjAngle,
    pub swing_angle: AdjAngle,
}

impl Path2DArcTo {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut width_radius = None;
        let mut height_radius = None;
        let mut start_angle = None;
        let mut swing_angle = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "wR" => width_radius = Some(value.parse()?),
                "hR" => height_radius = Some(value.parse()?),
                "stAng" => start_angle = Some(value.parse()?),
                "swAng" => swing_angle = Some(value.parse()?),
                _ => (),
            }
        }

        let width_radius = width_radius.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "wR"))?;
        let height_radius = height_radius.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "hR"))?;
        let start_angle = start_angle.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "stAng"))?;
        let swing_angle = swing_angle.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "swAng"))?;

        Ok(Self {
            width_radius,
            height_radius,
            start_angle,
            swing_angle,
        })
    }
}

pub struct Path2D {
    pub width: Option<PositiveCoordinate>,  // 0
    pub height: Option<PositiveCoordinate>, // 0
    pub fill_mode: Option<PathFillMode>,    // norm
    pub stroke: Option<bool>,               // true
    pub extrusion_ok: Option<bool>,         // true
    pub commands: Vec<Path2DCommand>,
}

impl Path2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut width = None;
        let mut height = None;
        let mut fill_mode = None;
        let mut stroke = None;
        let mut extrusion_ok = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "w" => width = Some(value.parse()?),
                "h" => height = Some(value.parse()?),
                "fill" => fill_mode = Some(value.parse()?),
                "stroke" => stroke = Some(parse_xml_bool(value)?),
                "extrusionOk" => extrusion_ok = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut commands = Vec::new();
        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "close" => commands.push(Path2DCommand::Close),
                "moveTo" => {
                    let pt_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "pt"))?;
                    commands.push(Path2DCommand::MoveTo(AdjPoint2D::from_xml_element(pt_node)?));
                }
                "lnTo" => {
                    let pt_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "pt"))?;
                    commands.push(Path2DCommand::LineTo(AdjPoint2D::from_xml_element(pt_node)?));
                }
                "arcTo" => commands.push(Path2DCommand::ArcTo(Path2DArcTo::from_xml_element(child_node)?)),
                "quadBezTo" => {
                    let pt1_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "pt"))?;
                    let pt2_node = child_node
                        .child_nodes
                        .get(1)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "pt"))?;
                    commands.push(Path2DCommand::QuadBezierTo(
                        AdjPoint2D::from_xml_element(pt1_node)?,
                        AdjPoint2D::from_xml_element(pt2_node)?,
                    ));
                }
                "cubicBezTo" => {
                    let pt1_node = child_node
                        .child_nodes
                        .get(0)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "pt"))?;
                    let pt2_node = child_node
                        .child_nodes
                        .get(1)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "pt"))?;
                    let pt3_node = child_node
                        .child_nodes
                        .get(2)
                        .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "pt"))?;
                    commands.push(Path2DCommand::CubicBezTo(
                        AdjPoint2D::from_xml_element(pt1_node)?,
                        AdjPoint2D::from_xml_element(pt2_node)?,
                        AdjPoint2D::from_xml_element(pt3_node)?,
                    ));
                }
                _ => (),
            }
        }

        Ok(Self {
            width,
            height,
            fill_mode,
            stroke,
            extrusion_ok,
            commands,
        })
    }
}

pub struct CustomGeometry2D {
    pub adjust_value_list: Vec<GeomGuide>,
    pub guide_list: Vec<GeomGuide>,
    pub adjust_handle_list: Vec<AdjustHandle>,
    pub connection_site_list: Vec<ConnectionSite>,
    pub rect: Option<Box<GeomRect>>,
    pub path_list: Vec<Box<Path2D>>,
}

impl CustomGeometry2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut adjust_value_list = Vec::new();
        let mut guide_list = Vec::new();
        let mut adjust_handle_list = Vec::new();
        let mut connection_site_list = Vec::new();
        let mut rect = None;
        let mut path_list = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "avLst" => {
                    for av_node in &child_node.child_nodes {
                        adjust_value_list.push(GeomGuide::from_xml_element(av_node)?);
                    }
                }
                "gdLst" => {
                    for gd_node in &child_node.child_nodes {
                        guide_list.push(GeomGuide::from_xml_element(gd_node)?);
                    }
                }
                "ahLst" => {
                    for ah_node in &child_node.child_nodes {
                        adjust_handle_list.push(AdjustHandle::from_xml_element(ah_node)?);
                    }
                }
                "cxnLst" => {
                    for cxn_node in &child_node.child_nodes {
                        connection_site_list.push(ConnectionSite::from_xml_element(cxn_node)?);
                    }
                }
                "rect" => rect = Some(Box::new(GeomRect::from_xml_element(child_node)?)),
                "pathLst" => {
                    for path_node in &child_node.child_nodes {
                        path_list.push(Box::new(Path2D::from_xml_element(path_node)?));
                    }
                }
                _ => (),
            }
        }

        Ok(Self {
            adjust_value_list,
            guide_list,
            adjust_handle_list,
            connection_site_list,
            rect,
            path_list,
        })
    }
}

pub struct PresetGeometry2D {
    pub adjust_value_list: Vec<GeomGuide>,
    pub preset: ShapeType,
}

impl PresetGeometry2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let preset_attr = xml_node
            .attribute("prst")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "prst"))?;
        let preset = preset_attr.parse()?;
        let mut adjust_value_list = Vec::new();

        for child_node in &xml_node.child_nodes {
            if child_node.local_name() == "avLst" {
                for av_node in &child_node.child_nodes {
                    adjust_value_list.push(GeomGuide::from_xml_element(av_node)?);
                }
            }
        }

        Ok(Self {
            preset,
            adjust_value_list,
        })
    }
}

pub struct ShapeProperties {
    pub black_and_white_mode: Option<BlackWhiteMode>,
    pub transform: Option<Box<Transform2D>>,
    pub geometry: Option<Geometry>,
    pub fill_properties: Option<FillProperties>,
    pub line_properties: Option<Box<LineProperties>>,
    pub effect_properties: Option<Box<EffectProperties>>,
    //pub scene_3d: Option<Scene3D>,
    //pub shape_3d: Option<Shape3D>,
}

impl ShapeProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let black_and_white_mode = match xml_node.attribute("bwMode") {
            Some(value) => Some(value.parse()?),
            None => None,
        };

        let mut transform = None;
        let mut geometry = None;
        let mut fill_properties = None;
        let mut line_properties = None;
        let effect_properties = None;

        for child_node in &xml_node.child_nodes {
            let child_local_name = child_node.local_name();
            if Geometry::is_choice_member(child_local_name) {
                geometry = Some(Geometry::from_xml_element(child_node)?);
            } else if FillProperties::is_choice_member(child_local_name) {
                fill_properties = Some(FillProperties::from_xml_element(child_node)?);
            //} else if EffectProperties::is_choice_member(child_local_name) {
            //    effect_properties = Some(EffectProperties::from_xml_element(child_node))?;
            } else {
                match child_local_name {
                    "xfrm" => transform = Some(Box::new(Transform2D::from_xml_element(child_node)?)),
                    "ln" => line_properties = Some(Box::new(LineProperties::from_xml_element(child_node)?)),
                    _ => (),
                }
            }
        }

        Ok(Self {
            black_and_white_mode,
            transform,
            geometry,
            fill_properties,
            line_properties,
            effect_properties,
        })
    }
}

pub struct ShapeStyle {
    pub line_reference: StyleMatrixReference,
    pub fill_reference: StyleMatrixReference,
    pub effect_reference: StyleMatrixReference,
    pub font_reference: FontReference,
}

impl ShapeStyle {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut line_reference = None;
        let mut fill_reference = None;
        let mut effect_reference = None;
        let mut font_reference = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "lnRef" => line_reference = Some(StyleMatrixReference::from_xml_element(child_node)?),
                "fillRef" => fill_reference = Some(StyleMatrixReference::from_xml_element(child_node)?),
                "effectRef" => effect_reference = Some(StyleMatrixReference::from_xml_element(child_node)?),
                "fontRef" => font_reference = Some(FontReference::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let line_reference =
            line_reference.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lnRef"))?;
        let fill_reference =
            fill_reference.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "fillRef"))?;
        let effect_reference =
            effect_reference.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "effectRef"))?;
        let font_reference =
            font_reference.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "fontRef"))?;

        Ok(Self {
            line_reference,
            fill_reference,
            effect_reference,
            font_reference,
        })
    }
}

pub struct FontReference {
    pub index: FontCollectionIndex,
    pub color: Option<Color>,
}

impl FontReference {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let index_attr = xml_node
            .attribute("idx")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "idx"))?;
        let index = index_attr.parse()?;

        let color = match xml_node.child_nodes.get(0) {
            Some(clr_node) => Some(Color::from_xml_element(clr_node)?),
            None => None,
        };

        Ok(Self { index, color })
    }
}

pub struct GraphicalObject {
    pub graphic_data: GraphicalObjectData,
}

pub struct GraphicalObjectData {
    //pub graphic_object: Vec<Any>,
    pub uri: String,
}

pub enum AnimationElementChoice {
    Diagram(AnimationDgmElement),
    Chart(AnimationChartElement),
}

pub struct AnimationDgmElement {
    pub id: Option<Guid>,                 // {00000000-0000-0000-0000-000000000000}
    pub build_step: Option<DgmBuildStep>, // sp
}

pub struct AnimationChartElement {
    pub series_index: Option<i32>,   // -1
    pub category_index: Option<i32>, // -1
    pub build_step: ChartBuildStep,
}

pub enum AnimationGraphicalObjectBuildProperties {
    BuildDiagram(AnimationDgmBuildProperties),
    BuildChart(AnimationChartBuildProperties),
}

pub struct AnimationDgmBuildProperties {
    pub build_type: Option<AnimationDgmBuildType>, // allAtOnce
    pub reverse: Option<bool>,                     // false
}

pub struct AnimationChartBuildProperties {
    pub build_type: Option<AnimationChartBuildType>, // allAtOnce
    pub animate_bg: Option<bool>,                    // true
}

pub struct OfficeStyleSheet {
    pub name: Option<String>, // ""
    pub theme_elements: Box<BaseStyles>,
    pub object_defaults: Option<ObjectStyleDefaults>,
    pub extra_color_scheme_list: Vec<ColorSchemeAndMapping>,
    pub custom_color_list: Vec<CustomColor>,
}

impl OfficeStyleSheet {
    pub fn from_zip_file(zip_file: &mut ZipFile<'_>) -> Result<Self> {
        let mut xml_string = String::new();
        zip_file.read_to_string(&mut xml_string)?;
        let xml_node = XmlNode::from_str(xml_string.as_str())?;

        match Self::from_xml_element(&xml_node) {
            Ok(v) => Ok(v),
            Err(err) => Err(err),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name = xml_node.attribute("name").cloned();
        let mut theme_elements = None;
        let mut object_defaults = None;
        let mut extra_color_scheme_list = Vec::new();
        let mut custom_color_list = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "themeElements" => theme_elements = Some(Box::new(BaseStyles::from_xml_element(child_node)?)),
                "objectDefaults" => object_defaults = Some(ObjectStyleDefaults::from_xml_element(child_node)?),
                "extraClrSchemeLst" => {
                    for extra_color_scheme_node in &child_node.child_nodes {
                        extra_color_scheme_list.push(
                            ColorSchemeAndMapping::from_xml_element(extra_color_scheme_node)?
                        );
                    }
                }
                "custClrLst" => {
                    for cust_color_node in &child_node.child_nodes {
                        custom_color_list.push(CustomColor::from_xml_element(cust_color_node)?);
                    }
                }
                _ => (),
            }
        }

        let theme_elements =
            theme_elements.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "themeElements"))?;

        Ok(Self {
            name,
            theme_elements,
            object_defaults,
            extra_color_scheme_list,
            custom_color_list,
        })
    }
}

pub struct BaseStyles {
    pub color_scheme: Box<ColorScheme>,
    pub font_scheme: FontScheme,
    pub format_scheme: Box<StyleMatrix>,
}

impl BaseStyles {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut color_scheme = None;
        let mut font_scheme = None;
        let mut format_scheme = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "clrScheme" => color_scheme = Some(Box::new(ColorScheme::from_xml_element(child_node)?)),
                "fontScheme" => font_scheme = Some(FontScheme::from_xml_element(child_node)?),
                "fmtScheme" => format_scheme = Some(Box::new(StyleMatrix::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let color_scheme =
            color_scheme.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "clrScheme"))?;
        let font_scheme = font_scheme.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "fontScheme"))?;
        let format_scheme =
            format_scheme.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "fmtScheme"))?;

        Ok(Self {
            color_scheme,
            font_scheme,
            format_scheme,
        })
    }
}

pub struct StyleMatrix {
    pub name: Option<String>,                      // ""
    pub fill_style_list: Vec<FillProperties>,      // minOccurs: 3
    pub line_style_list: Vec<Box<LineProperties>>, // minOccurs: 3
    pub effect_style_list: Vec<EffectStyleItem>,   // minOccurs: 3
    pub bg_fill_style_list: Vec<FillProperties>,   // minOccurs: 3
}

impl StyleMatrix {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name = xml_node.attribute("name").cloned();
        let mut fill_style_list = Vec::new();
        let mut line_style_list = Vec::new();
        let effect_style_list = Vec::new();
        let mut bg_fill_style_list = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "fillStyleLst" => {
                    for fill_style_node in &child_node.child_nodes {
                        fill_style_list.push(FillProperties::from_xml_element(fill_style_node)?);
                    }
                }
                "lnStyleLst" => {
                    for line_style_node in &child_node.child_nodes {
                        line_style_list.push(Box::new(LineProperties::from_xml_element(line_style_node)?));
                    }
                }
                // TODO: effect_style_list
                "bgFillStyleLst" => {
                    for bg_fill_style_node in &child_node.child_nodes {
                        bg_fill_style_list.push(FillProperties::from_xml_element(bg_fill_style_node)?);
                    }
                }
                _ => (),
            }
        }

        if fill_style_list.len() < 3 {
            return Err(LimitViolationError::new(
                xml_node.name.clone(),
                "fillStyleLst",
                Limit::Value(3),
                Limit::Unbounded,
                fill_style_list.len() as u32,
            )
            .into());
        }

        if line_style_list.len() < 3 {
            return Err(LimitViolationError::new(
                xml_node.name.clone(),
                "lnStyleLst",
                Limit::Value(3),
                Limit::Unbounded,
                line_style_list.len() as u32,
            )
            .into());
        }

        if bg_fill_style_list.len() < 3 {
            return Err(LimitViolationError::new(
                xml_node.name.clone(),
                "bgFillStyleLst",
                Limit::Value(3),
                Limit::Unbounded,
                bg_fill_style_list.len() as u32,
            )
            .into());
        }

        Ok(Self {
            name,
            fill_style_list,
            line_style_list,
            effect_style_list,
            bg_fill_style_list,
        })
    }
}

pub struct ObjectStyleDefaults {
    pub shape_definition: Option<Box<DefaultShapeDefinition>>,
    pub line_definition: Option<Box<DefaultShapeDefinition>>,
    pub text_definition: Option<Box<DefaultShapeDefinition>>,
}

impl ObjectStyleDefaults {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_definition = None;
        let mut line_definition = None;
        let mut text_definition = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "spDef" => shape_definition = Some(Box::new(DefaultShapeDefinition::from_xml_element(child_node)?)),
                "lnDef" => line_definition = Some(Box::new(DefaultShapeDefinition::from_xml_element(child_node)?)),
                "txDef" => text_definition = Some(Box::new(DefaultShapeDefinition::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        Ok(Self { shape_definition, line_definition, text_definition })
    }
}

pub struct DefaultShapeDefinition {
    pub shape_properties: Box<ShapeProperties>,
    pub text_body_properties: Box<TextBodyProperties>,
    pub text_list_style: Box<TextListStyle>,
    pub shape_style: Option<Box<ShapeStyle>>,
}

impl DefaultShapeDefinition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_properties = None;
        let mut text_body_properties = None;
        let mut text_list_style = None;
        let mut shape_style = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "spPr" => shape_properties = Some(Box::new(ShapeProperties::from_xml_element(child_node)?)),
                "bodyPr" => text_body_properties = Some(Box::new(
                    TextBodyProperties::from_xml_element(child_node)?
                )),
                "lstStyle" => text_list_style = Some(Box::new(TextListStyle::from_xml_element(child_node)?)),
                "style" => shape_style = Some(Box::new(ShapeStyle::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let shape_properties =
            shape_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "spPr"))?;
        let text_body_properties =
            text_body_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "bodyPr"))?;
        let text_list_style =
            text_list_style.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lstStyle"))?;

        Ok(Self {
            shape_properties,
            text_body_properties,
            text_list_style,
            shape_style,
        })
    }
}