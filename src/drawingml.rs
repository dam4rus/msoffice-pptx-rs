// TODO: This module defines shared types between different OOX file formats. It should be refactored into a different crate, if these types are needed.
use ::xml::{XmlNode, parse_optional_xml_attribute, parse_xml_bool};
use ::error::{NotGroupMemberError, MissingAttributeError, MissingChildNodeError, LimitViolationError, Limit, XmlError};
use ::relationship::RelationshipId;

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
            "tint" | "shade" | "comp" | "inv" | "gray"
            | "alpha" | "alphaOff" | "alphaMod"
            | "hue" | "hueOff" | "hueMod"
            | "sat" | "satOff" | "satMod"
            | "lum" | "lumOff" | "lumMod"
            | "red" | "redOff" | "redMod"
            | "green" | "greenOff" | "greenMod"
            | "blue" | "blueOff" | "blueMod"
            | "gamma" | "invGamma" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<ColorTransform, NotGroupMemberError> {
        match xml_node.local_name() {
            "tint" => Ok(ColorTransform::Tint(xml_node.attribute("val").unwrap().parse().unwrap())),
            "shade" => Ok(ColorTransform::Shade(xml_node.attribute("val").unwrap().parse().unwrap())),
            "comp" => Ok(ColorTransform::Complement),
            "inv" => Ok(ColorTransform::Inverse),
            "gray" => Ok(ColorTransform::Grayscale),
            "alpha" => Ok(ColorTransform::Alpha(xml_node.attribute("val").unwrap().parse().unwrap())),
            "alphaOff" => Ok(ColorTransform::AlphaOffset(xml_node.attribute("val").unwrap().parse().unwrap())),
            "alphaMod" => Ok(ColorTransform::AlphaModulate(xml_node.attribute("val").unwrap().parse().unwrap())),
            "hue" => Ok(ColorTransform::Hue(xml_node.attribute("val").unwrap().parse().unwrap())),
            "hueOff" => Ok(ColorTransform::HueOffset(xml_node.attribute("val").unwrap().parse().unwrap())),
            "hueMod" => Ok(ColorTransform::HueModulate(xml_node.attribute("val").unwrap().parse().unwrap())),
            "sat" => Ok(ColorTransform::Saturation(xml_node.attribute("val").unwrap().parse().unwrap())),
            "satOff" => Ok(ColorTransform::SaturationOffset(xml_node.attribute("val").unwrap().parse().unwrap())),
            "satMod" => Ok(ColorTransform::SaturationModulate(xml_node.attribute("val").unwrap().parse().unwrap())),
            "lum" => Ok(ColorTransform::Luminance(xml_node.attribute("val").unwrap().parse().unwrap())),
            "lumOff" => Ok(ColorTransform::LuminanceOffset(xml_node.attribute("val").unwrap().parse().unwrap())),
            "lumMod" => Ok(ColorTransform::LuminanceModulate(xml_node.attribute("val").unwrap().parse().unwrap())),
            "red" => Ok(ColorTransform::Red(xml_node.attribute("val").unwrap().parse().unwrap())),
            "redOff" => Ok(ColorTransform::RedOffset(xml_node.attribute("val").unwrap().parse().unwrap())),
            "redMod" => Ok(ColorTransform::RedModulate(xml_node.attribute("val").unwrap().parse().unwrap())),
            "green" => Ok(ColorTransform::Green(xml_node.attribute("val").unwrap().parse().unwrap())),
            "greenOff" => Ok(ColorTransform::GreenOffset(xml_node.attribute("val").unwrap().parse().unwrap())),
            "greenMod" => Ok(ColorTransform::GreenModulate(xml_node.attribute("val").unwrap().parse().unwrap())),
            "blue" => Ok(ColorTransform::Blue(xml_node.attribute("val").unwrap().parse().unwrap())),
            "blueOff" => Ok(ColorTransform::BlueOffset(xml_node.attribute("val").unwrap().parse().unwrap())),
            "blueMod" => Ok(ColorTransform::BlueModulate(xml_node.attribute("val").unwrap().parse().unwrap())),
            "gamma" => Ok(ColorTransform::Gamma),
            "invGamma" => Ok(ColorTransform::InverseGamma),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_ColorTransform")),
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
    fn from_rgb(r: Percentage, g: Percentage, b: Percentage) -> ScRgbColor {
        ScRgbColor {
            r: r,
            g: g,
            b: b,
            color_transforms: Vec::new(),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<ScRgbColor, MissingAttributeError> {
        let mut opt_r = None;
        let mut opt_g = None;
        let mut opt_b = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r" => opt_r = Some(value.parse().unwrap()),
                "g" => opt_g = Some(value.parse().unwrap()),
                "b" => opt_b = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        let mut instance = match (opt_r, opt_g, opt_b) {
            (Some(r), Some(g), Some(b)) => ScRgbColor::from_rgb(r, g, b),
            (None, _, _) => return Err(MissingAttributeError { attr: "r" }),
            (_, None, _) => return Err(MissingAttributeError { attr: "g" }),
            (_, _, None) => return Err(MissingAttributeError { attr: "b" }),
        };

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                instance.color_transforms.push(ColorTransform::from_xml_element(child_node).unwrap());
            }
        }

        Ok(instance)
    }
}

/// SRgbColor
pub struct SRgbColor {
    pub value: u32,
    pub color_transforms: Vec<ColorTransform>,
}

impl SRgbColor {
    fn from_rgb(rgb: u32) -> SRgbColor {
        SRgbColor {
            value: rgb,
            color_transforms: Vec::new(),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<SRgbColor, MissingAttributeError> {
        let mut instance = match xml_node.attribute("val") {
            Some(val) => SRgbColor::from_rgb(u32::from_str_radix(val, 16).unwrap()),
            None => return Err(MissingAttributeError { attr: "val" }),
        };

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                instance.color_transforms.push(ColorTransform::from_xml_element(child_node).unwrap());
            }
        }

        Ok(instance)
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
    fn from_hsl(
        hue: PositiveFixedAngle,
        saturation: Percentage,
        luminance: Percentage
    ) -> HslColor {
        HslColor {
            hue: hue,
            saturation: saturation,
            luminance: luminance,
            color_transforms: Vec::new(),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<HslColor, MissingAttributeError> {
        let mut opt_h = None;
        let mut opt_s = None;
        let mut opt_l = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "hue" => opt_h = Some(value.parse().unwrap()),
                "sat" => opt_s = Some(value.parse().unwrap()),
                "lum" => opt_l = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        let mut instance = match (opt_h, opt_s, opt_l) {
            (Some(h), Some(s), Some(l)) => HslColor::from_hsl(h, s, l),
            (None, _, _) => return Err(MissingAttributeError { attr: "hue" }),
            (_, None, _) => return Err(MissingAttributeError { attr: "sat" }),
            (_, _, None) => return Err(MissingAttributeError { attr: "lum" }),
        };

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                instance.color_transforms.push(ColorTransform::from_xml_element(child_node).unwrap());
            }
        }

        Ok(instance)
    }
}

/// SystemColor
pub struct SystemColor {
    pub value: SystemColorVal,
    pub last_color: Option<HexColorRGB>,
    pub color_transforms: Vec<ColorTransform>,
}

impl SystemColor {
    fn from_color_value(value: SystemColorVal) -> SystemColor {
        SystemColor {
            value: value,
            last_color: None,
            color_transforms: Vec::new(),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<SystemColor, MissingAttributeError> {
        let mut opt_val = None;
        let mut opt_last_color = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "val" => opt_val = Some(value.parse().unwrap()),
                "lastClr" => opt_last_color = Some(value.clone()),
                _ => (),
            }
        }

        let mut instance = match opt_val {
            Some(val) => SystemColor::from_color_value(val),
            None => return Err(MissingAttributeError { attr: "val" }),
        };

        instance.last_color = opt_last_color;

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                instance.color_transforms.push(ColorTransform::from_xml_element(child_node).unwrap());
            }
        }

        Ok(instance)
    }
}

/// PresetColor
pub struct PresetColor {
    pub value: PresetColorVal,
    pub color_transforms: Vec<ColorTransform>,
}

impl PresetColor {
    fn from_color_value(value: PresetColorVal) -> PresetColor {
        PresetColor {
            value: value,
            color_transforms: Vec::new(),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<PresetColor, MissingAttributeError> {
        let mut instance = match xml_node.attribute("val") {
            Some(val) => PresetColor::from_color_value(val.parse().unwrap()),
            None => return Err(MissingAttributeError { attr: "val" }),
        };

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                instance.color_transforms.push(ColorTransform::from_xml_element(child_node).unwrap());
            }
        }

        Ok(instance)
    }
}

/// SchemeColor
pub struct SchemeColor {
    pub value: SchemeColorVal,
    pub color_transforms: Vec<ColorTransform>,
}

impl SchemeColor {
    fn from_scheme_value(value: SchemeColorVal) -> SchemeColor {
        SchemeColor {
            value: value,
            color_transforms: Vec::new(),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<SchemeColor, MissingAttributeError> {
        let mut instance = match xml_node.attribute("val") {
            Some(val) => SchemeColor::from_scheme_value(val.parse().unwrap()),
            None => return Err(MissingAttributeError { attr: "val" }),
        };

        for child_node in &xml_node.child_nodes {
            if ColorTransform::is_choice_member(child_node.local_name()) {
                instance.color_transforms.push(ColorTransform::from_xml_element(child_node).unwrap());
            }
        }

        Ok(instance)
    }
}

/// Color
pub enum Color {
    ScRgbColor(ScRgbColor),
    SRgbColor(SRgbColor),
    HslColor(HslColor),
    SystemColor(SystemColor),
    SchemeColor(SchemeColor),
    PresetColor(PresetColor),
}

impl Color {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "scrgbClr" | "srgbClr" | "hslClr" | "sysClr" | "schemeClr" | "prstClr" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Color, NotGroupMemberError> {
        match xml_node.local_name() {
            "scrgbClr" => Ok(Color::ScRgbColor(ScRgbColor::from_xml_element(xml_node).unwrap())),
            "srgbClr" => Ok(Color::SRgbColor(SRgbColor::from_xml_element(xml_node).unwrap())),
            "hslClr" => Ok(Color::HslColor(HslColor::from_xml_element(xml_node).unwrap())),
            "sysClr" => Ok(Color::SystemColor(SystemColor::from_xml_element(xml_node).unwrap())),
            "schemeClr" => Ok(Color::SchemeColor(SchemeColor::from_xml_element(xml_node).unwrap())),
            "prstClr" => Ok(Color::PresetColor(PresetColor::from_xml_element(xml_node).unwrap())),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_ColorChoice")),
        }
    }
}

pub struct CustomColor {
    pub color: Color,
    pub name: Option<String>,
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
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self, XmlError> {
        let mut opt_name = None;
        let mut opt_dk1 = None;
        let mut opt_lt1 = None;
        let mut opt_dk2 = None;
        let mut opt_lt2 = None;
        let mut opt_accent1 = None;
        let mut opt_accent2 = None;
        let mut opt_accent3 = None;
        let mut opt_accent4 = None;
        let mut opt_accent5 = None;
        let mut opt_accent6 = None;
        let mut opt_hyperlink = None;
        let mut opt_follow_hyperlink = None;
        
        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "name" => opt_name = Some(value.clone()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "dk1" => opt_dk1 = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "lt1" => opt_lt1 = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "dk2" => opt_dk2 = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "lt2" => opt_lt2 = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "accent1" => opt_accent1 = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "accent2" => opt_accent2 = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "accent3" => opt_accent3 = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "accent4" => opt_accent4 = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "accent5" => opt_accent5 = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "accent6" => opt_accent6 = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "hlink" => opt_hyperlink = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                "folHlink" => opt_follow_hyperlink = Some(Color::from_xml_element(&child_node.child_nodes[0])?),
                _ => (),
            }
        }

        let name = opt_name.ok_or(XmlError::from(MissingAttributeError::new("name")))?;
        let dark1 = opt_dk1.ok_or(XmlError::from(MissingChildNodeError::new("dk1")))?;
        let light1 = opt_lt1.ok_or(XmlError::from(MissingChildNodeError::new("lt1")))?;
        let dark2 = opt_dk2.ok_or(XmlError::from(MissingChildNodeError::new("dk2")))?;
        let light2 = opt_lt2.ok_or(XmlError::from(MissingChildNodeError::new("lt2")))?;
        let accent1 = opt_accent1.ok_or(XmlError::from(MissingChildNodeError::new("accent1")))?;
        let accent2 = opt_accent2.ok_or(XmlError::from(MissingChildNodeError::new("accent2")))?;
        let accent3 = opt_accent3.ok_or(XmlError::from(MissingChildNodeError::new("accent3")))?;
        let accent4 = opt_accent4.ok_or(XmlError::from(MissingChildNodeError::new("accent4")))?;
        let accent5 = opt_accent5.ok_or(XmlError::from(MissingChildNodeError::new("accent5")))?;
        let accent6 = opt_accent6.ok_or(XmlError::from(MissingChildNodeError::new("accent6")))?;
        let hyperlink = opt_hyperlink.ok_or(XmlError::from(MissingChildNodeError::new("hlink")))?;
        let followed_hyperlink = opt_follow_hyperlink.ok_or(XmlError::from(MissingChildNodeError::new("folHlink")))?;

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
    UseMasterColorMapping,
    OverrideColorMapping(ColorMapping),
}

pub struct ColorSchemeAndMapping {
    pub color_scheme: ColorScheme,
    pub color_mapping: Option<ColorMapping>,
}

/// GradientStop
pub struct GradientStop {
    pub color: Color,
    pub position: PositiveFixedPercentage,
}

impl GradientStop {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<GradientStop, MissingAttributeError> {
        let pos = match xml_node.attribute("pos") {
            Some(val) => val.parse().unwrap(),
            None => return Err(MissingAttributeError::new("pos")),
        };

        let child_node = match xml_node.child_nodes.get(0) {
            Some(node) => node,
            None => return Err(MissingAttributeError::new("pos")),
        };

        if !Color::is_choice_member(child_node.local_name()) {
            return Err(MissingAttributeError::new("color"));
        }

        let color = match Color::from_xml_element(child_node) {
            Ok(clr) => clr,
            Err(_) => return Err(MissingAttributeError::new("color")),
        };

        Ok(GradientStop { color: color, position: pos })
    }
}

pub struct LinearShadeProperties {
    pub angle: Option<PositiveFixedAngle>,
    pub scaled: Option<bool>,
}

impl LinearShadeProperties {
    fn new() -> Self {
        Self {
            angle: None,
            scaled: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Self {
        let mut instance = Self::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "ang" => instance.angle = Some(value.parse::<i32>().unwrap()),
                "scaled" => instance.scaled = Some(parse_xml_bool(value.as_str()).unwrap()),
                _ => (),
            }
        }

        instance
    }
}

pub struct PathShadeProperties {
    pub fill_to_rect: Option<RelativeRect>,
    pub path: Option<PathShadeType>,
}

impl PathShadeProperties {
    fn new() -> Self {
        Self {
            fill_to_rect: None,
            path: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Self {
        let mut instance = Self::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "path" => instance.path = Some(value.parse::<PathShadeType>().unwrap()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "fillToRect" => instance.fill_to_rect = Some(RelativeRect::from_xml_element(child_node)),
                _ => (),
            }
        }

        instance
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

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self, NotGroupMemberError> {
        match xml_node.local_name() {
            "lin" => Ok(ShadeProperties::Linear(LinearShadeProperties::from_xml_element(xml_node))),
            "path" => Ok(ShadeProperties::Path(PathShadeProperties::from_xml_element(xml_node))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_ShadeProperties")),
        }
    }
}

/// GradientFillProperties
pub struct GradientFillProperties {
    pub gradient_stop_list: Vec<GradientStop>, // length: 2 <= n <= inf
    pub shade_properties: Option<ShadeProperties>,
    pub tile_rect: Option<RelativeRect>,
    pub flip: Option<TileFlipMode>,
    pub rotate_with_shape: Option<bool>,
}

impl GradientFillProperties {
    fn new() -> Self {
        Self {
            gradient_stop_list: Vec::new(),
            shade_properties: None,
            tile_rect: None,
            flip: None,
            rotate_with_shape: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Self {
        let mut instance = GradientFillProperties::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "flip" => instance.flip = Some(value.parse().unwrap()),
                "rotWithShape" => instance.rotate_with_shape = Some(parse_xml_bool(value).unwrap()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            let local_name = child_node.local_name();
            if ShadeProperties::is_choice_member(local_name) {
                instance.shade_properties = Some(ShadeProperties::from_xml_element(child_node).unwrap());
            } else {
                match child_node.local_name() {
                    "gsLst" => {
                        for gs_node in &child_node.child_nodes {
                            instance.gradient_stop_list.push(GradientStop::from_xml_element(gs_node).unwrap());
                        }
                    }
                    "tileRect" => instance.tile_rect = Some(RelativeRect::from_xml_element(child_node)),
                    _ => (),
                }
            }
        }

        instance
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

pub struct StretchInfoProperties {
    pub fill_rect: Option<RelativeRect>,
}

pub enum FillModeProperties {
    Tile(TileInfoProperties),
    Stretch(StretchInfoProperties),
}

pub struct BlipFillProperties {
    pub blip: Option<Blip>,
    pub source_rect: Option<RelativeRect>,
    pub fill_mode_properties: Option<FillModeProperties>,
    pub dpi: Option<u32>,
    pub rotate_with_shape: Option<bool>,
}

pub struct PatternFillProperties {
    pub fg_color: Option<Color>,
    pub bg_color: Option<Color>,
    pub preset: Option<PresetPatternVal>,
}

pub enum FillProperties {
    NoFill,
    SolidFill(Color),
    GradientFill(GradientFillProperties),
    BlipFill(BlipFillProperties),
    PatternFill(PatternFillProperties),
    GroupFill
}

impl FillProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self, NotGroupMemberError> {
        match xml_node.local_name() {
            "noFill" => Ok(FillProperties::NoFill),
            "solidFill" => Ok(FillProperties::SolidFill(Color::from_xml_element(&xml_node.child_nodes[0])?)),
            "gradFill" => Ok(FillProperties::GradientFill(GradientFillProperties::from_xml_element(xml_node))),
            // TODO: implement
            // <xsd:element name="blipFill" type="CT_BlipFillProperties" minOccurs="1" maxOccurs="1"/>
            // <xsd:element name="pattFill" type="CT_PatternFillProperties" minOccurs="1" maxOccurs="1"/>
            // <xsd:element name="grpFill" type="CT_GroupFillProperties" minOccurs="1" maxOccurs="1"/>
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_FillProperties")),
        }
    }
}

/// LineFillProperties
pub enum LineFillProperties {
    NoFill,
    SolidFill(Color),
    GradientFill(GradientFillProperties),
    PatternFill(PatternFillProperties),
}

impl LineFillProperties {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "noFill" | "solidFill" | "gradFill" | "pattFill" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<LineFillProperties, NotGroupMemberError> {
        match xml_node.local_name() {
            "noFill" => Ok(LineFillProperties::NoFill),
            "solidFill" => {
                let child_node = match xml_node.child_nodes.get(0) {
                    Some(node) => node,
                    None => return Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_Color")),
                };

                if !Color::is_choice_member(child_node.local_name()) {
                    return Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_Color"));
                }

                Ok(LineFillProperties::SolidFill(Color::from_xml_element(child_node).unwrap()))
            },
            "gradFill" => Ok(LineFillProperties::GradientFill(GradientFillProperties::from_xml_element(xml_node))),
            // TODO: implement pattFill
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_LineFillProperties")),
        }
    }
}

/// DashStop
pub struct DashStop {
    pub dash_length: PositivePercentage,
    pub space_length: PositivePercentage,
}

impl DashStop {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<DashStop, MissingAttributeError> {
        let mut opt_dash_length = None;
        let mut opt_space_length = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "d" => opt_dash_length = Some(value.parse().unwrap()),
                "sp" => opt_space_length = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        match (opt_dash_length, opt_space_length) {
            (Some(dash), Some(space)) => Ok(DashStop { dash_length: dash, space_length: space }),
            (None, _) => Err(MissingAttributeError { attr: "d"} ),
            (_, None) => Err(MissingAttributeError { attr: "sp"} ),
        }
    }
}

/// LineDashProperties
pub enum LineDashProperties {
    PresetDash(PresetLineDashVal),
    CustomDash(Vec<DashStop>)
}

impl LineDashProperties {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "prstDash" | "custDash" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<LineDashProperties, NotGroupMemberError> {
        match xml_node.local_name() {
            "prstDash" => Ok(LineDashProperties::PresetDash(xml_node.attribute("val").unwrap().parse().unwrap())),
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
            },
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_LineDashProperties")),
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

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<LineJoinProperties, NotGroupMemberError> {
        match xml_node.local_name() {
            "round" => Ok(LineJoinProperties::Round),
            "bevel" => Ok(LineJoinProperties::Bevel),
            "miter" => Ok(LineJoinProperties::Miter(parse_optional_xml_attribute(xml_node.attribute("lim").unwrap()))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_LineJoinProperties")),
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
    fn new() -> LineEndProperties {
        LineEndProperties {
            end_type: None,
            width: None,
            length: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> LineEndProperties {
        let mut instance = LineEndProperties::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => instance.end_type = Some(value.parse().unwrap()),
                "width" => instance.width = Some(value.parse().unwrap()),
                "length" => instance.length = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        instance
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
    fn new() -> LineProperties {
        LineProperties {
            width: None,
            cap: None,
            compound: None,
            pen_alignment: None,
            fill_properties: None,
            dash_properties: None,
            join_properties: None,
            head_end: None,
            tail_end: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> LineProperties {
        let mut instance = LineProperties::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "w" => instance.width = Some(value.parse().unwrap()),
                "cap" => instance.cap = Some(value.parse().unwrap()),
                "cmpd" => instance.compound = Some(value.parse().unwrap()),
                "algn" => instance.pen_alignment = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            if LineFillProperties::is_choice_member(child_node.local_name()) {
                instance.fill_properties = Some(LineFillProperties::from_xml_element(child_node).unwrap());
            } else if LineDashProperties::is_choice_member(child_node.local_name()) {
                instance.dash_properties = Some(LineDashProperties::from_xml_element(child_node).unwrap());
            } else if LineJoinProperties::is_choice_member(child_node.local_name()) {
                instance.join_properties = Some(LineJoinProperties::from_xml_element(child_node).unwrap());
            } else {
                match child_node.local_name() {
                    "headEnd" => instance.head_end = Some(LineEndProperties::from_xml_element(child_node)),
                    "tailEnd" => instance.tail_end = Some(LineEndProperties::from_xml_element(child_node)),
                    _ => (),
                }
            }
        }

        instance
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
    pub fn new() -> RelativeRect {
        RelativeRect {
            left: None,
            top: None,
            right: None,
            bottom: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> RelativeRect {
        let mut instance = RelativeRect::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "l" => instance.left = Some(value.parse().unwrap()),
                "t" => instance.top = Some(value.parse().unwrap()),
                "r" => instance.right = Some(value.parse().unwrap()),
                "b" => instance.bottom = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        instance
    }
}

pub struct Point2D {
    pub x: Coordinate,  
    pub y: Coordinate,
}

/// PositiveSize2D
pub struct PositiveSize2D {
    pub width: PositiveCoordinate,
    pub height: PositiveCoordinate,
}

impl PositiveSize2D {
    fn new(width: PositiveCoordinate, height: PositiveCoordinate) -> PositiveSize2D {
        PositiveSize2D {
            width: width,
            height: height,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<PositiveSize2D, MissingAttributeError> {
        let mut opt_width: Option<PositiveCoordinate> = None;
        let mut opt_height: Option<PositiveCoordinate> = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "cx" => opt_width = Some(value.parse().unwrap()),
                "cy" => opt_height = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        match (opt_width, opt_height) {
            (Some(w), Some(h)) => Ok(PositiveSize2D::new(w, h)),
            (None, _) => Err(MissingAttributeError { attr: "cx" }),
            (_, None) => Err(MissingAttributeError { attr: "cy" }),
        }
    }
}

pub struct StyleMatrixReference {
    pub index: StyleMatrixColumnIndex,
    pub color: Option<Color>,
}

/// EffectContainer
pub struct EffectContainer {
    pub effects: Vec<Effect>,
    pub container_type: Option<EffectContainerType>,
    pub name: Option<String>,
}

impl EffectContainer {
    pub fn new() -> EffectContainer {
        EffectContainer {
            effects: Vec::new(),
            container_type: None,
            name: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> EffectContainer {
        let mut instance = EffectContainer::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => instance.container_type = Some(value.parse().unwrap()),
                "name" => instance.name = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        // TODO: implement

        instance
    }
}

/// AlphaBiLevelEffect
pub struct AlphaBiLevelEffect {
    pub threshold: PositiveFixedPercentage,
}

impl AlphaBiLevelEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<AlphaBiLevelEffect, MissingAttributeError> {
        match xml_node.attribute("thresh") {
            Some(threshold) => Ok(AlphaBiLevelEffect { threshold: threshold.parse().unwrap() }),
            None => Err(MissingAttributeError { attr: "thresh"}),
        }
    }
}

/// AlphaInverseEffect
pub struct AlphaInverseEffect {
    pub color: Option<Color>,
}

impl AlphaInverseEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> AlphaInverseEffect {

        if let Some(child_node) = xml_node.child_nodes.get(0) {
            if Color::is_choice_member(child_node.local_name()) {
                match Color::from_xml_element(child_node) {
                    Ok(clr) => return AlphaInverseEffect { color: Some(clr) },
                    Err(err) => println!("Failed to parse EG_ColorChoice node of AlphaInverseEffect child. Error: {}", err),
                }
            }
        }

        AlphaInverseEffect { color: None }
    }
}

/// AlphaModulateEffect
pub struct AlphaModulateEffect {
    pub container: EffectContainer,
}

impl AlphaModulateEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<AlphaModulateEffect, MissingChildNodeError> {

        if let Some(child_node) = xml_node.child_nodes.get(0) {
            if child_node.local_name() == "cont" {
                return Ok(AlphaModulateEffect { container: EffectContainer::from_xml_element(child_node) } );
            }
        }

        Err(MissingChildNodeError { child_node: "cont" })
    }
}

/// AlphaModulateFixedEffect
pub struct AlphaModulateFixedEffect {
    pub amount: Option<PositivePercentage>, // 1.0
}

impl AlphaModulateFixedEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> AlphaModulateFixedEffect {
        match xml_node.attribute("amt") {
            Some(amt) => AlphaModulateFixedEffect { amount: Some(amt.parse().unwrap()) },
            None => AlphaModulateFixedEffect { amount: None },
        }
    }
}

pub struct AlphaOutsetEffect {
    pub radius: Coordinate,
}

pub struct AlphaReplaceEffect {
    pub alpha: PositiveFixedPercentage,
}

pub struct BiLevelEffect {
    pub treshold: PositiveFixedPercentage,
}

pub struct BlendEffect {
    pub container: EffectContainer,
    pub blend: BlendMode,
}

pub struct BlurEffect {
    pub radius: PositiveCoordinate, // 0
    pub grow: bool, // true
}

pub struct ColorChangeEffect {
    pub color_from: Color,
    pub color_to: Color,
    pub use_alpha: bool, // true
}

pub struct ColorReplaceEffect {
    pub color: Color,
}


pub struct LuminanceEffect {
    pub bright: Option<FixedPercentage>,
    pub contrast: Option<FixedPercentage>,
}

pub struct DuotoneEffect {
    pub colors: [Color; 2],
}

pub struct FillEffect {
    pub fill: FillProperties,
}

pub struct FillOverlayEffect {
    pub fill: FillProperties,
    pub blend_mode: BlendMode,
}

pub struct GlowEffect {
    pub color: Color,
    pub radius: Option<PositiveCoordinate>, // 0
}

pub struct HslEffect {
    pub hue: Option<PositiveFixedAngle>, // 0
    pub saturation: Option<FixedPercentage>, // 0%
    pub luminance: Option<FixedPercentage>, // 0%
}

pub struct InnerShadowEffect {
    pub color: Color,
    pub blur_radius: Option<PositiveCoordinate>, // 0
    pub distance: Option<PositiveCoordinate>, // 0
    pub direction: Option<PositiveFixedAngle>, // 0
}

pub struct OuterShadowEffect {
    pub color: Color,
    pub blur_radius: Option<PositiveCoordinate>, // 0
    pub distance: Option<PositiveCoordinate>, // 0
    pub direction: Option<PositiveFixedAngle>, // 0
    pub scale_x: Option<Percentage>, // 100000
    pub scale_y: Option<Percentage>, // 100000
    pub skew_x: Option<FixedAngle>, // 0
    pub skew_y: Option<FixedAngle>, // 0
    pub alignment: Option<RectAlignment>, // b
    pub rotate_with_shape: Option<bool>, // true
}

pub struct PresetShadowEffect {
    pub color: Color,
    pub preset: PresetShadowVal,
    pub distance: Option<PositiveCoordinate>, // 0
    pub direction: Option<PositiveFixedAngle>, // 0
}

pub struct ReflectionEffect {
    pub blur_radius: Option<PositiveCoordinate>, // 0
    pub start_opacity: Option<PositiveFixedPercentage>, // 100000
    pub start_position: Option<PositiveFixedPercentage>, // 0
    pub end_opacity: Option<PositiveFixedPercentage>, // 0
    pub end_position: Option<PositiveFixedPercentage>, // 100000
    pub distance: Option<PositiveCoordinate>, // 0
    pub direction: Option<PositiveFixedAngle>, // 0
    pub fade_direction: Option<PositiveFixedAngle>, // 5400000
    pub scale_x: Option<Percentage>, // 100000
    pub scale_y: Option<Percentage>, // 100000
    pub skew_x: Option<FixedAngle>, // 0
    pub skew_y: Option<FixedAngle>, // 0
    pub alignment: Option<RectAlignment>, // b
    pub rotate_with_shape: Option<bool>, // true
}

pub struct RelativeOffsetEffect {
    pub translate_x: Option<Percentage>, // 0
    pub translate_y: Option<Percentage>, // 0
}

pub struct SoftEdgesEffect {
    pub radius: PositiveCoordinate,
}

pub struct TintEffect {
    pub hue: Option<PositiveFixedAngle>, // 0
    pub amount: Option<FixedPercentage>, // 0
}

pub struct TransformEffect {
    pub scale_x: Option<Percentage>, // 100000
    pub scale_y: Option<Percentage>, // 100000
    pub translate_x: Option<Coordinate>, // 0
    pub translate_y: Option<Coordinate>, // 0
    pub skew_x: Option<FixedAngle>, // 0
    pub skew_y: Option<FixedAngle>, // 0
}

pub enum Effect {
    Cont(EffectContainer),
    EffectReference(String),
    AlphaBiLevel(AlphaBiLevelEffect),
    AlphaCeiling,
    ALphaFloor,
    AlphaInv(AlphaInverseEffect),
    AlphaMod(AlphaModulateEffect),
    AlphaModFix(AlphaModulateFixedEffect),
    AlphaOutset(AlphaOutsetEffect),
    AlphaRepl(AlphaReplaceEffect),
    BiLevel(BiLevelEffect),
    Blend(BlendEffect),
    Blur(BlurEffect),
    ClrChange(ColorChangeEffect),
    ClrRepl(ColorReplaceEffect),
    Duotone(DuotoneEffect),
    Fill(FillEffect),
    FillOverlay(FillOverlayEffect),
    Glow(GlowEffect),
    Grayscl,
    Hsl(HslEffect),
    InnerShdw(InnerShadowEffect),
    Lum(LuminanceEffect),
    OuterShdw(OuterShadowEffect),
    PrstShadow(PresetShadowEffect),
    Reflection(ReflectionEffect),
    RelOff(RelativeOffsetEffect),
    SoftEdge(SoftEdgesEffect),
    Tint(TintEffect),
    Xfrm(TransformEffect),
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
    EffectList(EffectList),
    EffectContainer(EffectContainer),
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

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<BlipEffect, NotGroupMemberError> {
        match xml_node.local_name() {
            "alphaBiLevel" => Ok(BlipEffect::AlphaBiLevel(AlphaBiLevelEffect::from_xml_element(xml_node).unwrap())),
            "alphaCeiling" => Ok(BlipEffect::AlphaCeiling),
            "alphaFloor" => Ok(BlipEffect::AlphaFloor),
            "alphaInv" => Ok(BlipEffect::AlphaInv(AlphaInverseEffect::from_xml_element(xml_node))),
            "alphaMod" => Ok(BlipEffect::AlphaMod(AlphaModulateEffect::from_xml_element(xml_node).unwrap())),
            "alphaModFixed" => Ok(BlipEffect::AlphaModFix(AlphaModulateFixedEffect::from_xml_element(xml_node))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_BlipEffect")),
        }
    }
}

/// Blip
pub struct Blip {
    pub effect_list: Vec<BlipEffect>,
    pub embed_rel_id: Option<RelationshipId>,
    pub linked_rel_id: Option<RelationshipId>,
    pub compression: Option<BlipCompression>,
}

impl Blip {
    fn new() -> Blip {
        Blip {
            effect_list: Vec::new(),
            embed_rel_id: None,
            linked_rel_id: None,
            compression: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Blip {
        /*
        <xsd:complexType name="CT_Blip">
    <xsd:sequence>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="alphaBiLevel" type="CT_AlphaBiLevelEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaCeiling" type="CT_AlphaCeilingEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaFloor" type="CT_AlphaFloorEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaInv" type="CT_AlphaInverseEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaMod" type="CT_AlphaModulateEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaModFix" type="CT_AlphaModulateFixedEffect" minOccurs="1"
          maxOccurs="1"/>
        <xsd:element name="alphaRepl" type="CT_AlphaReplaceEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="biLevel" type="CT_BiLevelEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="blur" type="CT_BlurEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="clrChange" type="CT_ColorChangeEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="clrRepl" type="CT_ColorReplaceEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="duotone" type="CT_DuotoneEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="fillOverlay" type="CT_FillOverlayEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="grayscl" type="CT_GrayscaleEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="hsl" type="CT_HSLEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="lum" type="CT_LuminanceEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="tint" type="CT_TintEffect" minOccurs="1" maxOccurs="1"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attributeGroup ref="AG_Blob"/>
    <xsd:attribute name="cstate" type="ST_BlipCompression" use="optional" default="none"/>
  </xsd:complexType>
        */

        let mut instance = Blip::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r:embed" => instance.embed_rel_id = Some(value.clone()),
                "r:link" => instance.linked_rel_id = Some(value.clone()),
                "cstate" => instance.compression = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            if BlipEffect::is_choice_member(child_node.local_name()) {
                instance.effect_list.push(BlipEffect::from_xml_element(child_node).unwrap());
            }
        }

        instance
    }
}

/// TextFont
pub struct TextFont {
    pub typeface: TextTypeFace,
    pub panose: Option<Panose>,
    pub pitch_family: Option<i32>, // 0
    pub charset: Option<i32>, // 1
}

impl TextFont {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextFont, MissingAttributeError> {
        let mut opt_typeface = None;
        let mut opt_panose = None;
        let mut opt_pitch_family = None;
        let mut opt_charset = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "typeface" => opt_typeface = Some(value.clone()),
                "panose" => opt_panose = Some(value.clone()),
                "pitchFamily" => opt_pitch_family = Some(value.parse().unwrap()),
                "charset" => opt_charset = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        match opt_typeface {
            Some(typeface) => Ok(TextFont {
                typeface: typeface,
                panose: opt_panose,
                pitch_family: opt_pitch_family,
                charset: opt_charset,
            }),
            None => Err(MissingAttributeError { attr: "typeface" }),
        }
    }
}

pub struct SupplementalFont {
    pub script: String,
    pub typeface: TextTypeFace,
}

impl SupplementalFont {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self, MissingAttributeError> {
        let mut opt_script = None;
        let mut opt_typeface = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "script" => opt_script = Some(value.clone()),
                "typeface" => opt_typeface = Some(value.clone()),
                _ => (),
            }
        }

        let script = opt_script.ok_or(MissingAttributeError::new("script"))?;
        let typeface = opt_typeface.ok_or(MissingAttributeError::new("typeface"))?;

        Ok(Self {
            script,
            typeface,
        })
    }
}

/// TextSpacing
pub enum TextSpacing {
    Percent(TextSpacingPercent),
    Point(TextSpacingPoint),
}

impl TextSpacing {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextSpacing, NotGroupMemberError> {
        match xml_node.local_name() {
            "spcPct" => Ok(TextSpacing::Percent(xml_node.attribute("val").unwrap().parse().unwrap())),
            "spcPts" => Ok(TextSpacing::Point(xml_node.attribute("val").unwrap().parse().unwrap())),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextSpacing")),
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

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextBulletColor, NotGroupMemberError> {
        match xml_node.local_name() {
            "buClrTx" => Ok(TextBulletColor::FollowText),
            "buClr" => {
                match xml_node.child_nodes.get(0) {
                    Some(child) => Ok(TextBulletColor::Color(Color::from_xml_element(child).unwrap())),
                    None => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBulletColor")), // TODO: return correct error
                }
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBulletColor")),
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

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextBulletSize, NotGroupMemberError> {
        match xml_node.local_name() {
            "buSzTx" => Ok(TextBulletSize::FollowText),
            "buSzPct" => Ok(TextBulletSize::Percent(xml_node.attribute("val").unwrap().parse().unwrap())),
            "buSzPts" => Ok(TextBulletSize::Point(xml_node.attribute("val").unwrap().parse().unwrap())),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBulletSize")),
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

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextBulletTypeface, &'static str> {
        match xml_node.local_name() {
            "buFontTx" => Ok(TextBulletTypeface::FollowText),
            "buFont" => Ok(TextBulletTypeface::Font(TextFont::from_xml_element(xml_node).unwrap())),
            _ => Err("Xml element is not a TextBulletTypeFace group choice"),
        }
    }
}

/// TextBullet
pub enum TextBullet {
    None,
    AutoNumbered(TextAutonumberedBullet),
    Character(String),
    Picture(Blip),
}

impl TextBullet {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "buNone" | "buAutoNum" | "buChar" | "buBlip" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextBullet, NotGroupMemberError> {
        match xml_node.local_name() {
            "buNone" => Ok(TextBullet::None),
            "buAutoNum" => Ok(TextBullet::AutoNumbered(TextAutonumberedBullet::from_xml_element(xml_node).unwrap())),
            "buChar" => Ok(TextBullet::Character(xml_node.attribute("char").unwrap().clone())),
            "buBlip" => {
                match xml_node.child_nodes.get(0) {
                    Some(child_node) => Ok(TextBullet::Picture(Blip::from_xml_element(child_node))),
                    None => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBullet")), // TODO: return correct error
                }
            },
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBullet")),
        }
    }
}


/// TextAutonumberedBullet
pub struct TextAutonumberedBullet {
    pub scheme: TextAutonumberScheme,
    pub start_at: Option<TextBulletStartAtNum>,
}

impl TextAutonumberedBullet {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextAutonumberedBullet, MissingAttributeError> {
        let mut opt_scheme = None;
        let mut opt_start_at = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => opt_scheme = Some(value.parse().unwrap()),
                "startAt" => opt_start_at = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        if let Some(scheme) = opt_scheme {
            Ok(TextAutonumberedBullet {
                scheme: scheme,
                start_at: opt_start_at,
            })
        } else {
            Err(MissingAttributeError { attr: "type"})
        }
    }
}

/// TextTabStop
pub struct TextTabStop {
    pub position: Option<Coordinate32>,
    pub alignment: Option<TextTabAlignType>,
}

impl TextTabStop {
    pub fn new() -> TextTabStop {
        TextTabStop {
            position: None,
            alignment: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> TextTabStop {
        let mut instance = TextTabStop::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "pos" => instance.position = Some(value.parse().unwrap()),
                "algn" => instance.alignment = Some(value.parse().unwrap()),
                _ => (),
            }
        }

        instance
    }
}

pub enum TextUnderlineLine {
    FollowText,
    Line(Option<LineProperties>),
}

pub enum TextUnderlineFill {
    FollowText,
    Fill(FillProperties),
}

pub struct Hyperlink {
    pub relationship_id: Option<RelationshipId>,
    pub invalid_url: Option<String>,
    pub action: Option<String>,
    pub target_frame: Option<String>,
    pub tooltip: Option<String>,
    pub history: Option<bool>, // true
    pub highlight_click: Option<bool>, // false
    pub end_sound: Option<bool>, // false
    pub sound: Option<EmbeddedWAVAudioFile>,
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
    pub dirty: Option<bool>, // true
    pub spelling_error: Option<bool>, // false
    pub smarttag_clean: Option<bool>, // true
    pub smarttag_id: Option<u32>, // 0
    pub bookmark_link_target: Option<String>,
    pub line_properties: Option<LineProperties>,
    pub fill_properties: Option<FillProperties>,
    pub effect_properties: Option<EffectProperties>,
    pub highlight_color: Option<Color>,
    pub text_underline_line: Option<TextUnderlineLine>,
    pub text_underline_fill: Option<TextUnderlineFill>,
    pub latin_font: Option<TextFont>,
    pub east_asian_font: Option<TextFont>,
    pub complex_script_font: Option<TextFont>,
    pub symbol_font: Option<TextFont>,
    pub hyperlink_click: Option<Hyperlink>,
    pub hyperlink_mouse_over: Option<Hyperlink>,
}

impl TextCharacterProperties {
    fn new() -> TextCharacterProperties {
        TextCharacterProperties {
            kumimoji: None,
            language: None,
            alternative_language: None,
            font_size: None,
            bold: None,
            italic: None,
            underline: None,
            strikethrough: None,
            kerning: None,
            caps_type: None,
            spacing: None,
            normalize_heights: None,
            baseline: None,
            no_proofing: None,
            dirty: None,
            spelling_error: None,
            smarttag_clean: None,
            smarttag_id: None,
            bookmark_link_target: None,
            line_properties: None,
            fill_properties: None,
            effect_properties: None,
            highlight_color: None,
            text_underline_line: None,
            text_underline_fill: None,
            latin_font: None,
            east_asian_font: None,
            complex_script_font: None,
            symbol_font: None,
            hyperlink_click: None,
            hyperlink_mouse_over: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> TextCharacterProperties {
        let mut instance = TextCharacterProperties::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "kumimoji" => instance.kumimoji = Some(value.parse().unwrap()),
                "lang" => instance.language = Some(value.clone()),
                "altLang" => instance.alternative_language = Some(value.clone()),
                "sz" => instance.font_size = Some(value.parse().unwrap()),
                "b" => instance.bold = Some(value.parse().unwrap()),
                "i" => instance.italic = Some(value.parse().unwrap()),
                "u" => instance.underline = Some(value.parse().unwrap()),
                "strike" => instance.strikethrough = Some(value.parse().unwrap()),
                "kern" => instance.kerning = Some(value.parse().unwrap()),
                "cap" => instance.caps_type = Some(value.parse().unwrap()),
                "spc" => instance.spacing = Some(value.parse().unwrap()),
                "normalizeH" => instance.normalize_heights = Some(value.parse().unwrap()),
                "baseline" => instance.baseline = Some(value.parse().unwrap()),
                "noProof" => instance.no_proofing = Some(value.parse().unwrap()),
                "dirty" => instance.dirty = Some(value.parse().unwrap()),
                "err" => instance.spelling_error = Some(value.parse().unwrap()),
                "smtClean" => instance.smarttag_clean = Some(value.parse().unwrap()),
                "smtId" => instance.smarttag_id = Some(value.parse().unwrap()),
                "bmk" => instance.bookmark_link_target = Some(value.clone()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "ln" => instance.line_properties = Some(LineProperties::from_xml_element(child_node)),
                "latin" => instance.latin_font = Some(TextFont::from_xml_element(child_node).unwrap()),
                "ea" => instance.east_asian_font = Some(TextFont::from_xml_element(child_node).unwrap()),
                "cs" => instance.complex_script_font = Some(TextFont::from_xml_element(child_node).unwrap()),
                _ => (),
            };
        }

        instance

        /*
          <xsd:complexType name="CT_TextCharacterProperties">
    <xsd:sequence>
      <xsd:element name="ln" type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_FillProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_EffectProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="highlight" type="CT_Color" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_TextUnderlineLine" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_TextUnderlineFill" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sym" type="CT_TextFont" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkClick" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkMouseOver" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="rtl" type="CT_Boolean" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>

  </xsd:complexType>
        */
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
    pub default_run_properties: Option<TextCharacterProperties>,
}

impl TextParagraphProperties {
    pub fn new() -> TextParagraphProperties {
        TextParagraphProperties {
            margin_left: None,
            margin_right: None,
            level: None,
            indent: None,
            align: None,
            default_tab_size: None,
            rtl: None,
            east_asian_line_break: None,
            font_align: None,
            latin_line_break: None,
            hanging_punctuations: None,
            line_spacing: None,
            space_before: None,
            space_after: None,
            bullet_color: None,
            bullet_size: None,
            bullet_typeface: None,
            bullet: None,
            tab_stop_list: Vec::new(),
            default_run_properties: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> TextParagraphProperties {
        let mut instance = TextParagraphProperties::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "marL" => instance.margin_left = Some(value.parse().unwrap()),
                "marR" => instance.margin_right = Some(value.parse().unwrap()),
                "lvl" => instance.level = Some(value.parse().unwrap()),
                "indent" => instance.indent = Some(value.parse().unwrap()),
                "algn" => instance.align = Some(value.parse().unwrap()),
                "defTabSz" => instance.default_tab_size = Some(value.parse().unwrap()),
                "rtl" => instance.rtl = Some(parse_xml_bool(value).unwrap()),
                "eaLnBrk" => instance.east_asian_line_break = Some(parse_xml_bool(value).unwrap()),
                "fontAlgn" => instance.font_align = Some(value.parse().unwrap()),
                "latinLnBrk" => instance.latin_line_break = Some(parse_xml_bool(value).unwrap()),
                "hangingPunct" => instance.hanging_punctuations = Some(parse_xml_bool(value).unwrap()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            if TextBulletColor::is_choice_member(child_node.local_name()) {
                instance.bullet_color = Some(TextBulletColor::from_xml_element(child_node).unwrap());
            } else if TextBulletColor::is_choice_member(child_node.local_name()) {
                instance.bullet_size = Some(TextBulletSize::from_xml_element(child_node).unwrap());
            } else if TextBulletTypeface::is_choice_member(child_node.local_name()) {
                instance.bullet_typeface = Some(TextBulletTypeface::from_xml_element(child_node).unwrap());
            } else if TextBullet::is_choice_member(child_node.local_name()) {
                instance.bullet = Some(TextBullet::from_xml_element(child_node).unwrap());
            } else {
                match child_node.local_name() {
                    "lnSpc" => {
                        match child_node.child_nodes.get(0) {
                            Some(ln_spc_node) => instance.line_spacing = Some(TextSpacing::from_xml_element(ln_spc_node).unwrap()),
                            None => println!("lnSpc element has no child node"),
                        }
                    },
                    "spcBef" => {
                        match child_node.child_nodes.get(0) {
                            Some(spc_bef_node) => instance.line_spacing = Some(TextSpacing::from_xml_element(spc_bef_node).unwrap()),
                            None => println!("spcBef element has no child node"),
                        }
                    },
                    "spcAft" => {
                        match child_node.child_nodes.get(0) {
                            Some(spc_after_node) => instance.line_spacing = Some(TextSpacing::from_xml_element(spc_after_node).unwrap()),
                            None => println!("spcAft element has no child node"),
                        }
                    },
                    "tabLst" => instance.tab_stop_list.push(TextTabStop::from_xml_element(child_node)),
                    "defRPr" => instance.default_run_properties = Some(TextCharacterProperties::from_xml_element(child_node)),
                    _ => (),
                }
            }
        }

        instance
    }
}

pub struct TextParagraph {
    pub properties: Option<TextParagraphProperties>,
    pub text_run_list: Vec<TextRun>,
    pub end_paragraph_char_properties: Option<TextCharacterProperties>,
}

pub enum TextRun {
    RegularTextRun(RegularTextRun),
    LineBreak(TextLineBreak),
    TextField(TextField),
}

pub struct RegularTextRun {
    pub char_properties: Option<TextCharacterProperties>,
    pub text: String,
}

pub struct TextLineBreak {
    pub char_properties: Option<TextCharacterProperties>,
}

pub struct TextField {
    pub char_properties: Option<TextCharacterProperties>,
    pub paragraph_properties: Option<TextParagraph>,
    pub text: String,
    pub id: Guid,
    pub field_type: Option<String>,
}

/// TextListStyle
pub struct TextListStyle {
    pub def_paragraph_props: Option<TextParagraphProperties>,
    pub lvl1_paragraph_props: Option<TextParagraphProperties>,
    pub lvl2_paragraph_props: Option<TextParagraphProperties>,
    pub lvl3_paragraph_props: Option<TextParagraphProperties>,
    pub lvl4_paragraph_props: Option<TextParagraphProperties>,
    pub lvl5_paragraph_props: Option<TextParagraphProperties>,
    pub lvl6_paragraph_props: Option<TextParagraphProperties>,
    pub lvl7_paragraph_props: Option<TextParagraphProperties>,
    pub lvl8_paragraph_props: Option<TextParagraphProperties>,
    pub lvl9_paragraph_props: Option<TextParagraphProperties>,
}

impl TextListStyle {
    fn new() -> TextListStyle {
        TextListStyle {
            def_paragraph_props: None,
            lvl1_paragraph_props: None,
            lvl2_paragraph_props: None,
            lvl3_paragraph_props: None,
            lvl4_paragraph_props: None,
            lvl5_paragraph_props: None,
            lvl6_paragraph_props: None,
            lvl7_paragraph_props: None,
            lvl8_paragraph_props: None,
            lvl9_paragraph_props: None,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> TextListStyle {
        let mut instance = TextListStyle::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "defPPr" => instance.def_paragraph_props = Some(TextParagraphProperties::from_xml_element(child_node)),
                "lvl1pPr" => instance.lvl1_paragraph_props = Some(TextParagraphProperties::from_xml_element(child_node)),
                "lvl2pPr" => instance.lvl2_paragraph_props = Some(TextParagraphProperties::from_xml_element(child_node)),
                "lvl3pPr" => instance.lvl3_paragraph_props = Some(TextParagraphProperties::from_xml_element(child_node)),
                "lvl4pPr" => instance.lvl4_paragraph_props = Some(TextParagraphProperties::from_xml_element(child_node)),
                "lvl5pPr" => instance.lvl5_paragraph_props = Some(TextParagraphProperties::from_xml_element(child_node)),
                "lvl6pPr" => instance.lvl6_paragraph_props = Some(TextParagraphProperties::from_xml_element(child_node)),
                "lvl7pPr" => instance.lvl7_paragraph_props = Some(TextParagraphProperties::from_xml_element(child_node)),
                "lvl8pPr" => instance.lvl8_paragraph_props = Some(TextParagraphProperties::from_xml_element(child_node)),
                "lvl9pPr" => instance.lvl9_paragraph_props = Some(TextParagraphProperties::from_xml_element(child_node)),
                _ => (),
            }
        }

        instance
    }
}

pub struct TextBody {
    pub body_properties: TextBodyProperties,
    pub list_style: Option<TextListStyle>,
    pub paragraph_array: Vec<TextParagraph>,
}

pub struct TextBodyProperties {
    pub preset_text_warp: Option<PresetTextShape>,
    pub auto_fit_type: Option<TextAutoFit>,
    //pub scene_3d: Option<Scene3D>,
    //pub text_3d: Option<Text3D>,
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
}

pub enum TextAutoFit {
    NoAutoFit,
    NormalAutoFit(TextNormalAutoFit),
    ShapeAutoFit,
}

pub struct TextNormalAutoFit {
    pub font_scale: Option<TextFontScalePercent>, // 100000
    pub line_spacing_reduction: Option<TextSpacingPercent>, // 0
}

pub struct PresetTextShape {
    pub adjust_value_list: Vec<GeomGuide>,
    pub preset: TextShapeType,
}

pub struct FontScheme {
    pub major_font: FontCollection,
    pub minor_font: FontCollection,
    pub name: String,
}

impl FontScheme {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self, XmlError> {
        let mut opt_name = None;
        let mut opt_major_font = None;
        let mut opt_minor_font = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "name" => opt_name = Some(value.clone()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "majorFont" => opt_major_font = Some(FontCollection::from_xml_element(child_node)?),
                "minorFont" => opt_minor_font = Some(FontCollection::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let name = opt_name.ok_or(XmlError::from(MissingAttributeError::new("name")))?;
        let major_font = opt_major_font.ok_or(XmlError::from(MissingChildNodeError::new("majorFont")))?;
        let minor_font = opt_minor_font.ok_or(XmlError::from(MissingChildNodeError::new("minorFont")))?;

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
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self, XmlError> {
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

        let latin = opt_latin.ok_or(XmlError::from(MissingChildNodeError::new("latin")))?;
        let east_asian = opt_ea.ok_or(XmlError::from(MissingChildNodeError::new("ea")))?;
        let complex_script = opt_cs.ok_or(XmlError::from(MissingChildNodeError::new("cs")))?;

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
    pub hyperlink_click: Option<Hyperlink>,
    pub hyperlink_hover: Option<Hyperlink>,
}

pub struct Locking {
    pub no_grouping: Option<bool>, // false
    pub no_select: Option<bool>, // false
    pub no_rotate: Option<bool>, // false
    pub no_change_aspect_ratio: Option<bool>, // false
    pub no_move: Option<bool>, // false
    pub no_resize: Option<bool>, // false
    pub no_edit_points: Option<bool>, // false
    pub no_adjust_handles: Option<bool>, // false
    pub no_change_arrowheads: Option<bool>, // false
    pub no_change_shape_type: Option<bool>, // false
}

pub struct ShapeLocking {
    pub locking: Locking,
    pub no_text_edit: Option<bool>, // false
}

pub struct GroupLocking {
    pub no_grouping: Option<bool>, // false
    pub no_ungrouping: Option<bool>, // false
    pub no_select: Option<bool>, // false
    pub no_rotate: Option<bool>, // false
    pub no_change_aspect_ratio: Option<bool>, // false
    pub no_move: Option<bool>, // false
    pub no_resize: Option<bool>, // false
}

pub struct GraphicalObjectFrameLocking {
    pub no_grouping: Option<bool>, // false
    pub no_drilldown: Option<bool>, // false
    pub no_select: Option<bool>, // false
    pub no_change_aspect: Option<bool>, // false
    pub no_move: Option<bool>, // false
    pub no_resize: Option<bool>, // false
}

pub struct ConnectorLocking {
    pub locking: Locking,
}

pub struct PictureLocking {
    pub locking: Locking,
    pub no_crop: Option<bool>, // false
}

pub struct NonVisualDrawingShapeProps {
    pub shape_locks: Option<ShapeLocking>,
    pub is_text_box: Option<bool>, // false
}

pub struct NonVisualGroupDrawingShapeProps {
    pub locks: Option<GroupLocking>,
}

pub struct NonVisualGraphicFrameProperties {
    pub graphic_frame_locks: Option<GraphicalObjectFrameLocking>,
}

pub struct NonVisualConnectorProperties {
    pub connector_locks: Option<ConnectorLocking>,
    pub start_connection: Option<Connection>,
    pub end_connection: Option<Connection>,
}

pub struct NonVisualPictureProperties {
    pub picture_locks: Option<PictureLocking>,
    pub prefer_relative_resize: Option<bool>, // true
}

pub struct Connection {
    pub id: DrawingElementId,
    pub shape_index: u32,
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
    pub offset: Option<Point2D>,
    pub extents:  Option<PositiveSize2D>,
    pub rotate_angle: Option<Angle>, // 0
    pub flip_horizontal: Option<bool>, // false
    pub flip_vertical: Option<bool>, // false
}

pub struct GroupTransform2D {
    pub offset: Option<Point2D>,
    pub extents:  Option<PositiveSize2D>,
    pub child_offset: Option<Point2D>,
    pub child_extents: Option<PositiveSize2D>,
    pub rotate_angle: Option<Angle>, // 0
    pub flip_horizontal: Option<bool>, // false
    pub flip_vertical: Option<bool>, // false
}

pub struct GroupShapeProperties {
    pub transform: Option<GroupTransform2D>,
    pub fill_properties: Option<FillProperties>,
    pub effect_properties: Option<EffectProperties>,
    //pub scene_3d: Option<Scene3D>,
    pub black_and_white_mode: Option<BlackWhiteMode>,
}

pub enum Geometry {
    Custom(CustomGeometry2D),
    Preset(PresetGeometry2D),
}

pub struct GeomGuide {
    pub name: GeomGuideName,
    pub formula: GeomGuideFormula,
}

pub enum AdjustHandle {
    XY(XYAdjustHandle),
    Polar(PolarAdjustHandle),
}

pub enum AdjCoordinate {
    Coordinate(Coordinate),
    GeomGuideName(GeomGuideName),
}

pub enum AdjAngle {
    Angle(Angle),
    GeomGuideName(GeomGuideName),
}

pub struct AdjPoint2D {
    pub x: AdjCoordinate,
    pub y: AdjCoordinate,
}

pub struct GeomRect {
    pub left: AdjCoordinate,
    pub top: AdjCoordinate,
    pub right: AdjCoordinate,
    pub bottom: AdjCoordinate,
}

pub struct XYAdjustHandle {
    pub position: AdjPoint2D,
    pub guide_reference_x: Option<GeomGuideName>,
    pub guide_reference_y: Option<GeomGuideName>,
    pub min_x: Option<AdjCoordinate>,
    pub max_x: Option<AdjCoordinate>,
    pub min_y: Option<AdjCoordinate>,
    pub max_y: Option<AdjCoordinate>,
}

pub struct PolarAdjustHandle {
    pub position: AdjPoint2D,
    pub guide_reference_radial: Option<GeomGuideName>,
    pub guide_reference_angle: Option<GeomGuideName>,
    pub min_radial: Option<AdjCoordinate>,
    pub max_radial: Option<AdjCoordinate>,
    pub min_angle: Option<AdjAngle>,
    pub max_angle: Option<AdjAngle>,
}

pub struct ConnectionSite {
    pub position: AdjPoint2D,
    pub angle: AdjAngle,
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

pub struct Path2D {
    pub commands: Vec<Path2DCommand>,
    pub width: Option<PositiveCoordinate>, // 0
    pub height: Option<PositiveCoordinate>, // 0
    pub fill_mode: Option<PathFillMode>, // norm
    pub stroke: Option<bool>, // true
    pub extrusion_ok: Option<bool>, // true
}

pub struct CustomGeometry2D {
    pub adjust_value_list: Vec<GeomGuide>,
    pub guide_list: Vec<GeomGuide>,
    pub adjust_handle_list: Vec<AdjustHandle>,
    pub connection_site_list: Vec<ConnectionSite>,
    pub rect: Option<GeomRect>,
    pub path_list: Vec<Path2D>,
}

pub struct PresetGeometry2D {
    pub adjust_value_list: Vec<GeomGuide>,
    pub preset: ShapeType,
}

pub struct ShapeProperties {
    pub transform: Option<Transform2D>,
    pub geometry: Option<Geometry>,
    pub fill_properties: Option<FillProperties>,
    pub line_properties: Option<LineProperties>,
    pub effect_properties: Option<EffectProperties>,
    //pub scene_3d: Option<Scene3D>,
    //pub shape_3d: Option<Shape3D>,
    pub black_and_white_mode: Option<BlackWhiteMode>,
}

pub struct ShapeStyle {
    pub line_reference: StyleMatrixReference,
    pub fill_reference: StyleMatrixReference,
    pub effect_reference: StyleMatrixReference,
    pub font_reference: FontReference,
}

pub struct FontReference {
    pub color: Option<Color>,
    pub index: FontCollectionIndex,
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
    pub id: Option<Guid>, // {00000000-0000-0000-0000-000000000000}
    pub build_step: Option<DgmBuildStep>, // sp
}

pub struct AnimationChartElement {
    pub series_index: Option<i32>, // -1
    pub category_index: Option<i32>, // -1
    pub build_step: ChartBuildStep,
}

pub enum AnimationGraphicalObjectBuildProperties {
    BuildDiagram(AnimationDgmBuildProperties),
    BuildChart(AnimationChartBuildProperties),
}


pub struct AnimationDgmBuildProperties {
    pub build_type: Option<AnimationDgmBuildType>, // allAtOnce
    pub reverse: Option<bool>, // false
}

pub struct AnimationChartBuildProperties {
    pub build_type: Option<AnimationChartBuildType>, // allAtOnce
    pub animate_bg: Option<bool>, // true
}

pub struct OfficeStyleSheet {
    pub name: Option<String>, // ""
    pub theme_elements: BaseStyles,
    pub object_defaults: Option<ObjectStyleDefaults>,
    pub extra_color_scheme_list: Vec<ColorSchemeAndMapping>,
    pub custom_color_list: Vec<CustomColor>,
}

impl OfficeStyleSheet {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self, XmlError> {
        let mut name = None;
        let mut opt_theme_elements = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "name" => name = Some(value.clone()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "themeElements" => opt_theme_elements = Some(BaseStyles::from_xml_element(child_node)?),
                // TODO: parse optional elements
                _ => (),
            }
        }

        let theme_elements = opt_theme_elements.ok_or(XmlError::from(MissingChildNodeError::new("themeElements")))?;

        Ok(Self {
            name,
            theme_elements,
            object_defaults: None,
            extra_color_scheme_list: Vec::new(),
            custom_color_list: Vec::new(),
        })
    }
}

pub struct BaseStyles {
    pub color_scheme: ColorScheme,
    pub font_scheme: FontScheme,
    pub format_scheme: StyleMatrix,
}

impl BaseStyles {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self, XmlError> {
        let mut opt_color_scheme = None;
        let mut opt_font_scheme = None;
        let mut opt_format_scheme = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "clrScheme" => opt_color_scheme = Some(ColorScheme::from_xml_element(child_node)?),
                "fontScheme" => opt_font_scheme = Some(FontScheme::from_xml_element(child_node)?),
                "fmtScheme" => opt_format_scheme = Some(StyleMatrix::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let color_scheme = opt_color_scheme.ok_or(XmlError::from(MissingChildNodeError::new("clrScheme")))?;
        let font_scheme = opt_font_scheme.ok_or(XmlError::from(MissingChildNodeError::new("fontScheme")))?;
        let format_scheme = opt_format_scheme.ok_or(XmlError::from(MissingChildNodeError::new("fmtScheme")))?;

        Ok(Self {
            color_scheme,
            font_scheme,
            format_scheme,
        })
    }
}

pub struct StyleMatrix {
    pub name: Option<String>, // ""
    pub fill_style_list: Vec<FillProperties>, // minOccurs: 3
    pub line_style_list: Vec<LineProperties>, // minOccurs: 3
    pub effect_style_list: Vec<EffectStyleItem>, // minOccurs: 3
    pub bg_fill_style_list: Vec<FillProperties>, // minOccurs: 3
}

impl StyleMatrix {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self, XmlError> {
        let mut name = None;
        let mut fill_style_list = Vec::new();
        let mut line_style_list = Vec::new();
        let mut effect_style_list = Vec::new();
        let mut bg_fill_style_list = Vec::new();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "name" => name = Some(value.clone()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "fillStyleLst" => {
                    for fill_style_node in &child_node.child_nodes {
                        fill_style_list.push(FillProperties::from_xml_element(fill_style_node)?);
                    }
                }
                "lnStyleLst" => {
                    for line_style_node in &child_node.child_nodes {
                        line_style_list.push(LineProperties::from_xml_element(line_style_node));
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
            let error = LimitViolationError::new(
                "fillStyleLst",
                Limit::Value(3),
                Limit::Unbounded,
                fill_style_list.len() as u32,
            );
            return Err(XmlError::from(error));
        }

        if line_style_list.len() < 3 {
            let error = LimitViolationError::new(
                "lnStyleLst",
                Limit::Value(3),
                Limit::Unbounded,
                line_style_list.len() as u32,
            );
            return Err(XmlError::from(error));
        }

        if bg_fill_style_list.len() < 3 {
            let error = LimitViolationError::new(
                "bgFillStyleLst",
                Limit::Value(3),
                Limit::Unbounded,
                bg_fill_style_list.len() as u32,
            );
            return Err(XmlError::from(error));
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
    pub shape_definition: Option<DefaultShapeDefinition>,
    pub line_definition: Option<DefaultShapeDefinition>,
    pub text_definition: Option<DefaultShapeDefinition>,
}

pub struct DefaultShapeDefinition {
    pub shape_properties: ShapeProperties,
    pub text_body_properties: TextBodyProperties,
    pub text_list_style: TextListStyle,
    pub shape_style: Option<ShapeStyle>,
}