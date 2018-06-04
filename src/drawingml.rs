// TODO: This module defines shared types between different OOX file formats. It should be refactored into a different crate, if these types are needed.
use relationship;

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
pub type PositiveFixedAngle = Angle; // TODO: 21600000
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
pub type Panose = String; // TODO: hex, length=10
pub type TextBulletStartAtNum = i32; // TODO: 1 <= n <= 32767
pub type Lang = String; 
pub type TextNonNegativePoint = i32; // TODO: 0 <= n <= 400000
pub type TextPoint = i32; // TODO: -400000 <= n <= 400000
pub type ShapeId = String;

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

decl_oox_enum! {
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

decl_oox_enum! {
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

decl_oox_enum! {
    pub enum SchemeColorVal {
        Bg1 = "bg1",
        Tx1 = "tx1",
        Bg2 = "bg2",
        Tx2 = "tx2",
        Accent1 = "accent1",
        Accent2 = "accent2",
        Accent3 = "accent3",
        Accent4 = "accent4",
        Accent5 = "accent5",
        Hlink = "hlink",
        FolHlink = "folHlink",
        PhClr = "phClr",
        Dk1 = "dk1",
        Lt1 = "lt1",
        Dk2 = "dk2",
        Lt2 = "lt2",
    }
}

decl_oox_enum! {
    pub enum TextAlignType {
        L = "l",
        Ctr = "ctr",
        R = "r",
        Just = "just",
        JustLow = "justLow",
        Dist = "dist",
        ThaiDist = "thaiDist",
    }
}

decl_oox_enum! {
    pub enum TextFontAlignType {
        Auto = "auto",
        T = "t",
        Ctr = "ctr",
        Base = "base",
        B = "b",
    }
}

decl_oox_enum! {
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

decl_oox_enum! {
    pub enum EffectContainerType {
        Sib = "sib",
        Tree = "tree",
    }
}

decl_oox_enum! {
    pub enum PathShadeType {
        Shape = "shape",
        Circle = "circle",
        Rect = "rect",
    }
}

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

pub struct ScRgbColor {
    pub r: Percentage,
    pub g: Percentage,
    pub b: Percentage,
    pub color_transforms: Vec<ColorTransform>,
}

pub struct SRgbColor {
    pub value: HexColorRGB,
    pub color_transforms: Vec<ColorTransform>,
}

pub struct HslColor {
    pub hue: PositiveFixedAngle,
    pub saturation: Percentage,
    pub luminance: Percentage,
    pub color_transforms: Vec<ColorTransform>,
}

pub struct SystemColor {
    pub value: SystemColorVal,
    pub last_color: HexColorRGB,
    pub color_transforms: Vec<ColorTransform>,
}

pub struct PresetColor {
    pub value: PresetColorVal,
    pub color_transforms: Vec<ColorTransform>,
}

pub struct SchemeColor {
    pub value: SchemeColorVal,
    pub color_transforms: Vec<ColorTransform>,
}

pub enum ColorGroup {
    ScRgbColor(ScRgbColor),
    SRgbColor(SRgbColor),
    HslColor(HslColor),
    SystemColor(SystemColor),
    SchemeColor(SchemeColor),
    PresetColor(PresetColor),
}

pub struct GradientStop {
    pub color: ColorGroup,
    pub position: PositiveFixedPercentage,
}

pub struct LinearShadeProperties {
    pub angle: Option<PositiveFixedAngle>,
    pub scaled: Option<bool>,
}

pub struct PathShadeProperties {
    pub fill_to_rect: Option<RelativeRect>,
    pub path: Option<PathShadeType>,
}

pub enum ShadePropertiesGroup {
    Linear(LinearShadeProperties),
    Path(PathShadeProperties),
}

pub struct GradientFillProperties {
    pub gradient_stop_list: Vec<GradientStop>, // length: 2 <= n <= inf
    pub shade_properties: Option<ShadePropertiesGroup>,
    pub tile_rect: Option<RelativeRect>,
    pub flip: Option<TileFlipMode>,
    pub rotate_with_shape: Option<bool>,
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

pub enum FillModePropertiesGroup {
    Tile(TileInfoProperties),
    Stretch(StretchInfoProperties),
}

pub struct BlipFillProperties {
    pub blip: Option<Blip>,
    pub source_rect: Option<RelativeRect>,
    pub fill_mode_properties: Option<FillModePropertiesGroup>,
    pub dpi: Option<u32>,
    pub rotate_with_shape: Option<bool>,
}

pub enum FillPropertiesGroup {
    NoFill,
    SolidFill(ColorGroup),
    GradientFill(GradientFillProperties),
    BlipFill(BlipFillProperties),
    PatternFill(PatternFillProperties),
    GroupFill(GroupFillProperties),
}

pub struct RelativeRect {
    pub left: Option<Percentage>,
    pub top: Option<Percentage>,
    pub right: Option<Percentage>,
    pub bottom: Option<Percentage>,
}

pub struct PositiveSize2D {
    pub width: PositiveCoordinate,
    pub height: PositiveCoordinate,
}

pub struct EffectContainer {
    pub effects: Vec<EffectGroup>,
    pub container_type: Option<EffectContainerType>,
    pub name: Option<String>,
}

pub struct AlphaBiLevelEffect {
    pub treshold: PositiveFixedPercentage,
}

pub struct AlphaInverseEffect {
    pub color: Option<ColorGroup>,
}

pub struct AlphaModulateEffect {
    pub container: EffectContainer,
}

pub struct AlphaModulateFixedEffect {
    pub amount: PositivePercentage,
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
    pub color_from: ColorGroup,
    pub color_to: ColorGroup,
    pub use_alpha: bool, // true
}

pub struct ColorReplaceEffect {
    pub color: ColorGroup,
}


pub struct LuminanceEffect {
    pub bright: Option<FixedPercentage>,
    pub contrast: Option<FixedPercentage>,
}

pub struct DuotoneEffect {
    pub colors: [ColorGroup; 2],
}

pub struct FillEffect {
    pub fill: FillPropertiesGroup,
}

pub enum EffectGroup {
    Cont(EffectContainer),
    EffectRefernce(String),
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
    Grayscl(GrayscaleEffect),
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

pub struct Blip {
    pub effect: Vec<EffectGroup>,
    pub embed_rel_id: Option<relationship::RelationshipId>,
    pub linked_rel_id: Option<relationship::RelationshipId>,
}

pub struct TextFont {
    pub typeface: TextTypeFace,
    pub panose: Option<Panose>,
    pub pitch_family: Option<i32>, // 0
    pub charset: Option<i32>, // 1
}

pub enum TextSpacingGroup {
    SpcPct(TextSpacingPercent),
    SpcPts(TextSpacingPoint),
}

pub enum TextBulletColorGroup {
    BuClrTx,
    BuClr(ColorGroup),
}

pub enum TextBulletSizeGroup {
    BuSzTx,
    BuSzPct(TextBulletSizePercent),
    BuSzPts(TextFontSize),
}

pub enum TextBulletTypefaceGroup {
    BuFontTx,
    BuFont(TextFont),
}

pub struct TextAutonumberedBullet {
    pub scheme: TextAutonumberScheme,
    pub start_t: Option<TextBulletStartAtNum>,
}

pub enum TextBulletGroup {
    BuNone,
    BuAutoNum(TextAutonumberedBullet),
    BuChar(String),
    BuBlip(Blip),
}

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
    pub line_spacing: Option<TextSpacingGroup>,
    pub space_before: Option<TextSpacingGroup>,
    pub space_after: Option<TextSpacingGroup>,
    pub text_bullet_color: Option<TextBulletColorGroup>,
    pub text_bullet_size: Option<TextBulletSizeGroup>,
    pub text_bullet_typeface: Option<TextBulletTypefaceGroup>,
    pub text_bullet: Option<TextBulletGroup>
}
/*
    class TextParagraphProperties
    {
    public:
        std::unique_ptr<TextBulletChoice> textBullet;
        TextTabStopList tabLst;
        std::unique_ptr<TextCharacterProperties> defRPr;
        //std::unique_ptr<OfficeArtExtensionList> extLst;

        static TextParagraphProperties* FromXmlNode(const MXmlNode2& xmlNode);

        TextParagraphProperties() = default;
        DISALLOW_COPY_AND_ASSIGN(TextParagraphProperties);

        void UpdateWith(const TextParagraphProperties &textParProps);
    };

*/

pub struct TextListStyle {
    pub def_paragraph_props: Option<TextParagraphProperties>,
}
/*
        std::unique_ptr<TextParagraphProperties> defPPr;
        std::unique_ptr<TextParagraphProperties> lvl1pPr;
        std::unique_ptr<TextParagraphProperties> lvl2pPr;
        std::unique_ptr<TextParagraphProperties> lvl3pPr;
        std::unique_ptr<TextParagraphProperties> lvl4pPr;
        std::unique_ptr<TextParagraphProperties> lvl5pPr;
        std::unique_ptr<TextParagraphProperties> lvl6pPr;
        std::unique_ptr<TextParagraphProperties> lvl7pPr;
        std::unique_ptr<TextParagraphProperties> lvl8pPr;
        std::unique_ptr<TextParagraphProperties> lvl9pPr;
        //std::unique_ptr<OfficeArtExtensionList> extLst;


*/