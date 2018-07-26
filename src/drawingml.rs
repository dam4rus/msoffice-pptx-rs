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
        Rnd = "rnd",
        Sq = "sq",
        Flat = "flat",
    }
}

decl_simple_type_enum! {
    pub enum CompoundLine {
        Sng = "sng",
        Dbl = "dbl",
        ThickThin = "thickThin",
        ThinThick = "thinThick",
        Tri = "tri",
    }
}

decl_simple_type_enum! {
    pub enum PenAlignment {
        Ctr = "ctr",
        In = "in",
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
        Sm = "sm",
        Med = "med",
        Lg = "lg",
    }
}

decl_simple_type_enum! {
    pub enum LineEndLength {
        Sm = "sm",
        Med = "med",
        Lg = "lg",
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
        L = "l",
        Ctr = "ctr",
        R = "r",
        Just = "just",
        JustLow = "justLow",
        Dist = "dist",
        ThaiDist = "thaiDist",
    }
}

decl_simple_type_enum! {
    pub enum TextFontAlignType {
        Auto = "auto",
        T = "t",
        Ctr = "ctr",
        Base = "base",
        B = "b",
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

pub enum Color {
    ScRgbColor(ScRgbColor),
    SRgbColor(SRgbColor),
    HslColor(HslColor),
    SystemColor(SystemColor),
    SchemeColor(SchemeColor),
    PresetColor(PresetColor),
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
    pub name: String,
}

pub enum ColorMappingOverride {
    UseMasterColorMapping,
    OverrideColorMapping(ColorMapping),
}

pub struct ColorSchemeAndMapping {
    pub color_scheme: ColorScheme,
    pub color_mapping: Option<ColorMapping>,
}

pub struct GradientStop {
    pub color: Color,
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

pub enum ShadeProperties {
    Linear(LinearShadeProperties),
    Path(PathShadeProperties),
}

pub struct GradientFillProperties {
    pub gradient_stop_list: Vec<GradientStop>, // length: 2 <= n <= inf
    pub shade_properties: Option<ShadeProperties>,
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

pub enum LineFillProperties {
    NoFill,
    SolidFill(Color),
    GradientFill(GradientFillProperties),
    PatternFill(PatternFillProperties),
}

pub struct DashStop {
    pub dash_length: PositivePercentage,
    pub space_length: PositivePercentage,
}

pub enum LineDashProperties {
    PresetDash(PresetLineDashVal),
    CustomDash(Vec<DashStop>)
}

pub struct LineJoinMiterProperties {
    pub limit: Option<PositivePercentage>,
}

pub enum LineJoinProperties {
    Round,
    Bevel,
    Miter(LineJoinMiterProperties),
}

pub struct LineEndProperties {
    pub end_type: Option<LineEndType>,
    pub width: Option<LineEndWidth>,
    pub length: Option<LineEndLength>,
}

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

pub struct RelativeRect {
    pub left: Option<Percentage>,
    pub top: Option<Percentage>,
    pub right: Option<Percentage>,
    pub bottom: Option<Percentage>,
}

pub struct Point2D {
    pub x: Coordinate,
    pub y: Coordinate,
}

pub struct PositiveSize2D {
    pub width: PositiveCoordinate,
    pub height: PositiveCoordinate,
}

pub struct StyleMatrixReference {
    pub index: StyleMatrixColumnIndex,
    pub color: Option<Color>,
}

pub struct EffectContainer {
    pub effects: Vec<Effect>,
    pub container_type: Option<EffectContainerType>,
    pub name: Option<String>,
}

pub struct AlphaBiLevelEffect {
    pub treshold: PositiveFixedPercentage,
}

pub struct AlphaInverseEffect {
    pub color: Option<Color>,
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

pub struct Blip {
    pub effect: Vec<Effect>,
    pub embed_rel_id: Option<relationship::RelationshipId>,
    pub linked_rel_id: Option<relationship::RelationshipId>,
}

pub struct TextFont {
    pub typeface: TextTypeFace,
    pub panose: Option<Panose>,
    pub pitch_family: Option<i32>, // 0
    pub charset: Option<i32>, // 1
}

pub struct SupplementalFont {
    pub script: String,
    pub typeface: TextTypeFace,
}

pub enum TextSpacing {
    SpcPct(TextSpacingPercent),
    SpcPts(TextSpacingPoint),
}

pub enum TextBulletColor {
    FollowText,
    Color(Color),
}

pub enum TextBulletSize {
    FollowText,
    Percent(TextBulletSizePercent),
    Point(TextFontSize),
}

pub enum TextBulletTypeface {
    FollowText,
    Font(TextFont),
}

pub struct TextAutonumberedBullet {
    pub scheme: TextAutonumberScheme,
    pub start_t: Option<TextBulletStartAtNum>,
}

pub enum TextBullet {
    None,
    AutoNumbered(TextAutonumberedBullet),
    Character(String),
    Picture(Blip),
}

pub struct TextTabStop {
    pub position: Option<Coordinate32>,
    pub alignment: Option<TextTabAlignType>,
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
    pub relationship_id: Option<relationship::RelationshipId>,
    pub invalid_url: Option<String>,
    pub action: Option<String>,
    pub target_frame: Option<String>,
    pub tooltip: Option<String>,
    pub history: Option<bool>, // true
    pub highlight_click: Option<bool>, // false
    pub end_sound: Option<bool>, // false
    pub sound: Option<EmbeddedWAVAudioFile>,
		//std::unique_ptr<OfficeArtExtensionList> extLst;
    
}

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

pub struct TextListStyle {
    pub def_paragraph_props: Option<TextParagraphProperties>,
    pub paragraph_level_props_array: [Option<TextParagraphProperties>; 9],
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

pub struct FontCollection {
    pub latin: TextFont,
    pub east_asian: TextFont,
    pub complex_script: TextFont,
    pub supplemental_font_list: Vec<SupplementalFont>,
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
    pub embed_rel_id: relationship::RelationshipId,
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
    pub link: relationship::RelationshipId,
    pub content_type: Option<String>,
}

pub struct VideoFile {
    pub link: relationship::RelationshipId,
    pub content_type: Option<String>,
}

pub struct QuickTimeFile {
    pub link: relationship::RelationshipId,
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
    pub theme_elements: BaseStyles,
    pub object_defaults: Option<ObjectStyleDefaults>,
    pub extra_color_scheme_list: Vec<ColorSchemeAndMapping>,
    pub custom_color_list: Vec<CustomColor>,
    pub name: Option<String>, // ""
}

pub struct BaseStyles {
    pub color_scheme: ColorScheme,
    pub font_scheme: FontScheme,
    pub format_scheme: StyleMatrix,
}

pub struct StyleMatrix {
    pub fill_style_list: Vec<FillProperties>, // minOccurs: 3
    pub line_style_list: Vec<LineProperties>, // minOccurs: 3
    pub effect_style_list: Vec<EffectStyleItem>, // minOccurs: 3
    pub bg_fill_style_list: Vec<FillProperties>, // minOccurs: 3
    pub name: Option<String>, // ""
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