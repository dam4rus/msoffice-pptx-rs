// TODO: This module defines shared types between different OOX file formats. It should be refactored into a different crate, if these types are needed.

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

pub type Color = i32;

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
    }
}

/*
<xsd:simpleType name="ST_PresetColorVal">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="dkBlue"/>
      <xsd:enumeration value="dkCyan"/>
      <xsd:enumeration value="dkGoldenrod"/>
      <xsd:enumeration value="dkGray"/>
      <xsd:enumeration value="dkGrey"/>
      <xsd:enumeration value="dkGreen"/>
      <xsd:enumeration value="dkKhaki"/>
      <xsd:enumeration value="dkMagenta"/>
      <xsd:enumeration value="dkOliveGreen"/>
      <xsd:enumeration value="dkOrange"/>
      <xsd:enumeration value="dkOrchid"/>
      <xsd:enumeration value="dkRed"/>
      <xsd:enumeration value="dkSalmon"/>
      <xsd:enumeration value="dkSeaGreen"/>
      <xsd:enumeration value="dkSlateBlue"/>
      <xsd:enumeration value="dkSlateGray"/>
      <xsd:enumeration value="dkSlateGrey"/>
      <xsd:enumeration value="dkTurquoise"/>
      <xsd:enumeration value="dkViolet"/>
      <xsd:enumeration value="deepPink"/>
      <xsd:enumeration value="deepSkyBlue"/>
      <xsd:enumeration value="dimGray"/>
      <xsd:enumeration value="dimGrey"/>
      <xsd:enumeration value="dodgerBlue"/>
      <xsd:enumeration value="firebrick"/>
      <xsd:enumeration value="floralWhite"/>
      <xsd:enumeration value="forestGreen"/>
      <xsd:enumeration value="fuchsia"/>
      <xsd:enumeration value="gainsboro"/>
      <xsd:enumeration value="ghostWhite"/>
      <xsd:enumeration value="gold"/>
      <xsd:enumeration value="goldenrod"/>
      <xsd:enumeration value="gray"/>
      <xsd:enumeration value="grey"/>
      <xsd:enumeration value="green"/>
      <xsd:enumeration value="greenYellow"/>
      <xsd:enumeration value="honeydew"/>
      <xsd:enumeration value="hotPink"/>
      <xsd:enumeration value="indianRed"/>
      <xsd:enumeration value="indigo"/>
      <xsd:enumeration value="ivory"/>
      <xsd:enumeration value="khaki"/>
      <xsd:enumeration value="lavender"/>
      <xsd:enumeration value="lavenderBlush"/>
      <xsd:enumeration value="lawnGreen"/>
      <xsd:enumeration value="lemonChiffon"/>
      <xsd:enumeration value="lightBlue"/>
      <xsd:enumeration value="lightCoral"/>
      <xsd:enumeration value="lightCyan"/>
      <xsd:enumeration value="lightGoldenrodYellow"/>
      <xsd:enumeration value="lightGray"/>
      <xsd:enumeration value="lightGrey"/>
      <xsd:enumeration value="lightGreen"/>
      <xsd:enumeration value="lightPink"/>
      <xsd:enumeration value="lightSalmon"/>
      <xsd:enumeration value="lightSeaGreen"/>
      <xsd:enumeration value="lightSkyBlue"/>
      <xsd:enumeration value="lightSlateGray"/>
      <xsd:enumeration value="lightSlateGrey"/>
      <xsd:enumeration value="lightSteelBlue"/>
      <xsd:enumeration value="lightYellow"/>
      <xsd:enumeration value="ltBlue"/>
      <xsd:enumeration value="ltCoral"/>
      <xsd:enumeration value="ltCyan"/>
      <xsd:enumeration value="ltGoldenrodYellow"/>
      <xsd:enumeration value="ltGray"/>
      <xsd:enumeration value="ltGrey"/>
      <xsd:enumeration value="ltGreen"/>
      <xsd:enumeration value="ltPink"/>
      <xsd:enumeration value="ltSalmon"/>
      <xsd:enumeration value="ltSeaGreen"/>
      <xsd:enumeration value="ltSkyBlue"/>
      <xsd:enumeration value="ltSlateGray"/>
      <xsd:enumeration value="ltSlateGrey"/>
      <xsd:enumeration value="ltSteelBlue"/>
      <xsd:enumeration value="ltYellow"/>
      <xsd:enumeration value="lime"/>
      <xsd:enumeration value="limeGreen"/>
      <xsd:enumeration value="linen"/>
      <xsd:enumeration value="magenta"/>
      <xsd:enumeration value="maroon"/>
      <xsd:enumeration value="medAquamarine"/>
      <xsd:enumeration value="medBlue"/>
      <xsd:enumeration value="medOrchid"/>
      <xsd:enumeration value="medPurple"/>
      <xsd:enumeration value="medSeaGreen"/>
      <xsd:enumeration value="medSlateBlue"/>
      <xsd:enumeration value="medSpringGreen"/>
      <xsd:enumeration value="medTurquoise"/>
      <xsd:enumeration value="medVioletRed"/>
      <xsd:enumeration value="mediumAquamarine"/>
      <xsd:enumeration value="mediumBlue"/>
      <xsd:enumeration value="mediumOrchid"/>
      <xsd:enumeration value="mediumPurple"/>
      <xsd:enumeration value="mediumSeaGreen"/>
      <xsd:enumeration value="mediumSlateBlue"/>
      <xsd:enumeration value="mediumSpringGreen"/>
      <xsd:enumeration value="mediumTurquoise"/>
      <xsd:enumeration value="mediumVioletRed"/>
      <xsd:enumeration value="midnightBlue"/>
      <xsd:enumeration value="mintCream"/>
      <xsd:enumeration value="mistyRose"/>
      <xsd:enumeration value="moccasin"/>
      <xsd:enumeration value="navajoWhite"/>
      <xsd:enumeration value="navy"/>
      <xsd:enumeration value="oldLace"/>
      <xsd:enumeration value="olive"/>
      <xsd:enumeration value="oliveDrab"/>
      <xsd:enumeration value="orange"/>
      <xsd:enumeration value="orangeRed"/>
      <xsd:enumeration value="orchid"/>
      <xsd:enumeration value="paleGoldenrod"/>
      <xsd:enumeration value="paleGreen"/>
      <xsd:enumeration value="paleTurquoise"/>
      <xsd:enumeration value="paleVioletRed"/>
      <xsd:enumeration value="papayaWhip"/>
      <xsd:enumeration value="peachPuff"/>
      <xsd:enumeration value="peru"/>
      <xsd:enumeration value="pink"/>
      <xsd:enumeration value="plum"/>
      <xsd:enumeration value="powderBlue"/>
      <xsd:enumeration value="purple"/>
      <xsd:enumeration value="red"/>
      <xsd:enumeration value="rosyBrown"/>
      <xsd:enumeration value="royalBlue"/>
      <xsd:enumeration value="saddleBrown"/>
      <xsd:enumeration value="salmon"/>
      <xsd:enumeration value="sandyBrown"/>
      <xsd:enumeration value="seaGreen"/>
      <xsd:enumeration value="seaShell"/>
      <xsd:enumeration value="sienna"/>
      <xsd:enumeration value="silver"/>
      <xsd:enumeration value="skyBlue"/>
      <xsd:enumeration value="slateBlue"/>
      <xsd:enumeration value="slateGray"/>
      <xsd:enumeration value="slateGrey"/>
      <xsd:enumeration value="snow"/>
      <xsd:enumeration value="springGreen"/>
      <xsd:enumeration value="steelBlue"/>
      <xsd:enumeration value="tan"/>
      <xsd:enumeration value="teal"/>
      <xsd:enumeration value="thistle"/>
      <xsd:enumeration value="tomato"/>
      <xsd:enumeration value="turquoise"/>
      <xsd:enumeration value="violet"/>
      <xsd:enumeration value="wheat"/>
      <xsd:enumeration value="white"/>
      <xsd:enumeration value="whiteSmoke"/>
      <xsd:enumeration value="yellow"/>
      <xsd:enumeration value="yellowGreen"/>
    </xsd:restriction>
  </xsd:simpleType>
*/

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
    pub color_transforms: Option<Vec<ColorTransform>>
}


pub struct SRgbColor {
    pub value: HexColorRGB,
    pub color_transforms: Option<Vec<ColorTransform>>
}


pub struct HslColor {
    pub hue: PositiveFixedAngle,
    pub saturation: Percentage,
    pub luminance: Percentage,
    pub color_transforms: Option<Vec<ColorTransform>>
}


pub struct SystemColor {
    pub value: SystemColorVal,
    pub last_color: HexColorRGB,
    pub color_transforms: Option<Vec<ColorTransform>>
}


pub enum ColorGroup {
    ScRgbColor(ScRgbColor),
    SRgbColor(SRgbColor),
    HslColor(HslColor),
    SystemColor(SystemColor),
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
    BuClr,
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

}
/*
    class TextParagraphProperties
    {
    public:
        Office::Optional<TextMargin> marL;
        Office::Optional<TextMargin> marR;
        Office::Optional<TextIndentLevelType> lvl;
        Office::Optional<TextIndent> indent;
        Office::Optional<TextAlignType> algn;
        Office::Optional<Coordinate32> defTabSz;
        Office::Optional<bool> rtl;
        Office::Optional<bool> eaLnBrk;
        Office::Optional<TextFontAlignType> fontAlgn;
        Office::Optional<bool> latinLnBrk;
        Office::Optional<bool> hangingPunct;
        std::unique_ptr<TextSpacingChoice> lnSpc;
        std::unique_ptr<TextSpacingChoice> spcBef;
        std::unique_ptr<TextSpacingChoice> spcAft;
        std::unique_ptr<TextBulletColorChoice> textBulletColor;
        std::unique_ptr<TextBulletSizeChoice> textBulletSize;
        std::unique_ptr<TextBulletTypefaceChoice> textBulletTypeface;
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