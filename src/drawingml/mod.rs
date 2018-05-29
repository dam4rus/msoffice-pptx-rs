type Percentage = f32;
type PositivePercentage = f32; // TODO: 0 <= n < inf
type PositiveFixedPercentage = f32; // TODO: 0 <= n <= 100000
type FixedPercentage = f32; // TODO: -100000 <= n <= 100000

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

macro_rules! str_to_enum {
    ($name:ident {
		$($variant:ident = $str_value:expr),*,
	}) => {
        impl $name {
            pub fn from_string(s: &String) -> $name {
                match s.as_str() {
                    $($str_value => $name::$variant),*,
                    _ => panic!("Cannot convert string to enum type $name"),
                }
            }
        }
	};
}

pub enum TileFlipMode {
	None,
	X,
	Y,
	XY,
}

str_to_enum!(TileFlipMode {
	None = "none",
	X = "x",
	Y = "y",
	XY = "xy",
});

/*
if (string == mT("none"))
			return TileFlipMode::None;
		else if (string == mT("x"))
			return TileFlipMode::X;
		else if (string == mT("y"))
			return TileFlipMode::Y;
		else if (string == mT("xy"))
			return TileFlipMode::XY;

		throw Office::EnumParseError("TileFlipMode", StringToAnsiString(string));
*/



/*
enum class TileFlipMode
	{
		None, // none
		X, // x
		Y, // y
		XY // xy
	};

	TileFlipMode TileFlipModeFromString(const MString &string);


	enum class RectAlignment
	{
		L,
		T,
		R,
		B,
		Tl,
		Tr,
		Bl,
		Br,
		Ctr,
	};

	RectAlignment RectAlignmentFromString(const MString &string);


	enum class BlackWhiteMode
	{
		Clr,
		Auto,
		Gray,
		LtGray,
		InvGray,
		GrayWhite,
		BlackGray,
		BlackWhite,
		Black,
		White,
		Hidden,
	};

	BlackWhiteMode BlackWhiteModeFromString(const MString &string);


	enum class PathFillMode
	{
		None,
		Norm,
		Lighten,
		LightenLess,
		Darken,
		DarkenLess,
	};

	PathFillMode PathFillModeFromString(const MString &string);


	enum class ShapeType
	{
		Line,
		LineInv,
		Triangle,
		RtTriangle,
		Rect,
		Diamond,
		Parallelogram,
		Trapezoid,
		NonIsoscelesTrapezoid,
		Pentagon,
		Hexagon,
		Heptagon,
		Octagon,
		Decagon,
		Dodecagon,
		Star4,
		Star5,
		Star6,
		Star7,
		Star8,
		Star10,
		Star12,
		Star16,
		Star24,
		Star32,
		RoundRect,
		Round1Rect,
		Round2SameRect,
		Round2DiagRect,
		SnipRoundRect,
		Snip1Rect,
		Snip2SameRect,
		Snip2DiagRect,
		Plaque,
		Ellipse,
		Teardrop,
		HomePlate,
		Chevron,
		PieWedge,
		Pie,
		BlockArc,
		Donut,
		NoSmoking,
		RightArrow,
		LeftArrow,
		UpArrow,
		DownArrow,
		StripedRightArrow,
		NotchedRightArrow,
		BentUpArrow,
		LeftRightArrow,
		UpDownArrow,
		LeftUpArrow,
		LeftRightUpArrow,
		QuadArrow,
		LeftArrowCallout,
		RightArrowCallout,
		UpArrowCallout,
		DownArrowCallout,
		LeftRightArrowCallout,
		UpDownArrowCallout,
		QuadArrowCallout,
		BentArrow,
		UturnArrow,
		CircularArrow,
		LeftCircularArrow,
		LeftRightCircularArrow,
		CurvedRightArrow,
		CurvedLeftArrow,
		CurvedUpArrow,
		CurvedDownArrow,
		SwooshArrow,
		Cube,
		Can,
		LightningBolt,
		Heart,
		Sun,
		Moon,
		SmileyFace,
		IrregularSeal1,
		IrregularSeal2,
		FoldedCorner,
		Bevel,
		Frame,
		HalfFrame,
		Corner,
		DiagStripe,
		Chord,
		Arc,
		LeftBracket,
		RightBracket,
		LeftBrace,
		RightBrace,
		BracketPair,
		BracePair,
		StraightConnector1,
		BentConnector2,
		BentConnector3,
		BentConnector4,
		BentConnector5,
		CurvedConnector2,
		CurvedConnector3,
		CurvedConnector4,
		CurvedConnector5,
		Callout1,
		Callout2,
		Callout3,
		AccentCallout1,
		AccentCallout2,
		AccentCallout3,
		BorderCallout1,
		BorderCallout2,
		BorderCallout3,
		AccentBorderCallout1,
		AccentBorderCallout2,
		AccentBorderCallout3,
		WedgeRectCallout,
		WedgeRoundRectCallout,
		WedgeEllipseCallout,
		CloudCallout,
		Cloud,
		Ribbon,
		Ribbon2,
		EllipseRibbon,
		EllipseRibbon2,
		LeftRightRibbon,
		VerticalScroll,
		HorizontalScroll,
		Wave,
		DoubleWave,
		Plus,
		FlowChartProcess,
		FlowChartDecision,
		FlowChartInputOutput,
		FlowChartPredefinedProcess,
		FlowChartInternalStorage,
		FlowChartDocument,
		FlowChartMultidocument,
		FlowChartTerminator,
		FlowChartPreparation,
		FlowChartManualInput,
		FlowChartManualOperation,
		FlowChartConnector,
		FlowChartPunchedCard,
		FlowChartPunchedTape,
		FlowChartSummingJunction,
		FlowChartOr,
		FlowChartCollate,
		FlowChartSort,
		FlowChartExtract,
		FlowChartMerge,
		FlowChartOfflineStorage,
		FlowChartOnlineStorage,
		FlowChartMagneticTape,
		FlowChartMagneticDisk,
		FlowChartMagneticDrum,
		FlowChartDisplay,
		FlowChartDelay,
		FlowChartAlternateProcess,
		FlowChartOffpageConnector,
		ActionButtonBlank,
		ActionButtonHome,
		ActionButtonHelp,
		ActionButtonInformation,
		ActionButtonForwardNext,
		ActionButtonBackPrevious,
		ActionButtonEnd,
		ActionButtonBeginning,
		ActionButtonReturn,
		ActionButtonDocument,
		ActionButtonSound,
		ActionButtonMovie,
		Gear6,
		Gear9,
		Funnel,
		MathPlus,
		MathMinus,
		MathMultiply,
		MathDivide,
		MathEqual,
		MathNotEqual,
		CornerTabs,
		SquareTabs,
		PlaqueTabs,
		ChartX,
		ChartStar,
		ChartPlus,
	};

	ShapeType ShapeTypeFromString(const MString &string);


	enum class LineCap
	{
		Rnd,
		Sq,
		Flat,
	};

	LineCap LineCapFromString(const MString &string);


	enum class CompoundLine
	{
		Sng,
		Dbl,
		ThickThin,
		ThinThick,
		Tri,
	};

	CompoundLine CompoundLineFromString(const MString &string);


	enum class PenAlignment
	{
		Ctr,
		In,
	};

	PenAlignment PenAlignmentFromString(const MString &string);


	enum class PresetLineDashVal
	{
		Solid,
		Dot,
		Dash,
		LgDash,
		DashDot,
		LgDashDot,
		LgDashDotDot,
		SysDash,
		SysDot,
		SysDashDot,
		SysDashDotDot,
	};

	PresetLineDashVal PresetLineDashValFromString(const MString &string);


	enum class LineEndType
	{
		None,
		Triangle,
		Stealth,
		Diamond,
		Oval,
		Arrow,
	};

	LineEndType LineEndTypeFromString(const MString &string);


	enum class LineEndWidth
	{
		Sm,
		Med,
		Lg,
	};

	LineEndWidth LineEndWidthFromString(const MString &string);


	enum class LineEndLength
	{
		Sm,
		Med,
		Lg,
	};

	LineEndLength LineEndLengthFromString(const MString &string);


	enum class BlendMode
	{
		Over,
		Mult,
		Screen,
		Darken,
		Lighten,
	};

	BlendMode BlendModeFromString(const MString &string);


	enum class PresetShadowVal
	{
		Shdw1,
		Shdw2,
		Shdw3,
		Shdw4,
		Shdw5,
		Shdw6,
		Shdw7,
		Shdw8,
		Shdw9,
		Shdw10,
		Shdw11,
		Shdw12,
		Shdw13,
		Shdw14,
		Shdw15,
		Shdw16,
		Shdw17,
		Shdw18,
		Shdw19,
		Shdw20,
	};

	PresetShadowVal PresetShadowValFromString(const MString &string);


	enum class EffectContainerType
	{
		Sib,
		Tree,
	};

	EffectContainerType EffectContainerTypeFromString(const MString &string);


	enum class FontCollectionIndex
	{
		Major,
		Minor,
		None,
	};

	FontCollectionIndex FontCollectionIndexFromString(const MString &string);


	enum class AnimationBuildType
	{
		AllAtOnce,
	};

	AnimationBuildType AnimationBuildTypeFromString(const MString &string);


	enum class AnimationDgmOnlyBuildType
	{
		One,
		LvlOne,
		LvlAtOnce,
	};

	AnimationDgmOnlyBuildType AnimationDgmOnlyBuildTypeFromString(const MString &string);


	enum class AnimationChartOnlyBuildType
	{
		Series,
		Category,
		SeriesEl,
		CategoryEl,
	};

	AnimationChartOnlyBuildType AnimationChartOnlyBuildTypeFromString(const MString &string);


	enum class DgmBuildStep
	{
		Sp,
		Bg,
	};

	DgmBuildStep DgmBuildStepFromString(const MString &string);


	enum class ChartBuildStep
	{
		Category,
		PtInCategory,
		Series,
		PtInSeries,
		AllPts,
		GridLegend,
	};

	ChartBuildStep ChartBuildStepFromString(const MString &string);


	enum class OnOffStyleType
	{
		On,
		Off,
		Def,
	};

	OnOffStyleType OnOffStyleTypeFromString(const MString &string);

*/