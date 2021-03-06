unit Excel_TLB_Constants;


// ************************************************************************ //
// NOTE
// -------
// This file is created by importing the Excel Type Library and then extracting
// the constants section of the TLB.pas file.  Other declaraitons are not required.
// ************************************************************************ //

// $Rev: 52393 $
// File generated on 27 Dec 2019 19:52:33 from Type Library described below.

// ************************************************************************
//86\Microsoft Shared\OFFICE16\MSO.DLL)
//   (2) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
//   (3) v5.3 VBIDE, (C:\Program Files (x86)\Microsoft Office\Root\VFS\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB)
// Type Lib: C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE (1)
// LIBID: {00020813-0000-0000-C000-000000000046}
// LCID: 0
// Helpfile: C:\Program Files (x86)\Microsoft Office\Root\Office16\VBAXL10.CHM
// SYS_KIND: SYS_WIN32
// Errors:
//   Hint: Symbol 'IFont' renamed to 'ExcelIFont'

Interface 
// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
//
// *********************************************************************//

const
// Constants for enum Constants
  xlAll = $FFFFEFF8;
  xlAutomatic = $FFFFEFF7;
  xlBoth = $00000001;
  xlCenter = $FFFFEFF4;
  xlChecker = $00000009;
  xlCircle = $00000008;
  xlCorner = $00000002;
  xlCrissCross = $00000010;
  xlCross = $00000004;
  xlDiamond = $00000002;
  xlDistributed = $FFFFEFEB;
  xlDoubleAccounting = $00000005;
  xlFixedValue = $00000001;
  xlFormats = $FFFFEFE6;
  xlGray16 = $00000011;
  xlGray8 = $00000012;
  xlGrid = $0000000F;
  xlHigh = $FFFFEFE1;
  xlInside = $00000002;
  xlJustify = $FFFFEFDE;
  xlLightDown = $0000000D;
  xlLightHorizontal = $0000000B;
  xlLightUp = $0000000E;
  xlLightVertical = $0000000C;
  xlLow = $FFFFEFDA;
  xlManual = $FFFFEFD9;
  xlMinusValues = $00000003;
  xlModule = $FFFFEFD3;
  xlNextToAxis = $00000004;
  xlNone = $FFFFEFD2;
  xlNotes = $FFFFEFD0;
  xlOff = $FFFFEFCE;
  xlOn = $00000001;
  xlPercent = $00000002;
  xlPlus = $00000009;
  xlPlusValues = $00000002;
  xlSemiGray75 = $0000000A;
  xlShowLabel = $00000004;
  xlShowLabelAndPercent = $00000005;
  xlShowPercent = $00000003;
  xlShowValue = $00000002;
  xlSimple = $FFFFEFC6;
  xlSingle = $00000002;
  xlSingleAccounting = $00000004;
  xlSolid = $00000001;
  xlSquare = $00000001;
  xlStar = $00000005;
  xlStError = $00000004;
  xlToolbarButton = $00000002;
  xlTriangle = $00000003;
  xlGray25 = $FFFFEFE4;
  xlGray50 = $FFFFEFE3;
  xlGray75 = $FFFFEFE2;
  xlBottom = $FFFFEFF5;
  xlLeft = $FFFFEFDD;
  xlRight = $FFFFEFC8;
  xlTop = $FFFFEFC0;
  xl3DBar = $FFFFEFFD;
  xl3DSurface = $FFFFEFF9;
  xlBar = $00000002;
  xlColumn = $00000003;
  xlCombination = $FFFFEFF1;
  xlCustom = $FFFFEFEE;
  xlDefaultAutoFormat = $FFFFFFFF;
  xlMaximum = $00000002;
  xlMinimum = $00000004;
  xlOpaque = $00000003;
  xlTransparent = $00000002;
  xlBidi = $FFFFEC78;
  xlLatin = $FFFFEC77;
  xlContext = $FFFFEC76;
  xlLTR = $FFFFEC75;
  xlRTL = $FFFFEC74;
  xlFullScript = $00000001;
  xlPartialScript = $00000002;
  xlMixedScript = $00000003;
  xlMixedAuthorizedScript = $00000004;
  xlVisualCursor = $00000002;
  xlLogicalCursor = $00000001;
  xlSystem = $00000001;
  xlPartial = $00000003;
  xlHindiNumerals = $00000003;
  xlBidiCalendar = $00000003;
  xlGregorian = $00000002;
  xlComplete = $00000004;
  xlScale = $00000003;
  xlClosed = $00000003;
  xlColor1 = $00000007;
  xlColor2 = $00000008;
  xlColor3 = $00000009;
  xlConstants = $00000002;
  xlContents = $00000002;
  xlBelow = $00000001;
  xlCascade = $00000007;
  xlCenterAcrossSelection = $00000007;
  xlChart4 = $00000002;
  xlChartSeries = $00000011;
  xlChartShort = $00000006;
  xlChartTitles = $00000012;
  xlClassic1 = $00000001;
  xlClassic2 = $00000002;
  xlClassic3 = $00000003;
  xl3DEffects1 = $0000000D;
  xl3DEffects2 = $0000000E;
  xlAbove = $00000000;
  xlAccounting1 = $00000004;
  xlAccounting2 = $00000005;
  xlAccounting3 = $00000006;
  xlAccounting4 = $00000011;
  xlAdd = $00000002;
  xlDebugCodePane = $0000000D;
  xlDesktop = $00000009;
  xlDirect = $00000001;
  xlDivide = $00000005;
  xlDoubleClosed = $00000005;
  xlDoubleOpen = $00000004;
  xlDoubleQuote = $00000001;
  xlEntireChart = $00000014;
  xlExcelMenus = $00000001;
  xlExtended = $00000003;
  xlFill = $00000005;
  xlFirst = $00000000;
  xlFloating = $00000005;
  xlFormula = $00000005;
  xlGeneral = $00000001;
  xlGridline = $00000016;
  xlIcons = $00000001;
  xlImmediatePane = $0000000C;
  xlInteger = $00000002;
  xlLast = $00000001;
  xlLastCell = $0000000B;
  xlList1 = $0000000A;
  xlList2 = $0000000B;
  xlList3 = $0000000C;
  xlLocalFormat1 = $0000000F;
  xlLocalFormat2 = $00000010;
  xlLong = $00000003;
  xlLotusHelp = $00000002;
  xlMacrosheetCell = $00000007;
  xlMixed = $00000002;
  xlMultiply = $00000004;
  xlNarrow = $00000001;
  xlNoDocuments = $00000003;
  xlOpen = $00000002;
  xlOutside = $00000003;
  xlReference = $00000004;
  xlSemiautomatic = $00000002;
  xlShort = $00000001;
  xlSingleQuote = $00000002;
  xlStrict = $00000002;
  xlSubtract = $00000003;
  xlTextBox = $00000010;
  xlTiled = $00000001;
  xlTitleBar = $00000008;
  xlToolbar = $00000001;
  xlVisible = $0000000C;
  xlWatchPane = $0000000B;
  xlWide = $00000003;
  xlWorkbookTab = $00000006;
  xlWorksheet4 = $00000001;
  xlWorksheetCell = $00000003;
  xlWorksheetShort = $00000005;
  xlAllExceptBorders = $00000007;
  xlLeftToRight = $00000002;
  xlTopToBottom = $00000001;
  xlVeryHidden = $00000002;
  xlDrawingObject = $0000000E;


// Constants for enum XlCreator
  xlCreatorCode = $5843454C;


// Constants for enum XlChartGallery
  xlBuiltIn = $00000015;
  xlUserDefined = $00000016;
  xlAnyGallery = $00000017;


// Constants for enum XlColorIndex
  xlColorIndexAutomatic = $FFFFEFF7;
  xlColorIndexNone = $FFFFEFD2;


// Constants for enum XlEndStyleCap
  xlCap = $00000001;
  xlNoCap = $00000002;


// Constants for enum XlRowCol
  xlColumns = $00000002;
  xlRows = $00000001;


// Constants for enum XlScaleType
  xlScaleLinear = $FFFFEFDC;
  xlScaleLogarithmic = $FFFFEFDB;


// Constants for enum XlDataSeriesType
  xlAutoFill = $00000004;
  xlChronological = $00000003;
  xlGrowth = $00000002;
  xlDataSeriesLinear = $FFFFEFDC;


// Constants for enum XlAxisCrosses
  xlAxisCrossesAutomatic = $FFFFEFF7;
  xlAxisCrossesCustom = $FFFFEFEE;
  xlAxisCrossesMaximum = $00000002;
  xlAxisCrossesMinimum = $00000004;


// Constants for enum XlAxisGroup
  xlPrimary = $00000001;
  xlSecondary = $00000002;


// Constants for enum XlBackground
  xlBackgroundAutomatic = $FFFFEFF7;
  xlBackgroundOpaque = $00000003;
  xlBackgroundTransparent = $00000002;


// Constants for enum XlWindowState
  xlMaximized = $FFFFEFD7;
  xlMinimized = $FFFFEFD4;
  xlNormal = $FFFFEFD1;


// Constants for enum XlAxisType
  xlCategory = $00000001;
  xlSeriesAxis = $00000003;
  xlValue = $00000002;


// Constants for enum XlArrowHeadLength
  xlArrowHeadLengthLong = $00000003;
  xlArrowHeadLengthMedium = $FFFFEFD6;
  xlArrowHeadLengthShort = $00000001;


// Constants for enum XlVAlign
  xlVAlignBottom = $FFFFEFF5;
  xlVAlignCenter = $FFFFEFF4;
  xlVAlignDistributed = $FFFFEFEB;
  xlVAlignJustify = $FFFFEFDE;
  xlVAlignTop = $FFFFEFC0;


// Constants for enum XlTickMark
  xlTickMarkCross = $00000004;
  xlTickMarkInside = $00000002;
  xlTickMarkNone = $FFFFEFD2;
  xlTickMarkOutside = $00000003;


// Constants for enum XlErrorBarDirection
  xlX = $FFFFEFB8;
  xlY = $00000001;


// Constants for enum XlErrorBarInclude
  xlErrorBarIncludeBoth = $00000001;
  xlErrorBarIncludeMinusValues = $00000003;
  xlErrorBarIncludeNone = $FFFFEFD2;
  xlErrorBarIncludePlusValues = $00000002;


// Constants for enum XlDisplayBlanksAs
  xlInterpolated = $00000003;
  xlNotPlotted = $00000001;
  xlZero = $00000002;


// Constants for enum XlArrowHeadStyle
  xlArrowHeadStyleClosed = $00000003;
  xlArrowHeadStyleDoubleClosed = $00000005;
  xlArrowHeadStyleDoubleOpen = $00000004;
  xlArrowHeadStyleNone = $FFFFEFD2;
  xlArrowHeadStyleOpen = $00000002;


// Constants for enum XlArrowHeadWidth
  xlArrowHeadWidthMedium = $FFFFEFD6;
  xlArrowHeadWidthNarrow = $00000001;
  xlArrowHeadWidthWide = $00000003;


// Constants for enum XlHAlign
  xlHAlignCenter = $FFFFEFF4;
  xlHAlignCenterAcrossSelection = $00000007;
  xlHAlignDistributed = $FFFFEFEB;
  xlHAlignFill = $00000005;
  xlHAlignGeneral = $00000001;
  xlHAlignJustify = $FFFFEFDE;
  xlHAlignLeft = $FFFFEFDD;
  xlHAlignRight = $FFFFEFC8;


// Constants for enum XlTickLabelPosition
  xlTickLabelPositionHigh = $FFFFEFE1;
  xlTickLabelPositionLow = $FFFFEFDA;
  xlTickLabelPositionNextToAxis = $00000004;
  xlTickLabelPositionNone = $FFFFEFD2;


// Constants for enum XlLegendPosition
  xlLegendPositionBottom = $FFFFEFF5;
  xlLegendPositionCorner = $00000002;
  xlLegendPositionLeft = $FFFFEFDD;
  xlLegendPositionRight = $FFFFEFC8;
  xlLegendPositionTop = $FFFFEFC0;
  xlLegendPositionCustom = $FFFFEFBF;


// Constants for enum XlChartPictureType
  xlStackScale = $00000003;
  xlStack = $00000002;
  xlStretch = $00000001;


// Constants for enum XlChartPicturePlacement
  xlSides = $00000001;
  xlEnd = $00000002;
  xlEndSides = $00000003;
  xlFront = $00000004;
  xlFrontSides = $00000005;
  xlFrontEnd = $00000006;
  xlAllFaces = $00000007;


// Constants for enum XlOrientation
  xlDownward = $FFFFEFB6;
  xlHorizontal = $FFFFEFE0;
  xlUpward = $FFFFEFB5;
  xlVertical = $FFFFEFBA;


// Constants for enum XlTickLabelOrientation
  xlTickLabelOrientationAutomatic = $FFFFEFF7;
  xlTickLabelOrientationDownward = $FFFFEFB6;
  xlTickLabelOrientationHorizontal = $FFFFEFE0;
  xlTickLabelOrientationUpward = $FFFFEFB5;
  xlTickLabelOrientationVertical = $FFFFEFBA;


// Constants for enum XlBorderWeight
  xlHairline = $00000001;
  xlMedium = $FFFFEFD6;
  xlThick = $00000004;
  xlThin = $00000002;


// Constants for enum XlDataSeriesDate
  xlDay = $00000001;
  xlMonth = $00000003;
  xlWeekday = $00000002;
  xlYear = $00000004;


// Constants for enum XlUnderlineStyle
  xlUnderlineStyleDouble = $FFFFEFE9;
  xlUnderlineStyleDoubleAccounting = $00000005;
  xlUnderlineStyleNone = $FFFFEFD2;
  xlUnderlineStyleSingle = $00000002;
  xlUnderlineStyleSingleAccounting = $00000004;


// Constants for enum XlErrorBarType
  xlErrorBarTypeCustom = $FFFFEFEE;
  xlErrorBarTypeFixedValue = $00000001;
  xlErrorBarTypePercent = $00000002;
  xlErrorBarTypeStDev = $FFFFEFC5;
  xlErrorBarTypeStError = $00000004;


// Constants for enum XlTrendlineType
  xlExponential = $00000005;
  xlLinear = $FFFFEFDC;
  xlLogarithmic = $FFFFEFDB;
  xlMovingAvg = $00000006;
  xlPolynomial = $00000003;
  xlPower = $00000004;


// Constants for enum XlLineStyle
  xlContinuous = $00000001;
  xlDash = $FFFFEFED;
  xlDashDot = $00000004;
  xlDashDotDot = $00000005;
  xlDot = $FFFFEFEA;
  xlDouble = $FFFFEFE9;
  xlSlantDashDot = $0000000D;
  xlLineStyleNone = $FFFFEFD2;


// Constants for enum XlDataLabelsType
  xlDataLabelsShowNone = $FFFFEFD2;
  xlDataLabelsShowValue = $00000002;
  xlDataLabelsShowPercent = $00000003;
  xlDataLabelsShowLabel = $00000004;
  xlDataLabelsShowLabelAndPercent = $00000005;
  xlDataLabelsShowBubbleSizes = $00000006;


// Constants for enum XlMarkerStyle
  xlMarkerStyleAutomatic = $FFFFEFF7;
  xlMarkerStyleCircle = $00000008;
  xlMarkerStyleDash = $FFFFEFED;
  xlMarkerStyleDiamond = $00000002;
  xlMarkerStyleDot = $FFFFEFEA;
  xlMarkerStyleNone = $FFFFEFD2;
  xlMarkerStylePicture = $FFFFEFCD;
  xlMarkerStylePlus = $00000009;
  xlMarkerStyleSquare = $00000001;
  xlMarkerStyleStar = $00000005;
  xlMarkerStyleTriangle = $00000003;
  xlMarkerStyleX = $FFFFEFB8;


// Constants for enum XlPictureConvertorType
  xlBMP = $00000001;
  xlCGM = $00000007;
  xlDRW = $00000004;
  xlDXF = $00000005;
  xlEPS = $00000008;
  xlHGL = $00000006;
  xlPCT = $0000000D;
  xlPCX = $0000000A;
  xlPIC = $0000000B;
  xlPLT = $0000000C;
  xlTIF = $00000009;
  xlWMF = $00000002;
  xlWPG = $00000003;


// Constants for enum XlPattern
  xlPatternAutomatic = $FFFFEFF7;
  xlPatternChecker = $00000009;
  xlPatternCrissCross = $00000010;
  xlPatternDown = $FFFFEFE7;
  xlPatternGray16 = $00000011;
  xlPatternGray25 = $FFFFEFE4;
  xlPatternGray50 = $FFFFEFE3;
  xlPatternGray75 = $FFFFEFE2;
  xlPatternGray8 = $00000012;
  xlPatternGrid = $0000000F;
  xlPatternHorizontal = $FFFFEFE0;
  xlPatternLightDown = $0000000D;
  xlPatternLightHorizontal = $0000000B;
  xlPatternLightUp = $0000000E;
  xlPatternLightVertical = $0000000C;
  xlPatternNone = $FFFFEFD2;
  xlPatternSemiGray75 = $0000000A;
  xlPatternSolid = $00000001;
  xlPatternUp = $FFFFEFBE;
  xlPatternVertical = $FFFFEFBA;
  xlPatternLinearGradient = $00000FA0;
  xlPatternRectangularGradient = $00000FA1;


// Constants for enum XlChartSplitType
  xlSplitByPosition = $00000001;
  xlSplitByPercentValue = $00000003;
  xlSplitByCustomSplit = $00000004;
  xlSplitByValue = $00000002;


// Constants for enum XlDisplayUnit
  xlHundreds = $FFFFFFFE;
  xlThousands = $FFFFFFFD;
  xlTenThousands = $FFFFFFFC;
  xlHundredThousands = $FFFFFFFB;
  xlMillions = $FFFFFFFA;
  xlTenMillions = $FFFFFFF9;
  xlHundredMillions = $FFFFFFF8;
  xlThousandMillions = $FFFFFFF7;
  xlMillionMillions = $FFFFFFF6;


// Constants for enum XlDataLabelPosition
  xlLabelPositionCenter = $FFFFEFF4;
  xlLabelPositionAbove = $00000000;
  xlLabelPositionBelow = $00000001;
  xlLabelPositionLeft = $FFFFEFDD;
  xlLabelPositionRight = $FFFFEFC8;
  xlLabelPositionOutsideEnd = $00000002;
  xlLabelPositionInsideEnd = $00000003;
  xlLabelPositionInsideBase = $00000004;
  xlLabelPositionBestFit = $00000005;
  xlLabelPositionMixed = $00000006;
  xlLabelPositionCustom = $00000007;


// Constants for enum XlTimeUnit
  xlDays = $00000000;
  xlMonths = $00000001;
  xlYears = $00000002;


// Constants for enum XlCategoryType
  xlCategoryScale = $00000002;
  xlTimeScale = $00000003;
  xlAutomaticScale = $FFFFEFF7;


// Constants for enum XlBarShape
  xlBox = $00000000;
  xlPyramidToPoint = $00000001;
  xlPyramidToMax = $00000002;
  xlCylinder = $00000003;
  xlConeToPoint = $00000004;
  xlConeToMax = $00000005;


// Constants for enum XlChartType
  xlColumnClustered = $00000033;
  xlColumnStacked = $00000034;
  xlColumnStacked100 = $00000035;
  xl3DColumnClustered = $00000036;
  xl3DColumnStacked = $00000037;
  xl3DColumnStacked100 = $00000038;
  xlBarClustered = $00000039;
  xlBarStacked = $0000003A;
  xlBarStacked100 = $0000003B;
  xl3DBarClustered = $0000003C;
  xl3DBarStacked = $0000003D;
  xl3DBarStacked100 = $0000003E;
  xlLineStacked = $0000003F;
  xlLineStacked100 = $00000040;
  xlLineMarkers = $00000041;
  xlLineMarkersStacked = $00000042;
  xlLineMarkersStacked100 = $00000043;
  xlPieOfPie = $00000044;
  xlPieExploded = $00000045;
  xl3DPieExploded = $00000046;
  xlBarOfPie = $00000047;
  xlXYScatterSmooth = $00000048;
  xlXYScatterSmoothNoMarkers = $00000049;
  xlXYScatterLines = $0000004A;
  xlXYScatterLinesNoMarkers = $0000004B;
  xlAreaStacked = $0000004C;
  xlAreaStacked100 = $0000004D;
  xl3DAreaStacked = $0000004E;
  xl3DAreaStacked100 = $0000004F;
  xlDoughnutExploded = $00000050;
  xlRadarMarkers = $00000051;
  xlRadarFilled = $00000052;
  xlSurface = $00000053;
  xlSurfaceWireframe = $00000054;
  xlSurfaceTopView = $00000055;
  xlSurfaceTopViewWireframe = $00000056;
  xlBubble = $0000000F;
  xlBubble3DEffect = $00000057;
  xlStockHLC = $00000058;
  xlStockOHLC = $00000059;
  xlStockVHLC = $0000005A;
  xlStockVOHLC = $0000005B;
  xlCylinderColClustered = $0000005C;
  xlCylinderColStacked = $0000005D;
  xlCylinderColStacked100 = $0000005E;
  xlCylinderBarClustered = $0000005F;
  xlCylinderBarStacked = $00000060;
  xlCylinderBarStacked100 = $00000061;
  xlCylinderCol = $00000062;
  xlConeColClustered = $00000063;
  xlConeColStacked = $00000064;
  xlConeColStacked100 = $00000065;
  xlConeBarClustered = $00000066;
  xlConeBarStacked = $00000067;
  xlConeBarStacked100 = $00000068;
  xlConeCol = $00000069;
  xlPyramidColClustered = $0000006A;
  xlPyramidColStacked = $0000006B;
  xlPyramidColStacked100 = $0000006C;
  xlPyramidBarClustered = $0000006D;
  xlPyramidBarStacked = $0000006E;
  xlPyramidBarStacked100 = $0000006F;
  xlPyramidCol = $00000070;
  xl3DColumn = $FFFFEFFC;
  xlLine = $00000004;
  xl3DLine = $FFFFEFFB;
  xl3DPie = $FFFFEFFA;
  xlPie = $00000005;
  xlXYScatter = $FFFFEFB7;
  xl3DArea = $FFFFEFFE;
  xlArea = $00000001;
  xlDoughnut = $FFFFEFE8;
  xlRadar = $FFFFEFC9;
  xlTreemap = $00000075;
  xlHistogram = $00000076;
  xlWaterfall = $00000077;
  xlSunburst = $00000078;
  xlBoxwhisker = $00000079;
  xlPareto = $0000007A;
  xlFunnel = $0000007B;
  xlRegionMap = $0000008C;


// Constants for enum XlChartItem
  xlDataLabel = $00000000;
  xlChartArea = $00000002;
  xlSeries = $00000003;
  xlChartTitle = $00000004;
  xlWalls = $00000005;
  xlCorners = $00000006;
  xlDataTable = $00000007;
  xlTrendline = $00000008;
  xlErrorBars = $00000009;
  xlXErrorBars = $0000000A;
  xlYErrorBars = $0000000B;
  xlLegendEntry = $0000000C;
  xlLegendKey = $0000000D;
  xlShape = $0000000E;
  xlMajorGridlines = $0000000F;
  xlMinorGridlines = $00000010;
  xlAxisTitle = $00000011;
  xlUpBars = $00000012;
  xlPlotArea = $00000013;
  xlDownBars = $00000014;
  xlAxis = $00000015;
  xlSeriesLines = $00000016;
  xlFloor = $00000017;
  xlLegend = $00000018;
  xlHiLoLines = $00000019;
  xlDropLines = $0000001A;
  xlRadarAxisLabels = $0000001B;
  xlNothing = $0000001C;
  xlLeaderLines = $0000001D;
  xlDisplayUnitLabel = $0000001E;
  xlPivotChartFieldButton = $0000001F;
  xlPivotChartDropZone = $00000020;
  xlPivotChartExpandEntireFieldButton = $00000021;
  xlPivotChartCollapseEntireFieldButton = $00000022;


// Constants for enum XlSizeRepresents
  xlSizeIsWidth = $00000002;
  xlSizeIsArea = $00000001;


// Constants for enum XlInsertShiftDirection
  xlShiftDown = $FFFFEFE7;
  xlShiftToRight = $FFFFEFBF;


// Constants for enum XlDeleteShiftDirection
  xlShiftToLeft = $FFFFEFC1;
  xlShiftUp = $FFFFEFBE;


// Constants for enum XlDirection
  xlDown = $FFFFEFE7;
  xlToLeft = $FFFFEFC1;
  xlToRight = $FFFFEFBF;
  xlUp = $FFFFEFBE;


// Constants for enum XlConsolidationFunction
  xlAverage = $FFFFEFF6;
  xlCount = $FFFFEFF0;
  xlCountNums = $FFFFEFEF;
  xlMax = $FFFFEFD8;
  xlMin = $FFFFEFD5;
  xlProduct = $FFFFEFCB;
  xlStDev = $FFFFEFC5;
  xlStDevP = $FFFFEFC4;
  xlSum = $FFFFEFC3;
  xlVar = $FFFFEFBC;
  xlVarP = $FFFFEFBB;
  xlUnknown = $000003E8;
  xlDistinctCount = $0000000B;


// Constants for enum XlSheetType
  xlChart = $FFFFEFF3;
  xlDialogSheet = $FFFFEFEC;
  xlExcel4IntlMacroSheet = $00000004;
  xlExcel4MacroSheet = $00000003;
  xlWorksheet = $FFFFEFB9;


// Constants for enum XlLocationInTable
  xlColumnHeader = $FFFFEFF2;
  xlColumnItem = $00000005;
  xlDataHeader = $00000003;
  xlDataItem = $00000007;
  xlPageHeader = $00000002;
  xlPageItem = $00000006;
  xlRowHeader = $FFFFEFC7;
  xlRowItem = $00000004;
  xlTableBody = $00000008;


// Constants for enum XlFindLookIn
  xlFormulas = $FFFFEFE5;
  xlComments = $FFFFEFD0;
  xlValues = $FFFFEFBD;
  xlCommentsThreaded = $FFFFEFA8;
  xlFormulas2 = $FFFFEFA7;


// Constants for enum XlWindowType
  xlChartAsWindow = $00000005;
  xlChartInPlace = $00000004;
  xlClipboard = $00000003;
  xlInfo = $FFFFEFDF;
  xlWorkbook = $00000001;


// Constants for enum XlPivotFieldDataType
  xlDate = $00000002;
  xlNumber = $FFFFEFCF;
  xlText = $FFFFEFC2;


// Constants for enum XlCopyPictureFormat
  xlBitmap = $00000002;
  xlPicture = $FFFFEFCD;


// Constants for enum XlPivotTableSourceType
  xlScenario = $00000004;
  xlConsolidation = $00000003;
  xlDatabase = $00000001;
  xlExternal = $00000002;
  xlPivotTable = $FFFFEFCC;


// Constants for enum XlReferenceStyle
  xlA1 = $00000001;
  xlR1C1 = $FFFFEFCA;


// Constants for enum XlMSApplication
  xlMicrosoftAccess = $00000004;
  xlMicrosoftFoxPro = $00000005;
  xlMicrosoftMail = $00000003;
  xlMicrosoftPowerPoint = $00000002;
  xlMicrosoftProject = $00000006;
  xlMicrosoftSchedulePlus = $00000007;
  xlMicrosoftWord = $00000001;


// Constants for enum XlMouseButton
  xlNoButton = $00000000;
  xlPrimaryButton = $00000001;
  xlSecondaryButton = $00000002;


// Constants for enum XlCutCopyMode
  xlCopy = $00000001;
  xlCut = $00000002;


// Constants for enum XlFillWith
  xlFillWithAll = $FFFFEFF8;
  xlFillWithContents = $00000002;
  xlFillWithFormats = $FFFFEFE6;


// Constants for enum XlFilterAction
  xlFilterCopy = $00000002;
  xlFilterInPlace = $00000001;


// Constants for enum XlOrder
  xlDownThenOver = $00000001;
  xlOverThenDown = $00000002;


// Constants for enum XlLinkType
  xlLinkTypeExcelLinks = $00000001;
  xlLinkTypeOLELinks = $00000002;


// Constants for enum XlApplyNamesOrder
  xlColumnThenRow = $00000002;
  xlRowThenColumn = $00000001;


// Constants for enum XlEnableCancelKey
  xlDisabled = $00000000;
  xlErrorHandler = $00000002;
  xlInterrupt = $00000001;


// Constants for enum XlPageBreak
  xlPageBreakAutomatic = $FFFFEFF7;
  xlPageBreakManual = $FFFFEFD9;
  xlPageBreakNone = $FFFFEFD2;


// Constants for enum XlOLEType
  xlOLEControl = $00000002;
  xlOLEEmbed = $00000001;
  xlOLELink = $00000000;


// Constants for enum XlPageOrientation
  xlLandscape = $00000002;
  xlPortrait = $00000001;


// Constants for enum XlLinkInfo
  xlEditionDate = $00000002;
  xlUpdateState = $00000001;
  xlLinkInfoStatus = $00000003;


// Constants for enum XlCommandUnderlines
  xlCommandUnderlinesAutomatic = $FFFFEFF7;
  xlCommandUnderlinesOff = $FFFFEFCE;
  xlCommandUnderlinesOn = $00000001;


// Constants for enum XlOLEVerb
  xlVerbOpen = $00000002;
  xlVerbPrimary = $00000001;


// Constants for enum XlCalculation
  xlCalculationAutomatic = $FFFFEFF7;
  xlCalculationManual = $FFFFEFD9;
  xlCalculationSemiautomatic = $00000002;


// Constants for enum XlFileAccess
  xlReadOnly = $00000003;
  xlReadWrite = $00000002;


// Constants for enum XlEditionType
  xlPublisher = $00000001;
  xlSubscriber = $00000002;


// Constants for enum XlObjectSize
  xlFitToPage = $00000002;
  xlFullPage = $00000003;
  xlScreenSize = $00000001;


// Constants for enum XlLookAt
  xlPart = $00000002;
  xlWhole = $00000001;


// Constants for enum XlMailSystem
  xlMAPI = $00000001;
  xlNoMailSystem = $00000000;
  xlPowerTalk = $00000002;


// Constants for enum XlLinkInfoType
  xlLinkInfoOLELinks = $00000002;
  xlLinkInfoPublishers = $00000005;
  xlLinkInfoSubscribers = $00000006;


// Constants for enum XlCVError
  xlErrBlocked = $000007FF;
  xlErrCalc = $00000802;
  xlErrConnect = $000007FE;
  xlErrDiv0 = $000007D7;
  xlErrField = $00000801;
  xlErrGettingData = $000007FB;
  xlErrNA = $000007FA;
  xlErrName = $000007ED;
  xlErrSpill = $000007FD;
  xlErrNull = $000007D0;
  xlErrNum = $000007F4;
  xlErrRef = $000007E7;
  xlErrUnknown = $00000800;
  xlErrValue = $000007DF;


// Constants for enum XlEditionFormat
  xlBIFF = $00000002;
  xlPICT = $00000001;
  xlRTF = $00000004;
  xlVALU = $00000008;


// Constants for enum XlLink
  xlExcelLinks = $00000001;
  xlOLELinks = $00000002;
  xlPublishers = $00000005;
  xlSubscribers = $00000006;


// Constants for enum XlCellType
  xlCellTypeBlanks = $00000004;
  xlCellTypeConstants = $00000002;
  xlCellTypeFormulas = $FFFFEFE5;
  xlCellTypeLastCell = $0000000B;
  xlCellTypeComments = $FFFFEFD0;
  xlCellTypeVisible = $0000000C;
  xlCellTypeAllFormatConditions = $FFFFEFB4;
  xlCellTypeSameFormatConditions = $FFFFEFB3;
  xlCellTypeAllValidation = $FFFFEFB2;
  xlCellTypeSameValidation = $FFFFEFB1;


// Constants for enum XlArrangeStyle
  xlArrangeStyleCascade = $00000007;
  xlArrangeStyleHorizontal = $FFFFEFE0;
  xlArrangeStyleTiled = $00000001;
  xlArrangeStyleVertical = $FFFFEFBA;


// Constants for enum XlMousePointer
  xlIBeam = $00000003;
  xlDefault = $FFFFEFD1;
  xlNorthwestArrow = $00000001;
  xlWait = $00000002;


// Constants for enum XlEditionOptionsOption
  xlAutomaticUpdate = $00000004;
  xlCancel = $00000001;
  xlChangeAttributes = $00000006;
  xlManualUpdate = $00000005;
  xlOpenSource = $00000003;
  xlSelect = $00000003;
  xlSendPublisher = $00000002;
  xlUpdateSubscriber = $00000002;


// Constants for enum XlAutoFillType
  xlFillCopy = $00000001;
  xlFillDays = $00000005;
  xlFillDefault = $00000000;
  xlFillFormats = $00000003;
  xlFillMonths = $00000007;
  xlFillSeries = $00000002;
  xlFillValues = $00000004;
  xlFillWeekdays = $00000006;
  xlFillYears = $00000008;
  xlGrowthTrend = $0000000A;
  xlLinearTrend = $00000009;
  xlFlashFill = $0000000B;


// Constants for enum XlAutoFilterOperator
  xlAnd = $00000001;
  xlBottom10Items = $00000004;
  xlBottom10Percent = $00000006;
  xlOr = $00000002;
  xlTop10Items = $00000003;
  xlTop10Percent = $00000005;
  xlFilterValues = $00000007;
  xlFilterCellColor = $00000008;
  xlFilterFontColor = $00000009;
  xlFilterIcon = $0000000A;
  xlFilterDynamic = $0000000B;
  xlFilterNoFill = $0000000C;
  xlFilterAutomaticFontColor = $0000000D;
  xlFilterNoIcon = $0000000E;


// Constants for enum XlClipboardFormat
  xlClipboardFormatBIFF12 = $0000003F;
  xlClipboardFormatBIFF = $00000008;
  xlClipboardFormatBIFF2 = $00000012;
  xlClipboardFormatBIFF3 = $00000014;
  xlClipboardFormatBIFF4 = $0000001E;
  xlClipboardFormatBinary = $0000000F;
  xlClipboardFormatBitmap = $00000009;
  xlClipboardFormatCGM = $0000000D;
  xlClipboardFormatCSV = $00000005;
  xlClipboardFormatDIF = $00000004;
  xlClipboardFormatDspText = $0000000C;
  xlClipboardFormatEmbeddedObject = $00000015;
  xlClipboardFormatEmbedSource = $00000016;
  xlClipboardFormatLink = $0000000B;
  xlClipboardFormatLinkSource = $00000017;
  xlClipboardFormatLinkSourceDesc = $00000020;
  xlClipboardFormatMovie = $00000018;
  xlClipboardFormatNative = $0000000E;
  xlClipboardFormatObjectDesc = $0000001F;
  xlClipboardFormatObjectLink = $00000013;
  xlClipboardFormatOwnerLink = $00000011;
  xlClipboardFormatPICT = $00000002;
  xlClipboardFormatPrintPICT = $00000003;
  xlClipboardFormatRTF = $00000007;
  xlClipboardFormatScreenPICT = $0000001D;
  xlClipboardFormatStandardFont = $0000001C;
  xlClipboardFormatStandardScale = $0000001B;
  xlClipboardFormatSYLK = $00000006;
  xlClipboardFormatTable = $00000010;
  xlClipboardFormatText = $00000000;
  xlClipboardFormatToolFace = $00000019;
  xlClipboardFormatToolFacePICT = $0000001A;
  xlClipboardFormatVALU = $00000001;
  xlClipboardFormatWK1 = $0000000A;


// Constants for enum XlFileFormat
  xlAddIn = $00000012;
  xlCSV = $00000006;
  xlCSVMac = $00000016;
  xlCSVMSDOS = $00000018;
  xlCSVWindows = $00000017;
  xlDBF2 = $00000007;
  xlDBF3 = $00000008;
  xlDBF4 = $0000000B;
  xlDIF = $00000009;
  xlExcel2 = $00000010;
  xlExcel2FarEast = $0000001B;
  xlExcel3 = $0000001D;
  xlExcel4 = $00000021;
  xlExcel5 = $00000027;
  xlExcel7 = $00000027;
  xlExcel9795 = $0000002B;
  xlExcel4Workbook = $00000023;
  xlIntlAddIn = $0000001A;
  xlIntlMacro = $00000019;
  xlWorkbookNormal = $FFFFEFD1;
  xlSYLK = $00000002;
  xlTemplate = $00000011;
  xlCurrentPlatformText = $FFFFEFC2;
  xlTextMac = $00000013;
  xlTextMSDOS = $00000015;
  xlTextPrinter = $00000024;
  xlTextWindows = $00000014;
  xlWJ2WD1 = $0000000E;
  xlWK1 = $00000005;
  xlWK1ALL = $0000001F;
  xlWK1FMT = $0000001E;
  xlWK3 = $0000000F;
  xlWK4 = $00000026;
  xlWK3FM3 = $00000020;
  xlWKS = $00000004;
  xlWorks2FarEast = $0000001C;
  xlWQ1 = $00000022;
  xlWJ3 = $00000028;
  xlWJ3FJ3 = $00000029;
  xlUnicodeText = $0000002A;
  xlHtml = $0000002C;
  xlWebArchive = $0000002D;
  xlXMLSpreadsheet = $0000002E;
  xlExcel12 = $00000032;
  xlOpenXMLWorkbook = $00000033;
  xlOpenXMLWorkbookMacroEnabled = $00000034;
  xlOpenXMLTemplateMacroEnabled = $00000035;
  xlTemplate8 = $00000011;
  xlOpenXMLTemplate = $00000036;
  xlAddIn8 = $00000012;
  xlOpenXMLAddIn = $00000037;
  xlExcel8 = $00000038;
  xlOpenDocumentSpreadsheet = $0000003C;
  xlOpenXMLStrictWorkbook = $0000003D;
  xlCSVUTF8 = $0000003E;
  xlWorkbookDefault = $00000033;


// Constants for enum XlApplicationInternational
  xl24HourClock = $00000021;
  xl4DigitYears = $0000002B;
  xlAlternateArraySeparator = $00000010;
  xlColumnSeparator = $0000000E;
  xlCountryCode = $00000001;
  xlCountrySetting = $00000002;
  xlCurrencyBefore = $00000025;
  xlCurrencyCode = $00000019;
  xlCurrencyDigits = $0000001B;
  xlCurrencyLeadingZeros = $00000028;
  xlCurrencyMinusSign = $00000026;
  xlCurrencyNegative = $0000001C;
  xlCurrencySpaceBefore = $00000024;
  xlCurrencyTrailingZeros = $00000027;
  xlDateOrder = $00000020;
  xlDateSeparator = $00000011;
  xlDayCode = $00000015;
  xlDayLeadingZero = $0000002A;
  xlDecimalSeparator = $00000003;
  xlGeneralFormatName = $0000001A;
  xlHourCode = $00000016;
  xlLeftBrace = $0000000C;
  xlLeftBracket = $0000000A;
  xlListSeparator = $00000005;
  xlLowerCaseColumnLetter = $00000009;
  xlLowerCaseRowLetter = $00000008;
  xlMDY = $0000002C;
  xlMetric = $00000023;
  xlMinuteCode = $00000017;
  xlMonthCode = $00000014;
  xlMonthLeadingZero = $00000029;
  xlMonthNameChars = $0000001E;
  xlNoncurrencyDigits = $0000001D;
  xlNonEnglishFunctions = $00000022;
  xlRightBrace = $0000000D;
  xlRightBracket = $0000000B;
  xlRowSeparator = $0000000F;
  xlSecondCode = $00000018;
  xlThousandsSeparator = $00000004;
  xlTimeLeadingZero = $0000002D;
  xlTimeSeparator = $00000012;
  xlUpperCaseColumnLetter = $00000007;
  xlUpperCaseRowLetter = $00000006;
  xlWeekdayNameChars = $0000001F;
  xlYearCode = $00000013;
  xlUICultureTag = $0000002E;


// Constants for enum XlPageBreakExtent
  xlPageBreakFull = $00000001;
  xlPageBreakPartial = $00000002;


// Constants for enum XlCellInsertionMode
  xlOverwriteCells = $00000000;
  xlInsertDeleteCells = $00000001;
  xlInsertEntireRows = $00000002;


// Constants for enum XlFormulaLabel
  xlNoLabels = $FFFFEFD2;
  xlRowLabels = $00000001;
  xlColumnLabels = $00000002;
  xlMixedLabels = $00000003;


// Constants for enum XlHighlightChangesTime
  xlSinceMyLastSave = $00000001;
  xlAllChanges = $00000002;
  xlNotYetReviewed = $00000003;


// Constants for enum XlCommentDisplayMode
  xlNoIndicator = $00000000;
  xlCommentIndicatorOnly = $FFFFFFFF;
  xlCommentAndIndicator = $00000001;


// Constants for enum XlFormatConditionType
  xlCellValue = $00000001;
  xlExpression = $00000002;
  xlColorScale = $00000003;
  xlDatabar = $00000004;
  xlTop10 = $00000005;
  xlIconSets = $00000006;
  xlUniqueValues = $00000008;
  xlTextString = $00000009;
  xlBlanksCondition = $0000000A;
  xlTimePeriod = $0000000B;
  xlAboveAverageCondition = $0000000C;
  xlNoBlanksCondition = $0000000D;
  xlErrorsCondition = $00000010;
  xlNoErrorsCondition = $00000011;


// Constants for enum XlFormatConditionOperator
  xlBetween = $00000001;
  xlNotBetween = $00000002;
  xlEqual = $00000003;
  xlNotEqual = $00000004;
  xlGreater = $00000005;
  xlLess = $00000006;
  xlGreaterEqual = $00000007;
  xlLessEqual = $00000008;


// Constants for enum XlEnableSelection
  xlNoRestrictions = $00000000;
  xlUnlockedCells = $00000001;
  xlNoSelection = $FFFFEFD2;


// Constants for enum XlDVType
  xlValidateInputOnly = $00000000;
  xlValidateWholeNumber = $00000001;
  xlValidateDecimal = $00000002;
  xlValidateList = $00000003;
  xlValidateDate = $00000004;
  xlValidateTime = $00000005;
  xlValidateTextLength = $00000006;
  xlValidateCustom = $00000007;


// Constants for enum XlIMEMode
  xlIMEModeNoControl = $00000000;
  xlIMEModeOn = $00000001;
  xlIMEModeOff = $00000002;
  xlIMEModeDisable = $00000003;
  xlIMEModeHiragana = $00000004;
  xlIMEModeKatakana = $00000005;
  xlIMEModeKatakanaHalf = $00000006;
  xlIMEModeAlphaFull = $00000007;
  xlIMEModeAlpha = $00000008;
  xlIMEModeHangulFull = $00000009;
  xlIMEModeHangul = $0000000A;


// Constants for enum XlDVAlertStyle
  xlValidAlertStop = $00000001;
  xlValidAlertWarning = $00000002;
  xlValidAlertInformation = $00000003;


// Constants for enum XlChartLocation
  xlLocationAsNewSheet = $00000001;
  xlLocationAsObject = $00000002;
  xlLocationAutomatic = $00000003;


// Constants for enum XlPaperSize
  xlPaper10x14 = $00000010;
  xlPaper11x17 = $00000011;
  xlPaperA3 = $00000008;
  xlPaperA4 = $00000009;
  xlPaperA4Small = $0000000A;
  xlPaperA5 = $0000000B;
  xlPaperB4 = $0000000C;
  xlPaperB5 = $0000000D;
  xlPaperCsheet = $00000018;
  xlPaperDsheet = $00000019;
  xlPaperEnvelope10 = $00000014;
  xlPaperEnvelope11 = $00000015;
  xlPaperEnvelope12 = $00000016;
  xlPaperEnvelope14 = $00000017;
  xlPaperEnvelope9 = $00000013;
  xlPaperEnvelopeB4 = $00000021;
  xlPaperEnvelopeB5 = $00000022;
  xlPaperEnvelopeB6 = $00000023;
  xlPaperEnvelopeC3 = $0000001D;
  xlPaperEnvelopeC4 = $0000001E;
  xlPaperEnvelopeC5 = $0000001C;
  xlPaperEnvelopeC6 = $0000001F;
  xlPaperEnvelopeC65 = $00000020;
  xlPaperEnvelopeDL = $0000001B;
  xlPaperEnvelopeItaly = $00000024;
  xlPaperEnvelopeMonarch = $00000025;
  xlPaperEnvelopePersonal = $00000026;
  xlPaperEsheet = $0000001A;
  xlPaperExecutive = $00000007;
  xlPaperFanfoldLegalGerman = $00000029;
  xlPaperFanfoldStdGerman = $00000028;
  xlPaperFanfoldUS = $00000027;
  xlPaperFolio = $0000000E;
  xlPaperLedger = $00000004;
  xlPaperLegal = $00000005;
  xlPaperLetter = $00000001;
  xlPaperLetterSmall = $00000002;
  xlPaperNote = $00000012;
  xlPaperQuarto = $0000000F;
  xlPaperStatement = $00000006;
  xlPaperTabloid = $00000003;
  xlPaperUser = $00000100;


// Constants for enum XlPasteSpecialOperation
  xlPasteSpecialOperationAdd = $00000002;
  xlPasteSpecialOperationDivide = $00000005;
  xlPasteSpecialOperationMultiply = $00000004;
  xlPasteSpecialOperationNone = $FFFFEFD2;
  xlPasteSpecialOperationSubtract = $00000003;


// Constants for enum XlPasteType
  xlPasteAll = $FFFFEFF8;
  xlPasteAllUsingSourceTheme = $0000000D;
  xlPasteAllMergingConditionalFormats = $0000000E;
  xlPasteAllExceptBorders = $00000007;
  xlPasteFormats = $FFFFEFE6;
  xlPasteFormulas = $FFFFEFE5;
  xlPasteComments = $FFFFEFD0;
  xlPasteValues = $FFFFEFBD;
  xlPasteColumnWidths = $00000008;
  xlPasteValidation = $00000006;
  xlPasteFormulasAndNumberFormats = $0000000B;
  xlPasteValuesAndNumberFormats = $0000000C;


// Constants for enum XlPhoneticCharacterType
  xlKatakanaHalf = $00000000;
  xlKatakana = $00000001;
  xlHiragana = $00000002;
  xlNoConversion = $00000003;


// Constants for enum XlPhoneticAlignment
  xlPhoneticAlignNoControl = $00000000;
  xlPhoneticAlignLeft = $00000001;
  xlPhoneticAlignCenter = $00000002;
  xlPhoneticAlignDistributed = $00000003;


// Constants for enum XlPictureAppearance
  xlPrinter = $00000002;
  xlScreen = $00000001;


// Constants for enum XlPivotFieldOrientation
  xlColumnField = $00000002;
  xlDataField = $00000004;
  xlHidden = $00000000;
  xlPageField = $00000003;
  xlRowField = $00000001;


// Constants for enum XlPivotFieldCalculation
  xlDifferenceFrom = $00000002;
  xlIndex = $00000009;
  xlNoAdditionalCalculation = $FFFFEFD1;
  xlPercentDifferenceFrom = $00000004;
  xlPercentOf = $00000003;
  xlPercentOfColumn = $00000007;
  xlPercentOfRow = $00000006;
  xlPercentOfTotal = $00000008;
  xlRunningTotal = $00000005;
  xlPercentOfParentRow = $0000000A;
  xlPercentOfParentColumn = $0000000B;
  xlPercentOfParent = $0000000C;
  xlPercentRunningTotal = $0000000D;
  xlRankAscending = $0000000E;
  xlRankDecending = $0000000F;


// Constants for enum XlPlacement
  xlFreeFloating = $00000003;
  xlMove = $00000002;
  xlMoveAndSize = $00000001;


// Constants for enum XlPlatform
  xlMacintosh = $00000001;
  xlMSDOS = $00000003;
  xlWindows = $00000002;


// Constants for enum XlPrintLocation
  xlPrintSheetEnd = $00000001;
  xlPrintInPlace = $00000010;
  xlPrintNoComments = $FFFFEFD2;


// Constants for enum XlPriority
  xlPriorityHigh = $FFFFEFE1;
  xlPriorityLow = $FFFFEFDA;
  xlPriorityNormal = $FFFFEFD1;


// Constants for enum XlPTSelectionMode
  xlLabelOnly = $00000001;
  xlDataAndLabel = $00000000;
  xlDataOnly = $00000002;
  xlOrigin = $00000003;
  xlButton = $0000000F;
  xlBlanks = $00000004;
  xlFirstRow = $00000100;


// Constants for enum XlRangeAutoFormat
  xlRangeAutoFormat3DEffects1 = $0000000D;
  xlRangeAutoFormat3DEffects2 = $0000000E;
  xlRangeAutoFormatAccounting1 = $00000004;
  xlRangeAutoFormatAccounting2 = $00000005;
  xlRangeAutoFormatAccounting3 = $00000006;
  xlRangeAutoFormatAccounting4 = $00000011;
  xlRangeAutoFormatClassic1 = $00000001;
  xlRangeAutoFormatClassic2 = $00000002;
  xlRangeAutoFormatClassic3 = $00000003;
  xlRangeAutoFormatColor1 = $00000007;
  xlRangeAutoFormatColor2 = $00000008;
  xlRangeAutoFormatColor3 = $00000009;
  xlRangeAutoFormatList1 = $0000000A;
  xlRangeAutoFormatList2 = $0000000B;
  xlRangeAutoFormatList3 = $0000000C;
  xlRangeAutoFormatLocalFormat1 = $0000000F;
  xlRangeAutoFormatLocalFormat2 = $00000010;
  xlRangeAutoFormatLocalFormat3 = $00000013;
  xlRangeAutoFormatLocalFormat4 = $00000014;
  xlRangeAutoFormatReport1 = $00000015;
  xlRangeAutoFormatReport2 = $00000016;
  xlRangeAutoFormatReport3 = $00000017;
  xlRangeAutoFormatReport4 = $00000018;
  xlRangeAutoFormatReport5 = $00000019;
  xlRangeAutoFormatReport6 = $0000001A;
  xlRangeAutoFormatReport7 = $0000001B;
  xlRangeAutoFormatReport8 = $0000001C;
  xlRangeAutoFormatReport9 = $0000001D;
  xlRangeAutoFormatReport10 = $0000001E;
  xlRangeAutoFormatClassicPivotTable = $0000001F;
  xlRangeAutoFormatTable1 = $00000020;
  xlRangeAutoFormatTable2 = $00000021;
  xlRangeAutoFormatTable3 = $00000022;
  xlRangeAutoFormatTable4 = $00000023;
  xlRangeAutoFormatTable5 = $00000024;
  xlRangeAutoFormatTable6 = $00000025;
  xlRangeAutoFormatTable7 = $00000026;
  xlRangeAutoFormatTable8 = $00000027;
  xlRangeAutoFormatTable9 = $00000028;
  xlRangeAutoFormatTable10 = $00000029;
  xlRangeAutoFormatPTNone = $0000002A;
  xlRangeAutoFormatNone = $FFFFEFD2;
  xlRangeAutoFormatSimple = $FFFFEFC6;


// Constants for enum XlReferenceType
  xlAbsolute = $00000001;
  xlAbsRowRelColumn = $00000002;
  xlRelative = $00000004;
  xlRelRowAbsColumn = $00000003;


// Constants for enum XlLayoutFormType
  xlTabular = $00000000;
  xlOutline = $00000001;


// Constants for enum XlRoutingSlipDelivery
  xlAllAtOnce = $00000002;
  xlOneAfterAnother = $00000001;


// Constants for enum XlRoutingSlipStatus
  xlNotYetRouted = $00000000;
  xlRoutingComplete = $00000002;
  xlRoutingInProgress = $00000001;


// Constants for enum XlRunAutoMacro
  xlAutoActivate = $00000003;
  xlAutoClose = $00000002;
  xlAutoDeactivate = $00000004;
  xlAutoOpen = $00000001;


// Constants for enum XlSaveAction
  xlDoNotSaveChanges = $00000002;
  xlSaveChanges = $00000001;


// Constants for enum XlSaveAsAccessMode
  xlExclusive = $00000003;
  xlNoChange = $00000001;
  xlShared = $00000002;


// Constants for enum XlSaveConflictResolution
  xlLocalSessionChanges = $00000002;
  xlOtherSessionChanges = $00000003;
  xlUserResolution = $00000001;


// Constants for enum XlSearchDirection
  xlNext = $00000001;
  xlPrevious = $00000002;


// Constants for enum XlSearchOrder
  xlByColumns = $00000002;
  xlByRows = $00000001;


// Constants for enum XlSheetVisibility
  xlSheetVisible = $FFFFFFFF;
  xlSheetHidden = $00000000;
  xlSheetVeryHidden = $00000002;


// Constants for enum XlSortMethod
  xlPinYin = $00000001;
  xlStroke = $00000002;


// Constants for enum XlSortMethodOld
  xlCodePage = $00000002;
  xlSyllabary = $00000001;


// Constants for enum XlSortOrder
  xlAscending = $00000001;
  xlDescending = $00000002;


// Constants for enum XlSortOrientation
  xlSortRows = $00000002;
  xlSortColumns = $00000001;


// Constants for enum XlSortType
  xlSortLabels = $00000002;
  xlSortValues = $00000001;


// Constants for enum XlSpecialCellsValue
  xlErrors = $00000010;
  xlLogical = $00000004;
  xlNumbers = $00000001;
  xlTextValues = $00000002;


// Constants for enum XlSubscribeToFormat
  xlSubscribeToPicture = $FFFFEFCD;
  xlSubscribeToText = $FFFFEFC2;


// Constants for enum XlSummaryRow
  xlSummaryAbove = $00000000;
  xlSummaryBelow = $00000001;


// Constants for enum XlSummaryColumn
  xlSummaryOnLeft = $FFFFEFDD;
  xlSummaryOnRight = $FFFFEFC8;


// Constants for enum XlSummaryReportType
  xlSummaryPivotTable = $FFFFEFCC;
  xlStandardSummary = $00000001;


// Constants for enum XlTabPosition
  xlTabPositionFirst = $00000000;
  xlTabPositionLast = $00000001;


// Constants for enum XlTextParsingType
  xlDelimited = $00000001;
  xlFixedWidth = $00000002;


// Constants for enum XlTextQualifier
  xlTextQualifierDoubleQuote = $00000001;
  xlTextQualifierNone = $FFFFEFD2;
  xlTextQualifierSingleQuote = $00000002;


// Constants for enum XlWBATemplate
  xlWBATChart = $FFFFEFF3;
  xlWBATExcel4IntlMacroSheet = $00000004;
  xlWBATExcel4MacroSheet = $00000003;
  xlWBATWorksheet = $FFFFEFB9;


// Constants for enum XlWindowView
  xlNormalView = $00000001;
  xlPageBreakPreview = $00000002;
  xlPageLayoutView = $00000003;


// Constants for enum XlXLMMacroType
  xlCommand = $00000002;
  xlFunction = $00000001;
  xlNotXLM = $00000003;


// Constants for enum XlYesNoGuess
  xlGuess = $00000000;
  xlNo = $00000002;
  xlYes = $00000001;


// Constants for enum XlBordersIndex
  xlInsideHorizontal = $0000000C;
  xlInsideVertical = $0000000B;
  xlDiagonalDown = $00000005;
  xlDiagonalUp = $00000006;
  xlEdgeBottom = $00000009;
  xlEdgeLeft = $00000007;
  xlEdgeRight = $0000000A;
  xlEdgeTop = $00000008;


// Constants for enum XlToolbarProtection
  xlNoButtonChanges = $00000001;
  xlNoChanges = $00000004;
  xlNoDockingChanges = $00000003;
  xlToolbarProtectionNone = $FFFFEFD1;
  xlNoShapeChanges = $00000002;


// Constants for enum XlBuiltInDialog
  xlDialogOpen = $00000001;
  xlDialogOpenLinks = $00000002;
  xlDialogSaveAs = $00000005;
  xlDialogFileDelete = $00000006;
  xlDialogPageSetup = $00000007;
  xlDialogPrint = $00000008;
  xlDialogPrinterSetup = $00000009;
  xlDialogArrangeAll = $0000000C;
  xlDialogWindowSize = $0000000D;
  xlDialogWindowMove = $0000000E;
  xlDialogRun = $00000011;
  xlDialogSetPrintTitles = $00000017;
  xlDialogFont = $0000001A;
  xlDialogDisplay = $0000001B;
  xlDialogProtectDocument = $0000001C;
  xlDialogCalculation = $00000020;
  xlDialogExtract = $00000023;
  xlDialogDataDelete = $00000024;
  xlDialogSort = $00000027;
  xlDialogDataSeries = $00000028;
  xlDialogTable = $00000029;
  xlDialogFormatNumber = $0000002A;
  xlDialogAlignment = $0000002B;
  xlDialogStyle = $0000002C;
  xlDialogBorder = $0000002D;
  xlDialogCellProtection = $0000002E;
  xlDialogColumnWidth = $0000002F;
  xlDialogClear = $00000034;
  xlDialogPasteSpecial = $00000035;
  xlDialogEditDelete = $00000036;
  xlDialogInsert = $00000037;
  xlDialogPasteNames = $0000003A;
  xlDialogDefineName = $0000003D;
  xlDialogCreateNames = $0000003E;
  xlDialogFormulaGoto = $0000003F;
  xlDialogFormulaFind = $00000040;
  xlDialogGalleryArea = $00000043;
  xlDialogGalleryBar = $00000044;
  xlDialogGalleryColumn = $00000045;
  xlDialogGalleryLine = $00000046;
  xlDialogGalleryPie = $00000047;
  xlDialogGalleryScatter = $00000048;
  xlDialogCombination = $00000049;
  xlDialogGridlines = $0000004C;
  xlDialogAxes = $0000004E;
  xlDialogAttachText = $00000050;
  xlDialogPatterns = $00000054;
  xlDialogMainChart = $00000055;
  xlDialogOverlay = $00000056;
  xlDialogScale = $00000057;
  xlDialogFormatLegend = $00000058;
  xlDialogFormatText = $00000059;
  xlDialogParse = $0000005B;
  xlDialogUnhide = $0000005E;
  xlDialogWorkspace = $0000005F;
  xlDialogActivate = $00000067;
  xlDialogCopyPicture = $0000006C;
  xlDialogDeleteName = $0000006E;
  xlDialogDeleteFormat = $0000006F;
  xlDialogNew = $00000077;
  xlDialogRowHeight = $0000007F;
  xlDialogFormatMove = $00000080;
  xlDialogFormatSize = $00000081;
  xlDialogFormulaReplace = $00000082;
  xlDialogSelectSpecial = $00000084;
  xlDialogApplyNames = $00000085;
  xlDialogReplaceFont = $00000086;
  xlDialogSplit = $00000089;
  xlDialogOutline = $0000008E;
  xlDialogSaveWorkbook = $00000091;
  xlDialogCopyChart = $00000093;
  xlDialogFormatFont = $00000096;
  xlDialogNote = $0000009A;
  xlDialogSetUpdateStatus = $0000009F;
  xlDialogColorPalette = $000000A1;
  xlDialogChangeLink = $000000A6;
  xlDialogAppMove = $000000AA;
  xlDialogAppSize = $000000AB;
  xlDialogMainChartType = $000000B9;
  xlDialogOverlayChartType = $000000BA;
  xlDialogOpenMail = $000000BC;
  xlDialogSendMail = $000000BD;
  xlDialogStandardFont = $000000BE;
  xlDialogConsolidate = $000000BF;
  xlDialogSortSpecial = $000000C0;
  xlDialogGallery3dArea = $000000C1;
  xlDialogGallery3dColumn = $000000C2;
  xlDialogGallery3dLine = $000000C3;
  xlDialogGallery3dPie = $000000C4;
  xlDialogView3d = $000000C5;
  xlDialogGoalSeek = $000000C6;
  xlDialogWorkgroup = $000000C7;
  xlDialogFillGroup = $000000C8;
  xlDialogUpdateLink = $000000C9;
  xlDialogPromote = $000000CA;
  xlDialogDemote = $000000CB;
  xlDialogShowDetail = $000000CC;
  xlDialogObjectProperties = $000000CF;
  xlDialogSaveNewObject = $000000D0;
  xlDialogApplyStyle = $000000D4;
  xlDialogAssignToObject = $000000D5;
  xlDialogObjectProtection = $000000D6;
  xlDialogCreatePublisher = $000000D9;
  xlDialogSubscribeTo = $000000DA;
  xlDialogShowToolbar = $000000DC;
  xlDialogPrintPreview = $000000DE;
  xlDialogEditColor = $000000DF;
  xlDialogFormatMain = $000000E1;
  xlDialogFormatOverlay = $000000E2;
  xlDialogEditSeries = $000000E4;
  xlDialogDefineStyle = $000000E5;
  xlDialogGalleryRadar = $000000F9;
  xlDialogEditionOptions = $000000FB;
  xlDialogZoom = $00000100;
  xlDialogInsertObject = $00000103;
  xlDialogSize = $00000105;
  xlDialogMove = $00000106;
  xlDialogFormatAuto = $0000010D;
  xlDialogGallery3dBar = $00000110;
  xlDialogGallery3dSurface = $00000111;
  xlDialogCustomizeToolbar = $00000114;
  xlDialogWorkbookAdd = $00000119;
  xlDialogWorkbookMove = $0000011A;
  xlDialogWorkbookCopy = $0000011B;
  xlDialogWorkbookOptions = $0000011C;
  xlDialogSaveWorkspace = $0000011D;
  xlDialogChartWizard = $00000120;
  xlDialogAssignToTool = $00000125;
  xlDialogPlacement = $0000012C;
  xlDialogFillWorkgroup = $0000012D;
  xlDialogWorkbookNew = $0000012E;
  xlDialogScenarioCells = $00000131;
  xlDialogScenarioAdd = $00000133;
  xlDialogScenarioEdit = $00000134;
  xlDialogScenarioSummary = $00000137;
  xlDialogPivotTableWizard = $00000138;
  xlDialogPivotFieldProperties = $00000139;
  xlDialogOptionsCalculation = $0000013E;
  xlDialogOptionsEdit = $0000013F;
  xlDialogOptionsView = $00000140;
  xlDialogAddinManager = $00000141;
  xlDialogMenuEditor = $00000142;
  xlDialogAttachToolbars = $00000143;
  xlDialogOptionsChart = $00000145;
  xlDialogVbaInsertFile = $00000148;
  xlDialogVbaProcedureDefinition = $0000014A;
  xlDialogRoutingSlip = $00000150;
  xlDialogMailLogon = $00000153;
  xlDialogInsertPicture = $00000156;
  xlDialogGalleryDoughnut = $00000158;
  xlDialogChartTrend = $0000015E;
  xlDialogWorkbookInsert = $00000162;
  xlDialogOptionsTransition = $00000163;
  xlDialogOptionsGeneral = $00000164;
  xlDialogFilterAdvanced = $00000172;
  xlDialogMailNextLetter = $0000017A;
  xlDialogDataLabel = $0000017B;
  xlDialogInsertTitle = $0000017C;
  xlDialogFontProperties = $0000017D;
  xlDialogMacroOptions = $0000017E;
  xlDialogWorkbookUnhide = $00000180;
  xlDialogWorkbookName = $00000182;
  xlDialogGalleryCustom = $00000184;
  xlDialogAddChartAutoformat = $00000186;
  xlDialogChartAddData = $00000188;
  xlDialogTabOrder = $0000018A;
  xlDialogSubtotalCreate = $0000018E;
  xlDialogWorkbookTabSplit = $0000019F;
  xlDialogWorkbookProtect = $000001A1;
  xlDialogScrollbarProperties = $000001A4;
  
  xlDialogPivotShowPages = $000001A5;
  xlDialogTextToColumns = $000001A6;
  xlDialogCheckboxProperties = $000001B3;
  xlDialogLabelProperties = $000001B4;
  xlDialogListboxProperties = $000001B5;
  xlDialogEditboxProperties = $000001B6;
  xlDialogOpenText = $000001B9;
  xlDialogPushbuttonProperties = $000001BD;
  xlDialogFilter = $000001BF;
  xlDialogFunctionWizard = $000001C2;
  xlDialogSaveCopyAs = $000001C8;
  xlDialogOptionsListsAdd = $000001CA;
  xlDialogSeriesAxes = $000001CC;
  xlDialogSeriesX = $000001CD;
  xlDialogSeriesY = $000001CE;
  xlDialogErrorbarX = $000001CF;
  xlDialogErrorbarY = $000001D0;
  xlDialogFormatChart = $000001D1;
  xlDialogSeriesOrder = $000001D2;
  xlDialogMailEditMailer = $000001D6;
  xlDialogStandardWidth = $000001D8;
  xlDialogScenarioMerge = $000001D9;
  xlDialogProperties = $000001DA;
  xlDialogSummaryInfo = $000001DA;
  xlDialogFindFile = $000001DB;
  xlDialogActiveCellFont = $000001DC;
  xlDialogVbaMakeAddin = $000001DE;
  xlDialogFileSharing = $000001E1;
  xlDialogAutoCorrect = $000001E5;
  xlDialogCustomViews = $000001ED;
  xlDialogInsertNameLabel = $000001F0;
  xlDialogSeriesShape = $000001F8;
  xlDialogChartOptionsDataLabels = $000001F9;
  xlDialogChartOptionsDataTable = $000001FA;
  xlDialogSetBackgroundPicture = $000001FD;
  xlDialogDataValidation = $0000020D;
  xlDialogChartType = $0000020E;
  xlDialogChartLocation = $0000020F;
  _xlDialogPhonetic = $0000021A;
  xlDialogChartSourceData = $0000021C;
  _xlDialogChartSourceData = $0000021D;
  xlDialogSeriesOptions = $0000022D;
  xlDialogPivotTableOptions = $00000237;
  xlDialogPivotSolveOrder = $00000238;
  xlDialogPivotCalculatedField = $0000023A;
  xlDialogPivotCalculatedItem = $0000023C;
  xlDialogConditionalFormatting = $00000247;
  xlDialogInsertHyperlink = $00000254;
  xlDialogProtectSharing = $0000026C;
  xlDialogOptionsME = $00000287;
  xlDialogPublishAsWebPage = $0000028D;
  xlDialogPhonetic = $00000290;
  xlDialogNewWebQuery = $0000029B;
  xlDialogImportTextFile = $0000029A;
  xlDialogExternalDataProperties = $00000212;
  xlDialogWebOptionsGeneral = $000002AB;
  xlDialogWebOptionsFiles = $000002AC;
  xlDialogWebOptionsPictures = $000002AD;
  xlDialogWebOptionsEncoding = $000002AE;
  xlDialogWebOptionsFonts = $000002AF;
  xlDialogPivotClientServerSet = $000002B1;
  xlDialogPropertyFields = $000002F2;
  xlDialogSearch = $000002DB;
  xlDialogEvaluateFormula = $000002C5;
  xlDialogDataLabelMultiple = $000002D3;
  xlDialogChartOptionsDataLabelMultiple = $000002D4;
  xlDialogErrorChecking = $000002DC;
  xlDialogWebOptionsBrowsers = $00000305;
  xlDialogCreateList = $0000031C;
  xlDialogPermission = $00000340;
  xlDialogMyPermission = $00000342;
  xlDialogDocumentInspector = $0000035E;
  xlDialogNameManager = $000003D1;
  xlDialogNewName = $000003D2;
  xlDialogSparklineInsertLine = $0000046D;
  xlDialogSparklineInsertColumn = $0000046E;
  xlDialogSparklineInsertWinLoss = $0000046F;
  xlDialogSlicerSettings = $0000049B;
  xlDialogSlicerCreation = $0000049E;
  xlDialogSlicerPivotTableConnections = $000004A0;
  xlDialogPivotTableSlicerConnections = $0000049F;
  xlDialogPivotTableWhatIfAnalysisSettings = $00000481;
  xlDialogSetManager = $00000455;
  xlDialogSetMDXEditor = $000004B8;
  xlDialogSetTupleEditorOnRows = $00000453;
  xlDialogSetTupleEditorOnColumns = $00000454;
  xlDialogManageRelationships = $000004F7;
  xlDialogCreateRelationship = $000004F8;
  xlDialogRecommendedPivotTables = $000004EA;
  xlDialogForecastETS = $00000514;
  xlDialogPivotDefaultLayout = $00000550;


// Constants for enum XlParameterType
  xlPrompt = $00000000;
  xlConstant = $00000001;
  xlRange = $00000002;


// Constants for enum XlParameterDataType
  xlParamTypeUnknown = $00000000;
  xlParamTypeChar = $00000001;
  xlParamTypeNumeric = $00000002;
  xlParamTypeDecimal = $00000003;
  xlParamTypeInteger = $00000004;
  xlParamTypeSmallInt = $00000005;
  xlParamTypeFloat = $00000006;
  xlParamTypeReal = $00000007;
  xlParamTypeDouble = $00000008;
  xlParamTypeVarChar = $0000000C;
  xlParamTypeDate = $00000009;
  xlParamTypeTime = $0000000A;
  xlParamTypeTimestamp = $0000000B;
  xlParamTypeLongVarChar = $FFFFFFFF;
  xlParamTypeBinary = $FFFFFFFE;
  xlParamTypeVarBinary = $FFFFFFFD;
  xlParamTypeLongVarBinary = $FFFFFFFC;
  xlParamTypeBigInt = $FFFFFFFB;
  xlParamTypeTinyInt = $FFFFFFFA;
  xlParamTypeBit = $FFFFFFF9;
  xlParamTypeWChar = $FFFFFFF8;


// Constants for enum XlFormControl
  xlButtonControl = $00000000;
  xlCheckBox = $00000001;
  xlDropDown = $00000002;
  xlEditBox = $00000003;
  xlGroupBox = $00000004;
  xlLabel = $00000005;
  xlListBox = $00000006;
  xlOptionButton = $00000007;
  xlScrollBar = $00000008;
  xlSpinner = $00000009;


// Constants for enum XlSourceType
  xlSourceWorkbook = $00000000;
  xlSourceSheet = $00000001;
  xlSourcePrintArea = $00000002;
  xlSourceAutoFilter = $00000003;
  xlSourceRange = $00000004;
  xlSourceChart = $00000005;
  xlSourcePivotTable = $00000006;
  xlSourceQuery = $00000007;


// Constants for enum XlHtmlType
  xlHtmlStatic = $00000000;
  xlHtmlCalc = $00000001;
  xlHtmlList = $00000002;
  xlHtmlChart = $00000003;


// Constants for enum XlPivotFormatType
  xlReport1 = $00000000;
  xlReport2 = $00000001;
  xlReport3 = $00000002;
  xlReport4 = $00000003;
  xlReport5 = $00000004;
  xlReport6 = $00000005;
  xlReport7 = $00000006;
  xlReport8 = $00000007;
  xlReport9 = $00000008;
  xlReport10 = $00000009;
  xlTable1 = $0000000A;
  xlTable2 = $0000000B;
  xlTable3 = $0000000C;
  xlTable4 = $0000000D;
  xlTable5 = $0000000E;
  xlTable6 = $0000000F;
  xlTable7 = $00000010;
  xlTable8 = $00000011;
  xlTable9 = $00000012;
  xlTable10 = $00000013;
  xlPTClassic = $00000014;
  xlPTNone = $00000015;


// Constants for enum XlCmdType
  xlCmdCube = $00000001;
  xlCmdSql = $00000002;
  xlCmdTable = $00000003;
  xlCmdDefault = $00000004;
  xlCmdList = $00000005;
  xlCmdTableCollection = $00000006;
  xlCmdExcel = $00000007;
  xlCmdDAX = $00000008;


// Constants for enum XlColumnDataType
  xlGeneralFormat = $00000001;
  xlTextFormat = $00000002;
  xlMDYFormat = $00000003;
  xlDMYFormat = $00000004;
  xlYMDFormat = $00000005;
  xlMYDFormat = $00000006;
  xlDYMFormat = $00000007;
  xlYDMFormat = $00000008;
  xlSkipColumn = $00000009;
  xlEMDFormat = $0000000A;


// Constants for enum XlQueryType
  xlODBCQuery = $00000001;
  xlDAORecordset = $00000002;
  xlWebQuery = $00000004;
  xlOLEDBQuery = $00000005;
  xlTextImport = $00000006;
  xlADORecordset = $00000007;


// Constants for enum XlWebSelectionType
  xlEntirePage = $00000001;
  xlAllTables = $00000002;
  xlSpecifiedTables = $00000003;


// Constants for enum XlCubeFieldType
  xlHierarchy = $00000001;
  xlMeasure = $00000002;
  xlSet = $00000003;


// Constants for enum XlWebFormatting
  xlWebFormattingAll = $00000001;
  xlWebFormattingRTF = $00000002;
  xlWebFormattingNone = $00000003;


// Constants for enum XlDisplayDrawingObjects
  xlDisplayShapes = $FFFFEFF8;
  xlHide = $00000003;
  xlPlaceholders = $00000002;


// Constants for enum XlSubtototalLocationType
  xlAtTop = $00000001;
  xlAtBottom = $00000002;


// Constants for enum XlPivotTableVersionList
  xlPivotTableVersion2000 = $00000000;
  xlPivotTableVersion10 = $00000001;
  xlPivotTableVersion11 = $00000002;
  xlPivotTableVersion12 = $00000003;
  xlPivotTableVersion14 = $00000004;
  xlPivotTableVersion15 = $00000005;
  xlPivotTableVersionCurrent = $FFFFFFFF;


// Constants for enum XlPrintErrors
  xlPrintErrorsDisplayed = $00000000;
  xlPrintErrorsBlank = $00000001;
  xlPrintErrorsDash = $00000002;
  xlPrintErrorsNA = $00000003;


// Constants for enum XlPivotCellType
  xlPivotCellValue = $00000000;
  xlPivotCellPivotItem = $00000001;
  xlPivotCellSubtotal = $00000002;
  xlPivotCellGrandTotal = $00000003;
  xlPivotCellDataField = $00000004;
  xlPivotCellPivotField = $00000005;
  xlPivotCellPageFieldItem = $00000006;
  xlPivotCellCustomSubtotal = $00000007;
  xlPivotCellDataPivotField = $00000008;
  xlPivotCellBlankCell = $00000009;


// Constants for enum XlPivotTableMissingItems
  xlMissingItemsDefault = $FFFFFFFF;
  xlMissingItemsNone = $00000000;
  xlMissingItemsMax = $00007EF4;
  xlMissingItemsMax2 = $00100000;


// Constants for enum XlCalculationState
  xlDone = $00000000;
  xlCalculating = $00000001;
  xlPending = $00000002;


// Constants for enum XlCalculationInterruptKey
  xlNoKey = $00000000;
  xlEscKey = $00000001;
  xlAnyKey = $00000002;


// Constants for enum XlSortDataOption
  xlSortNormal = $00000000;
  xlSortTextAsNumbers = $00000001;


// Constants for enum XlUpdateLinks
  xlUpdateLinksUserSetting = $00000001;
  xlUpdateLinksNever = $00000002;
  xlUpdateLinksAlways = $00000003;


// Constants for enum XlLinkStatus
  xlLinkStatusOK = $00000000;
  xlLinkStatusMissingFile = $00000001;
  xlLinkStatusMissingSheet = $00000002;
  xlLinkStatusOld = $00000003;
  xlLinkStatusSourceNotCalculated = $00000004;
  xlLinkStatusIndeterminate = $00000005;
  xlLinkStatusNotStarted = $00000006;
  xlLinkStatusInvalidName = $00000007;
  xlLinkStatusSourceNotOpen = $00000008;
  xlLinkStatusSourceOpen = $00000009;
  xlLinkStatusCopiedValues = $0000000A;


// Constants for enum XlSearchWithin
  xlWithinSheet = $00000001;
  xlWithinWorkbook = $00000002;


// Constants for enum XlCorruptLoad
  xlNormalLoad = $00000000;
  xlRepairFile = $00000001;
  xlExtractData = $00000002;


// Constants for enum XlRobustConnect
  xlAsRequired = $00000000;
  xlAlways = $00000001;
  xlNever = $00000002;


// Constants for enum XlErrorChecks
  xlEvaluateToError = $00000001;
  xlTextDate = $00000002;
  xlNumberAsText = $00000003;
  xlInconsistentFormula = $00000004;
  xlOmittedCells = $00000005;
  xlUnlockedFormulaCells = $00000006;
  xlEmptyCellReferences = $00000007;
  xlListDataValidation = $00000008;
  xlInconsistentListFormula = $00000009;
  xlMisleadingFormat = $0000000A;


// Constants for enum XlDataLabelSeparator
  xlDataLabelSeparatorDefault = $00000001;


// Constants for enum XlSmartTagDisplayMode
  xlIndicatorAndButton = $00000000;
  xlDisplayNone = $00000001;
  xlButtonOnly = $00000002;


// Constants for enum XlRangeValueDataType
  xlRangeValueDefault = $0000000A;
  xlRangeValueXMLSpreadsheet = $0000000B;
  xlRangeValueMSPersistXML = $0000000C;


// Constants for enum XlSpeakDirection
  xlSpeakByRows = $00000000;
  xlSpeakByColumns = $00000001;


// Constants for enum XlInsertFormatOrigin
  xlFormatFromLeftOrAbove = $00000000;
  xlFormatFromRightOrBelow = $00000001;


// Constants for enum XlArabicModes
  xlArabicNone = $00000000;
  xlArabicStrictAlefHamza = $00000001;
  xlArabicStrictFinalYaa = $00000002;
  xlArabicBothStrict = $00000003;


// Constants for enum XlImportDataAs
  xlQueryTable = $00000000;
  xlPivotTableReport = $00000001;
  xlTable = $00000002;


// Constants for enum XlCalculatedMemberType
  xlCalculatedMember = $00000000;
  xlCalculatedSet = $00000001;
  xlCalculatedMeasure = $00000002;


// Constants for enum XlHebrewModes
  xlHebrewFullScript = $00000000;
  xlHebrewPartialScript = $00000001;
  xlHebrewMixedScript = $00000002;
  xlHebrewMixedAuthorizedScript = $00000003;


// Constants for enum XlListObjectSourceType
  xlSrcExternal = $00000000;
  xlSrcRange = $00000001;
  xlSrcXml = $00000002;
  xlSrcQuery = $00000003;
  xlSrcModel = $00000004;


// Constants for enum XlTextVisualLayoutType
  xlTextVisualLTR = $00000001;
  xlTextVisualRTL = $00000002;


// Constants for enum XlListDataType
  xlListDataTypeNone = $00000000;
  xlListDataTypeText = $00000001;
  xlListDataTypeMultiLineText = $00000002;
  xlListDataTypeNumber = $00000003;
  xlListDataTypeCurrency = $00000004;
  xlListDataTypeDateTime = $00000005;
  xlListDataTypeChoice = $00000006;
  xlListDataTypeChoiceMulti = $00000007;
  xlListDataTypeListLookup = $00000008;
  xlListDataTypeCheckbox = $00000009;
  xlListDataTypeHyperLink = $0000000A;
  xlListDataTypeCounter = $0000000B;
  xlListDataTypeMultiLineRichText = $0000000C;


// Constants for enum XlTotalsCalculation
  xlTotalsCalculationNone = $00000000;
  xlTotalsCalculationSum = $00000001;
  xlTotalsCalculationAverage = $00000002;
  xlTotalsCalculationCount = $00000003;
  xlTotalsCalculationCountNums = $00000004;
  xlTotalsCalculationMin = $00000005;
  xlTotalsCalculationMax = $00000006;
  xlTotalsCalculationStdDev = $00000007;
  xlTotalsCalculationVar = $00000008;
  xlTotalsCalculationCustom = $00000009;


// Constants for enum XlXmlLoadOption
  xlXmlLoadPromptUser = $00000000;
  xlXmlLoadOpenXml = $00000001;
  xlXmlLoadImportToList = $00000002;
  xlXmlLoadMapXml = $00000003;


// Constants for enum XlSmartTagControlType
  xlSmartTagControlSmartTag = $00000001;
  xlSmartTagControlLink = $00000002;
  xlSmartTagControlHelp = $00000003;
  xlSmartTagControlHelpURL = $00000004;
  xlSmartTagControlSeparator = $00000005;
  xlSmartTagControlButton = $00000006;
  xlSmartTagControlLabel = $00000007;
  xlSmartTagControlImage = $00000008;
  xlSmartTagControlCheckbox = $00000009;
  xlSmartTagControlTextbox = $0000000A;
  xlSmartTagControlListbox = $0000000B;
  xlSmartTagControlCombo = $0000000C;
  xlSmartTagControlActiveX = $0000000D;
  xlSmartTagControlRadioGroup = $0000000E;


// Constants for enum XlListConflict
  xlListConflictDialog = $00000000;
  xlListConflictRetryAllConflicts = $00000001;
  xlListConflictDiscardAllConflicts = $00000002;
  xlListConflictError = $00000003;


// Constants for enum XlXmlExportResult
  xlXmlExportSuccess = $00000000;
  xlXmlExportValidationFailed = $00000001;


// Constants for enum XlXmlImportResult
  xlXmlImportSuccess = $00000000;
  xlXmlImportElementsTruncated = $00000001;
  xlXmlImportValidationFailed = $00000002;


// Constants for enum XlRemoveDocInfoType
  xlRDIComments = $00000001;
  xlRDIRemovePersonalInformation = $00000004;
  xlRDIEmailHeader = $00000005;
  xlRDIRoutingSlip = $00000006;
  xlRDISendForReview = $00000007;
  xlRDIDocumentProperties = $00000008;
  xlRDIDocumentWorkspace = $0000000A;
  xlRDIInkAnnotations = $0000000B;
  xlRDIScenarioComments = $0000000C;
  xlRDIPublishInfo = $0000000D;
  xlRDIDocumentServerProperties = $0000000E;
  xlRDIDocumentManagementPolicy = $0000000F;
  xlRDIContentType = $00000010;
  xlRDIDefinedNameComments = $00000012;
  xlRDIInactiveDataConnections = $00000013;
  xlRDIPrinterPath = $00000014;
  xlRDIInlineWebExtensions = $00000015;
  xlRDITaskpaneWebExtensions = $00000016;
  xlRDIExcelDataModel = $00000017;
  xlRDIAll = $00000063;


// Constants for enum XlRgbColor
  rgbAliceBlue = $00FFF8F0;
  rgbAntiqueWhite = $00D7EBFA;
  rgbAqua = $00FFFF00;
  rgbAquamarine = $00D4FF7F;
  rgbAzure = $00FFFFF0;
  rgbBeige = $00DCF5F5;
  rgbBisque = $00C4E4FF;
  rgbBlack = $00000000;
  rgbBlanchedAlmond = $00CDEBFF;
  rgbBlue = $00FF0000;
  rgbBlueViolet = $00E22B8A;
  rgbBrown = $002A2AA5;
  rgbBurlyWood = $0087B8DE;
  rgbCadetBlue = $00A09E5F;
  rgbChartreuse = $0000FF7F;
  rgbCoral = $00507FFF;
  rgbCornflowerBlue = $00ED9564;
  rgbCornsilk = $00DCF8FF;
  rgbCrimson = $003C14DC;
  rgbDarkBlue = $008B0000;
  rgbDarkCyan = $008B8B00;
  rgbDarkGoldenrod = $000B86B8;
  rgbDarkGreen = $00006400;
  rgbDarkGray = $00A9A9A9;
  rgbDarkGrey = $00A9A9A9;
  rgbDarkKhaki = $006BB7BD;
  rgbDarkMagenta = $008B008B;
  rgbDarkOliveGreen = $002F6B55;
  rgbDarkOrange = $00008CFF;
  rgbDarkOrchid = $00CC3299;
  rgbDarkRed = $0000008B;
  rgbDarkSalmon = $007A96E9;
  rgbDarkSeaGreen = $008FBC8F;
  rgbDarkSlateBlue = $008B3D48;
  rgbDarkSlateGray = $004F4F2F;
  rgbDarkSlateGrey = $004F4F2F;
  rgbDarkTurquoise = $00D1CE00;
  rgbDarkViolet = $00D30094;
  rgbDeepPink = $009314FF;
  rgbDeepSkyBlue = $00FFBF00;
  rgbDimGray = $00696969;
  rgbDimGrey = $00696969;
  rgbDodgerBlue = $00FF901E;
  rgbFireBrick = $002222B2;
  rgbFloralWhite = $00F0FAFF;
  rgbForestGreen = $00228B22;
  rgbFuchsia = $00FF00FF;
  rgbGainsboro = $00DCDCDC;
  rgbGhostWhite = $00FFF8F8;
  rgbGold = $0000D7FF;
  rgbGoldenrod = $0020A5DA;
  rgbGray = $00808080;
  rgbGreen = $00008000;
  rgbGrey = $00808080;
  rgbGreenYellow = $002FFFAD;
  rgbHoneydew = $00F0FFF0;
  rgbHotPink = $00B469FF;
  rgbIndianRed = $005C5CCD;
  rgbIndigo = $0082004B;
  rgbIvory = $00F0FFFF;
  rgbKhaki = $008CE6F0;
  rgbLavender = $00FAE6E6;
  rgbLavenderBlush = $00F5F0FF;
  rgbLawnGreen = $0000FC7C;
  rgbLemonChiffon = $00CDFAFF;
  rgbLightBlue = $00E6D8AD;
  rgbLightCoral = $008080F0;
  rgbLightCyan = $008B8B00;
  rgbLightGoldenrodYellow = $00D2FAFA;
  rgbLightGray = $00D3D3D3;
  rgbLightGreen = $0090EE90;
  rgbLightGrey = $00D3D3D3;
  rgbLightPink = $00C1B6FF;
  rgbLightSalmon = $007AA0FF;
  rgbLightSeaGreen = $00AAB220;
  rgbLightSkyBlue = $00FACE87;
  rgbLightSlateGray = $00998877;
  rgbLightSlateGrey = $00998877;
  rgbLightSteelBlue = $00DEC4B0;
  rgbLightYellow = $00E0FFFF;
  rgbLime = $0000FF00;
  rgbLimeGreen = $0032CD32;
  rgbLinen = $00E6F0FA;
  rgbMaroon = $00000080;
  rgbMediumAquamarine = $00AAFF66;
  rgbMediumBlue = $00CD0000;
  rgbMediumOrchid = $00D355BA;
  rgbMediumPurple = $00DB7093;
  rgbMediumSeaGreen = $0071B33C;
  rgbMediumSlateBlue = $00EE687B;
  rgbMediumSpringGreen = $009AFA00;
  rgbMediumTurquoise = $00CCD148;
  rgbMediumVioletRed = $008515C7;
  rgbMidnightBlue = $00701919;
  rgbMintCream = $00FAFFF5;
  rgbMistyRose = $00E1E4FF;
  rgbMoccasin = $00B5E4FF;
  rgbNavajoWhite = $00ADDEFF;
  rgbNavy = $00800000;
  rgbNavyBlue = $00800000;
  rgbOldLace = $00E6F5FD;
  rgbOlive = $00008080;
  rgbOliveDrab = $00238E6B;
  rgbOrange = $0000A5FF;
  rgbOrangeRed = $000045FF;
  rgbOrchid = $00D670DA;
  rgbPaleGoldenrod = $006BE8EE;
  rgbPaleGreen = $0098FB98;
  rgbPaleTurquoise = $00EEEEAF;
  rgbPaleVioletRed = $009370DB;
  rgbPapayaWhip = $00D5EFFF;
  rgbPeachPuff = $00B9DAFF;
  rgbPeru = $003F85CD;
  rgbPink = $00CBC0FF;
  rgbPlum = $00DDA0DD;
  rgbPowderBlue = $00E6E0B0;
  rgbPurple = $00800080;
  rgbRed = $000000FF;
  rgbRosyBrown = $008F8FBC;
  rgbRoyalBlue = $00E16941;
  rgbSalmon = $007280FA;
  rgbSandyBrown = $0060A4F4;
  rgbSeaGreen = $00578B2E;
  rgbSeashell = $00EEF5FF;
  rgbSienna = $002D52A0;
  rgbSilver = $00C0C0C0;
  rgbSkyBlue = $00EBCE87;
  rgbSlateBlue = $00CD5A6A;
  rgbSlateGray = $00908070;
  rgbSlateGrey = $00908070;
  rgbSnow = $00FAFAFF;
  rgbSpringGreen = $007FFF00;
  rgbSteelBlue = $00B48246;
  rgbTan = $008CB4D2;
  rgbTeal = $00808000;
  rgbThistle = $00D8BFD8;
  rgbTomato = $004763FF;
  rgbTurquoise = $00D0E040;
  rgbYellow = $0000FFFF;
  rgbYellowGreen = $0032CD9A;
  rgbViolet = $00EE82EE;
  rgbWheat = $00B3DEF5;
  rgbWhite = $00FFFFFF;
  rgbWhiteSmoke = $00F5F5F5;


// Constants for enum XlStdColorScale
  xlColorScaleRYG = $00000001;
  xlColorScaleGYR = $00000002;
  xlColorScaleBlackWhite = $00000003;
  xlColorScaleWhiteBlack = $00000004;


// Constants for enum XlConditionValueTypes
  xlConditionValueNone = $FFFFFFFF;
  xlConditionValueNumber = $00000000;
  xlConditionValueLowestValue = $00000001;
  xlConditionValueHighestValue = $00000002;
  xlConditionValuePercent = $00000003;
  xlConditionValueFormula = $00000004;
  xlConditionValuePercentile = $00000005;
  xlConditionValueAutomaticMin = $00000006;
  xlConditionValueAutomaticMax = $00000007;


// Constants for enum XlFormatFilterTypes
  xlFilterBottom = $00000000;
  xlFilterTop = $00000001;
  xlFilterBottomPercent = $00000002;
  xlFilterTopPercent = $00000003;


// Constants for enum XlContainsOperator
  xlContains = $00000000;
  xlDoesNotContain = $00000001;
  xlBeginsWith = $00000002;
  xlEndsWith = $00000003;


// Constants for enum XlAboveBelow
  xlAboveAverage = $00000000;
  xlBelowAverage = $00000001;
  xlEqualAboveAverage = $00000002;
  xlEqualBelowAverage = $00000003;
  xlAboveStdDev = $00000004;
  xlBelowStdDev = $00000005;


// Constants for enum XlLookFor
  xlLookForBlanks = $00000000;
  xlLookForErrors = $00000001;
  xlLookForFormulas = $00000002;


// Constants for enum XlTimePeriods
  xlToday = $00000000;
  xlYesterday = $00000001;
  xlLast7Days = $00000002;
  xlThisWeek = $00000003;
  xlLastWeek = $00000004;
  xlLastMonth = $00000005;
  xlTomorrow = $00000006;
  xlNextWeek = $00000007;
  xlNextMonth = $00000008;
  xlThisMonth = $00000009;


// Constants for enum XlDupeUnique
  xlUnique = $00000000;
  xlDuplicate = $00000001;


// Constants for enum XlTopBottom
  xlTop10Top = $00000001;
  xlTop10Bottom = $00000000;


// Constants for enum XlIconSet
  xlCustomSet = $FFFFFFFF;
  xl3Arrows = $00000001;
  xl3ArrowsGray = $00000002;
  xl3Flags = $00000003;
  xl3TrafficLights1 = $00000004;
  xl3TrafficLights2 = $00000005;
  xl3Signs = $00000006;
  xl3Symbols = $00000007;
  xl3Symbols2 = $00000008;
  xl4Arrows = $00000009;
  xl4ArrowsGray = $0000000A;
  xl4RedToBlack = $0000000B;
  xl4CRV = $0000000C;
  xl4TrafficLights = $0000000D;
  xl5Arrows = $0000000E;
  xl5ArrowsGray = $0000000F;
  xl5CRV = $00000010;
  xl5Quarters = $00000011;
  xl3Stars = $00000012;
  xl3Triangles = $00000013;
  xl5Boxes = $00000014;


// Constants for enum XlThemeFont
  xlThemeFontNone = $00000000;
  xlThemeFontMajor = $00000001;
  xlThemeFontMinor = $00000002;


// Constants for enum XlPivotLineType
  xlPivotLineRegular = $00000000;
  xlPivotLineSubtotal = $00000001;
  xlPivotLineGrandTotal = $00000002;
  xlPivotLineBlank = $00000003;


// Constants for enum XlCheckInVersionType
  xlCheckInMinorVersion = $00000000;
  xlCheckInMajorVersion = $00000001;
  xlCheckInOverwriteVersion = $00000002;


// Constants for enum XlPropertyDisplayedIn
  xlDisplayPropertyInPivotTable = $00000001;
  xlDisplayPropertyInTooltip = $00000002;
  xlDisplayPropertyInPivotTableAndTooltip = $00000003;


// Constants for enum XlConnectionType
  xlConnectionTypeOLEDB = $00000001;
  xlConnectionTypeODBC = $00000002;
  xlConnectionTypeXMLMAP = $00000003;
  xlConnectionTypeTEXT = $00000004;
  xlConnectionTypeWEB = $00000005;
  xlConnectionTypeDATAFEED = $00000006;
  xlConnectionTypeMODEL = $00000007;
  xlConnectionTypeWORKSHEET = $00000008;
  xlConnectionTypeNOSOURCE = $00000009;


// Constants for enum XlActionType
  xlActionTypeUrl = $00000001;
  xlActionTypeRowset = $00000010;
  xlActionTypeReport = $00000080;
  xlActionTypeDrillthrough = $00000100;


// Constants for enum XlLayoutRowType
  xlCompactRow = $00000000;
  xlTabularRow = $00000001;
  xlOutlineRow = $00000002;


// Constants for enum XlMeasurementUnits
  xlInches = $00000000;
  xlCentimeters = $00000001;
  xlMillimeters = $00000002;


// Constants for enum XlPivotFilterType
  xlTopCount = $00000001;
  xlBottomCount = $00000002;
  xlTopPercent = $00000003;
  xlBottomPercent = $00000004;
  xlTopSum = $00000005;
  xlBottomSum = $00000006;
  xlValueEquals = $00000007;
  xlValueDoesNotEqual = $00000008;
  xlValueIsGreaterThan = $00000009;
  xlValueIsGreaterThanOrEqualTo = $0000000A;
  xlValueIsLessThan = $0000000B;
  xlValueIsLessThanOrEqualTo = $0000000C;
  xlValueIsBetween = $0000000D;
  xlValueIsNotBetween = $0000000E;
  xlCaptionEquals = $0000000F;
  xlCaptionDoesNotEqual = $00000010;
  xlCaptionBeginsWith = $00000011;
  xlCaptionDoesNotBeginWith = $00000012;
  xlCaptionEndsWith = $00000013;
  xlCaptionDoesNotEndWith = $00000014;
  xlCaptionContains = $00000015;
  xlCaptionDoesNotContain = $00000016;
  xlCaptionIsGreaterThan = $00000017;
  xlCaptionIsGreaterThanOrEqualTo = $00000018;
  xlCaptionIsLessThan = $00000019;
  xlCaptionIsLessThanOrEqualTo = $0000001A;
  xlCaptionIsBetween = $0000001B;
  xlCaptionIsNotBetween = $0000001C;
  xlSpecificDate = $0000001D;
  xlNotSpecificDate = $0000001E;
  xlBefore = $0000001F;
  xlBeforeOrEqualTo = $00000020;
  xlAfter = $00000021;
  xlAfterOrEqualTo = $00000022;
  xlDateBetween = $00000023;
  xlDateNotBetween = $00000024;
  xlDateTomorrow = $00000025;
  xlDateToday = $00000026;
  xlDateYesterday = $00000027;
  xlDateNextWeek = $00000028;
  xlDateThisWeek = $00000029;
  xlDateLastWeek = $0000002A;
  xlDateNextMonth = $0000002B;
  xlDateThisMonth = $0000002C;
  xlDateLastMonth = $0000002D;
  xlDateNextQuarter = $0000002E;
  xlDateThisQuarter = $0000002F;
  xlDateLastQuarter = $00000030;
  xlDateNextYear = $00000031;
  xlDateThisYear = $00000032;
  xlDateLastYear = $00000033;
  xlYearToDate = $00000034;
  xlAllDatesInPeriodQuarter1 = $00000035;
  xlAllDatesInPeriodQuarter2 = $00000036;
  xlAllDatesInPeriodQuarter3 = $00000037;
  xlAllDatesInPeriodQuarter4 = $00000038;
  xlAllDatesInPeriodJanuary = $00000039;
  xlAllDatesInPeriodFebruary = $0000003A;
  xlAllDatesInPeriodMarch = $0000003B;
  xlAllDatesInPeriodApril = $0000003C;
  xlAllDatesInPeriodMay = $0000003D;
  xlAllDatesInPeriodJune = $0000003E;
  xlAllDatesInPeriodJuly = $0000003F;
  xlAllDatesInPeriodAugust = $00000040;
  xlAllDatesInPeriodSeptember = $00000041;
  xlAllDatesInPeriodOctober = $00000042;
  xlAllDatesInPeriodNovember = $00000043;
  xlAllDatesInPeriodDecember = $00000044;


// Constants for enum XlCredentialsMethod
  xlCredentialsMethodIntegrated = $00000000;
  xlCredentialsMethodNone = $00000001;
  xlCredentialsMethodStored = $00000002;


// Constants for enum XlCubeFieldSubType
  xlCubeHierarchy = $00000001;
  xlCubeMeasure = $00000002;
  xlCubeSet = $00000003;
  xlCubeAttribute = $00000004;
  xlCubeCalculatedMeasure = $00000005;
  xlCubeKPIValue = $00000006;
  xlCubeKPIGoal = $00000007;
  xlCubeKPIStatus = $00000008;
  xlCubeKPITrend = $00000009;
  xlCubeKPIWeight = $0000000A;
  xlCubeImplicitMeasure = $0000000B;


// Constants for enum XlSortOn
  xlSortOnValues = $00000000;
  xlSortOnCellColor = $00000001;
  xlSortOnFontColor = $00000002;
  xlSortOnIcon = $00000003;


// Constants for enum XlDynamicFilterCriteria
  xlFilterToday = $00000001;
  xlFilterYesterday = $00000002;
  xlFilterTomorrow = $00000003;
  xlFilterThisWeek = $00000004;
  xlFilterLastWeek = $00000005;
  xlFilterNextWeek = $00000006;
  xlFilterThisMonth = $00000007;
  xlFilterLastMonth = $00000008;
  xlFilterNextMonth = $00000009;
  xlFilterThisQuarter = $0000000A;
  xlFilterLastQuarter = $0000000B;
  xlFilterNextQuarter = $0000000C;
  xlFilterThisYear = $0000000D;
  xlFilterLastYear = $0000000E;
  xlFilterNextYear = $0000000F;
  xlFilterYearToDate = $00000010;
  xlFilterAllDatesInPeriodQuarter1 = $00000011;
  xlFilterAllDatesInPeriodQuarter2 = $00000012;
  xlFilterAllDatesInPeriodQuarter3 = $00000013;
  xlFilterAllDatesInPeriodQuarter4 = $00000014;
  xlFilterAllDatesInPeriodJanuary = $00000015;
  xlFilterAllDatesInPeriodFebruray = $00000016;
  xlFilterAllDatesInPeriodMarch = $00000017;
  xlFilterAllDatesInPeriodApril = $00000018;
  xlFilterAllDatesInPeriodMay = $00000019;
  xlFilterAllDatesInPeriodJune = $0000001A;
  xlFilterAllDatesInPeriodJuly = $0000001B;
  xlFilterAllDatesInPeriodAugust = $0000001C;
  xlFilterAllDatesInPeriodSeptember = $0000001D;
  xlFilterAllDatesInPeriodOctober = $0000001E;
  xlFilterAllDatesInPeriodNovember = $0000001F;
  xlFilterAllDatesInPeriodDecember = $00000020;
  xlFilterAboveAverage = $00000021;
  xlFilterBelowAverage = $00000022;


// Constants for enum XlFilterAllDatesInPeriod
  xlFilterAllDatesInPeriodYear = $00000000;
  xlFilterAllDatesInPeriodMonth = $00000001;
  xlFilterAllDatesInPeriodDay = $00000002;
  xlFilterAllDatesInPeriodHour = $00000003;
  xlFilterAllDatesInPeriodMinute = $00000004;
  xlFilterAllDatesInPeriodSecond = $00000005;


// Constants for enum XlTableStyleElementType
  xlWholeTable = $00000000;
  xlHeaderRow = $00000001;
  xlTotalRow = $00000002;
  xlGrandTotalRow = $00000002;
  xlFirstColumn = $00000003;
  xlLastColumn = $00000004;
  xlGrandTotalColumn = $00000004;
  xlRowStripe1 = $00000005;
  xlRowStripe2 = $00000006;
  xlColumnStripe1 = $00000007;
  xlColumnStripe2 = $00000008;
  xlFirstHeaderCell = $00000009;
  xlLastHeaderCell = $0000000A;
  xlFirstTotalCell = $0000000B;
  xlLastTotalCell = $0000000C;
  xlSubtotalColumn1 = $0000000D;
  xlSubtotalColumn2 = $0000000E;
  xlSubtotalColumn3 = $0000000F;
  xlSubtotalRow1 = $00000010;
  xlSubtotalRow2 = $00000011;
  xlSubtotalRow3 = $00000012;
  xlBlankRow = $00000013;
  xlColumnSubheading1 = $00000014;
  xlColumnSubheading2 = $00000015;
  xlColumnSubheading3 = $00000016;
  xlRowSubheading1 = $00000017;
  xlRowSubheading2 = $00000018;
  xlRowSubheading3 = $00000019;
  xlPageFieldLabels = $0000001A;
  xlPageFieldValues = $0000001B;
  xlSlicerUnselectedItemWithData = $0000001C;
  xlSlicerUnselectedItemWithNoData = $0000001D;
  xlSlicerSelectedItemWithData = $0000001E;
  xlSlicerSelectedItemWithNoData = $0000001F;
  xlSlicerHoveredUnselectedItemWithData = $00000020;
  xlSlicerHoveredSelectedItemWithData = $00000021;
  xlSlicerHoveredUnselectedItemWithNoData = $00000022;
  xlSlicerHoveredSelectedItemWithNoData = $00000023;
  xlTimelineSelectionLabel = $00000024;
  xlTimelineTimeLevel = $00000025;
  xlTimelinePeriodLabels1 = $00000026;
  xlTimelinePeriodLabels2 = $00000027;
  xlTimelineSelectedTimeBlock = $00000028;
  xlTimelineUnselectedTimeBlock = $00000029;
  xlTimelineSelectedTimeBlockSpace = $0000002A;


// Constants for enum XlPivotConditionScope
  xlSelectionScope = $00000000;
  xlFieldsScope = $00000001;
  xlDataFieldScope = $00000002;


// Constants for enum XlCalcFor
  xlAllValues = $00000000;
  xlRowGroups = $00000001;
  xlColGroups = $00000002;


// Constants for enum XlThemeColor
  xlThemeColorDark1 = $00000001;
  xlThemeColorLight1 = $00000002;
  xlThemeColorDark2 = $00000003;
  xlThemeColorLight2 = $00000004;
  xlThemeColorAccent1 = $00000005;
  xlThemeColorAccent2 = $00000006;
  xlThemeColorAccent3 = $00000007;
  xlThemeColorAccent4 = $00000008;
  xlThemeColorAccent5 = $00000009;
  xlThemeColorAccent6 = $0000000A;
  xlThemeColorHyperlink = $0000000B;
  xlThemeColorFollowedHyperlink = $0000000C;


// Constants for enum XlFixedFormatType
// Modified.
  XlFixedFormatType_xlTypePDF = $00000000;
  XlFixedFormatType_xlTypeXPS = $00000001;


// Constants for enum XlFixedFormatQuality
  xlQualityStandard = $00000000;
  xlQualityMinimum = $00000001;


// Constants for enum XlChartElementPosition
  xlChartElementPositionAutomatic = $FFFFEFF7;
  xlChartElementPositionCustom = $FFFFEFEE;


// Constants for enum XlGenerateTableRefs
  xlGenerateTableRefA1 = $00000000;
  xlGenerateTableRefStruct = $00000001;


// Constants for enum XlGradientFillType
  xlGradientFillLinear = $00000000;
  xlGradientFillPath = $00000001;


// Constants for enum XlThreadMode
  xlThreadModeAutomatic = $00000000;
  xlThreadModeManual = $00000001;


// Constants for enum XlOartHorizontalOverflow
  xlOartHorizontalOverflowOverflow = $00000000;
  xlOartHorizontalOverflowClip = $00000001;


// Constants for enum XlOartVerticalOverflow
  xlOartVerticalOverflowOverflow = $00000000;
  xlOartVerticalOverflowClip = $00000001;
  xlOartVerticalOverflowEllipsis = $00000002;


// Constants for enum XlSparkScale
  xlSparkScaleGroup = $00000001;
  xlSparkScaleSingle = $00000002;
  xlSparkScaleCustom = $00000003;


// Constants for enum XlSparkType
  xlSparkLine = $00000001;
  xlSparkColumn = $00000002;
  xlSparkColumnStacked100 = $00000003;


// Constants for enum XlSparklineRowCol
  xlSparklineNonSquare = $00000000;
  xlSparklineRowsSquare = $00000001;
  xlSparklineColumnsSquare = $00000002;


// Constants for enum XlDataBarFillType
  xlDataBarFillSolid = $00000000;
  xlDataBarFillGradient = $00000001;


// Constants for enum XlDataBarBorderType
  xlDataBarBorderNone = $00000000;
  xlDataBarBorderSolid = $00000001;


// Constants for enum XlDataBarAxisPosition
  xlDataBarAxisAutomatic = $00000000;
  xlDataBarAxisMidpoint = $00000001;
  xlDataBarAxisNone = $00000002;


// Constants for enum XlDataBarNegativeColorType
  xlDataBarColor = $00000000;
  xlDataBarSameAsPositive = $00000001;


// Constants for enum XlAllocation
  xlManualAllocation = $00000001;
  xlAutomaticAllocation = $00000002;


// Constants for enum XlAllocationValue
  xlAllocateValue = $00000001;
  xlAllocateIncrement = $00000002;


// Constants for enum XlAllocationMethod
  xlEqualAllocation = $00000001;
  xlWeightedAllocation = $00000002;


// Constants for enum XlCellChangedState
  xlCellNotChanged = $00000001;
  xlCellChanged = $00000002;
  xlCellChangeApplied = $00000003;


// Constants for enum XlPivotFieldRepeatLabels
  xlDoNotRepeatLabels = $00000001;
  xlRepeatLabels = $00000002;


// Constants for enum XlPieSliceIndex
  xlOuterCounterClockwisePoint = $00000001;
  xlOuterCenterPoint = $00000002;
  xlOuterClockwisePoint = $00000003;
  xlMidClockwiseRadiusPoint = $00000004;
  xlCenterPoint = $00000005;
  xlMidCounterClockwiseRadiusPoint = $00000006;
  xlInnerClockwisePoint = $00000007;
  xlInnerCenterPoint = $00000008;
  xlInnerCounterClockwisePoint = $00000009;


// Constants for enum XlSpanishModes
  xlSpanishTuteoOnly = $00000000;
  xlSpanishTuteoAndVoseo = $00000001;
  xlSpanishVoseoOnly = $00000002;


// Constants for enum XlSlicerCrossFilterType
  xlSlicerNoCrossFilter = $00000001;
  xlSlicerCrossFilterShowItemsWithDataAtTop = $00000002;
  xlSlicerCrossFilterShowItemsWithNoData = $00000003;
  xlSlicerCrossFilterHideButtonsWithNoData = $00000004;


// Constants for enum XlSlicerSort
  xlSlicerSortDataSourceOrder = $00000001;
  xlSlicerSortAscending = $00000002;
  xlSlicerSortDescending = $00000003;


// Constants for enum XlIcon
  xlIconNoCellIcon = $FFFFFFFF;
  xlIconGreenUpArrow = $00000001;
  xlIconYellowSideArrow = $00000002;
  xlIconRedDownArrow = $00000003;
  xlIconGrayUpArrow = $00000004;
  xlIconGraySideArrow = $00000005;
  xlIconGrayDownArrow = $00000006;
  xlIconGreenFlag = $00000007;
  xlIconYellowFlag = $00000008;
  xlIconRedFlag = $00000009;
  xlIconGreenCircle = $0000000A;
  xlIconYellowCircle = $0000000B;
  xlIconRedCircleWithBorder = $0000000C;
  xlIconBlackCircleWithBorder = $0000000D;
  xlIconGreenTrafficLight = $0000000E;
  xlIconYellowTrafficLight = $0000000F;
  xlIconRedTrafficLight = $00000010;
  xlIconYellowTriangle = $00000011;
  xlIconRedDiamond = $00000012;
  xlIconGreenCheckSymbol = $00000013;
  xlIconYellowExclamationSymbol = $00000014;
  xlIconRedCrossSymbol = $00000015;
  xlIconGreenCheck = $00000016;
  xlIconYellowExclamation = $00000017;
  xlIconRedCross = $00000018;
  xlIconYellowUpInclineArrow = $00000019;
  xlIconYellowDownInclineArrow = $0000001A;
  xlIconGrayUpInclineArrow = $0000001B;
  xlIconGrayDownInclineArrow = $0000001C;
  xlIconRedCircle = $0000001D;
  xlIconPinkCircle = $0000001E;
  xlIconGrayCircle = $0000001F;
  xlIconBlackCircle = $00000020;
  xlIconCircleWithOneWhiteQuarter = $00000021;
  xlIconCircleWithTwoWhiteQuarters = $00000022;
  xlIconCircleWithThreeWhiteQuarters = $00000023;
  xlIconWhiteCircleAllWhiteQuarters = $00000024;
  xlIcon0Bars = $00000025;
  xlIcon1Bar = $00000026;
  xlIcon2Bars = $00000027;
  xlIcon3Bars = $00000028;
  xlIcon4Bars = $00000029;
  xlIconGoldStar = $0000002A;
  xlIconHalfGoldStar = $0000002B;
  xlIconSilverStar = $0000002C;
  xlIconGreenUpTriangle = $0000002D;
  xlIconYellowDash = $0000002E;
  xlIconRedDownTriangle = $0000002F;
  xlIcon4FilledBoxes = $00000030;
  xlIcon3FilledBoxes = $00000031;
  xlIcon2FilledBoxes = $00000032;
  xlIcon1FilledBox = $00000033;
  xlIcon0FilledBoxes = $00000034;


// Constants for enum XlProtectedViewCloseReason
  xlProtectedViewCloseNormal = $00000000;
  xlProtectedViewCloseEdit = $00000001;
  xlProtectedViewCloseForced = $00000002;


// Constants for enum XlProtectedViewWindowState
  xlProtectedViewWindowNormal = $00000000;
  xlProtectedViewWindowMinimized = $00000001;
  xlProtectedViewWindowMaximized = $00000002;


// Constants for enum XlFileValidationPivotMode
  xlFileValidationPivotDefault = $00000000;
  xlFileValidationPivotRun = $00000001;
  xlFileValidationPivotSkip = $00000002;


// Constants for enum XlPieSliceLocation
  xlHorizontalCoordinate = $00000001;
  xlVerticalCoordinate = $00000002;


// Constants for enum XlPortugueseReform
  xlPortuguesePreReform = $00000001;
  xlPortuguesePostReform = $00000002;
  xlPortugueseBoth = $00000003;


// Constants for enum XlQuickAnalysisMode
  xlLensOnly = $00000000;
  xlFormatConditions = $00000001;
  xlRecommendedCharts = $00000002;
  xlTotals = $00000003;
  xlTables = $00000004;
  xlSparklines = $00000005;


// Constants for enum XlSlicerCacheType
  xlSlicer = $00000001;
  xlTimeline = $00000002;


// Constants for enum XlCategoryLabelLevel
  xlCategoryLabelLevelNone = $FFFFFFFD;
  xlCategoryLabelLevelCustom = $FFFFFFFE;
  xlCategoryLabelLevelAll = $FFFFFFFF;


// Constants for enum XlSeriesNameLevel
  xlSeriesNameLevelNone = $FFFFFFFD;
  xlSeriesNameLevelCustom = $FFFFFFFE;
  xlSeriesNameLevelAll = $FFFFFFFF;


// Constants for enum XlCalcMemNumberFormatType
  xlNumberFormatTypeDefault = $00000000;
  xlNumberFormatTypeNumber = $00000001;
  xlNumberFormatTypePercent = $00000002;


// Constants for enum XlTimelineLevel
  xlTimelineLevelYears = $00000000;
  xlTimelineLevelQuarters = $00000001;
  xlTimelineLevelMonths = $00000002;
  xlTimelineLevelDays = $00000003;


// Constants for enum XlFilterStatus
  xlFilterStatusOK = $00000000;
  xlFilterStatusDateWrongOrder = $00000001;
  xlFilterStatusDateHasTime = $00000002;
  xlFilterStatusInvalidDate = $00000003;


// Constants for enum XlModelChangeSource
  xlChangeByExcel = $00000000;
  xlChangeByPowerPivotAddIn = $00000001;


// Constants for enum XlParentDataLabelOptions
  xlParentDataLabelOptionsNone = $00000000;
  xlParentDataLabelOptionsBanner = $00000001;
  xlParentDataLabelOptionsOverlapping = $00000002;


// Constants for enum XlBinsType
  xlBinsTypeAutomatic = $00000000;
  xlBinsTypeCategorical = $00000001;
  xlBinsTypeManual = $00000002;
  xlBinsTypeBinSize = $00000003;
  xlBinsTypeBinCount = $00000004;


// Constants for enum XlForecastDataCompletion
  xlForecastDataCompletionZeros = $00000000;
  xlForecastDataCompletionInterpolate = $00000001;


// Constants for enum XlForecastAggregation
  xlForecastAggregationAverage = $00000001;
  xlForecastAggregationCount = $00000002;
  xlForecastAggregationCountA = $00000003;
  xlForecastAggregationMax = $00000004;
  xlForecastAggregationMedian = $00000005;
  xlForecastAggregationMin = $00000006;
  xlForecastAggregationSum = $00000007;


// Constants for enum XlForecastChartType
  xlForecastChartTypeLine = $00000000;
  xlForecastChartTypeColumn = $00000001;


// Constants for enum XlPublishToDocsDisclosureScope
  msoPublic = $00000000;
  msoLimited = $00000001;
  msoOrganization = $00000002;
  msoNoOverwrite = $00000003;


// Constants for enum XlCategorySortOrder
  xlIndexAscending = $00000000;
  xlIndexDescending = $00000001;
  xlCategoryAscending = $00000002;
  xlCategoryDescending = $00000003;


// Constants for enum XlValueSortOrder
  xlValueNone = $00000000;
  xlValueAscending = $00000001;
  xlValueDescending = $00000002;


// Constants for enum XlGeoProjectionType
  xlGeoProjectionTypeAutomatic = $00000000;
  xlGeoProjectionTypeMercator = $00000001;
  xlGeoProjectionTypeMiller = $00000002;
  xlGeoProjectionTypeAlbers = $00000003;
  xlGeoProjectionTypeRobinson = $00000004;


// Constants for enum XlGeoMappingLevel
  xlGeoMappingLevelAutomatic = $00000000;
  xlGeoMappingLevelDataOnly = $00000001;
  xlGeoMappingLevelPostalCode = $00000002;
  xlGeoMappingLevelCounty = $00000003;
  xlGeoMappingLevelState = $00000004;
  xlGeoMappingLevelCountryRegion = $00000005;
  xlGeoMappingLevelCountryRegionList = $00000006;
  xlGeoMappingLevelWorld = $00000007;


// Constants for enum XlRegionLabelOptions
  xlRegionLabelOptionsNone = $00000000;
  xlRegionLabelOptionsBestFitOnly = $00000001;
  xlRegionLabelOptionsShowAll = $00000002;


// Constants for enum XlPublishToPBIPublishType
  msoPBIExport = $00000000;
  msoPBIUpload = $00000001;


// Constants for enum XlPublishToPBINameConflictAction
  msoPBIIgnore = $00000000;
  msoPBIAbort = $00000001;
  msoPBIOverwrite = $00000002;


// Constants for enum XlSeriesColorGradientStyle
  xlSeriesColorGradientStyleSequential = $00000000;
  xlSeriesColorGradientStyleDiverging = $00000001;


// Constants for enum XlGradientStopPositionType
  xlGradientStopPositionTypeExtremeValue = $00000000;
  xlGradientStopPositionTypeNumber = $00000001;
  xlGradientStopPositionTypePercent = $00000002;


// Constants for enum XlLinkedDataTypeState
  xlLinkedDataTypeStateNone = $00000000;
  xlLinkedDataTypeStateValidLinkedData = $00000001;
  xlLinkedDataTypeStateDisambiguationNeeded = $00000002;
  xlLinkedDataTypeStateBrokenLinkedData = $00000003;
  xlLinkedDataTypeStateFetchingData = $00000004;


// Constants for enum XlFormulaVersion
  xlReplaceFormula = $00000000;
  xlReplaceFormula2 = $00000001;

implementation

end.