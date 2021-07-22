/* PptxGenJS 3.7.1 @ 2021-07-22T03:01:25.399Z */
import JSZip from 'jszip';

/*! *****************************************************************************
Copyright (c) Microsoft Corporation.

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH
REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY
AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,
INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM
LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR
OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR
PERFORMANCE OF THIS SOFTWARE.
***************************************************************************** */

var __assign = function() {
    __assign = Object.assign || function __assign(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};

function __spreadArray(to, from) {
    for (var i = 0, il = from.length, j = to.length; i < il; i++, j++)
        to[j] = from[i];
    return to;
}

/**
 * PptxGenJS Enums
 * NOTE: `enum` wont work for objects, so use `Object.freeze`
 */
// CONST
var EMU = 914400; // One (1) inch (OfficeXML measures in EMU (English Metric Units))
var ONEPT = 12700; // One (1) point (pt)
var CRLF = '\r\n'; // AKA: Chr(13) & Chr(10)
var LAYOUT_IDX_SERIES_BASE = 2147483649;
var REGEX_HEX_COLOR = /^[0-9a-fA-F]{6}$/;
var LINEH_MODIFIER = 1.67; // AKA: Golden Ratio Typography
var DEF_BULLET_MARGIN = 27;
var DEF_CELL_BORDER = { type: 'solid', color: '666666', pt: 1 };
var DEF_CELL_MARGIN_PT = [3, 3, 3, 3]; // TRBL-style
var DEF_CHART_GRIDLINE = { color: '888888', style: 'solid', size: 1 };
var DEF_FONT_COLOR = '000000';
var DEF_FONT_SIZE = 12;
var DEF_FONT_TITLE_SIZE = 18;
var DEF_PRES_LAYOUT = 'LAYOUT_16x9';
var DEF_PRES_LAYOUT_NAME = 'DEFAULT';
var DEF_SHAPE_LINE_COLOR = '333333';
var DEF_SHAPE_SHADOW = { type: 'outer', blur: 3, offset: 23000 / 12700, angle: 90, color: '000000', opacity: 0.35, rotateWithShape: true };
var DEF_SLIDE_MARGIN_IN = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
var DEF_TEXT_SHADOW = { type: 'outer', blur: 8, offset: 4, angle: 270, color: '000000', opacity: 0.75 };
var DEF_TEXT_GLOW = { size: 8, color: 'FFFFFF', opacity: 0.75 };
var AXIS_ID_VALUE_PRIMARY = '2094734552';
var AXIS_ID_VALUE_SECONDARY = '2094734553';
var AXIS_ID_CATEGORY_PRIMARY = '2094734554';
var AXIS_ID_CATEGORY_SECONDARY = '2094734555';
var AXIS_ID_SERIES_PRIMARY = '2094734556';
var LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
var BARCHART_COLORS = [
    'C0504D',
    '4F81BD',
    '9BBB59',
    '8064A2',
    '4BACC6',
    'F79646',
    '628FC6',
    'C86360',
    'C0504D',
    '4F81BD',
    '9BBB59',
    '8064A2',
    '4BACC6',
    'F79646',
    '628FC6',
    'C86360',
];
var PIECHART_COLORS = [
    '5DA5DA',
    'FAA43A',
    '60BD68',
    'F17CB0',
    'B2912F',
    'B276B2',
    'DECF3F',
    'F15854',
    'A7A7A7',
    '5DA5DA',
    'FAA43A',
    '60BD68',
    'F17CB0',
    'B2912F',
    'B276B2',
    'DECF3F',
    'F15854',
    'A7A7A7',
];
var TEXT_HALIGN;
(function (TEXT_HALIGN) {
    TEXT_HALIGN["left"] = "left";
    TEXT_HALIGN["center"] = "center";
    TEXT_HALIGN["right"] = "right";
    TEXT_HALIGN["justify"] = "justify";
})(TEXT_HALIGN || (TEXT_HALIGN = {}));
var TEXT_VALIGN;
(function (TEXT_VALIGN) {
    TEXT_VALIGN["b"] = "b";
    TEXT_VALIGN["ctr"] = "ctr";
    TEXT_VALIGN["t"] = "t";
})(TEXT_VALIGN || (TEXT_VALIGN = {}));
var SLDNUMFLDID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}';
// ENUM
// TODO: 3.5 or v4.0: rationalize ts-def exported enum names/case!
// NOTE: First tsdef enum named correctly (shapes -> 'Shape', colors -> 'Color'), etc.
var OutputType;
(function (OutputType) {
    OutputType["arraybuffer"] = "arraybuffer";
    OutputType["base64"] = "base64";
    OutputType["binarystring"] = "binarystring";
    OutputType["blob"] = "blob";
    OutputType["nodebuffer"] = "nodebuffer";
    OutputType["uint8array"] = "uint8array";
})(OutputType || (OutputType = {}));
var ChartType;
(function (ChartType) {
    ChartType["area"] = "area";
    ChartType["bar"] = "bar";
    ChartType["bar3d"] = "bar3D";
    ChartType["bubble"] = "bubble";
    ChartType["doughnut"] = "doughnut";
    ChartType["line"] = "line";
    ChartType["pie"] = "pie";
    ChartType["radar"] = "radar";
    ChartType["scatter"] = "scatter";
})(ChartType || (ChartType = {}));
var ShapeType;
(function (ShapeType) {
    ShapeType["accentBorderCallout1"] = "accentBorderCallout1";
    ShapeType["accentBorderCallout2"] = "accentBorderCallout2";
    ShapeType["accentBorderCallout3"] = "accentBorderCallout3";
    ShapeType["accentCallout1"] = "accentCallout1";
    ShapeType["accentCallout2"] = "accentCallout2";
    ShapeType["accentCallout3"] = "accentCallout3";
    ShapeType["actionButtonBackPrevious"] = "actionButtonBackPrevious";
    ShapeType["actionButtonBeginning"] = "actionButtonBeginning";
    ShapeType["actionButtonBlank"] = "actionButtonBlank";
    ShapeType["actionButtonDocument"] = "actionButtonDocument";
    ShapeType["actionButtonEnd"] = "actionButtonEnd";
    ShapeType["actionButtonForwardNext"] = "actionButtonForwardNext";
    ShapeType["actionButtonHelp"] = "actionButtonHelp";
    ShapeType["actionButtonHome"] = "actionButtonHome";
    ShapeType["actionButtonInformation"] = "actionButtonInformation";
    ShapeType["actionButtonMovie"] = "actionButtonMovie";
    ShapeType["actionButtonReturn"] = "actionButtonReturn";
    ShapeType["actionButtonSound"] = "actionButtonSound";
    ShapeType["arc"] = "arc";
    ShapeType["bentArrow"] = "bentArrow";
    ShapeType["bentUpArrow"] = "bentUpArrow";
    ShapeType["bevel"] = "bevel";
    ShapeType["blockArc"] = "blockArc";
    ShapeType["borderCallout1"] = "borderCallout1";
    ShapeType["borderCallout2"] = "borderCallout2";
    ShapeType["borderCallout3"] = "borderCallout3";
    ShapeType["bracePair"] = "bracePair";
    ShapeType["bracketPair"] = "bracketPair";
    ShapeType["callout1"] = "callout1";
    ShapeType["callout2"] = "callout2";
    ShapeType["callout3"] = "callout3";
    ShapeType["can"] = "can";
    ShapeType["chartPlus"] = "chartPlus";
    ShapeType["chartStar"] = "chartStar";
    ShapeType["chartX"] = "chartX";
    ShapeType["chevron"] = "chevron";
    ShapeType["chord"] = "chord";
    ShapeType["circularArrow"] = "circularArrow";
    ShapeType["cloud"] = "cloud";
    ShapeType["cloudCallout"] = "cloudCallout";
    ShapeType["corner"] = "corner";
    ShapeType["cornerTabs"] = "cornerTabs";
    ShapeType["cube"] = "cube";
    ShapeType["curvedDownArrow"] = "curvedDownArrow";
    ShapeType["curvedLeftArrow"] = "curvedLeftArrow";
    ShapeType["curvedRightArrow"] = "curvedRightArrow";
    ShapeType["curvedUpArrow"] = "curvedUpArrow";
    ShapeType["custGeom"] = "custGeom";
    ShapeType["decagon"] = "decagon";
    ShapeType["diagStripe"] = "diagStripe";
    ShapeType["diamond"] = "diamond";
    ShapeType["dodecagon"] = "dodecagon";
    ShapeType["donut"] = "donut";
    ShapeType["doubleWave"] = "doubleWave";
    ShapeType["downArrow"] = "downArrow";
    ShapeType["downArrowCallout"] = "downArrowCallout";
    ShapeType["ellipse"] = "ellipse";
    ShapeType["ellipseRibbon"] = "ellipseRibbon";
    ShapeType["ellipseRibbon2"] = "ellipseRibbon2";
    ShapeType["flowChartAlternateProcess"] = "flowChartAlternateProcess";
    ShapeType["flowChartCollate"] = "flowChartCollate";
    ShapeType["flowChartConnector"] = "flowChartConnector";
    ShapeType["flowChartDecision"] = "flowChartDecision";
    ShapeType["flowChartDelay"] = "flowChartDelay";
    ShapeType["flowChartDisplay"] = "flowChartDisplay";
    ShapeType["flowChartDocument"] = "flowChartDocument";
    ShapeType["flowChartExtract"] = "flowChartExtract";
    ShapeType["flowChartInputOutput"] = "flowChartInputOutput";
    ShapeType["flowChartInternalStorage"] = "flowChartInternalStorage";
    ShapeType["flowChartMagneticDisk"] = "flowChartMagneticDisk";
    ShapeType["flowChartMagneticDrum"] = "flowChartMagneticDrum";
    ShapeType["flowChartMagneticTape"] = "flowChartMagneticTape";
    ShapeType["flowChartManualInput"] = "flowChartManualInput";
    ShapeType["flowChartManualOperation"] = "flowChartManualOperation";
    ShapeType["flowChartMerge"] = "flowChartMerge";
    ShapeType["flowChartMultidocument"] = "flowChartMultidocument";
    ShapeType["flowChartOfflineStorage"] = "flowChartOfflineStorage";
    ShapeType["flowChartOffpageConnector"] = "flowChartOffpageConnector";
    ShapeType["flowChartOnlineStorage"] = "flowChartOnlineStorage";
    ShapeType["flowChartOr"] = "flowChartOr";
    ShapeType["flowChartPredefinedProcess"] = "flowChartPredefinedProcess";
    ShapeType["flowChartPreparation"] = "flowChartPreparation";
    ShapeType["flowChartProcess"] = "flowChartProcess";
    ShapeType["flowChartPunchedCard"] = "flowChartPunchedCard";
    ShapeType["flowChartPunchedTape"] = "flowChartPunchedTape";
    ShapeType["flowChartSort"] = "flowChartSort";
    ShapeType["flowChartSummingJunction"] = "flowChartSummingJunction";
    ShapeType["flowChartTerminator"] = "flowChartTerminator";
    ShapeType["folderCorner"] = "folderCorner";
    ShapeType["frame"] = "frame";
    ShapeType["funnel"] = "funnel";
    ShapeType["gear6"] = "gear6";
    ShapeType["gear9"] = "gear9";
    ShapeType["halfFrame"] = "halfFrame";
    ShapeType["heart"] = "heart";
    ShapeType["heptagon"] = "heptagon";
    ShapeType["hexagon"] = "hexagon";
    ShapeType["homePlate"] = "homePlate";
    ShapeType["horizontalScroll"] = "horizontalScroll";
    ShapeType["irregularSeal1"] = "irregularSeal1";
    ShapeType["irregularSeal2"] = "irregularSeal2";
    ShapeType["leftArrow"] = "leftArrow";
    ShapeType["leftArrowCallout"] = "leftArrowCallout";
    ShapeType["leftBrace"] = "leftBrace";
    ShapeType["leftBracket"] = "leftBracket";
    ShapeType["leftCircularArrow"] = "leftCircularArrow";
    ShapeType["leftRightArrow"] = "leftRightArrow";
    ShapeType["leftRightArrowCallout"] = "leftRightArrowCallout";
    ShapeType["leftRightCircularArrow"] = "leftRightCircularArrow";
    ShapeType["leftRightRibbon"] = "leftRightRibbon";
    ShapeType["leftRightUpArrow"] = "leftRightUpArrow";
    ShapeType["leftUpArrow"] = "leftUpArrow";
    ShapeType["lightningBolt"] = "lightningBolt";
    ShapeType["line"] = "line";
    ShapeType["lineInv"] = "lineInv";
    ShapeType["mathDivide"] = "mathDivide";
    ShapeType["mathEqual"] = "mathEqual";
    ShapeType["mathMinus"] = "mathMinus";
    ShapeType["mathMultiply"] = "mathMultiply";
    ShapeType["mathNotEqual"] = "mathNotEqual";
    ShapeType["mathPlus"] = "mathPlus";
    ShapeType["moon"] = "moon";
    ShapeType["noSmoking"] = "noSmoking";
    ShapeType["nonIsoscelesTrapezoid"] = "nonIsoscelesTrapezoid";
    ShapeType["notchedRightArrow"] = "notchedRightArrow";
    ShapeType["octagon"] = "octagon";
    ShapeType["parallelogram"] = "parallelogram";
    ShapeType["pentagon"] = "pentagon";
    ShapeType["pie"] = "pie";
    ShapeType["pieWedge"] = "pieWedge";
    ShapeType["plaque"] = "plaque";
    ShapeType["plaqueTabs"] = "plaqueTabs";
    ShapeType["plus"] = "plus";
    ShapeType["quadArrow"] = "quadArrow";
    ShapeType["quadArrowCallout"] = "quadArrowCallout";
    ShapeType["rect"] = "rect";
    ShapeType["ribbon"] = "ribbon";
    ShapeType["ribbon2"] = "ribbon2";
    ShapeType["rightArrow"] = "rightArrow";
    ShapeType["rightArrowCallout"] = "rightArrowCallout";
    ShapeType["rightBrace"] = "rightBrace";
    ShapeType["rightBracket"] = "rightBracket";
    ShapeType["round1Rect"] = "round1Rect";
    ShapeType["round2DiagRect"] = "round2DiagRect";
    ShapeType["round2SameRect"] = "round2SameRect";
    ShapeType["roundRect"] = "roundRect";
    ShapeType["rtTriangle"] = "rtTriangle";
    ShapeType["smileyFace"] = "smileyFace";
    ShapeType["snip1Rect"] = "snip1Rect";
    ShapeType["snip2DiagRect"] = "snip2DiagRect";
    ShapeType["snip2SameRect"] = "snip2SameRect";
    ShapeType["snipRoundRect"] = "snipRoundRect";
    ShapeType["squareTabs"] = "squareTabs";
    ShapeType["star10"] = "star10";
    ShapeType["star12"] = "star12";
    ShapeType["star16"] = "star16";
    ShapeType["star24"] = "star24";
    ShapeType["star32"] = "star32";
    ShapeType["star4"] = "star4";
    ShapeType["star5"] = "star5";
    ShapeType["star6"] = "star6";
    ShapeType["star7"] = "star7";
    ShapeType["star8"] = "star8";
    ShapeType["stripedRightArrow"] = "stripedRightArrow";
    ShapeType["sun"] = "sun";
    ShapeType["swooshArrow"] = "swooshArrow";
    ShapeType["teardrop"] = "teardrop";
    ShapeType["trapezoid"] = "trapezoid";
    ShapeType["triangle"] = "triangle";
    ShapeType["upArrow"] = "upArrow";
    ShapeType["upArrowCallout"] = "upArrowCallout";
    ShapeType["upDownArrow"] = "upDownArrow";
    ShapeType["upDownArrowCallout"] = "upDownArrowCallout";
    ShapeType["uturnArrow"] = "uturnArrow";
    ShapeType["verticalScroll"] = "verticalScroll";
    ShapeType["wave"] = "wave";
    ShapeType["wedgeEllipseCallout"] = "wedgeEllipseCallout";
    ShapeType["wedgeRectCallout"] = "wedgeRectCallout";
    ShapeType["wedgeRoundRectCallout"] = "wedgeRoundRectCallout";
})(ShapeType || (ShapeType = {}));
var SchemeColor;
(function (SchemeColor) {
    SchemeColor["text1"] = "tx1";
    SchemeColor["text2"] = "tx2";
    SchemeColor["background1"] = "bg1";
    SchemeColor["background2"] = "bg2";
    SchemeColor["accent1"] = "accent1";
    SchemeColor["accent2"] = "accent2";
    SchemeColor["accent3"] = "accent3";
    SchemeColor["accent4"] = "accent4";
    SchemeColor["accent5"] = "accent5";
    SchemeColor["accent6"] = "accent6";
})(SchemeColor || (SchemeColor = {}));
var AlignH;
(function (AlignH) {
    AlignH["left"] = "left";
    AlignH["center"] = "center";
    AlignH["right"] = "right";
    AlignH["justify"] = "justify";
})(AlignH || (AlignH = {}));
var AlignV;
(function (AlignV) {
    AlignV["top"] = "top";
    AlignV["middle"] = "middle";
    AlignV["bottom"] = "bottom";
})(AlignV || (AlignV = {}));
var SHAPE_TYPE;
(function (SHAPE_TYPE) {
    SHAPE_TYPE["ACTION_BUTTON_BACK_OR_PREVIOUS"] = "actionButtonBackPrevious";
    SHAPE_TYPE["ACTION_BUTTON_BEGINNING"] = "actionButtonBeginning";
    SHAPE_TYPE["ACTION_BUTTON_CUSTOM"] = "actionButtonBlank";
    SHAPE_TYPE["ACTION_BUTTON_DOCUMENT"] = "actionButtonDocument";
    SHAPE_TYPE["ACTION_BUTTON_END"] = "actionButtonEnd";
    SHAPE_TYPE["ACTION_BUTTON_FORWARD_OR_NEXT"] = "actionButtonForwardNext";
    SHAPE_TYPE["ACTION_BUTTON_HELP"] = "actionButtonHelp";
    SHAPE_TYPE["ACTION_BUTTON_HOME"] = "actionButtonHome";
    SHAPE_TYPE["ACTION_BUTTON_INFORMATION"] = "actionButtonInformation";
    SHAPE_TYPE["ACTION_BUTTON_MOVIE"] = "actionButtonMovie";
    SHAPE_TYPE["ACTION_BUTTON_RETURN"] = "actionButtonReturn";
    SHAPE_TYPE["ACTION_BUTTON_SOUND"] = "actionButtonSound";
    SHAPE_TYPE["ARC"] = "arc";
    SHAPE_TYPE["BALLOON"] = "wedgeRoundRectCallout";
    SHAPE_TYPE["BENT_ARROW"] = "bentArrow";
    SHAPE_TYPE["BENT_UP_ARROW"] = "bentUpArrow";
    SHAPE_TYPE["BEVEL"] = "bevel";
    SHAPE_TYPE["BLOCK_ARC"] = "blockArc";
    SHAPE_TYPE["CAN"] = "can";
    SHAPE_TYPE["CHART_PLUS"] = "chartPlus";
    SHAPE_TYPE["CHART_STAR"] = "chartStar";
    SHAPE_TYPE["CHART_X"] = "chartX";
    SHAPE_TYPE["CHEVRON"] = "chevron";
    SHAPE_TYPE["CHORD"] = "chord";
    SHAPE_TYPE["CIRCULAR_ARROW"] = "circularArrow";
    SHAPE_TYPE["CLOUD"] = "cloud";
    SHAPE_TYPE["CLOUD_CALLOUT"] = "cloudCallout";
    SHAPE_TYPE["CORNER"] = "corner";
    SHAPE_TYPE["CORNER_TABS"] = "cornerTabs";
    SHAPE_TYPE["CROSS"] = "plus";
    SHAPE_TYPE["CUBE"] = "cube";
    SHAPE_TYPE["CURVED_DOWN_ARROW"] = "curvedDownArrow";
    SHAPE_TYPE["CURVED_DOWN_RIBBON"] = "ellipseRibbon";
    SHAPE_TYPE["CURVED_LEFT_ARROW"] = "curvedLeftArrow";
    SHAPE_TYPE["CURVED_RIGHT_ARROW"] = "curvedRightArrow";
    SHAPE_TYPE["CURVED_UP_ARROW"] = "curvedUpArrow";
    SHAPE_TYPE["CURVED_UP_RIBBON"] = "ellipseRibbon2";
    SHAPE_TYPE["CUSTOM_GEOMETRY"] = "custGeom";
    SHAPE_TYPE["DECAGON"] = "decagon";
    SHAPE_TYPE["DIAGONAL_STRIPE"] = "diagStripe";
    SHAPE_TYPE["DIAMOND"] = "diamond";
    SHAPE_TYPE["DODECAGON"] = "dodecagon";
    SHAPE_TYPE["DONUT"] = "donut";
    SHAPE_TYPE["DOUBLE_BRACE"] = "bracePair";
    SHAPE_TYPE["DOUBLE_BRACKET"] = "bracketPair";
    SHAPE_TYPE["DOUBLE_WAVE"] = "doubleWave";
    SHAPE_TYPE["DOWN_ARROW"] = "downArrow";
    SHAPE_TYPE["DOWN_ARROW_CALLOUT"] = "downArrowCallout";
    SHAPE_TYPE["DOWN_RIBBON"] = "ribbon";
    SHAPE_TYPE["EXPLOSION1"] = "irregularSeal1";
    SHAPE_TYPE["EXPLOSION2"] = "irregularSeal2";
    SHAPE_TYPE["FLOWCHART_ALTERNATE_PROCESS"] = "flowChartAlternateProcess";
    SHAPE_TYPE["FLOWCHART_CARD"] = "flowChartPunchedCard";
    SHAPE_TYPE["FLOWCHART_COLLATE"] = "flowChartCollate";
    SHAPE_TYPE["FLOWCHART_CONNECTOR"] = "flowChartConnector";
    SHAPE_TYPE["FLOWCHART_DATA"] = "flowChartInputOutput";
    SHAPE_TYPE["FLOWCHART_DECISION"] = "flowChartDecision";
    SHAPE_TYPE["FLOWCHART_DELAY"] = "flowChartDelay";
    SHAPE_TYPE["FLOWCHART_DIRECT_ACCESS_STORAGE"] = "flowChartMagneticDrum";
    SHAPE_TYPE["FLOWCHART_DISPLAY"] = "flowChartDisplay";
    SHAPE_TYPE["FLOWCHART_DOCUMENT"] = "flowChartDocument";
    SHAPE_TYPE["FLOWCHART_EXTRACT"] = "flowChartExtract";
    SHAPE_TYPE["FLOWCHART_INTERNAL_STORAGE"] = "flowChartInternalStorage";
    SHAPE_TYPE["FLOWCHART_MAGNETIC_DISK"] = "flowChartMagneticDisk";
    SHAPE_TYPE["FLOWCHART_MANUAL_INPUT"] = "flowChartManualInput";
    SHAPE_TYPE["FLOWCHART_MANUAL_OPERATION"] = "flowChartManualOperation";
    SHAPE_TYPE["FLOWCHART_MERGE"] = "flowChartMerge";
    SHAPE_TYPE["FLOWCHART_MULTIDOCUMENT"] = "flowChartMultidocument";
    SHAPE_TYPE["FLOWCHART_OFFLINE_STORAGE"] = "flowChartOfflineStorage";
    SHAPE_TYPE["FLOWCHART_OFFPAGE_CONNECTOR"] = "flowChartOffpageConnector";
    SHAPE_TYPE["FLOWCHART_OR"] = "flowChartOr";
    SHAPE_TYPE["FLOWCHART_PREDEFINED_PROCESS"] = "flowChartPredefinedProcess";
    SHAPE_TYPE["FLOWCHART_PREPARATION"] = "flowChartPreparation";
    SHAPE_TYPE["FLOWCHART_PROCESS"] = "flowChartProcess";
    SHAPE_TYPE["FLOWCHART_PUNCHED_TAPE"] = "flowChartPunchedTape";
    SHAPE_TYPE["FLOWCHART_SEQUENTIAL_ACCESS_STORAGE"] = "flowChartMagneticTape";
    SHAPE_TYPE["FLOWCHART_SORT"] = "flowChartSort";
    SHAPE_TYPE["FLOWCHART_STORED_DATA"] = "flowChartOnlineStorage";
    SHAPE_TYPE["FLOWCHART_SUMMING_JUNCTION"] = "flowChartSummingJunction";
    SHAPE_TYPE["FLOWCHART_TERMINATOR"] = "flowChartTerminator";
    SHAPE_TYPE["FOLDED_CORNER"] = "folderCorner";
    SHAPE_TYPE["FRAME"] = "frame";
    SHAPE_TYPE["FUNNEL"] = "funnel";
    SHAPE_TYPE["GEAR_6"] = "gear6";
    SHAPE_TYPE["GEAR_9"] = "gear9";
    SHAPE_TYPE["HALF_FRAME"] = "halfFrame";
    SHAPE_TYPE["HEART"] = "heart";
    SHAPE_TYPE["HEPTAGON"] = "heptagon";
    SHAPE_TYPE["HEXAGON"] = "hexagon";
    SHAPE_TYPE["HORIZONTAL_SCROLL"] = "horizontalScroll";
    SHAPE_TYPE["ISOSCELES_TRIANGLE"] = "triangle";
    SHAPE_TYPE["LEFT_ARROW"] = "leftArrow";
    SHAPE_TYPE["LEFT_ARROW_CALLOUT"] = "leftArrowCallout";
    SHAPE_TYPE["LEFT_BRACE"] = "leftBrace";
    SHAPE_TYPE["LEFT_BRACKET"] = "leftBracket";
    SHAPE_TYPE["LEFT_CIRCULAR_ARROW"] = "leftCircularArrow";
    SHAPE_TYPE["LEFT_RIGHT_ARROW"] = "leftRightArrow";
    SHAPE_TYPE["LEFT_RIGHT_ARROW_CALLOUT"] = "leftRightArrowCallout";
    SHAPE_TYPE["LEFT_RIGHT_CIRCULAR_ARROW"] = "leftRightCircularArrow";
    SHAPE_TYPE["LEFT_RIGHT_RIBBON"] = "leftRightRibbon";
    SHAPE_TYPE["LEFT_RIGHT_UP_ARROW"] = "leftRightUpArrow";
    SHAPE_TYPE["LEFT_UP_ARROW"] = "leftUpArrow";
    SHAPE_TYPE["LIGHTNING_BOLT"] = "lightningBolt";
    SHAPE_TYPE["LINE_CALLOUT_1"] = "borderCallout1";
    SHAPE_TYPE["LINE_CALLOUT_1_ACCENT_BAR"] = "accentCallout1";
    SHAPE_TYPE["LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR"] = "accentBorderCallout1";
    SHAPE_TYPE["LINE_CALLOUT_1_NO_BORDER"] = "callout1";
    SHAPE_TYPE["LINE_CALLOUT_2"] = "borderCallout2";
    SHAPE_TYPE["LINE_CALLOUT_2_ACCENT_BAR"] = "accentCallout2";
    SHAPE_TYPE["LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR"] = "accentBorderCallout2";
    SHAPE_TYPE["LINE_CALLOUT_2_NO_BORDER"] = "callout2";
    SHAPE_TYPE["LINE_CALLOUT_3"] = "borderCallout3";
    SHAPE_TYPE["LINE_CALLOUT_3_ACCENT_BAR"] = "accentCallout3";
    SHAPE_TYPE["LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR"] = "accentBorderCallout3";
    SHAPE_TYPE["LINE_CALLOUT_3_NO_BORDER"] = "callout3";
    SHAPE_TYPE["LINE_CALLOUT_4"] = "borderCallout3";
    SHAPE_TYPE["LINE_CALLOUT_4_ACCENT_BAR"] = "accentCallout3";
    SHAPE_TYPE["LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR"] = "accentBorderCallout3";
    SHAPE_TYPE["LINE_CALLOUT_4_NO_BORDER"] = "callout3";
    SHAPE_TYPE["LINE"] = "line";
    SHAPE_TYPE["LINE_INVERSE"] = "lineInv";
    SHAPE_TYPE["MATH_DIVIDE"] = "mathDivide";
    SHAPE_TYPE["MATH_EQUAL"] = "mathEqual";
    SHAPE_TYPE["MATH_MINUS"] = "mathMinus";
    SHAPE_TYPE["MATH_MULTIPLY"] = "mathMultiply";
    SHAPE_TYPE["MATH_NOT_EQUAL"] = "mathNotEqual";
    SHAPE_TYPE["MATH_PLUS"] = "mathPlus";
    SHAPE_TYPE["MOON"] = "moon";
    SHAPE_TYPE["NON_ISOSCELES_TRAPEZOID"] = "nonIsoscelesTrapezoid";
    SHAPE_TYPE["NOTCHED_RIGHT_ARROW"] = "notchedRightArrow";
    SHAPE_TYPE["NO_SYMBOL"] = "noSmoking";
    SHAPE_TYPE["OCTAGON"] = "octagon";
    SHAPE_TYPE["OVAL"] = "ellipse";
    SHAPE_TYPE["OVAL_CALLOUT"] = "wedgeEllipseCallout";
    SHAPE_TYPE["PARALLELOGRAM"] = "parallelogram";
    SHAPE_TYPE["PENTAGON"] = "homePlate";
    SHAPE_TYPE["PIE"] = "pie";
    SHAPE_TYPE["PIE_WEDGE"] = "pieWedge";
    SHAPE_TYPE["PLAQUE"] = "plaque";
    SHAPE_TYPE["PLAQUE_TABS"] = "plaqueTabs";
    SHAPE_TYPE["QUAD_ARROW"] = "quadArrow";
    SHAPE_TYPE["QUAD_ARROW_CALLOUT"] = "quadArrowCallout";
    SHAPE_TYPE["RECTANGLE"] = "rect";
    SHAPE_TYPE["RECTANGULAR_CALLOUT"] = "wedgeRectCallout";
    SHAPE_TYPE["REGULAR_PENTAGON"] = "pentagon";
    SHAPE_TYPE["RIGHT_ARROW"] = "rightArrow";
    SHAPE_TYPE["RIGHT_ARROW_CALLOUT"] = "rightArrowCallout";
    SHAPE_TYPE["RIGHT_BRACE"] = "rightBrace";
    SHAPE_TYPE["RIGHT_BRACKET"] = "rightBracket";
    SHAPE_TYPE["RIGHT_TRIANGLE"] = "rtTriangle";
    SHAPE_TYPE["ROUNDED_RECTANGLE"] = "roundRect";
    SHAPE_TYPE["ROUNDED_RECTANGULAR_CALLOUT"] = "wedgeRoundRectCallout";
    SHAPE_TYPE["ROUND_1_RECTANGLE"] = "round1Rect";
    SHAPE_TYPE["ROUND_2_DIAG_RECTANGLE"] = "round2DiagRect";
    SHAPE_TYPE["ROUND_2_SAME_RECTANGLE"] = "round2SameRect";
    SHAPE_TYPE["SMILEY_FACE"] = "smileyFace";
    SHAPE_TYPE["SNIP_1_RECTANGLE"] = "snip1Rect";
    SHAPE_TYPE["SNIP_2_DIAG_RECTANGLE"] = "snip2DiagRect";
    SHAPE_TYPE["SNIP_2_SAME_RECTANGLE"] = "snip2SameRect";
    SHAPE_TYPE["SNIP_ROUND_RECTANGLE"] = "snipRoundRect";
    SHAPE_TYPE["SQUARE_TABS"] = "squareTabs";
    SHAPE_TYPE["STAR_10_POINT"] = "star10";
    SHAPE_TYPE["STAR_12_POINT"] = "star12";
    SHAPE_TYPE["STAR_16_POINT"] = "star16";
    SHAPE_TYPE["STAR_24_POINT"] = "star24";
    SHAPE_TYPE["STAR_32_POINT"] = "star32";
    SHAPE_TYPE["STAR_4_POINT"] = "star4";
    SHAPE_TYPE["STAR_5_POINT"] = "star5";
    SHAPE_TYPE["STAR_6_POINT"] = "star6";
    SHAPE_TYPE["STAR_7_POINT"] = "star7";
    SHAPE_TYPE["STAR_8_POINT"] = "star8";
    SHAPE_TYPE["STRIPED_RIGHT_ARROW"] = "stripedRightArrow";
    SHAPE_TYPE["SUN"] = "sun";
    SHAPE_TYPE["SWOOSH_ARROW"] = "swooshArrow";
    SHAPE_TYPE["TEAR"] = "teardrop";
    SHAPE_TYPE["TRAPEZOID"] = "trapezoid";
    SHAPE_TYPE["UP_ARROW"] = "upArrow";
    SHAPE_TYPE["UP_ARROW_CALLOUT"] = "upArrowCallout";
    SHAPE_TYPE["UP_DOWN_ARROW"] = "upDownArrow";
    SHAPE_TYPE["UP_DOWN_ARROW_CALLOUT"] = "upDownArrowCallout";
    SHAPE_TYPE["UP_RIBBON"] = "ribbon2";
    SHAPE_TYPE["U_TURN_ARROW"] = "uturnArrow";
    SHAPE_TYPE["VERTICAL_SCROLL"] = "verticalScroll";
    SHAPE_TYPE["WAVE"] = "wave";
})(SHAPE_TYPE || (SHAPE_TYPE = {}));
var CHART_TYPE;
(function (CHART_TYPE) {
    CHART_TYPE["AREA"] = "area";
    CHART_TYPE["BAR"] = "bar";
    CHART_TYPE["BAR3D"] = "bar3D";
    CHART_TYPE["BUBBLE"] = "bubble";
    CHART_TYPE["DOUGHNUT"] = "doughnut";
    CHART_TYPE["LINE"] = "line";
    CHART_TYPE["PIE"] = "pie";
    CHART_TYPE["RADAR"] = "radar";
    CHART_TYPE["SCATTER"] = "scatter";
})(CHART_TYPE || (CHART_TYPE = {}));
var SCHEME_COLOR_NAMES;
(function (SCHEME_COLOR_NAMES) {
    SCHEME_COLOR_NAMES["TEXT1"] = "tx1";
    SCHEME_COLOR_NAMES["TEXT2"] = "tx2";
    SCHEME_COLOR_NAMES["BACKGROUND1"] = "bg1";
    SCHEME_COLOR_NAMES["BACKGROUND2"] = "bg2";
    SCHEME_COLOR_NAMES["ACCENT1"] = "accent1";
    SCHEME_COLOR_NAMES["ACCENT2"] = "accent2";
    SCHEME_COLOR_NAMES["ACCENT3"] = "accent3";
    SCHEME_COLOR_NAMES["ACCENT4"] = "accent4";
    SCHEME_COLOR_NAMES["ACCENT5"] = "accent5";
    SCHEME_COLOR_NAMES["ACCENT6"] = "accent6";
})(SCHEME_COLOR_NAMES || (SCHEME_COLOR_NAMES = {}));
var MASTER_OBJECTS;
(function (MASTER_OBJECTS) {
    MASTER_OBJECTS["chart"] = "chart";
    MASTER_OBJECTS["image"] = "image";
    MASTER_OBJECTS["line"] = "line";
    MASTER_OBJECTS["rect"] = "rect";
    MASTER_OBJECTS["text"] = "text";
    MASTER_OBJECTS["placeholder"] = "placeholder";
})(MASTER_OBJECTS || (MASTER_OBJECTS = {}));
var SLIDE_OBJECT_TYPES;
(function (SLIDE_OBJECT_TYPES) {
    SLIDE_OBJECT_TYPES["chart"] = "chart";
    SLIDE_OBJECT_TYPES["hyperlink"] = "hyperlink";
    SLIDE_OBJECT_TYPES["image"] = "image";
    SLIDE_OBJECT_TYPES["media"] = "media";
    SLIDE_OBJECT_TYPES["online"] = "online";
    SLIDE_OBJECT_TYPES["placeholder"] = "placeholder";
    SLIDE_OBJECT_TYPES["table"] = "table";
    SLIDE_OBJECT_TYPES["tablecell"] = "tablecell";
    SLIDE_OBJECT_TYPES["text"] = "text";
    SLIDE_OBJECT_TYPES["notes"] = "notes";
})(SLIDE_OBJECT_TYPES || (SLIDE_OBJECT_TYPES = {}));
var PLACEHOLDER_TYPES;
(function (PLACEHOLDER_TYPES) {
    PLACEHOLDER_TYPES["title"] = "title";
    PLACEHOLDER_TYPES["body"] = "body";
    PLACEHOLDER_TYPES["image"] = "pic";
    PLACEHOLDER_TYPES["chart"] = "chart";
    PLACEHOLDER_TYPES["table"] = "tbl";
    PLACEHOLDER_TYPES["media"] = "media";
})(PLACEHOLDER_TYPES || (PLACEHOLDER_TYPES = {}));
/**
 * NOTE: 20170304: BULLET_TYPES: Only default is used so far. I'd like to combine the two pieces of code that use these before implementing these as options
 * Since we close <p> within the text object bullets, its slightly more difficult than combining into a func and calling to get the paraProp
 * and i'm not sure if anyone will even use these... so, skipping for now.
 */
var BULLET_TYPES;
(function (BULLET_TYPES) {
    BULLET_TYPES["DEFAULT"] = "&#x2022;";
    BULLET_TYPES["CHECK"] = "&#x2713;";
    BULLET_TYPES["STAR"] = "&#x2605;";
    BULLET_TYPES["TRIANGLE"] = "&#x25B6;";
})(BULLET_TYPES || (BULLET_TYPES = {}));
// IMAGES (base64)
var IMG_BROKEN = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGQAAAB3CAYAAAD1oOVhAAAGAUlEQVR4Xu2dT0xcRRzHf7tAYSsc0EBSIq2xEg8mtTGebVzEqOVIolz0siRE4gGTStqKwdpWsXoyGhMuyAVJOHBgqyvLNgonDkabeCBYW/8kTUr0wsJC+Wfm0bfuvn37Znbem9mR9303mJnf/Pb7ed95M7PDI5JIJPYJV5EC7e3t1N/fT62trdqViQCIu+bVgpIHEo/Hqbe3V/sdYVKHyWSSZmZm8ilVA0oeyNjYmEnaVC2Xvr6+qg5fAOJAz4DU1dURGzFSqZRVqtMpAFIGyMjICC0vL9PExIRWKADiAYTNshYWFrRCARAOEFZcCKWtrY0GBgaUTYkBRACIE4rKZwqACALR5RQAqQCIDqcASIVAVDsFQCSAqHQKgEgCUeUUAPEBRIVTAMQnEBvK5OQkbW9vk991CoAEAMQJxc86BUACAhKUUwAkQCBBOAVAAgbi1ykAogCIH6cAiCIgsk4BEIVAZJwCIIqBVLqiBxANQFgXS0tLND4+zl08AogmIG5OSSQS1gGKwgtANAIRcQqAaAbCe6YASBWA2E6xDyeyDUl7+AKQMkDYYevm5mZHabA/Li4uUiaTsYLau8QA4gLE/hU7wajyYtv1hReDAiAOxQcHBymbzark4BkbQKom/X8dp9Npmpqasn4BIAYAYSnYp+4BBEAMUcCwNOCQsAKZnp62NtQOw8WmwT09PUo+ijaHsOMx7GppaaH6+nolH0Z10K2tLVpdXbW6UfV3mNqBdHd3U1NTk2rtlMRfW1uj2dlZAFGirkRQAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAGHqrm8caPzQ0WC1logbeiC7X3xJm0PvUmRzh45cuki1588FAmVn9BO6P3yF9utrqGH0MtW82S8UN9RA9v/4k7InjhcJFTs/TLVXLwmJV67S7vD7tHF5pKi46fYdosdOcOOGG8j1OcqefbFEJD9Q3GCwDhqT31HklS4A8VRgfYM2Op6k3bt/BQJl58J7lPvwg5JYNccepaMry0LPqFA7hCm39+NNyp2J0172b19QysGINj5CsRtpij57musOViH0QPJQXn6J9u7dlYJSFkbrMYolrwvDAJAC+WWdEpQz7FTgECeUCpzi6YxvvqXoM6eEhqnCSgDikEzUKUE7Aw7xuHctKB5OYU3dZlNR9syQdAaAcAYTC0pXF+39c09o2Ik+3EqxVKqiB7hbYAxZkk4pbBaEM+AQofv+wTrFwylBOQNABIGwavdfe4O2pg5elO+86l99nY58/VUF0byrYsjiSFluNlXYrOHcBar7+EogUADEQ0YRGHbzoKAASBkg2+9cpM1rV0tK2QOcXW7bLEFAARAXIF4w2DrDWoeUWaf4hQIgDiA8GPZ2iNfi0Q8UACkAIgrDbrJ385eDxaPLLrEsFAB5oG6lMPJQPLZZZKAACBGVhcG2Q+bmuLu2nk55e4jqPv1IeEoceiBeX7s2zCa5MAqdstl91vfXwaEGsv/rb5TtOFk6tWXOuJGh6KmnhO9sayrMninPx103JBtXblHkice58cINZP4Hyr5wpkgkdiChEmc4FWazLzenNKa/p0jncwDiqcD6BuWePk07t1asatZGoYQzSqA4nFJ7soNiP/+EUyfc25GI2GG53dHPrKo1g/1Cw4pIXLrzO+1c+/wg7tBbFDle/EbQcjFCPWQJCau5EoBoFpzXHYDwFNJcDiCaBed1ByA8hTSXA4hmwXndAQhPIc3lAKJZcF53AMJTSHM5gGgWnNcdgPAU0lwOIJoF53UHIDyFNJcfSiCdnZ0Ui8U0SxlMd7lcjubn561gh+Y1scFIU/0o/3sgeLO12E2k7UXKYumgFoAYdg8ACIAYpoBh6cAhAGKYAoalA4cAiGEKGJYOHAIghilgWDpwCIAYpoBh6cAhAGKYAoalA4cAiGEKGJYOHAIghilgWDpwCIAYpoBh6ZQ4JB6PKzviYthnNy4d9h+1M5mMlVckkUjsG5dhiBMCEMPg/wuOfrZZ/RSywQAAAABJRU5ErkJggg==';
var IMG_PLAYBTN = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAyAAAAHCCAYAAAAXY63IAAAACXBIWXMAAAsTAAALEwEAmpwYAAAKT2lDQ1BQaG90b3Nob3AgSUNDIHByb2ZpbGUAAHjanVNnVFPpFj333vRCS4iAlEtvUhUIIFJCi4AUkSYqIQkQSoghodkVUcERRUUEG8igiAOOjoCMFVEsDIoK2AfkIaKOg6OIisr74Xuja9a89+bN/rXXPues852zzwfACAyWSDNRNYAMqUIeEeCDx8TG4eQuQIEKJHAAEAizZCFz/SMBAPh+PDwrIsAHvgABeNMLCADATZvAMByH/w/qQplcAYCEAcB0kThLCIAUAEB6jkKmAEBGAYCdmCZTAKAEAGDLY2LjAFAtAGAnf+bTAICd+Jl7AQBblCEVAaCRACATZYhEAGg7AKzPVopFAFgwABRmS8Q5ANgtADBJV2ZIALC3AMDOEAuyAAgMADBRiIUpAAR7AGDIIyN4AISZABRG8lc88SuuEOcqAAB4mbI8uSQ5RYFbCC1xB1dXLh4ozkkXKxQ2YQJhmkAuwnmZGTKBNA/g88wAAKCRFRHgg/P9eM4Ors7ONo62Dl8t6r8G/yJiYuP+5c+rcEAAAOF0ftH+LC+zGoA7BoBt/qIl7gRoXgugdfeLZrIPQLUAoOnaV/Nw+H48PEWhkLnZ2eXk5NhKxEJbYcpXff5nwl/AV/1s+X48/Pf14L7iJIEyXYFHBPjgwsz0TKUcz5IJhGLc5o9H/LcL//wd0yLESWK5WCoU41EScY5EmozzMqUiiUKSKcUl0v9k4t8s+wM+3zUAsGo+AXuRLahdYwP2SycQWHTA4vcAAPK7b8HUKAgDgGiD4c93/+8//UegJQCAZkmScQAAXkQkLlTKsz/HCAAARKCBKrBBG/TBGCzABhzBBdzBC/xgNoRCJMTCQhBCCmSAHHJgKayCQiiGzbAdKmAv1EAdNMBRaIaTcA4uwlW4Dj1wD/phCJ7BKLyBCQRByAgTYSHaiAFiilgjjggXmYX4IcFIBBKLJCDJiBRRIkuRNUgxUopUIFVIHfI9cgI5h1xGupE7yAAygvyGvEcxlIGyUT3UDLVDuag3GoRGogvQZHQxmo8WoJvQcrQaPYw2oefQq2gP2o8+Q8cwwOgYBzPEbDAuxsNCsTgsCZNjy7EirAyrxhqwVqwDu4n1Y8+xdwQSgUXACTYEd0IgYR5BSFhMWE7YSKggHCQ0EdoJNwkDhFHCJyKTqEu0JroR+cQYYjIxh1hILCPWEo8TLxB7iEPENyQSiUMyJ7mQAkmxpFTSEtJG0m5SI+ksqZs0SBojk8naZGuyBzmULCAryIXkneTD5DPkG+Qh8lsKnWJAcaT4U+IoUspqShnlEOU05QZlmDJBVaOaUt2ooVQRNY9aQq2htlKvUYeoEzR1mjnNgxZJS6WtopXTGmgXaPdpr+h0uhHdlR5Ol9BX0svpR+iX6AP0dwwNhhWDx4hnKBmbGAcYZxl3GK+YTKYZ04sZx1QwNzHrmOeZD5lvVVgqtip8FZHKCpVKlSaVGyovVKmqpqreqgtV81XLVI+pXlN9rkZVM1PjqQnUlqtVqp1Q61MbU2epO6iHqmeob1Q/pH5Z/YkGWcNMw09DpFGgsV/jvMYgC2MZs3gsIWsNq4Z1gTXEJrHN2Xx2KruY/R27iz2qqaE5QzNKM1ezUvOUZj8H45hx+Jx0TgnnKKeX836K3hTvKeIpG6Y0TLkxZVxrqpaXllirSKtRq0frvTau7aedpr1Fu1n7gQ5Bx0onXCdHZ4/OBZ3nU9lT3acKpxZNPTr1ri6qa6UbobtEd79up+6Ynr5egJ5Mb6feeb3n+hx9L/1U/W36p/VHDFgGswwkBtsMzhg8xTVxbzwdL8fb8VFDXcNAQ6VhlWGX4YSRudE8o9VGjUYPjGnGXOMk423GbcajJgYmISZLTepN7ppSTbmmKaY7TDtMx83MzaLN1pk1mz0x1zLnm+eb15vft2BaeFostqi2uGVJsuRaplnutrxuhVo5WaVYVVpds0atna0l1rutu6cRp7lOk06rntZnw7Dxtsm2qbcZsOXYBtuutm22fWFnYhdnt8Wuw+6TvZN9un2N/T0HDYfZDqsdWh1+c7RyFDpWOt6azpzuP33F9JbpL2dYzxDP2DPjthPLKcRpnVOb00dnF2e5c4PziIuJS4LLLpc+Lpsbxt3IveRKdPVxXeF60vWdm7Obwu2o26/uNu5p7ofcn8w0nymeWTNz0MPIQ+BR5dE/C5+VMGvfrH5PQ0+BZ7XnIy9jL5FXrdewt6V3qvdh7xc+9j5yn+M+4zw33jLeWV/MN8C3yLfLT8Nvnl+F30N/I/9k/3r/0QCngCUBZwOJgUGBWwL7+Hp8Ib+OPzrbZfay2e1BjKC5QRVBj4KtguXBrSFoyOyQrSH355jOkc5pDoVQfujW0Adh5mGLw34MJ4WHhVeGP45wiFga0TGXNXfR3ENz30T6RJZE3ptnMU85ry1KNSo+qi5qPNo3ujS6P8YuZlnM1VidWElsSxw5LiquNm5svt/87fOH4p3iC+N7F5gvyF1weaHOwvSFpxapLhIsOpZATIhOOJTwQRAqqBaMJfITdyWOCnnCHcJnIi/RNtGI2ENcKh5O8kgqTXqS7JG8NXkkxTOlLOW5hCepkLxMDUzdmzqeFpp2IG0yPTq9MYOSkZBxQqohTZO2Z+pn5mZ2y6xlhbL+xW6Lty8elQfJa7OQrAVZLQq2QqboVFoo1yoHsmdlV2a/zYnKOZarnivN7cyzytuQN5zvn//tEsIS4ZK2pYZLVy0dWOa9rGo5sjxxedsK4xUFK4ZWBqw8uIq2Km3VT6vtV5eufr0mek1rgV7ByoLBtQFr6wtVCuWFfevc1+1dT1gvWd+1YfqGnRs+FYmKrhTbF5cVf9go3HjlG4dvyr+Z3JS0qavEuWTPZtJm6ebeLZ5bDpaql+aXDm4N2dq0Dd9WtO319kXbL5fNKNu7g7ZDuaO/PLi8ZafJzs07P1SkVPRU+lQ27tLdtWHX+G7R7ht7vPY07NXbW7z3/T7JvttVAVVN1WbVZftJ+7P3P66Jqun4lvttXa1ObXHtxwPSA/0HIw6217nU1R3SPVRSj9Yr60cOxx++/p3vdy0NNg1VjZzG4iNwRHnk6fcJ3/ceDTradox7rOEH0x92HWcdL2pCmvKaRptTmvtbYlu6T8w+0dbq3nr8R9sfD5w0PFl5SvNUyWna6YLTk2fyz4ydlZ19fi753GDborZ752PO32oPb++6EHTh0kX/i+c7vDvOXPK4dPKy2+UTV7hXmq86X23qdOo8/pPTT8e7nLuarrlca7nuer21e2b36RueN87d9L158Rb/1tWeOT3dvfN6b/fF9/XfFt1+cif9zsu72Xcn7q28T7xf9EDtQdlD3YfVP1v+3Njv3H9qwHeg89HcR/cGhYPP/pH1jw9DBY+Zj8uGDYbrnjg+OTniP3L96fynQ89kzyaeF/6i/suuFxYvfvjV69fO0ZjRoZfyl5O/bXyl/erA6xmv28bCxh6+yXgzMV70VvvtwXfcdx3vo98PT+R8IH8o/2j5sfVT0Kf7kxmTk/8EA5jz/GMzLdsAAAAgY0hSTQAAeiUAAICDAAD5/wAAgOkAAHUwAADqYAAAOpgAABdvkl/FRgAAFRdJREFUeNrs3WFz2lbagOEnkiVLxsYQsP//z9uZZmMswJIlS3k/tPb23U3TOAUM6Lpm8qkzbXM4A7p1dI4+/etf//oWAAAAB3ARETGdTo0EAACwV1VVRWIYAACAQxEgAACAAAEAAAQIAACAAAEAAAQIAACAAAEAAAQIAAAgQAAAAAQIAAAgQAAAAAQIAAAgQAAAAAECAAAgQAAAAAECAAAgQAAAAAECAAAIEAAAAAECAAAIEAAAAAECAAAIEAAAQIAAAAAIEAAAQIAAAAAIEAAAQIAAAAACBAAAQIAAAAACBAAAQIAAAAACBAAAQIAAAAACBAAAECAAAAACBAAAECAAAAACBAAAECAAAIAAAQAAECAAAIAAAQAAECAAAIAAAQAABAgAAIAAAQAABAgAAIAAAQAABAgAACBAAAAABAgAACBAAAAABAgAACBAAAAAAQIAACBAAAAAAQIAACBAAAAAAQIAACBAAAAAAQIAAAgQAAAAAQIAAAgQAAAAAQIAAAgQAABAgAAAAAgQAABAgAAAAAgQAABAgAAAAAIEAABAgAAAAAIEAABAgAAAAAIEAAAQIAAAAAIEAAAQIAAAAAIEAAAQIAAAgAABAAAQIAAAgAABAAAQIAAAgAABAAAQIAAAgAABAAAECAAAgAABAAAECAAAgAABAAAECAAAIEAAAAAECAAAIEAAAAAECAAAIEAAAAABAgAAIEAAAAABAgAAIEAAAAABAgAACBAAAAABAgAACBAAAAABAgAACBAAAECAAAAACBAAAECAAAAACBAAAECAAAAAAgQAAECAAAAAAgQAAECAAAAAAgQAAECAAAAAAgQAABAgAAAAAgQAABAgAAAAAgQAABAgAACAAAEAABAgAACAAAEAABAgAACAAAEAAAQIAACAAAEAAAQIAACAAAEAAAQIAAAgQAAAAPbnwhAA8CuGYYiXl5fv/7hcXESSuMcFgAAB4G90XRffvn2L5+fniIho2zYiIvq+j77vf+nfmaZppGkaERF5nkdExOXlZXz69CmyLDPoAAIEgDFo2zaen5/j5eUl+r6Pruv28t/5c7y8Bs1ms3n751mWRZqmcXFxEZeXl2+RAoAAAeBEDcMQbdu+/dlXbPyKruve/n9ewyTLssjz/O2PR7oABAgAR67v+2iaJpqmeVt5OBWvUbLdbiPi90e3iqKIoijeHucCQIAAcATRsd1uo2maX96zcYxeV26qqoo0TaMoiphMJmIEQIAAcGjDMERd11HX9VE9WrXvyNput5FlWZRlGWVZekwLQIAAsE+vjyjVdT3qMei6LqqqirIsYzKZOFkLQIAAsEt1XcfT09PJ7es4xLjUdR15nsfV1VWUZWlQAAQIAP/kAnu9Xp/V3o59eN0vsl6v4+bmRogACBAAhMf+9X0fq9VKiAAIEAB+RtM0UVWV8NhhiEyn0yiKwqAACBAAXr1uqrbHY/ch8vDwEHmex3Q6tVkdQIAAjNswDLHZbN5evsd+tG0bX758iclkEtfX147vBRAgAOPTNE08Pj7GMAwG40BejzC+vb31WBaAAAEYh9f9CR63+hjDMLw9ljWfz62GAOyZb1mAD9Q0TXz58kV8HIG2beO3336LpmkMBsAeWQEB+ADDMERVVaN+g/mxfi4PDw9RlmVMp1OrIQACBOD0dV0XDw8PjtY9YnVdR9u2MZ/PnZQFsGNu7QAc+ML269ev4uME9H0fX79+tUoFsGNWQAAOZLVauZg9McMwxGq1iufn55jNZgYEQIAAnMZF7MPDg43mJ6yu6+j73ilZADvgWxRgj7qui69fv4qPM9C2rcfnAAQIwPHHR9d1BuOMPtMvX774TAEECMBxxoe3mp+fYRiEJYAAATgeryddiY/zjxAvLQQQIAAfHh+r1Up8jCRCHh4enGwGIEAAPkbTNLFarQzEyKxWKyshAAIE4LC6rovHx0cDMVKPj4/2hAAIEIDDxYc9H+NmYzqAAAEQH4gQAAECcF4XnI+Pj+IDcwJAgADs38PDg7vd/I+u6+Lh4cFAAAgQgN1ZrVbRtq2B4LvatnUiGoAAAdiNuq69+wHzBECAAOxf13VRVZWB4KdUVeUxPQABAvBrXt98bYMx5gyAAAHYu6qqou97A8G79H1v1QxAgAC8T9M0nufnl9V1HU3TGAgAAQLw9/q+j8fHx5P6f86yLMqy9OEdEe8HARAgAD9ltVqd3IXjp0+fYjabxWKxiDzPfYhH4HU/CIAAAeAvNU1z0u/7yPM8FotFzGazSBJf+R+tbVuPYgECxBAAfN8wDCf36NVfKcsy7u7u4vr62gf7wTyKBQgQAL5rs9mc1YVikiRxc3MT9/f3URSFD/gDw3az2RgIQIAA8B9d18V2uz3Lv1uapjGfz2OxWESWZT7sD7Ddbr2gEBAgAPzHGN7bkOd5LJfLmE6n9oeYYwACBOCjnPrG8/eaTCZxd3cXk8nEh39ANqQDAgSAiBjnnekkSWI6ncb9/b1je801AAECcCh1XUff96P9+6dpGovFIhaLRaRpakLsWd/3Ude1gQAECMBYrddrgxC/7w+5v7+P6+tr+0PMOQABArAPY1/9+J6bm5u4u7uLsiwNxp5YBQEECMBIuRP9Fz8USRKz2SyWy6X9IeYegAAB2AWrH38vy7JYLBYxn8/tD9kxqyCAAAEYmaenJ4Pwk4qiiOVyaX+IOQggQAB+Rdd1o3rvx05+PJIkbm5uYrlc2h+yI23bejs6IEAAxmC73RqEX5Smacxms1gsFpFlmQExFwEECMCPDMPg2fsdyPM8lstlzGYzj2X9A3VdxzAMBgIQIADnfMHH7pRlGXd3d3F9fW0wzEkAAQLgYu8APyx/7A+5v7+PoigMiDkJIEAAIn4/+tSm3/1J0zTm83ksFgvH9r5D13WOhAYECMA5suH3MPI8j/v7+5hOp/aHmJsAAgQYr6ZpDMIBTSaTuLu7i8lkYjDMTUCAAIxL3/cec/mIH50kiel0Gvf395HnuQExPwEBAjAO7jB/rDRNY7FYxHw+tz/EHAUECICLOw6jKIq4v7+P6+tr+0PMUUCAAJynYRiibVsDcURubm7i7u4uyrI0GH9o29ZLCQEBAnAuF3Yc4Q9SksRsNovlcml/iLkKCBAAF3UcRpZlsVgsYjabjX5/iLkKnKMLQwC4qOMYlWUZl5eXsd1u4+npaZSPI5mrwDmyAgKMjrefn9CPVJLEzc1NLJfLUe4PMVcBAQJw4txRPk1pmsZsNovFYhFZlpmzAAIE4DQ8Pz8bhBOW53ksl8uYzWajObbXnAXOjT0gwKi8vLwYhDPw5/0hm83GnAU4IVZAgFHp+94gnMsP2B/7Q+7v78/62F5zFhAgACfMpt7zk6ZpLBaLWCwWZ3lsrzkLCBAAF3IcoTzP4/7+PqbT6dntDzF3AQECcIK+fftmEEZgMpnE3d1dTCYTcxdAgAB8HKcJjejHLUliOp3Gcrk8i/0h5i4gQADgBGRZFovFIubz+VnuDwE4RY7hBUbDC93GqyiKKIoi1ut1PD09xTAM5i7AB7ECAsBo3NzcxN3dXZRlaTAABAjAfnmfAhG/7w+ZzWaxWCxOZn+IuQsIEAABwonL8zwWi0XMZrOj3x9i7gLnxB4QAEatLMu4vLyM7XZ7kvtDAE6NFRAA/BgmSdzc3MRyuYyiKAwIgAAB+Gfc1eZnpGka8/k8FotFZFlmDgMIEIBf8/LyYhD4aXmex3K5jNlsFkmSmMMAO2QPCAD8hT/vD9lsNgYEYAesgADAj34o/9gfcn9/fzLH9gIIEAAAgPAIFgD80DAMsdlsYrvdGgwAAQIA+/O698MJVAACBOB9X3YXvu74eW3bRlVV0XWdOQwgQADe71iOUuW49X0fVVVF0zTmMIAAAYD9GIbBUbsAAgQA9q+u61iv19H3vcEAECAAu5OmqYtM3rRtG+v1Otq2PYm5CyBAAAQIJ6jv+1iv11HX9UnNXQABAgAnZr1ex9PTk2N1AQQIwP7leX4Sj9uwe03TRFVVJ7sClue5DxEQIABw7Lqui6qqhCeAAAE4vMvLS8esjsQwDLHZbGK73Z7N3AUQIAAn5tOnTwZhBF7f53FO+zzMXUCAAJygLMsMwhlr2zZWq9VZnnRm7gICBOCEL+S6rjMQZ6Tv+1itVme7z0N8AAIE4ISlaSpAzsQwDG+PW537nAUQIACn+qV34WvvHNR1HVVVjeJ9HuYsIEAATpiTsE5b27ZRVdWoVrGcgAUIEIBT/tJzN/kk9X0fVVVF0zSj+7t7CSEgQABOWJIkNqKfkNd9Hk9PT6N43Oq/2YAOCBCAM5DnuQA5AXVdx3q9Pstjdd8zVwEECMAZXNSdyxuyz1HXdVFV1dkeqytAAAEC4KKOIzAMQ1RVFXVdGwxzFRAgAOcjSZLI89wd9iOyXq9Hu8/jR/GRJImBAAQIwDkoikKAHIGmaaKqqlHv8/jRHAUQIABndHFXVZWB+CB938dqtRKBAgQQIADjkKZppGnqzvuBDcMQm83GIQA/OT8BBAjAGSmKwoXwAW2329hsNvZ5/OTcBBAgAGdmMpkIkANo2zZWq5XVpnfOTQABAnBm0jT1VvQ96vs+qqqKpmkMxjtkWebxK0CAAJyrsiwFyI4Nw/D2uBW/NicBBAjAGV/sOQ1rd+q6jqqq7PMQIAACBOB7kiSJsiy9ffsfats2qqqymrSD+PDyQUCAAJy5q6srAfKL+r6P9Xpt/HY4FwEECMCZy/M88jz3Urx3eN3n8fT05HGrHc9DAAECMAJXV1cC5CfVdR3r9dqxunuYgwACBGAkyrJ0Uf03uq6LqqqE2h6kaWrzOSBAAMbm5uYmVquVgfgvwzBEVVX2eex57gEIEICRsQryv9brtX0ee2b1AxAgACNmFeR3bdvGarUSYweacwACBGCkxr4K0vd9rFYr+zwOxOoHIEAAGOUqyDAMsdlsYrvdmgAHnmsAAgRg5MqyjKenp9GsAmy329hsNvZ5HFie51Y/gFFKDAHA/xrDnem2bePLly9RVZX4MMcADsYKCMB3vN6dPsejZ/u+j6qqomkaH/QHKcvSW88BAQLA/zedTuP5+flsVgeGYXh73IqPkyRJTKdTAwGM93vQEAD89YXi7e3tWfxd6rqO3377TXwcgdvb20gSP7/AeFkBAfiBoigiz/OT3ZDetm2s12vH6h6JPM+jKAoDAYyaWzAAf2M2m53cHetv377FarWKf//73+LjWH5wkyRms5mBAHwfGgKAH0vT9OQexeq67iw30J+y29vbSNPUQAACxBAA/L2iKDw6g/kDIEAADscdbH7FKa6gAQgQgGP4wkySmM/nBoJ3mc/nTr0CECAAvybLMhuJ+Wmz2SyyLDMQAAIE4NeVZRllWRoIzBMAAQJwGO5s8yNWygAECMDOff78WYTw3fj4/PmzgQAQIAA7/gJNkri9vbXBGHMCQIAAHMbr3W4XnCRJYlUMQIAAiBDEB4AAATjDCJlOpwZipKbTqfgAECAAh1WWpZOPRmg2mzluF+AdLgwBwG4jJCKiqqoYhsGAnLEkSWI6nYoPgPd+fxoCgN1HiD0h5x8fnz9/Fh8AAgTgONiYfv7xYc8HgAABOMoIcaHqMwVAgAC4YOVd8jz3WQIIEIAT+KJNklgul/YLnLCyLGOxWHikDkCAAJyO2WzmmF6fG8DoOYYX4IDKsoyLi4t4eHiIvu8NyBFL0zTm87lHrgB2zAoIwIFlWRbL5TKKojAYR6ooilgul+IDYA+sgAB8gCRJYj6fR9M08fj46KWFR/S53N7eikMAAQJwnoqiiCzLYrVaRdu2BuQD5Xkes9ks0jQ1GAACBOB8pWkai8XCasgHseoBIEAARqkoisjzPKqqirquDcgBlGUZ0+nU8boAAgRgnJIkidlsFldXV7Ferz2WtSd5nsd0OrXJHECAAPB6gbxYLKKu61iv147s3ZE0TWM6nXrcCkCAAPA9ZVlGWZZCZAfhcXNz4230AAIEACEiPAAECABHHyJPT0/2iPyFPM/j6upKeAAIEAB2GSJt28bT05NTs/40LpPJxOZyAAECwD7kef52olNd11HXdXRdN6oxyLLsLcgcpwsgQAA4gCRJYjKZxGQyib7vY7vdRtM0Z7tXJE3TKIoiJpOJN5cDCBAAPvrifDqdxnQ6jb7vo2maaJrm5PeL5HkeRVFEURSiA0CAAHCsMfK6MjIMQ7Rt+/bn2B/VyrLs7RGzPM89XgUgQAA4JUmSvK0gvGrbNp6fn+Pl5SX6vv+wKMmyLNI0jYuLi7i8vIw8z31gAAIEgHPzurrwZ13Xxbdv3+L5+fktUiIi+r7/5T0laZq+PTb1+t+7vLyMT58+ObEKQIAAMGavQfB3qxDDMMTLy8v3f1wuLjwyBYAAAWB3kiTxqBQA7//9MAQAAIAAAQAABAgAAIAAAQAABAgAAIAAAQAABAgAACBAAAAABAgAACBAAAAABAgAACBAAAAAAQIAACBAAAAAAQIAACBAAAAAAQIAAAgQAAAAAQIAAAgQAAAAAQIAAAgQAABAgAAAAAgQAABAgAAAAAgQAABAgAAAAAIEAABAgAAAAAIEAABAgAAAAAIEAABAgAAAAAIEAAAQIAAAAAIEAAAQIAAAAAIEAAAQIAAAgAABAAAQIAAAgAABAAAQIAAAgAABAAAECAAAgAABAAAECAAAgAABAAAECAAAIEAAAAAECAAAIEAAAAAECAAAIEAAAAABAgAAIEAAAAABAgAAIEAAAAABAgAAIEAAAAABAgAACBAAAAABAgAACBAAAAABAgAACBAAAECAAAAACBAAAECAAAAACBAAAECAAAAAAgQAAECAAAAAAgQAAECAAAAAAgQAABAgAAAAAgQAABAgAAAAAgQAABAgAACAAAEAABAgAACAAAEAABAgAACAAAEAAASIIQAAAAQIAAAgQAAAAAQIAAAgQAAAAAQIAAAgQAAAAAECAAAgQAAAAAECAAAgQAAAAAECAAAIEAAAAAECAAAIEAAAAAECAAAIEAAAQIAAAAAIEAAAQIAAAAAIEAAAQIAAAAACBAAAQIAAAAACBAAAQIAAAAACBAAAECAAAAACBAAAECAAAAACBAAAECAAAAACBAAAECAAAIAAAQAAECAAAIAAAQAAECAAAIAAAQAABAgAAIAAAQAABAgAAIAAAQAABAgAACBAAAAAdu0iIqKqKiMBAADs3f8NAFFjCf5mB+leAAAAAElFTkSuQmCC';

/**
 * PptxGenJS: Utility Methods
 */
/**
 * Convert string percentages to number relative to slide size
 * @param {number|string} size - numeric ("5.5") or percentage ("90%")
 * @param {'X' | 'Y'} xyDir - direction
 * @param {PresLayout} layout - presentation layout
 * @returns {number} calculated size
 */
function getSmartParseNumber(size, xyDir, layout) {
    // FIRST: Convert string numeric value if reqd
    if (typeof size === 'string' && !isNaN(Number(size)))
        size = Number(size);
    // CASE 1: Number in inches
    // Assume any number less than 100 is inches
    if (typeof size === 'number' && size < 100)
        return inch2Emu(size);
    // CASE 2: Number is already converted to something other than inches
    // Assume any number greater than 100 is not inches! Just return it (its EMU already i guess??)
    if (typeof size === 'number' && size >= 100)
        return size;
    // CASE 3: Percentage (ex: '50%')
    if (typeof size === 'string' && size.indexOf('%') > -1) {
        if (xyDir && xyDir === 'X')
            return Math.round((parseFloat(size) / 100) * layout.width);
        if (xyDir && xyDir === 'Y')
            return Math.round((parseFloat(size) / 100) * layout.height);
        // Default: Assume width (x/cx)
        return Math.round((parseFloat(size) / 100) * layout.width);
    }
    // LAST: Default value
    return 0;
}
/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
function getUuid(uuidFormat) {
    return uuidFormat.replace(/[xy]/g, function (c) {
        var r = (Math.random() * 16) | 0, v = c === 'x' ? r : (r & 0x3) | 0x8;
        return v.toString(16);
    });
}
/**
 * TODO: What does this method do again??
 * shallow mix, returns new object
 */
function getMix(o1, o2, etc) {
    var objMix = {};
    var _loop_1 = function (i) {
        var oN = arguments_1[i];
        if (oN)
            Object.keys(oN).forEach(function (key) {
                objMix[key] = oN[key];
            });
    };
    var arguments_1 = arguments;
    for (var i = 0; i <= arguments.length; i++) {
        _loop_1(i);
    }
    return objMix;
}
/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string} xml - XML string to encode
 * @returns {string} escaped XML
 */
function encodeXmlEntities(xml) {
    // NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
    if (typeof xml === 'undefined' || xml == null)
        return '';
    return xml.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
/**
 * Convert inches into EMU
 * @param {number|string} inches - as string or number
 * @returns {number} EMU value
 */
function inch2Emu(inches) {
    // FIRST: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
    // Any value over 100 damn sure isnt inches, must be EMU already, so just return it
    if (typeof inches === 'number' && inches > 100)
        return inches;
    if (typeof inches === 'string')
        inches = Number(inches.replace(/in*/gi, ''));
    return Math.round(EMU * inches);
}
/**
 * Convert `pt` into points (using `ONEPT`)
 *
 * @param {number|string} pt
 * @returns {number} value in points (`ONEPT`)
 */
function valToPts(pt) {
    var points = Number(pt) || 0;
    return isNaN(points) ? 0 : Math.round(points * ONEPT);
}
/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 *
 * @param {number} d - degrees
 * @returns {number} rot - value
 */
function convertRotationDegrees(d) {
    d = d || 0;
    return Math.round((d > 360 ? d - 360 : d) * 60000);
}
/**
 * Converts component value to hex value
 * @param {number} c - component color
 * @returns {string} hex string
 */
function componentToHex(c) {
    var hex = c.toString(16);
    return hex.length === 1 ? '0' + hex : hex;
}
/**
 * Converts RGB colors from css selectors to Hex for Presentation colors
 * @param {number} r - red value
 * @param {number} g - green value
 * @param {number} b - blue value
 * @returns {string} XML string
 */
function rgbToHex(r, g, b) {
    return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase();
}
/**
 * Create either a `a:schemeClr` - (scheme color) or `a:srgbClr` (hexa representation).
 * @param {string|SCHEME_COLORS} colorStr - hexa representation (eg. "FFFF00") or a scheme color constant (eg. pptx.SchemeColor.ACCENT1)
 * @param {string} innerElements - additional elements that adjust the color and are enclosed by the color element
 * @returns {string} XML string
 */
function createColorElement(colorStr, innerElements) {
    var colorVal = (colorStr || '').replace('#', '');
    var isHexaRgb = REGEX_HEX_COLOR.test(colorVal);
    if (!isHexaRgb &&
        colorVal !== SchemeColor.background1 &&
        colorVal !== SchemeColor.background2 &&
        colorVal !== SchemeColor.text1 &&
        colorVal !== SchemeColor.text2 &&
        colorVal !== SchemeColor.accent1 &&
        colorVal !== SchemeColor.accent2 &&
        colorVal !== SchemeColor.accent3 &&
        colorVal !== SchemeColor.accent4 &&
        colorVal !== SchemeColor.accent5 &&
        colorVal !== SchemeColor.accent6) {
        console.warn("\"" + colorVal + "\" is not a valid scheme color or hexa RGB! \"" + DEF_FONT_COLOR + "\" is used as a fallback. Pass 6-digit RGB or 'pptx.SchemeColor' values");
        colorVal = DEF_FONT_COLOR;
    }
    var tagName = isHexaRgb ? 'srgbClr' : 'schemeClr';
    var colorAttr = 'val="' + (isHexaRgb ? colorVal.toUpperCase() : colorVal) + '"';
    return innerElements ? "<a:" + tagName + " " + colorAttr + ">" + innerElements + "</a:" + tagName + ">" : "<a:" + tagName + " " + colorAttr + "/>";
}
/**
 * Creates `a:glow` element
 * @param {TextGlowProps} options glow properties
 * @param {TextGlowProps} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 *	{ size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
function createGlowElement(options, defaults) {
    var strXml = '', opts = getMix(defaults, options), size = Math.round(opts['size'] * ONEPT), color = opts['color'], opacity = Math.round(opts['opacity'] * 100000);
    strXml += "<a:glow rad=\"" + size + "\">";
    strXml += createColorElement(color, "<a:alpha val=\"" + opacity + "\"/>");
    strXml += "</a:glow>";
    return strXml;
}
/**
 * Create color selection
 * @param {Color | ShapeFillProps | ShapeLineProps} props fill props
 * @returns XML string
 */
function genXmlColorSelection(props) {
    var fillType = 'solid';
    var colorVal = '';
    var internalElements = '';
    var outText = '';
    if (props) {
        if (typeof props === 'string')
            colorVal = props;
        else {
            if (props.type)
                fillType = props.type;
            if (props.color)
                colorVal = props.color;
            if (props.alpha)
                internalElements += "<a:alpha val=\"" + Math.round((100 - props.alpha) * 1000) + "\"/>"; // DEPRECATED: @deprecated v3.3.0
            if (props.transparency)
                internalElements += "<a:alpha val=\"" + Math.round((100 - props.transparency) * 1000) + "\"/>";
        }
        switch (fillType) {
            case 'solid':
                outText += "<a:solidFill>" + createColorElement(colorVal, internalElements) + "</a:solidFill>";
                break;
            default:
                outText += ''; // @note need a statement as having only "break" is removed by rollup, then tiggers "no-default" js-linter
                break;
        }
    }
    return outText;
}
/**
 * Get a new rel ID (rId) for charts, media, etc.
 * @param {PresSlide} target - the slide to use
 * @returns {number} count of all current rels plus 1 for the caller to use as its "rId"
 */
function getNewRelId(target) {
    return target._rels.length + target._relsChart.length + target._relsMedia.length + 1;
}

/**
 * PptxGenJS: Table Generation
 */
/**
 * Break text paragraphs into lines based upon table column width (e.g.: Magic Happens Here(tm))
 * @param {TableCell} cell - table cell
 * @param {number} colWidth - table column width
 * @return {string[]} XML
 */
function parseTextToLines(cell, colWidth) {
    var CHAR = 2.2 + (cell.options && cell.options.autoPageCharWeight ? cell.options.autoPageCharWeight : 0); // Character Constant (An approximation of the Golden Ratio)
    var CPL = (colWidth * EMU) / (((cell.options && cell.options.fontSize) || DEF_FONT_SIZE) / CHAR); // Chars-Per-Line
    var arrLines = [];
    var strCurrLine = '';
    // A: Allow a single space/whitespace as cell text (user-requested feature)
    if (cell.text && cell.text.toString().trim().length === 0)
        return [' '];
    // B: Remove leading/trailing spaces
    var inStr = (cell.text || '').toString().trim();
    // C: Build line array
    inStr.split('\n').forEach(function (line) {
        line.split(' ').forEach(function (word) {
            if (strCurrLine.length + word.length + 1 < CPL) {
                strCurrLine += word + ' ';
            }
            else {
                if (strCurrLine)
                    arrLines.push(strCurrLine);
                strCurrLine = word + ' ';
            }
        });
        // All words for this line have been exhausted, flush buffer to new line, clear line var
        if (strCurrLine)
            arrLines.push(strCurrLine.trim() + CRLF);
        strCurrLine = '';
    });
    // D: Remove trailing linebreak
    arrLines[arrLines.length - 1] = arrLines[arrLines.length - 1].trim();
    return arrLines;
}
/**
 * Takes an array of table rows and breaks into an array of slides, which contain the calculated amount of table rows that fit on that slide
 * @param {TableCell[][]} tableRows - HTMLElementID of the table
 * @param {ITableToSlidesOpts} tabOpts - array of options (e.g.: tabsize)
 * @param {PresLayout} presLayout - Presentation layout
 * @param {SlideLayout} masterSlide - master slide (if any)
 * @return {TableRowSlide[]} array of table rows
 */
function getSlidesForTableRows(tableRows, tabOpts, presLayout, masterSlide) {
    if (tableRows === void 0) { tableRows = []; }
    if (tabOpts === void 0) { tabOpts = {}; }
    var arrInchMargins = DEF_SLIDE_MARGIN_IN, emuTabCurrH = 0, emuSlideTabW = EMU * 1, emuSlideTabH = EMU * 1, numCols = 0, tableRowSlides = [
        {
            rows: [],
        },
    ];
    if (tabOpts.verbose) {
        console.log("-- VERBOSE MODE ----------------------------------");
        console.log(".. (PARAMETERS)");
        console.log("presLayout.height ......... = " + presLayout.height / EMU);
        console.log("tabOpts.h ................. = " + tabOpts.h);
        console.log("tabOpts.w ................. = " + tabOpts.w);
        console.log("tabOpts.colW .............. = " + tabOpts.colW);
        console.log("tabOpts.slideMargin ....... = " + (tabOpts.slideMargin || ''));
        console.log(".. (/PARAMETERS)");
    }
    // STEP 1: Calculate margins
    {
        // Important: Use default size as zero cell margin is causing our tables to be too large and touch bottom of slide!
        if (!tabOpts.slideMargin && tabOpts.slideMargin !== 0)
            tabOpts.slideMargin = DEF_SLIDE_MARGIN_IN[0];
        if (masterSlide && typeof masterSlide._margin !== 'undefined') {
            if (Array.isArray(masterSlide._margin))
                arrInchMargins = masterSlide._margin;
            else if (!isNaN(Number(masterSlide._margin)))
                arrInchMargins = [Number(masterSlide._margin), Number(masterSlide._margin), Number(masterSlide._margin), Number(masterSlide._margin)];
        }
        else if (tabOpts.slideMargin || tabOpts.slideMargin === 0) {
            if (Array.isArray(tabOpts.slideMargin))
                arrInchMargins = tabOpts.slideMargin;
            else if (!isNaN(tabOpts.slideMargin))
                arrInchMargins = [tabOpts.slideMargin, tabOpts.slideMargin, tabOpts.slideMargin, tabOpts.slideMargin];
        }
        if (tabOpts.verbose)
            console.log('arrInchMargins ......... = ' + arrInchMargins.toString());
    }
    // STEP 2: Calculate number of columns
    {
        // NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
        // ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
        var firstRow = tableRows[0] || [];
        firstRow.forEach(function (cell) {
            if (!cell)
                cell = { _type: SLIDE_OBJECT_TYPES.tablecell };
            var cellOpts = cell.options || null;
            numCols += Number(cellOpts && cellOpts.colspan ? cellOpts.colspan : 1);
        });
        if (tabOpts.verbose)
            console.log('numCols ................ = ' + numCols);
    }
    // STEP 3: Calculate tabOpts.w if tabOpts.colW was provided
    if (!tabOpts.w && tabOpts.colW) {
        if (Array.isArray(tabOpts.colW))
            tabOpts.colW.forEach(function (val) {
                typeof tabOpts.w !== 'number' ? (tabOpts.w = 0 + val) : (tabOpts.w += val);
            });
        else {
            tabOpts.w = tabOpts.colW * numCols;
        }
    }
    // STEP 4: Calculate usable space/table size (now that total usable space is known)
    {
        emuSlideTabW =
            typeof tabOpts.w === 'number'
                ? inch2Emu(tabOpts.w)
                : presLayout.width - inch2Emu((typeof tabOpts.x === 'number' ? tabOpts.x : arrInchMargins[1]) + arrInchMargins[3]);
        if (tabOpts.verbose)
            console.log('emuSlideTabW (in) ...... = ' + (emuSlideTabW / EMU).toFixed(1));
    }
    // STEP 5: Calculate column widths if not provided (emuSlideTabW will be used below to determine lines-per-col)
    if (!tabOpts.colW || !Array.isArray(tabOpts.colW)) {
        if (tabOpts.colW && !isNaN(Number(tabOpts.colW))) {
            var arrColW_1 = [];
            var firstRow = tableRows[0] || [];
            firstRow.forEach(function () { return arrColW_1.push(tabOpts.colW); });
            tabOpts.colW = [];
            arrColW_1.forEach(function (val) {
                if (Array.isArray(tabOpts.colW))
                    tabOpts.colW.push(val);
            });
        }
        // No column widths provided? Then distribute cols.
        else {
            tabOpts.colW = [];
            for (var iCol = 0; iCol < numCols; iCol++) {
                tabOpts.colW.push(emuSlideTabW / EMU / numCols);
            }
        }
    }
    // STEP 6: **MAIN** Iterate over rows, add table content, create new slides as rows overflow
    var iRow = 0;
    var _loop_1 = function () {
        var row = tableRows.shift();
        iRow++;
        // A: Row variables
        var maxLineHeight = 0;
        var linesRow = [];
        var maxCellMarTopEmu = 0;
        var maxCellMarBtmEmu = 0;
        // B: Create new row in data model
        var currSlide = tableRowSlides[tableRowSlides.length - 1];
        var newRowSlide = [];
        row.forEach(function (cell) {
            newRowSlide.push({
                type: SLIDE_OBJECT_TYPES.tablecell,
                text: '',
                options: cell.options,
            });
            if (cell.options.margin && cell.options.margin[0] && valToPts(cell.options.margin[0]) > maxCellMarTopEmu)
                maxCellMarTopEmu = valToPts(cell.options.margin[0]);
            else if (tabOpts.margin && tabOpts.margin[0] && valToPts(tabOpts.margin[0]) > maxCellMarTopEmu)
                maxCellMarTopEmu = valToPts(tabOpts.margin[0]);
            if (cell.options.margin && cell.options.margin[2] && valToPts(cell.options.margin[2]) > maxCellMarBtmEmu)
                maxCellMarBtmEmu = valToPts(cell.options.margin[2]);
            else if (tabOpts.margin && tabOpts.margin[2] && valToPts(tabOpts.margin[2]) > maxCellMarBtmEmu)
                maxCellMarBtmEmu = valToPts(tabOpts.margin[2]);
        });
        // C: Calc usable vertical space/table height. Set default value first, adjust below when necessary.
        emuSlideTabH =
            tabOpts.h && typeof tabOpts.h === 'number'
                ? tabOpts.h
                : presLayout.height - inch2Emu(arrInchMargins[0] + arrInchMargins[2]) - (tabOpts.y && typeof tabOpts.y === 'number' ? tabOpts.y : 0);
        if (tabOpts.verbose)
            console.log('emuSlideTabH (in) ...... = ' + (emuSlideTabH / EMU).toFixed(1));
        // D: RULE: Use margins for starting point after the initial Slide, not `opt.y` (ISSUE#43, ISSUE#47, ISSUE#48)
        if (tableRowSlides.length > 1 && typeof tabOpts.autoPageSlideStartY === 'number') {
            emuSlideTabH = tabOpts.h && typeof tabOpts.h === 'number' ? tabOpts.h : presLayout.height - inch2Emu(tabOpts.autoPageSlideStartY + arrInchMargins[2]);
        }
        else if (tableRowSlides.length > 1 && typeof tabOpts.newSlideStartY === 'number') {
            // @deprecated v3.3.0
            emuSlideTabH = tabOpts.h && typeof tabOpts.h === 'number' ? tabOpts.h : presLayout.height - inch2Emu(tabOpts.newSlideStartY + arrInchMargins[2]);
        }
        else if (tableRowSlides.length > 1 && typeof tabOpts.y === 'number') {
            emuSlideTabH = presLayout.height - inch2Emu((tabOpts.y / EMU < arrInchMargins[0] ? tabOpts.y / EMU : arrInchMargins[0]) + arrInchMargins[2]);
            // Use whichever is greater: area between margins or the table H provided (dont shrink usable area - the whole point of over-riding X on paging is to *increarse* usable space)
            if (typeof tabOpts.h === 'number' && emuSlideTabH < tabOpts.h)
                emuSlideTabH = tabOpts.h;
        }
        else if (typeof tabOpts.h === 'number' && typeof tabOpts.y === 'number')
            emuSlideTabH = tabOpts.h ? tabOpts.h : presLayout.height - inch2Emu((tabOpts.y / EMU || arrInchMargins[0]) + arrInchMargins[2]);
        //if (tabOpts.verbose) console.log(`- SLIDE [${tableRowSlides.length}]: emuSlideTabH .. = ${(emuSlideTabH / EMU).toFixed(1)}`)
        // E: **BUILD DATA SET** | Iterate over cells: split text into lines[], set `lineHeight`
        row.forEach(function (cell, iCell) {
            var newCell = {
                _type: SLIDE_OBJECT_TYPES.tablecell,
                _lines: [],
                _lineHeight: inch2Emu(((cell.options && cell.options.fontSize ? cell.options.fontSize : tabOpts.fontSize ? tabOpts.fontSize : DEF_FONT_SIZE) *
                    (LINEH_MODIFIER + (tabOpts.autoPageLineWeight ? tabOpts.autoPageLineWeight : 0))) /
                    100),
                text: '',
                options: cell.options,
            };
            //if (tabOpts.verbose) console.log(`- CELL [${iCell}]: newCell.lineHeight ..... = ${(newCell.lineHeight / EMU).toFixed(2)}`)
            // 1: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
            if (newCell.options.rowspan)
                newCell._lineHeight = 0;
            // 2: The parseTextToLines method uses `autoPageCharWeight`, so inherit from table options
            newCell.options.autoPageCharWeight = tabOpts.autoPageCharWeight ? tabOpts.autoPageCharWeight : null;
            // 3: **MAIN** Parse cell contents into lines based upon col width, font, etc
            var totalColW = tabOpts.colW[iCell];
            if (cell.options.colspan && Array.isArray(tabOpts.colW)) {
                totalColW = tabOpts.colW.filter(function (_cell, idx) { return idx >= iCell && idx < idx + cell.options.colspan; }).reduce(function (prev, curr) { return prev + curr; });
            }
            newCell._lines = parseTextToLines(cell, totalColW / ONEPT);
            // 4: Add to array
            linesRow.push(newCell);
        });
        // F: Start row height with margins
        if (tabOpts.verbose)
            console.log("- SLIDE [" + tableRowSlides.length + "]: ROW [" + iRow + "]: maxCellMarTopEmu=" + maxCellMarTopEmu + " / maxCellMarBtmEmu=" + maxCellMarBtmEmu);
        emuTabCurrH += maxCellMarTopEmu + maxCellMarBtmEmu;
        // G: Only create a new row if there is room, otherwise, it'll be an empty row as "A:" below will create a new Slide before loop can populate this row
        if (emuTabCurrH + maxLineHeight <= emuSlideTabH)
            currSlide.rows.push(newRowSlide);
        /* H: **PAGE DATA SET**
         * Add text one-line-a-time to this row's cells until: lines are exhausted OR table height limit is hit
         * Design: Building cells L-to-R/loop style wont work as one could be 100 lines and another 1 line.
         * Therefore, build the whole row, 1-line-at-a-time, spanning all columns.
         * That way, when the vertical size limit is hit, all lines pick up where they need to on the subsequent slide.
         */
        if (tabOpts.verbose)
            console.log("- SLIDE [" + tableRowSlides.length + "]: ROW [" + iRow + "]: START...");
        var _loop_2 = function () {
            // A: Add new Slide if there is no more space to fix 1 new line
            if (emuTabCurrH + maxLineHeight > emuSlideTabH) {
                if (tabOpts.verbose)
                    console.log("** NEW SLIDE CREATED *****************************************" +
                        (" (why?): " + (emuTabCurrH / EMU).toFixed(2) + "+" + (maxLineHeight / EMU).toFixed(2) + " > " + emuSlideTabH / EMU));
                // 1: Add a new slide
                tableRowSlides.push({
                    rows: [],
                });
                // 2: Reset current table height for new Slide
                emuTabCurrH = 0; // This row's emuRowH w/b added below
                // 3: Handle repeat headers option /or/ Add new empty row to continue current lines into
                if ((tabOpts.addHeaderToEach || tabOpts.autoPageRepeatHeader) && tabOpts._arrObjTabHeadRows) {
                    // A: Add remaining cell lines
                    var newRowSlide_1 = [];
                    linesRow.forEach(function (cell) {
                        newRowSlide_1.push({
                            type: SLIDE_OBJECT_TYPES.tablecell,
                            text: cell._lines.join(''),
                            options: cell.options,
                        });
                    });
                    tableRows.unshift(newRowSlide_1);
                    // B: Add header row(s)
                    var tableHeadRows_1 = [];
                    tabOpts._arrObjTabHeadRows.forEach(function (row) {
                        var newHeadRow = [];
                        row.forEach(function (cell) { return newHeadRow.push(cell); });
                        tableHeadRows_1.push(newHeadRow);
                    });
                    tableRows = __spreadArray(__spreadArray([], tableHeadRows_1), tableRows);
                    return "break";
                }
                else {
                    // A: Add new row to new slide
                    var currSlide_1 = tableRowSlides[tableRowSlides.length - 1];
                    var newRowSlide_2 = [];
                    row.forEach(function (cell) {
                        newRowSlide_2.push({
                            type: SLIDE_OBJECT_TYPES.tablecell,
                            text: '',
                            options: cell.options,
                        });
                    });
                    currSlide_1.rows.push(newRowSlide_2);
                }
            }
            // B: Add a line of text to 1-N cells that still have `lines`
            linesRow.forEach(function (cell, idxR) {
                if (cell._lines.length > 0) {
                    // 1
                    var currSlide_2 = tableRowSlides[tableRowSlides.length - 1];
                    // NOTE: TableCell.text type c/b string|IText (for conversion in method that calls this one), but we can guarantee it always string b/c we craft it, hence this TS workaround
                    var rowCell = currSlide_2.rows[currSlide_2.rows.length - 1][idxR];
                    var currText = rowCell.text.toString();
                    rowCell.text += (currText.length > 0 && !RegExp(/\n$/g).test(currText) ? CRLF : '').replace(/[\r\n]+$/g, CRLF) + cell._lines.shift();
                    // 2
                    if (cell._lineHeight > maxLineHeight)
                        maxLineHeight = cell._lineHeight;
                }
            });
            // C: Increase table height by one line height as 1-N cells below are
            emuTabCurrH += maxLineHeight;
            if (tabOpts.verbose)
                console.log("- SLIDE [" + tableRowSlides.length + "]: ROW [" + iRow + "]: one line added ... emuTabCurrH = " + (emuTabCurrH / EMU).toFixed(2));
        };
        while (linesRow.filter(function (cell) { return cell._lines.length > 0; }).length > 0) {
            var state_1 = _loop_2();
            if (state_1 === "break")
                break;
        }
        if (tabOpts.verbose)
            console.log("- SLIDE [" + tableRowSlides.length + "]: ROW [" + iRow + "]: ...COMPLETE ...... emuTabCurrH = " + (emuTabCurrH / EMU).toFixed(2) + " ( emuSlideTabH = " + (emuSlideTabH / EMU).toFixed(2) + " )");
    };
    while (tableRows.length > 0) {
        _loop_1();
    }
    if (tabOpts.verbose) {
        console.log("\n|================================================|\n| FINAL: tableRowSlides.length = " + tableRowSlides.length);
        console.log(tableRowSlides);
        //console.log(JSON.stringify(tableRowSlides,null,2))
        console.log("|================================================|\n\n");
    }
    return tableRowSlides;
}
/**
 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
 * @param {PptxGenJS} pptx - pptxgenjs instance
 * @param {string} tabEleId - HTMLElementID of the table
 * @param {ITableToSlidesOpts} options - array of options (e.g.: tabsize)
 * @param {SlideLayout} masterSlide - masterSlide
 */
function genTableToSlides(pptx, tabEleId, options, masterSlide) {
    if (options === void 0) { options = {}; }
    var opts = options || {};
    opts.slideMargin = opts.slideMargin || opts.slideMargin === 0 ? opts.slideMargin : 0.5;
    var emuSlideTabW = opts.w || pptx.presLayout.width;
    var arrObjTabHeadRows = [];
    var arrObjTabBodyRows = [];
    var arrObjTabFootRows = [];
    var arrColW = [];
    var arrTabColW = [];
    var arrInchMargins = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
    var intTabW = 0;
    // REALITY-CHECK:
    if (!document.getElementById(tabEleId))
        throw new Error('tableToSlides: Table ID "' + tabEleId + '" does not exist!');
    // STEP 1: Set margins
    if (masterSlide && masterSlide._margin) {
        if (Array.isArray(masterSlide._margin))
            arrInchMargins = masterSlide._margin;
        else if (!isNaN(masterSlide._margin))
            arrInchMargins = [masterSlide._margin, masterSlide._margin, masterSlide._margin, masterSlide._margin];
        opts.slideMargin = arrInchMargins;
    }
    else if (opts && opts.slideMargin) {
        if (Array.isArray(opts.slideMargin))
            arrInchMargins = opts.slideMargin;
        else if (!isNaN(opts.slideMargin))
            arrInchMargins = [opts.slideMargin, opts.slideMargin, opts.slideMargin, opts.slideMargin];
    }
    emuSlideTabW = (opts.w ? inch2Emu(opts.w) : pptx.presLayout.width) - inch2Emu(arrInchMargins[1] + arrInchMargins[3]);
    if (opts.verbose)
        console.log('-- VERBOSE MODE ----------------------------------');
    if (opts.verbose)
        console.log("opts.h ................. = " + opts.h);
    if (opts.verbose)
        console.log("opts.w ................. = " + opts.w);
    if (opts.verbose)
        console.log("pptx.presLayout.width .. = " + pptx.presLayout.width / EMU);
    if (opts.verbose)
        console.log("emuSlideTabW (in)....... = " + emuSlideTabW / EMU);
    // STEP 2: Grab table col widths - just find the first availble row, either thead/tbody/tfoot, others may have colspsna,s who cares, we only need col widths from 1
    var firstRowCells = document.querySelectorAll("#" + tabEleId + " tr:first-child th");
    if (firstRowCells.length === 0)
        firstRowCells = document.querySelectorAll("#" + tabEleId + " tr:first-child td");
    firstRowCells.forEach(function (cell) {
        if (cell.getAttribute('colspan')) {
            // Guesstimate (divide evenly) col widths
            // NOTE: both j$query and vanilla selectors return {0} when table is not visible)
            for (var idxc = 0; idxc < Number(cell.getAttribute('colspan')); idxc++) {
                arrTabColW.push(Math.round(cell.offsetWidth / Number(cell.getAttribute('colspan'))));
            }
        }
        else {
            arrTabColW.push(cell.offsetWidth);
        }
    });
    arrTabColW.forEach(function (colW) {
        intTabW += colW;
    });
    // STEP 3: Calc/Set column widths by using same column width percent from HTML table
    arrTabColW.forEach(function (colW, idxW) {
        var intCalcWidth = Number(((Number(emuSlideTabW) * ((colW / intTabW) * 100)) / 100 / EMU).toFixed(2));
        var intMinWidth = 0;
        var colSelectorMin = document.querySelector("#" + tabEleId + " thead tr:first-child th:nth-child(" + (idxW + 1) + ")");
        if (colSelectorMin)
            intMinWidth = Number(colSelectorMin.getAttribute('data-pptx-min-width'));
        var colSelectorSet = document.querySelector("#" + tabEleId + " thead tr:first-child th:nth-child(" + (idxW + 1) + ")");
        if (colSelectorSet)
            intMinWidth = Number(colSelectorSet.getAttribute('data-pptx-width'));
        arrColW.push(intMinWidth > intCalcWidth ? intMinWidth : intCalcWidth);
    });
    if (opts.verbose) {
        console.log("arrColW ................ = " + arrColW.toString());
    }
    ['thead', 'tbody', 'tfoot'].forEach(function (part) {
        document.querySelectorAll("#" + tabEleId + " " + part + " tr").forEach(function (row) {
            var arrObjTabCells = [];
            Array.from(row.cells).forEach(function (cell) {
                // A: Get RGB text/bkgd colors
                var arrRGB1 = window.getComputedStyle(cell).getPropertyValue('color').replace(/\s+/gi, '').replace('rgba(', '').replace('rgb(', '').replace(')', '').split(',');
                var arrRGB2 = window
                    .getComputedStyle(cell)
                    .getPropertyValue('background-color')
                    .replace(/\s+/gi, '')
                    .replace('rgba(', '')
                    .replace('rgb(', '')
                    .replace(')', '')
                    .split(',');
                if (
                // NOTE: (ISSUE#57): Default for unstyled tables is black bkgd, so use white instead
                window.getComputedStyle(cell).getPropertyValue('background-color') === 'rgba(0, 0, 0, 0)' ||
                    window.getComputedStyle(cell).getPropertyValue('transparent')) {
                    arrRGB2 = ['255', '255', '255'];
                }
                // B: Create option object
                var cellOpts = {
                    align: null,
                    bold: window.getComputedStyle(cell).getPropertyValue('font-weight') === 'bold' ||
                        Number(window.getComputedStyle(cell).getPropertyValue('font-weight')) >= 500
                        ? true
                        : false,
                    border: null,
                    color: rgbToHex(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
                    fill: { color: rgbToHex(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2])) },
                    fontFace: (window.getComputedStyle(cell).getPropertyValue('font-family') || '').split(',')[0].replace(/"/g, '').replace('inherit', '').replace('initial', '') ||
                        null,
                    fontSize: Number(window.getComputedStyle(cell).getPropertyValue('font-size').replace(/[a-z]/gi, '')),
                    margin: null,
                    colspan: Number(cell.getAttribute('colspan')) || null,
                    rowspan: Number(cell.getAttribute('rowspan')) || null,
                    valign: null,
                };
                if (['left', 'center', 'right', 'start', 'end'].indexOf(window.getComputedStyle(cell).getPropertyValue('text-align')) > -1) {
                    var align = window.getComputedStyle(cell).getPropertyValue('text-align').replace('start', 'left').replace('end', 'right');
                    cellOpts.align = align === 'center' ? 'center' : align === 'left' ? 'left' : align === 'right' ? 'right' : null;
                }
                if (['top', 'middle', 'bottom'].indexOf(window.getComputedStyle(cell).getPropertyValue('vertical-align')) > -1) {
                    var valign = window.getComputedStyle(cell).getPropertyValue('vertical-align');
                    cellOpts.valign = valign === 'top' ? 'top' : valign === 'middle' ? 'middle' : valign === 'bottom' ? 'bottom' : null;
                }
                // C: Add padding [margin] (if any)
                // NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
                if (window.getComputedStyle(cell).getPropertyValue('padding-left')) {
                    cellOpts.margin = [0, 0, 0, 0];
                    var sidesPad = ['padding-top', 'padding-right', 'padding-bottom', 'padding-left'];
                    sidesPad.forEach(function (val, idxs) {
                        cellOpts.margin[idxs] = Math.round(Number(window.getComputedStyle(cell).getPropertyValue(val).replace(/\D/gi, '')));
                    });
                }
                // D: Add border (if any)
                if (window.getComputedStyle(cell).getPropertyValue('border-top-width') ||
                    window.getComputedStyle(cell).getPropertyValue('border-right-width') ||
                    window.getComputedStyle(cell).getPropertyValue('border-bottom-width') ||
                    window.getComputedStyle(cell).getPropertyValue('border-left-width')) {
                    cellOpts.border = [null, null, null, null];
                    var sidesBor = ['top', 'right', 'bottom', 'left'];
                    sidesBor.forEach(function (val, idxb) {
                        var intBorderW = Math.round(Number(window
                            .getComputedStyle(cell)
                            .getPropertyValue('border-' + val + '-width')
                            .replace('px', '')));
                        var arrRGB = [];
                        arrRGB = window
                            .getComputedStyle(cell)
                            .getPropertyValue('border-' + val + '-color')
                            .replace(/\s+/gi, '')
                            .replace('rgba(', '')
                            .replace('rgb(', '')
                            .replace(')', '')
                            .split(',');
                        var strBorderC = rgbToHex(Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2]));
                        cellOpts.border[idxb] = { pt: intBorderW, color: strBorderC };
                    });
                }
                // LAST: Add cell
                arrObjTabCells.push({
                    _type: SLIDE_OBJECT_TYPES.tablecell,
                    text: cell.innerText,
                    options: cellOpts,
                });
            });
            switch (part) {
                case 'thead':
                    arrObjTabHeadRows.push(arrObjTabCells);
                    break;
                case 'tbody':
                    arrObjTabBodyRows.push(arrObjTabCells);
                    break;
                case 'tfoot':
                    arrObjTabFootRows.push(arrObjTabCells);
                    break;
                default:
                    console.log("table parsing: unexpected table part: " + part);
                    break;
            }
        });
    });
    // STEP 5: Break table into Slides as needed
    // Pass head-rows as there is an option to add to each table and the parse func needs this data to fulfill that option
    opts._arrObjTabHeadRows = arrObjTabHeadRows || null;
    opts.colW = arrColW;
    getSlidesForTableRows(__spreadArray(__spreadArray(__spreadArray([], arrObjTabHeadRows), arrObjTabBodyRows), arrObjTabFootRows), opts, pptx.presLayout, masterSlide).forEach(function (slide, idxTr) {
        // A: Create new Slide
        var newSlide = pptx.addSlide({ masterName: opts.masterSlideName || null });
        // B: DESIGN: Reset `y` to startY or margin after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
        if (idxTr === 0)
            opts.y = opts.y || arrInchMargins[0];
        if (idxTr > 0)
            opts.y = opts.autoPageSlideStartY || opts.newSlideStartY || arrInchMargins[0];
        if (opts.verbose)
            console.log('opts.autoPageSlideStartY:' + opts.autoPageSlideStartY + ' / arrInchMargins[0]:' + arrInchMargins[0] + ' => opts.y = ' + opts.y);
        // C: Add table to Slide
        newSlide.addTable(slide.rows, { x: opts.x || arrInchMargins[3], y: opts.y, w: Number(emuSlideTabW) / EMU, colW: arrColW, autoPage: false });
        // D: Add any additional objects
        if (opts.addImage)
            newSlide.addImage({ path: opts.addImage.url, x: opts.addImage.x, y: opts.addImage.y, w: opts.addImage.w, h: opts.addImage.h });
        if (opts.addShape)
            newSlide.addShape(opts.addShape.shape, opts.addShape.options || {});
        if (opts.addTable)
            newSlide.addTable(opts.addTable.rows, opts.addTable.options || {});
        if (opts.addText)
            newSlide.addText(opts.addText.text, opts.addText.options || {});
    });
}

/**
 * PptxGenJS: XML Generation
 */
var imageSizingXml = {
    cover: function (imgSize, boxDim) {
        var imgRatio = imgSize.h / imgSize.w, boxRatio = boxDim.h / boxDim.w, isBoxBased = boxRatio > imgRatio, width = isBoxBased ? boxDim.h / imgRatio : boxDim.w, height = isBoxBased ? boxDim.h : boxDim.w * imgRatio, hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)), vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height));
        return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '"/><a:stretch/>';
    },
    contain: function (imgSize, boxDim) {
        var imgRatio = imgSize.h / imgSize.w, boxRatio = boxDim.h / boxDim.w, widthBased = boxRatio > imgRatio, width = widthBased ? boxDim.w : boxDim.h / imgRatio, height = widthBased ? boxDim.w * imgRatio : boxDim.h, hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)), vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height));
        return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '"/><a:stretch/>';
    },
    crop: function (imageSize, boxDim) {
        var l = boxDim.x, r = imageSize.w - (boxDim.x + boxDim.w), t = boxDim.y, b = imageSize.h - (boxDim.y + boxDim.h), lPerc = Math.round(1e5 * (l / imageSize.w)), rPerc = Math.round(1e5 * (r / imageSize.w)), tPerc = Math.round(1e5 * (t / imageSize.h)), bPerc = Math.round(1e5 * (b / imageSize.h));
        return '<a:srcRect l="' + lPerc + '" r="' + rPerc + '" t="' + tPerc + '" b="' + bPerc + '"/><a:stretch/>';
    },
};
/**
 * Transforms a slide or slideLayout to resulting XML string - Creates `ppt/slide*.xml`
 * @param {PresSlide|SlideLayout} slideObject - slide object created within createSlideObject
 * @return {string} XML string with <p:cSld> as the root
 */
function slideObjectToXml(slide) {
    var strSlideXml = slide._name ? '<p:cSld name="' + slide._name + '">' : '<p:cSld>';
    var intTableNum = 1;
    // STEP 1: Add background color/image (ensure only a single `<p:bg>` tag is created, ex: when master-baskground has both `color` and `path`)
    if (slide._bkgdImgRid) {
        strSlideXml += "<p:bg><p:bgPr><a:blipFill dpi=\"0\" rotWithShape=\"1\"><a:blip r:embed=\"rId" + slide._bkgdImgRid + "\"><a:lum/></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>";
    }
    else if (slide.background && slide.background.color) {
        strSlideXml += "<p:bg><p:bgPr>" + genXmlColorSelection(slide.background) + "</p:bgPr></p:bg>";
    }
    else if (!slide.bkgd && slide._name && slide._name === DEF_PRES_LAYOUT_NAME) {
        // NOTE: Default [white] background is needed on slideMaster1.xml to avoid gray background in Keynote (and Finder previews)
        strSlideXml += "<p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg>";
    }
    // STEP 2: Continue slide by starting spTree node
    strSlideXml += '<p:spTree>';
    strSlideXml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>';
    strSlideXml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>';
    strSlideXml += '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
    // STEP 3: Loop over all Slide.data objects and add them to this slide
    slide._slideObjects.forEach(function (slideItemObj, idx) {
        var _a;
        var x = 0, y = 0, cx = getSmartParseNumber('75%', 'X', slide._presLayout), cy = 0;
        var placeholderObj;
        var locationAttr = '';
        if (slide._slideLayout !== undefined &&
            slide._slideLayout._slideObjects !== undefined &&
            slideItemObj.options &&
            slideItemObj.options.placeholder) {
            placeholderObj = slide._slideLayout._slideObjects.filter(function (object) { return object.options.placeholder === slideItemObj.options.placeholder; })[0];
        }
        // A: Set option vars
        slideItemObj.options = slideItemObj.options || {};
        if (typeof slideItemObj.options.x !== 'undefined')
            x = getSmartParseNumber(slideItemObj.options.x, 'X', slide._presLayout);
        if (typeof slideItemObj.options.y !== 'undefined')
            y = getSmartParseNumber(slideItemObj.options.y, 'Y', slide._presLayout);
        if (typeof slideItemObj.options.w !== 'undefined')
            cx = getSmartParseNumber(slideItemObj.options.w, 'X', slide._presLayout);
        if (typeof slideItemObj.options.h !== 'undefined')
            cy = getSmartParseNumber(slideItemObj.options.h, 'Y', slide._presLayout);
        // If using a placeholder then inherit it's position
        if (placeholderObj) {
            if (placeholderObj.options.x || placeholderObj.options.x === 0)
                x = getSmartParseNumber(placeholderObj.options.x, 'X', slide._presLayout);
            if (placeholderObj.options.y || placeholderObj.options.y === 0)
                y = getSmartParseNumber(placeholderObj.options.y, 'Y', slide._presLayout);
            if (placeholderObj.options.w || placeholderObj.options.w === 0)
                cx = getSmartParseNumber(placeholderObj.options.w, 'X', slide._presLayout);
            if (placeholderObj.options.h || placeholderObj.options.h === 0)
                cy = getSmartParseNumber(placeholderObj.options.h, 'Y', slide._presLayout);
        }
        //
        if (slideItemObj.options.flipH)
            locationAttr += ' flipH="1"';
        if (slideItemObj.options.flipV)
            locationAttr += ' flipV="1"';
        if (slideItemObj.options.rotate)
            locationAttr += ' rot="' + convertRotationDegrees(slideItemObj.options.rotate) + '"';
        // B: Add OBJECT to the current Slide
        switch (slideItemObj._type) {
            case SLIDE_OBJECT_TYPES.table:
                var arrTabRows_1 = slideItemObj.arrTabRows;
                var objTabOpts_1 = slideItemObj.options;
                var intColCnt_1 = 0, intColW = 0;
                var cellOpts_1;
                // Calc number of columns
                // NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
                // ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
                arrTabRows_1[0].forEach(function (cell) {
                    cellOpts_1 = cell.options || null;
                    intColCnt_1 += cellOpts_1 && cellOpts_1.colspan ? Number(cellOpts_1.colspan) : 1;
                });
                // STEP 1: Start Table XML
                // NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
                var strXml_1 = "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"" + (intTableNum * slide._slideNum + 1) + "\" name=\"Table " + intTableNum * slide._slideNum + "\"/>";
                strXml_1 +=
                    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
                        '  <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>' +
                        '</p:nvGraphicFramePr>';
                strXml_1 += "<p:xfrm><a:off x=\"" + (x || (x === 0 ? 0 : EMU)) + "\" y=\"" + (y || (y === 0 ? 0 : EMU)) + "\"/><a:ext cx=\"" + (cx || (cx === 0 ? 0 : EMU)) + "\" cy=\"" + (cy || EMU) + "\"/></p:xfrm>";
                strXml_1 += '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>';
                // + '        <a:tblPr bandRow="1"/>';
                // TODO: Support banded rows, first/last row, etc.
                // NOTE: Banding, etc. only shows when using a table style! (or set alt row color if banding)
                // <a:tblPr firstCol="0" firstRow="0" lastCol="0" lastRow="0" bandCol="0" bandRow="1">
                // STEP 2: Set column widths
                // Evenly distribute cols/rows across size provided when applicable (calc them if only overall dimensions were provided)
                // A: Col widths provided?
                if (Array.isArray(objTabOpts_1.colW)) {
                    strXml_1 += '<a:tblGrid>';
                    for (var col = 0; col < intColCnt_1; col++) {
                        var w = inch2Emu(objTabOpts_1.colW[col]);
                        if (w == null || isNaN(w)) {
                            w = (typeof slideItemObj.options.w === 'number' ? slideItemObj.options.w : 1) / intColCnt_1;
                        }
                        strXml_1 += '<a:gridCol w="' + Math.round(w) + '"/>';
                    }
                    strXml_1 += '</a:tblGrid>';
                }
                // B: Table Width provided without colW? Then distribute cols
                else {
                    intColW = objTabOpts_1.colW ? objTabOpts_1.colW : EMU;
                    if (slideItemObj.options.w && !objTabOpts_1.colW)
                        intColW = Math.round((typeof slideItemObj.options.w === 'number' ? slideItemObj.options.w : 1) / intColCnt_1);
                    strXml_1 += '<a:tblGrid>';
                    for (var colw = 0; colw < intColCnt_1; colw++) {
                        strXml_1 += '<a:gridCol w="' + intColW + '"/>';
                    }
                    strXml_1 += '</a:tblGrid>';
                }
                // STEP 3: Build our row arrays into an actual grid to match the XML we will be building next (ISSUE #36)
                // Note row arrays can arrive "lopsided" as in row1:[1,2,3] row2:[3] when first two cols rowspan!,
                // so a simple loop below in XML building wont suffice to build table correctly.
                // We have to build an actual grid now
                /*
                    EX: (A0:rowspan=3, B1:rowspan=2, C1:colspan=2)

                    /------|------|------|------\
                    |  A0  |  B0  |  C0  |  D0  |
                    |      |  B1  |  C1  |      |
                    |      |      |  C2  |  D2  |
                    \------|------|------|------/
                */
                // A: add _hmerge cell for colspan. should reserve rowspan
                arrTabRows_1.forEach(function (cells) {
                    var _a, _b;
                    var _loop_1 = function (cIdx) {
                        var cell = cells[cIdx];
                        var colspan = (_a = cell.options) === null || _a === void 0 ? void 0 : _a.colspan;
                        var rowspan = (_b = cell.options) === null || _b === void 0 ? void 0 : _b.rowspan;
                        if (colspan && colspan > 1) {
                            var vMergeCells = new Array(colspan - 1).fill(undefined).map(function (_) {
                                return { _type: SLIDE_OBJECT_TYPES.tablecell, options: { rowspan: rowspan }, _hmerge: true };
                            });
                            cells.splice.apply(cells, __spreadArray([cIdx + 1, 0], vMergeCells));
                            cIdx += colspan;
                        }
                        else {
                            cIdx += 1;
                        }
                        out_cIdx_1 = cIdx;
                    };
                    var out_cIdx_1;
                    for (var cIdx = 0; cIdx < cells.length;) {
                        _loop_1(cIdx);
                        cIdx = out_cIdx_1;
                    }
                });
                // B: add _vmerge cell for rowspan. should reserve colspan/_hmerge
                arrTabRows_1.forEach(function (cells, rIdx) {
                    var nextRow = arrTabRows_1[rIdx + 1];
                    if (!nextRow)
                        return;
                    cells.forEach(function (cell, cIdx) {
                        var _a, _b;
                        var rowspan = cell._rowContinue || ((_a = cell.options) === null || _a === void 0 ? void 0 : _a.rowspan);
                        var colspan = (_b = cell.options) === null || _b === void 0 ? void 0 : _b.colspan;
                        var _hmerge = cell._hmerge;
                        if (rowspan && rowspan > 1) {
                            var hMergeCell = { _type: SLIDE_OBJECT_TYPES.tablecell, options: { colspan: colspan }, _rowContinue: rowspan - 1, _vmerge: true, _hmerge: _hmerge };
                            nextRow.splice(cIdx, 0, hMergeCell);
                        }
                    });
                });
                // STEP 4: Build table rows/cells
                arrTabRows_1.forEach(function (cells, rIdx) {
                    // A: Table Height provided without rowH? Then distribute rows
                    var intRowH = 0; // IMPORTANT: Default must be zero for auto-sizing to work
                    if (Array.isArray(objTabOpts_1.rowH) && objTabOpts_1.rowH[rIdx])
                        intRowH = inch2Emu(Number(objTabOpts_1.rowH[rIdx]));
                    else if (objTabOpts_1.rowH && !isNaN(Number(objTabOpts_1.rowH)))
                        intRowH = inch2Emu(Number(objTabOpts_1.rowH));
                    else if (slideItemObj.options.cy || slideItemObj.options.h)
                        intRowH = Math.round((slideItemObj.options.h ? inch2Emu(slideItemObj.options.h) : typeof slideItemObj.options.cy === 'number' ? slideItemObj.options.cy : 1) /
                            arrTabRows_1.length);
                    // B: Start row
                    strXml_1 += "<a:tr h=\"" + intRowH + "\">";
                    // C: Loop over each CELL
                    cells.forEach(function (cellObj) {
                        var _a, _b;
                        var cell = cellObj;
                        var cellSpanAttrs = {
                            rowSpan: ((_a = cell.options) === null || _a === void 0 ? void 0 : _a.rowspan) > 1 ? cell.options.rowspan : undefined,
                            gridSpan: ((_b = cell.options) === null || _b === void 0 ? void 0 : _b.colspan) > 1 ? cell.options.colspan : undefined,
                            vMerge: cell._vmerge ? 1 : undefined,
                            hMerge: cell._hmerge ? 1 : undefined,
                        };
                        var cellSpanAttrStr = Object.keys(cellSpanAttrs)
                            .map(function (k) { return [k, cellSpanAttrs[k]]; })
                            .filter(function (_a) {
                            _a[0]; var v = _a[1];
                            return !!v;
                        })
                            .map(function (_a) {
                            var k = _a[0], v = _a[1];
                            return k + "=\"" + v + "\"";
                        })
                            .join(' ');
                        if (cellSpanAttrStr)
                            cellSpanAttrStr = ' ' + cellSpanAttrStr;
                        // 1: COLSPAN/ROWSPAN: Add dummy cells for any active colspan/rowspan
                        if (cell._hmerge || cell._vmerge) {
                            strXml_1 += "<a:tc" + cellSpanAttrStr + "><a:tcPr/></a:tc>";
                            return;
                        }
                        // 2: OPTIONS: Build/set cell options
                        var cellOpts = cell.options || {};
                        cell.options = cellOpts;
                        ['align', 'bold', 'border', 'color', 'fill', 'fontFace', 'fontSize', 'margin', 'underline', 'valign'].forEach(function (name) {
                            if (objTabOpts_1[name] && !cellOpts[name] && cellOpts[name] !== 0)
                                cellOpts[name] = objTabOpts_1[name];
                        });
                        var cellValign = cellOpts.valign
                            ? ' anchor="' +
                                cellOpts.valign
                                    .replace(/^c$/i, 'ctr')
                                    .replace(/^m$/i, 'ctr')
                                    .replace('center', 'ctr')
                                    .replace('middle', 'ctr')
                                    .replace('top', 't')
                                    .replace('btm', 'b')
                                    .replace('bottom', 'b') +
                                '"'
                            : '';
                        var fillColor = cell._optImp && cell._optImp.fill && cell._optImp.fill.color
                            ? cell._optImp.fill.color
                            : cell._optImp && cell._optImp.fill && typeof cell._optImp.fill === 'string'
                                ? cell._optImp.fill
                                : '';
                        fillColor =
                            fillColor || (cellOpts.fill && cellOpts.fill.color) ? cellOpts.fill.color : cellOpts.fill && typeof cellOpts.fill === 'string' ? cellOpts.fill : '';
                        var cellFill = fillColor ? "<a:solidFill>" + createColorElement(fillColor) + "</a:solidFill>" : '';
                        var cellMargin = cellOpts.margin === 0 || cellOpts.margin ? cellOpts.margin : DEF_CELL_MARGIN_PT;
                        if (!Array.isArray(cellMargin) && typeof cellMargin === 'number')
                            cellMargin = [cellMargin, cellMargin, cellMargin, cellMargin];
                        var cellMarginXml = " marL=\"" + valToPts(cellMargin[3]) + "\" marR=\"" + valToPts(cellMargin[1]) + "\" marT=\"" + valToPts(cellMargin[0]) + "\" marB=\"" + valToPts(cellMargin[2]) + "\"";
                        // FUTURE: Cell NOWRAP property (textwrap: add to a:tcPr (horzOverflow="overflow" or whatever options exist)
                        // 4: Set CELL content and properties ==================================
                        strXml_1 += "<a:tc" + cellSpanAttrStr + ">" + genXmlTextBody(cell) + "<a:tcPr" + cellMarginXml + cellValign + ">";
                        //strXml += `<a:tc${cellColspan}${cellRowspan}>${genXmlTextBody(cell)}<a:tcPr${cellMarginXml}${cellValign}${cellTextDir}>`
                        // FIXME: 20200525: ^^^
                        // <a:tcPr marL="38100" marR="38100" marT="38100" marB="38100" vert="vert270">
                        // 5: Borders: Add any borders
                        if (cellOpts.border && Array.isArray(cellOpts.border)) {
                            [
                                { idx: 3, name: 'lnL' },
                                { idx: 1, name: 'lnR' },
                                { idx: 0, name: 'lnT' },
                                { idx: 2, name: 'lnB' },
                            ].forEach(function (obj) {
                                if (cellOpts.border[obj.idx].type !== 'none') {
                                    strXml_1 += "<a:" + obj.name + " w=\"" + valToPts(cellOpts.border[obj.idx].pt) + "\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">";
                                    strXml_1 += "<a:solidFill>" + createColorElement(cellOpts.border[obj.idx].color) + "</a:solidFill>";
                                    strXml_1 += "<a:prstDash val=\"" + (cellOpts.border[obj.idx].type === 'dash' ? 'sysDash' : 'solid') + "\"/><a:round/><a:headEnd type=\"none\" w=\"med\" len=\"med\"/><a:tailEnd type=\"none\" w=\"med\" len=\"med\"/>";
                                    strXml_1 += "</a:" + obj.name + ">";
                                }
                                else {
                                    strXml_1 += "<a:" + obj.name + " w=\"0\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:noFill/></a:" + obj.name + ">";
                                }
                            });
                        }
                        // 6: Close cell Properties & Cell
                        strXml_1 += cellFill;
                        strXml_1 += '  </a:tcPr>';
                        strXml_1 += ' </a:tc>';
                    });
                    // D: Complete row
                    strXml_1 += '</a:tr>';
                });
                // STEP 5: Complete table
                strXml_1 += '      </a:tbl>';
                strXml_1 += '    </a:graphicData>';
                strXml_1 += '  </a:graphic>';
                strXml_1 += '</p:graphicFrame>';
                // STEP 6: Set table XML
                strSlideXml += strXml_1;
                // LAST: Increment counter
                intTableNum++;
                break;
            case SLIDE_OBJECT_TYPES.text:
            case SLIDE_OBJECT_TYPES.placeholder:
                var shapeName = slideItemObj.options.shapeName ? encodeXmlEntities(slideItemObj.options.shapeName) : "Object" + (idx + 1);
                // Lines can have zero cy, but text should not
                if (!slideItemObj.options.line && cy === 0)
                    cy = EMU * 0.3;
                // Margin/Padding/Inset for textboxes
                if (!slideItemObj.options._bodyProp)
                    slideItemObj.options._bodyProp = {};
                if (slideItemObj.options.margin && Array.isArray(slideItemObj.options.margin)) {
                    slideItemObj.options._bodyProp.lIns = valToPts(slideItemObj.options.margin[0] || 0);
                    slideItemObj.options._bodyProp.rIns = valToPts(slideItemObj.options.margin[1] || 0);
                    slideItemObj.options._bodyProp.bIns = valToPts(slideItemObj.options.margin[2] || 0);
                    slideItemObj.options._bodyProp.tIns = valToPts(slideItemObj.options.margin[3] || 0);
                }
                else if (typeof slideItemObj.options.margin === 'number') {
                    slideItemObj.options._bodyProp.lIns = valToPts(slideItemObj.options.margin);
                    slideItemObj.options._bodyProp.rIns = valToPts(slideItemObj.options.margin);
                    slideItemObj.options._bodyProp.bIns = valToPts(slideItemObj.options.margin);
                    slideItemObj.options._bodyProp.tIns = valToPts(slideItemObj.options.margin);
                }
                // A: Start SHAPE =======================================================
                strSlideXml += '<p:sp>';
                // B: The addition of the "txBox" attribute is the sole determiner of if an object is a shape or textbox
                strSlideXml += "<p:nvSpPr><p:cNvPr id=\"" + (idx + 2) + "\" name=\"" + shapeName + "\">";
                // <Hyperlink>
                if (slideItemObj.options.hyperlink && slideItemObj.options.hyperlink.url)
                    strSlideXml +=
                        '<a:hlinkClick r:id="rId' +
                            slideItemObj.options.hyperlink._rId +
                            '" tooltip="' +
                            (slideItemObj.options.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.options.hyperlink.tooltip) : '') +
                            '"/>';
                if (slideItemObj.options.hyperlink && slideItemObj.options.hyperlink.slide)
                    strSlideXml +=
                        '<a:hlinkClick r:id="rId' +
                            slideItemObj.options.hyperlink._rId +
                            '" tooltip="' +
                            (slideItemObj.options.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.options.hyperlink.tooltip) : '') +
                            '" action="ppaction://hlinksldjump"/>';
                // </Hyperlink>
                strSlideXml += '</p:cNvPr>';
                strSlideXml += '<p:cNvSpPr' + (slideItemObj.options && slideItemObj.options.isTextBox ? ' txBox="1"/>' : '/>');
                strSlideXml += "<p:nvPr>" + (slideItemObj._type === 'placeholder' ? genXmlPlaceholder(slideItemObj) : genXmlPlaceholder(placeholderObj)) + "</p:nvPr>";
                strSlideXml += '</p:nvSpPr><p:spPr>';
                strSlideXml += "<a:xfrm" + locationAttr + ">";
                strSlideXml += "<a:off x=\"" + x + "\" y=\"" + y + "\"/>";
                strSlideXml += "<a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/></a:xfrm>";
                if (slideItemObj.shape === 'custGeom') {
                    strSlideXml += '<a:custGeom><a:avLst />';
                    strSlideXml += '<a:gdLst>';
                    strSlideXml += '</a:gdLst>';
                    strSlideXml += '<a:ahLst />';
                    strSlideXml += '<a:cxnLst>';
                    strSlideXml += '</a:cxnLst>';
                    strSlideXml += '<a:rect l="l" t="t" r="r" b="b" />';
                    strSlideXml += '<a:pathLst>';
                    strSlideXml += "<a:path w=\"" + cx + "\" h=\"" + cy + "\">";
                    (_a = slideItemObj.options.points) === null || _a === void 0 ? void 0 : _a.map(function (point, i) {
                        if ('curve' in point) {
                            switch (point.curve.type) {
                                case 'arc':
                                    strSlideXml += "<a:arcTo hR=\"" + getSmartParseNumber(point.curve.hR, 'Y', slide._presLayout) + "\" wR=\"" + getSmartParseNumber(point.curve.wR, 'X', slide._presLayout) + "\" stAng=\"" + convertRotationDegrees(point.curve.stAng) + "\" swAng=\"" + convertRotationDegrees(point.curve.swAng) + "\" />";
                                    break;
                                case 'cubic':
                                    strSlideXml += "<a:cubicBezTo>\n\t\t\t\t\t\t\t\t\t<a:pt x=\"" + getSmartParseNumber(point.curve.x1, 'X', slide._presLayout) + "\" y=\"" + getSmartParseNumber(point.curve.y1, 'Y', slide._presLayout) + "\" />\n\t\t\t\t\t\t\t\t\t<a:pt x=\"" + getSmartParseNumber(point.curve.x2, 'X', slide._presLayout) + "\" y=\"" + getSmartParseNumber(point.curve.y2, 'Y', slide._presLayout) + "\" />\n\t\t\t\t\t\t\t\t\t<a:pt x=\"" + getSmartParseNumber(point.x, 'X', slide._presLayout) + "\" y=\"" + getSmartParseNumber(point.y, 'Y', slide._presLayout) + "\" />\n\t\t\t\t\t\t\t\t\t</a:cubicBezTo>";
                                    break;
                                case 'quadratic':
                                    strSlideXml += "<a:quadBezTo>\n\t\t\t\t\t\t\t\t\t<a:pt x=\"" + getSmartParseNumber(point.curve.x1, 'X', slide._presLayout) + "\" y=\"" + getSmartParseNumber(point.curve.y1, 'Y', slide._presLayout) + "\" />\n\t\t\t\t\t\t\t\t\t<a:pt x=\"" + getSmartParseNumber(point.x, 'X', slide._presLayout) + "\" y=\"" + getSmartParseNumber(point.y, 'Y', slide._presLayout) + "\" />\n\t\t\t\t\t\t\t\t\t</a:quadBezTo>";
                                    break;
                            }
                        }
                        else if ('close' in point) {
                            strSlideXml += "<a:close />";
                        }
                        else if (point.moveTo || i === 0) {
                            strSlideXml += "<a:moveTo><a:pt x=\"" + getSmartParseNumber(point.x, 'X', slide._presLayout) + "\" y=\"" + getSmartParseNumber(point.y, 'Y', slide._presLayout) + "\" /></a:moveTo>";
                        }
                        else {
                            strSlideXml += "<a:lnTo><a:pt x=\"" + getSmartParseNumber(point.x, 'X', slide._presLayout) + "\" y=\"" + getSmartParseNumber(point.y, 'Y', slide._presLayout) + "\" /></a:lnTo>";
                        }
                    });
                    strSlideXml += '</a:path>';
                    strSlideXml += '</a:pathLst>';
                    strSlideXml += '</a:custGeom>';
                }
                else {
                    strSlideXml += '<a:prstGeom prst="' + slideItemObj.shape + '"><a:avLst>';
                    if (slideItemObj.options.rectRadius) {
                        strSlideXml += "<a:gd name=\"adj\" fmla=\"val " + Math.round((slideItemObj.options.rectRadius * EMU * 100000) / Math.min(cx, cy)) + "\"/>";
                    }
                    else if (slideItemObj.options.angleRange) {
                        for (var i = 0; i < 2; i++) {
                            var angle = slideItemObj.options.angleRange[i];
                            strSlideXml += "<a:gd name=\"adj" + (i + 1) + "\" fmla=\"val " + convertRotationDegrees(angle) + "\" />";
                        }
                        if (slideItemObj.options.arcThicknessRatio) {
                            strSlideXml += "<a:gd name=\"adj3\" fmla=\"val " + Math.round(slideItemObj.options.arcThicknessRatio * 50000) + "\" />";
                        }
                    }
                    strSlideXml += '</a:avLst></a:prstGeom>';
                }
                // Option: FILL
                strSlideXml += slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : '<a:noFill/>';
                // shape Type: LINE: line color
                if (slideItemObj.options.line) {
                    strSlideXml += slideItemObj.options.line.width ? "<a:ln w=\"" + valToPts(slideItemObj.options.line.width) + "\">" : '<a:ln>';
                    if (slideItemObj.options.line.color)
                        strSlideXml += genXmlColorSelection(slideItemObj.options.line);
                    if (slideItemObj.options.line.dashType)
                        strSlideXml += "<a:prstDash val=\"" + slideItemObj.options.line.dashType + "\"/>";
                    if (slideItemObj.options.line.beginArrowType)
                        strSlideXml += "<a:headEnd type=\"" + slideItemObj.options.line.beginArrowType + "\"/>";
                    if (slideItemObj.options.line.endArrowType)
                        strSlideXml += "<a:tailEnd type=\"" + slideItemObj.options.line.endArrowType + "\"/>";
                    // FUTURE: `endArrowSize` < a: headEnd type = "arrow" w = "lg" len = "lg" /> 'sm' | 'med' | 'lg'(values are 1 - 9, making a 3x3 grid of w / len possibilities)
                    strSlideXml += '</a:ln>';
                }
                // EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
                if (slideItemObj.options.shadow) {
                    slideItemObj.options.shadow.type = slideItemObj.options.shadow.type || 'outer';
                    slideItemObj.options.shadow.blur = valToPts(slideItemObj.options.shadow.blur || 8);
                    slideItemObj.options.shadow.offset = valToPts(slideItemObj.options.shadow.offset || 4);
                    slideItemObj.options.shadow.angle = Math.round((slideItemObj.options.shadow.angle || 270) * 60000);
                    slideItemObj.options.shadow.opacity = Math.round((slideItemObj.options.shadow.opacity || 0.75) * 100000);
                    slideItemObj.options.shadow.color = slideItemObj.options.shadow.color || DEF_TEXT_SHADOW.color;
                    strSlideXml += '<a:effectLst>';
                    strSlideXml += '<a:' + slideItemObj.options.shadow.type + 'Shdw sx="100000" sy="100000" kx="0" ky="0" ';
                    strSlideXml += ' algn="bl" rotWithShape="0" blurRad="' + slideItemObj.options.shadow.blur + '" ';
                    strSlideXml += ' dist="' + slideItemObj.options.shadow.offset + '" dir="' + slideItemObj.options.shadow.angle + '">';
                    strSlideXml += '<a:srgbClr val="' + slideItemObj.options.shadow.color + '">';
                    strSlideXml += '<a:alpha val="' + slideItemObj.options.shadow.opacity + '"/></a:srgbClr>';
                    strSlideXml += '</a:outerShdw>';
                    strSlideXml += '</a:effectLst>';
                }
                /* TODO: FUTURE: Text wrapping (copied from MS-PPTX export)
                    // Commented out b/c i'm not even sure this works - current code produces text that wraps in shapes and textboxes, so...
                    if ( slideItemObj.options.textWrap ) {
                        strSlideXml += '<a:extLst>'
                                    + '<a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">'
                                    + '<ma14:wrappingTextBoxFlag xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main" val="1"/>'
                                    + '</a:ext>'
                                    + '</a:extLst>';
                    }
                    */
                // B: Close shape Properties
                strSlideXml += '</p:spPr>';
                // C: Add formatted text (text body "bodyPr")
                strSlideXml += genXmlTextBody(slideItemObj);
                // LAST: Close SHAPE =======================================================
                strSlideXml += '</p:sp>';
                break;
            case SLIDE_OBJECT_TYPES.image:
                var imageOpts = slideItemObj.options;
                var sizing = imageOpts.sizing, rounding = imageOpts.rounding, width = cx, height = cy;
                strSlideXml += '<p:pic>';
                strSlideXml += '  <p:nvPicPr>';
                strSlideXml += "<p:cNvPr id=\"" + (idx + 2) + "\" name=\"Object " + (idx + 1) + "\" descr=\"" + encodeXmlEntities(imageOpts.altText || slideItemObj.image) + "\">";
                if (slideItemObj.hyperlink && slideItemObj.hyperlink.url)
                    strSlideXml += "<a:hlinkClick r:id=\"rId" + slideItemObj.hyperlink._rId + "\" tooltip=\"" + (slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : '') + "\"/>";
                if (slideItemObj.hyperlink && slideItemObj.hyperlink.slide)
                    strSlideXml += "<a:hlinkClick r:id=\"rId" + slideItemObj.hyperlink._rId + "\" tooltip=\"" + (slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : '') + "\" action=\"ppaction://hlinksldjump\"/>";
                strSlideXml += '    </p:cNvPr>';
                strSlideXml += '    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>';
                strSlideXml += '    <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>';
                strSlideXml += '  </p:nvPicPr>';
                strSlideXml += '<p:blipFill>';
                // NOTE: This works for both cases: either `path` or `data` contains the SVG
                if ((slide._relsMedia || []).filter(function (rel) { return rel.rId === slideItemObj.imageRid; })[0] &&
                    (slide._relsMedia || []).filter(function (rel) { return rel.rId === slideItemObj.imageRid; })[0]['extn'] === 'svg') {
                    strSlideXml += '<a:blip r:embed="rId' + (slideItemObj.imageRid - 1) + '">';
                    strSlideXml += ' <a:extLst>';
                    strSlideXml += '  <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">';
                    strSlideXml += '   <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId' + slideItemObj.imageRid + '"/>';
                    strSlideXml += '  </a:ext>';
                    strSlideXml += ' </a:extLst>';
                    strSlideXml += '</a:blip>';
                }
                else {
                    strSlideXml += '<a:blip r:embed="rId' + slideItemObj.imageRid + '"/>';
                }
                if (sizing && sizing.type) {
                    var boxW = sizing.w ? getSmartParseNumber(sizing.w, 'X', slide._presLayout) : cx, boxH = sizing.h ? getSmartParseNumber(sizing.h, 'Y', slide._presLayout) : cy, boxX = getSmartParseNumber(sizing.x || 0, 'X', slide._presLayout), boxY = getSmartParseNumber(sizing.y || 0, 'Y', slide._presLayout);
                    strSlideXml += imageSizingXml[sizing.type]({ w: width, h: height }, { w: boxW, h: boxH, x: boxX, y: boxY });
                    width = boxW;
                    height = boxH;
                }
                else {
                    strSlideXml += '  <a:stretch><a:fillRect/></a:stretch>';
                }
                strSlideXml += '</p:blipFill>';
                strSlideXml += '<p:spPr>';
                strSlideXml += ' <a:xfrm' + locationAttr + '>';
                strSlideXml += '  <a:off x="' + x + '" y="' + y + '"/>';
                strSlideXml += '  <a:ext cx="' + width + '" cy="' + height + '"/>';
                strSlideXml += ' </a:xfrm>';
                strSlideXml += ' <a:prstGeom prst="' + (rounding ? 'ellipse' : 'rect') + '"><a:avLst/></a:prstGeom>';
                strSlideXml += '</p:spPr>';
                strSlideXml += '</p:pic>';
                break;
            case SLIDE_OBJECT_TYPES.media:
                if (slideItemObj.mtype === 'online') {
                    strSlideXml += '<p:pic>';
                    strSlideXml += ' <p:nvPicPr>';
                    // IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preview image rId, PowerPoint throws error!
                    strSlideXml += ' <p:cNvPr id="' + (slideItemObj.mediaRid + 2) + '" name="Picture' + (idx + 1) + '"/>';
                    strSlideXml += ' <p:cNvPicPr/>';
                    strSlideXml += ' <p:nvPr>';
                    strSlideXml += '  <a:videoFile r:link="rId' + slideItemObj.mediaRid + '"/>';
                    strSlideXml += ' </p:nvPr>';
                    strSlideXml += ' </p:nvPicPr>';
                    // NOTE: `blip` is diferent than videos; also there's no preview "p:extLst" above but exists in videos
                    strSlideXml += ' <p:blipFill><a:blip r:embed="rId' + (slideItemObj.mediaRid + 1) + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'; // NOTE: Preview image is required!
                    strSlideXml += ' <p:spPr>';
                    strSlideXml += '  <a:xfrm' + locationAttr + '>';
                    strSlideXml += '   <a:off x="' + x + '" y="' + y + '"/>';
                    strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>';
                    strSlideXml += '  </a:xfrm>';
                    strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
                    strSlideXml += ' </p:spPr>';
                    strSlideXml += '</p:pic>';
                }
                else {
                    strSlideXml += '<p:pic>';
                    strSlideXml += ' <p:nvPicPr>';
                    // IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preiew image rId, PowerPoint throws error!
                    strSlideXml +=
                        ' <p:cNvPr id="' +
                            (slideItemObj.mediaRid + 2) +
                            '" name="' +
                            slideItemObj.media.split('/').pop().split('.').shift() +
                            '"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>';
                    strSlideXml += ' <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>';
                    strSlideXml += ' <p:nvPr>';
                    strSlideXml += '  <a:videoFile r:link="rId' + slideItemObj.mediaRid + '"/>';
                    strSlideXml += '  <p:extLst>';
                    strSlideXml += '   <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">';
                    strSlideXml += '    <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId' + (slideItemObj.mediaRid + 1) + '"/>';
                    strSlideXml += '   </p:ext>';
                    strSlideXml += '  </p:extLst>';
                    strSlideXml += ' </p:nvPr>';
                    strSlideXml += ' </p:nvPicPr>';
                    strSlideXml += ' <p:blipFill><a:blip r:embed="rId' + (slideItemObj.mediaRid + 2) + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'; // NOTE: Preview image is required!
                    strSlideXml += ' <p:spPr>';
                    strSlideXml += '  <a:xfrm' + locationAttr + '>';
                    strSlideXml += '   <a:off x="' + x + '" y="' + y + '"/>';
                    strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>';
                    strSlideXml += '  </a:xfrm>';
                    strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
                    strSlideXml += ' </p:spPr>';
                    strSlideXml += '</p:pic>';
                }
                break;
            case SLIDE_OBJECT_TYPES.chart:
                var chartOpts = slideItemObj.options;
                strSlideXml += '<p:graphicFrame>';
                strSlideXml += ' <p:nvGraphicFramePr>';
                strSlideXml += "   <p:cNvPr id=\"" + (idx + 2) + "\" name=\"Chart " + (idx + 1) + "\" descr=\"" + encodeXmlEntities(chartOpts.altText || '') + "\"/>";
                strSlideXml += '   <p:cNvGraphicFramePr/>';
                strSlideXml += "   <p:nvPr>" + genXmlPlaceholder(placeholderObj) + "</p:nvPr>";
                strSlideXml += ' </p:nvGraphicFramePr>';
                strSlideXml += " <p:xfrm><a:off x=\"" + x + "\" y=\"" + y + "\"/><a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/></p:xfrm>";
                strSlideXml += ' <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
                strSlideXml += '  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">';
                strSlideXml += "   <c:chart r:id=\"rId" + slideItemObj.chartRid + "\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>";
                strSlideXml += '  </a:graphicData>';
                strSlideXml += ' </a:graphic>';
                strSlideXml += '</p:graphicFrame>';
                break;
            default:
                strSlideXml += '';
                break;
        }
    });
    // STEP 4: Add slide numbers (if any) last
    if (slide._slideNumberProps) {
        // Set some defaults (done here b/c SlideNumber canbe added to masters or slides and has numerous entry points)
        if (!slide._slideNumberProps.align)
            slide._slideNumberProps.align = 'left';
        strSlideXml +=
            '<p:sp>' +
                '  <p:nvSpPr>' +
                '    <p:cNvPr id="25" name="Slide Number Placeholder 24"/>' +
                '    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
                '    <p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr>' +
                '  </p:nvSpPr>' +
                '  <p:spPr>' +
                '    <a:xfrm>' +
                '      <a:off x="' +
                getSmartParseNumber(slide._slideNumberProps.x, 'X', slide._presLayout) +
                '" y="' +
                getSmartParseNumber(slide._slideNumberProps.y, 'Y', slide._presLayout) +
                '"/>' +
                '      <a:ext cx="' +
                (slide._slideNumberProps.w ? getSmartParseNumber(slide._slideNumberProps.w, 'X', slide._presLayout) : 800000) +
                '" cy="' +
                (slide._slideNumberProps.h ? getSmartParseNumber(slide._slideNumberProps.h, 'Y', slide._presLayout) : 300000) +
                '"/>' +
                '    </a:xfrm>' +
                '    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
                '    <a:extLst><a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}"><ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext></a:extLst>' +
                '  </p:spPr>';
        strSlideXml += '<p:txBody>';
        strSlideXml += '<a:bodyPr';
        if (slide._slideNumberProps.margin && Array.isArray(slide._slideNumberProps.margin)) {
            strSlideXml += " lIns=\"" + valToPts(slide._slideNumberProps.margin[3] || 0) + "\"";
            strSlideXml += " tIns=\"" + valToPts(slide._slideNumberProps.margin[0] || 0) + "\"";
            strSlideXml += " rIns=\"" + valToPts(slide._slideNumberProps.margin[1] || 0) + "\"";
            strSlideXml += " bIns=\"" + valToPts(slide._slideNumberProps.margin[2] || 0) + "\"";
        }
        else if (typeof slide._slideNumberProps.margin === 'number') {
            strSlideXml += " lIns=\"" + valToPts(slide._slideNumberProps.margin || 0) + "\"";
            strSlideXml += " tIns=\"" + valToPts(slide._slideNumberProps.margin || 0) + "\"";
            strSlideXml += " rIns=\"" + valToPts(slide._slideNumberProps.margin || 0) + "\"";
            strSlideXml += " bIns=\"" + valToPts(slide._slideNumberProps.margin || 0) + "\"";
        }
        strSlideXml += '/>';
        strSlideXml += '  <a:lstStyle><a:lvl1pPr>';
        if (slide._slideNumberProps.fontFace || slide._slideNumberProps.fontSize || slide._slideNumberProps.color) {
            strSlideXml += "<a:defRPr sz=\"" + Math.round((slide._slideNumberProps.fontSize || 12) * 100) + "\">";
            if (slide._slideNumberProps.color)
                strSlideXml += genXmlColorSelection(slide._slideNumberProps.color);
            if (slide._slideNumberProps.fontFace)
                strSlideXml += "<a:latin typeface=\"" + slide._slideNumberProps.fontFace + "\"/><a:ea typeface=\"" + slide._slideNumberProps.fontFace + "\"/><a:cs typeface=\"" + slide._slideNumberProps.fontFace + "\"/>";
            strSlideXml += '</a:defRPr>';
        }
        strSlideXml += '</a:lvl1pPr></a:lstStyle>';
        strSlideXml += "<a:p><a:fld id=\"" + SLDNUMFLDID + "\" type=\"slidenum\"><a:rPr lang=\"en-US\"/>";
        if (slide._slideNumberProps.align.startsWith('l'))
            strSlideXml += '<a:pPr algn="l"/>';
        else if (slide._slideNumberProps.align.startsWith('c'))
            strSlideXml += '<a:pPr algn="ctr"/>';
        else if (slide._slideNumberProps.align.startsWith('r'))
            strSlideXml += '<a:pPr algn="r"/>';
        else
            strSlideXml += "<a:pPr algn=\"l\"/>";
        strSlideXml += "<a:t></a:t></a:fld><a:endParaRPr lang=\"en-US\"/></a:p>";
        strSlideXml += '</p:txBody></p:sp>';
    }
    // STEP 5: Close spTree and finalize slide XML
    strSlideXml += '</p:spTree>';
    strSlideXml += '</p:cSld>';
    // LAST: Return
    return strSlideXml;
}
/**
 * Transforms slide relations to XML string.
 * Extra relations that are not dynamic can be passed using the 2nd arg (e.g. theme relation in master file).
 * These relations use rId series that starts with 1-increased maximum of rIds used for dynamic relations.
 * @param {PresSlide | SlideLayout} slide - slide object whose relations are being transformed
 * @param {{ target: string; type: string }[]} defaultRels - array of default relations
 * @return {string} XML
 */
function slideObjectRelationsToXml(slide, defaultRels) {
    var lastRid = 0; // stores maximum rId used for dynamic relations
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    // STEP 1: Add all rels for this Slide
    slide._rels.forEach(function (rel) {
        lastRid = Math.max(lastRid, rel.rId);
        if (rel.type.toLowerCase().indexOf('hyperlink') > -1) {
            if (rel.data === 'slide') {
                strXml +=
                    '<Relationship Id="rId' +
                        rel.rId +
                        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"' +
                        ' Target="slide' +
                        rel.Target +
                        '.xml"/>';
            }
            else {
                strXml +=
                    '<Relationship Id="rId' +
                        rel.rId +
                        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"' +
                        ' Target="' +
                        rel.Target +
                        '" TargetMode="External"/>';
            }
        }
        else if (rel.type.toLowerCase().indexOf('notesSlide') > -1) {
            strXml +=
                '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"/>';
        }
    });
    (slide._relsChart || []).forEach(function (rel) {
        lastRid = Math.max(lastRid, rel.rId);
        strXml += '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="' + rel.Target + '"/>';
    });
    (slide._relsMedia || []).forEach(function (rel) {
        lastRid = Math.max(lastRid, rel.rId);
        if (rel.type.toLowerCase().indexOf('image') > -1) {
            strXml += '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + rel.Target + '"/>';
        }
        else if (rel.type.toLowerCase().indexOf('audio') > -1) {
            // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
            if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                strXml += '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' + rel.Target + '"/>';
            else
                strXml +=
                    '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="' + rel.Target + '"/>';
        }
        else if (rel.type.toLowerCase().indexOf('video') > -1) {
            // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
            if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                strXml += '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' + rel.Target + '"/>';
            else
                strXml +=
                    '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="' + rel.Target + '"/>';
        }
        else if (rel.type.toLowerCase().indexOf('online') > -1) {
            // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
            if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                strXml += '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.microsoft.com/office/2007/relationships/image" Target="' + rel.Target + '"/>';
            else
                strXml +=
                    '<Relationship Id="rId' +
                        rel.rId +
                        '" Target="' +
                        rel.Target +
                        '" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>';
        }
    });
    // STEP 2: Add default rels
    defaultRels.forEach(function (rel, idx) {
        strXml += '<Relationship Id="rId' + (lastRid + idx + 1) + '" Type="' + rel.type + '" Target="' + rel.target + '"/>';
    });
    strXml += '</Relationships>';
    return strXml;
}
/**
 * Generate XML Paragraph Properties
 * @param {ISlideObject|TextProps} textObj - text object
 * @param {boolean} isDefault - array of default relations
 * @return {string} XML
 */
function genXmlParagraphProperties(textObj, isDefault) {
    var strXmlBullet = '', strXmlLnSpc = '', strXmlParaSpc = '', strXmlTabStops = '';
    var tag = isDefault ? 'a:lvl1pPr' : 'a:pPr';
    var bulletMarL = valToPts(DEF_BULLET_MARGIN);
    var paragraphPropXml = "<" + tag + (textObj.options.rtlMode ? ' rtl="1" ' : '');
    // A: Build paragraphProperties
    {
        // OPTION: align
        if (textObj.options.align) {
            switch (textObj.options.align) {
                case 'left':
                    paragraphPropXml += ' algn="l"';
                    break;
                case 'right':
                    paragraphPropXml += ' algn="r"';
                    break;
                case 'center':
                    paragraphPropXml += ' algn="ctr"';
                    break;
                case 'justify':
                    paragraphPropXml += ' algn="just"';
                    break;
                default:
                    paragraphPropXml += '';
                    break;
            }
        }
        if (textObj.options.lineSpacing) {
            strXmlLnSpc = "<a:lnSpc><a:spcPts val=\"" + Math.round(textObj.options.lineSpacing * 100) + "\"/></a:lnSpc>";
        }
        else if (textObj.options.lineSpacingMultiple) {
            strXmlLnSpc = "<a:lnSpc><a:spcPct val=\"" + Math.round(textObj.options.lineSpacingMultiple * 100000) + "\"/></a:lnSpc>";
        }
        // OPTION: indent
        if (textObj.options.indentLevel && !isNaN(Number(textObj.options.indentLevel)) && textObj.options.indentLevel > 0) {
            paragraphPropXml += " lvl=\"" + textObj.options.indentLevel + "\"";
        }
        // OPTION: Paragraph Spacing: Before/After
        if (textObj.options.paraSpaceBefore && !isNaN(Number(textObj.options.paraSpaceBefore)) && textObj.options.paraSpaceBefore > 0) {
            strXmlParaSpc += "<a:spcBef><a:spcPts val=\"" + Math.round(textObj.options.paraSpaceBefore * 100) + "\"/></a:spcBef>";
        }
        if (textObj.options.paraSpaceAfter && !isNaN(Number(textObj.options.paraSpaceAfter)) && textObj.options.paraSpaceAfter > 0) {
            strXmlParaSpc += "<a:spcAft><a:spcPts val=\"" + Math.round(textObj.options.paraSpaceAfter * 100) + "\"/></a:spcAft>";
        }
        // OPTION: bullet
        // NOTE: OOXML uses the unicode character set for Bullets
        // EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
        if (typeof textObj.options.bullet === 'object') {
            if (textObj && textObj.options && textObj.options.bullet && textObj.options.bullet.indent)
                bulletMarL = valToPts(textObj.options.bullet.indent);
            if (textObj.options.bullet.type) {
                if (textObj.options.bullet.type.toString().toLowerCase() === 'number') {
                    paragraphPropXml += " marL=\"" + (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletMarL + bulletMarL * textObj.options.indentLevel : bulletMarL) + "\" indent=\"-" + bulletMarL + "\"";
                    strXmlBullet = "<a:buSzPct val=\"100000\"/><a:buFont typeface=\"+mj-lt\"/><a:buAutoNum type=\"" + (textObj.options.bullet.style || 'arabicPeriod') + "\" startAt=\"" + (textObj.options.bullet.numberStartAt || textObj.options.bullet.startAt || '1') + "\"/>";
                }
            }
            else if (textObj.options.bullet.characterCode) {
                var bulletCode = "&#x" + textObj.options.bullet.characterCode + ";";
                // Check value for hex-ness (s/b 4 char hex)
                if (/^[0-9A-Fa-f]{4}$/.test(textObj.options.bullet.characterCode) === false) {
                    console.warn('Warning: `bullet.characterCode should be a 4-digit unicode charatcer (ex: 22AB)`!');
                    bulletCode = BULLET_TYPES['DEFAULT'];
                }
                paragraphPropXml += " marL=\"" + (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletMarL + bulletMarL * textObj.options.indentLevel : bulletMarL) + "\" indent=\"-" + bulletMarL + "\"";
                strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + bulletCode + '"/>';
            }
            else if (textObj.options.bullet.code) {
                // @deprecated `bullet.code` v3.3.0
                var bulletCode = "&#x" + textObj.options.bullet.code + ";";
                // Check value for hex-ness (s/b 4 char hex)
                if (/^[0-9A-Fa-f]{4}$/.test(textObj.options.bullet.code) === false) {
                    console.warn('Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!');
                    bulletCode = BULLET_TYPES['DEFAULT'];
                }
                paragraphPropXml += " marL=\"" + (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletMarL + bulletMarL * textObj.options.indentLevel : bulletMarL) + "\" indent=\"-" + bulletMarL + "\"";
                strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + bulletCode + '"/>';
            }
            else {
                paragraphPropXml += " marL=\"" + (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletMarL + bulletMarL * textObj.options.indentLevel : bulletMarL) + "\" indent=\"-" + bulletMarL + "\"";
                strXmlBullet = "<a:buSzPct val=\"100000\"/><a:buChar char=\"" + BULLET_TYPES['DEFAULT'] + "\"/>";
            }
        }
        else if (textObj.options.bullet === true) {
            paragraphPropXml += " marL=\"" + (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletMarL + bulletMarL * textObj.options.indentLevel : bulletMarL) + "\" indent=\"-" + bulletMarL + "\"";
            strXmlBullet = "<a:buSzPct val=\"100000\"/><a:buChar char=\"" + BULLET_TYPES['DEFAULT'] + "\"/>";
        }
        else if (textObj.options.bullet === false) {
            // We only add this when the user explicitely asks for no bullet, otherwise, it can override the master defaults!
            paragraphPropXml += " indent=\"0\" marL=\"0\""; // FIX: ISSUE#589 - specify zero indent and marL or default will be hanging paragraph
            strXmlBullet = '<a:buNone/>';
        }
        // OPTION: tabStops
        if (textObj.options.tabStops && Array.isArray(textObj.options.tabStops)) {
            var tabStopsXml = textObj.options.tabStops.map(function (stop) { return "<a:tab pos=\"" + inch2Emu(stop.position || 1) + "\" algn=\"" + (stop.alignment || 'l') + "\"/>"; }).join('');
            strXmlTabStops = "<a:tabLst>" + tabStopsXml + "</a:tabLst>";
        }
        // B: Close Paragraph-Properties
        // IMPORTANT: strXmlLnSpc, strXmlParaSpc, and strXmlBullet require strict ordering - anything out of order is ignored. (PPT-Online, PPT for Mac)
        paragraphPropXml += '>' + strXmlLnSpc + strXmlParaSpc + strXmlBullet + strXmlTabStops;
        if (isDefault)
            paragraphPropXml += genXmlTextRunProperties(textObj.options, true);
        paragraphPropXml += '</' + tag + '>';
    }
    return paragraphPropXml;
}
/**
 * Generate XML Text Run Properties (`a:rPr`)
 * @param {ObjectOptions|TextPropsOptions} opts - text options
 * @param {boolean} isDefault - whether these are the default text run properties
 * @return {string} XML
 */
function genXmlTextRunProperties(opts, isDefault) {
    var _a;
    var runProps = '';
    var runPropsTag = isDefault ? 'a:defRPr' : 'a:rPr';
    // BEGIN runProperties (ex: `<a:rPr lang="en-US" sz="1600" b="1" dirty="0">`)
    runProps += '<' + runPropsTag + ' lang="' + (opts.lang ? opts.lang : 'en-US') + '"' + (opts.lang ? ' altLang="en-US"' : '');
    runProps += opts.fontSize ? ' sz="' + Math.round(opts.fontSize) + '00"' : ''; // NOTE: Use round so sizes like '7.5' wont cause corrupt pres.
    runProps += opts.hasOwnProperty('bold') ? " b=\"" + (opts.bold ? 1 : 0) + "\"" : '';
    runProps += opts.hasOwnProperty('italic') ? " i=\"" + (opts.italic ? 1 : 0) + "\"" : '';
    runProps += opts.hasOwnProperty('strike') ? " strike=\"" + (typeof opts.strike === 'string' ? opts.strike : 'sngStrike') + "\"" : '';
    if (typeof opts.underline === 'object' && ((_a = opts.underline) === null || _a === void 0 ? void 0 : _a.style)) {
        runProps += " u=\"" + opts.underline.style + "\"";
    }
    else if (typeof opts.underline === 'string') {
        // DEPRECATED: opts.underline is an object in v3.5.0
        runProps += " u=\"" + opts.underline + "\"";
    }
    else if (opts.hyperlink) {
        runProps += ' u="sng"';
    }
    if (opts.baseline) {
        runProps += " baseline=\"" + Math.round(opts.baseline * 50) + "\"";
    }
    else if (opts.subscript) {
        runProps += ' baseline="-40000"';
    }
    else if (opts.superscript) {
        runProps += ' baseline="30000"';
    }
    runProps += opts.charSpacing ? " spc=\"" + Math.round(opts.charSpacing * 100) + "\" kern=\"0\"" : ''; // IMPORTANT: Also disable kerning; otherwise text won't actually expand
    runProps += ' dirty="0">';
    // Color / Font / Highlight / Outline are children of <a:rPr>, so add them now before closing the runProperties tag
    if (opts.color || opts.fontFace || opts.outline || (typeof opts.underline === 'object' && opts.underline.color)) {
        if (opts.outline && typeof opts.outline === 'object') {
            runProps += "<a:ln w=\"" + valToPts(opts.outline.size || 0.75) + "\">" + genXmlColorSelection(opts.outline.color || 'FFFFFF') + "</a:ln>";
        }
        if (opts.color)
            runProps += genXmlColorSelection(opts.color);
        if (opts.highlight)
            runProps += "<a:highlight>" + createColorElement(opts.highlight) + "</a:highlight>";
        if (typeof opts.underline === 'object' && opts.underline.color)
            runProps += "<a:uFill>" + genXmlColorSelection(opts.underline.color) + "</a:uFill>";
        if (opts.glow)
            runProps += "<a:effectLst>" + createGlowElement(opts.glow, DEF_TEXT_GLOW) + "</a:effectLst>";
        if (opts.fontFace) {
            // NOTE: 'cs' = Complex Script, 'ea' = East Asian (use "-120" instead of "0" - per Issue #174); ea must come first (Issue #174)
            runProps += "<a:latin typeface=\"" + opts.fontFace + "\" pitchFamily=\"34\" charset=\"0\"/><a:ea typeface=\"" + opts.fontFace + "\" pitchFamily=\"34\" charset=\"-122\"/><a:cs typeface=\"" + opts.fontFace + "\" pitchFamily=\"34\" charset=\"-120\"/>";
        }
    }
    // Hyperlink support
    if (opts.hyperlink) {
        if (typeof opts.hyperlink !== 'object')
            throw new Error("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` ");
        else if (!opts.hyperlink.url && !opts.hyperlink.slide)
            throw new Error("ERROR: 'hyperlink requires either `url` or `slide`'");
        else if (opts.hyperlink.url) {
            //runProps += '<a:uFill>'+ genXmlColorSelection('0000FF') +'</a:uFill>'; // Breaks PPT2010! (Issue#74)
            runProps += "<a:hlinkClick r:id=\"rId" + opts.hyperlink._rId + "\" invalidUrl=\"\" action=\"\" tgtFrame=\"\" tooltip=\"" + (opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : '') + "\" history=\"1\" highlightClick=\"0\" endSnd=\"0\"" + (opts.color ? '>' : '/>');
        }
        else if (opts.hyperlink.slide) {
            runProps += "<a:hlinkClick r:id=\"rId" + opts.hyperlink._rId + "\" action=\"ppaction://hlinksldjump\" tooltip=\"" + (opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : '') + "\"" + (opts.color ? '>' : '/>');
        }
        if (opts.color) {
            runProps += '	<a:extLst>';
            runProps += '		<a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}">';
            runProps += '			<ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/>';
            runProps += '		</a:ext>';
            runProps += '	</a:extLst>';
            runProps += '</a:hlinkClick>';
        }
    }
    // END runProperties
    runProps += "</" + runPropsTag + ">";
    return runProps;
}
/**
 * Build textBody text runs [`<a:r></a:r>`] for paragraphs [`<a:p>`]
 * @param {TextProps} textObj - Text object
 * @return {string} XML string
 */
function genXmlTextRun(textObj) {
    // NOTE: Dont create full rPr runProps for empty [lineBreak] runs
    // Why? The size of the lineBreak wont match (eg: below it will be 18px instead of the correct 36px)
    // Do this:
    /*
        <a:p>
            <a:pPr algn="r"/>
            <a:endParaRPr lang="en-US" sz="3600" dirty="0"/>
        </a:p>
    */
    // NOT this:
    /*
        <a:p>
            <a:pPr algn="r"/>
            <a:r>
                <a:rPr lang="en-US" sz="3600" dirty="0">
                    <a:solidFill>
                        <a:schemeClr val="accent5"/>
                    </a:solidFill>
                    <a:latin typeface="Times" pitchFamily="34" charset="0"/>
                    <a:ea typeface="Times" pitchFamily="34" charset="-122"/>
                    <a:cs typeface="Times" pitchFamily="34" charset="-120"/>
                </a:rPr>
                <a:t></a:t>
            </a:r>
            <a:endParaRPr lang="en-US" dirty="0"/>
        </a:p>
    */
    // Return paragraph with text run
    return textObj.text ? "<a:r>" + genXmlTextRunProperties(textObj.options, false) + "<a:t>" + encodeXmlEntities(textObj.text) + "</a:t></a:r>" : '';
}
/**
 * Builds `<a:bodyPr></a:bodyPr>` tag for "genXmlTextBody()"
 * @param {ISlideObject | TableCell} slideObject - various options
 * @return {string} XML string
 */
function genXmlBodyProperties(slideObject) {
    var bodyProperties = '<a:bodyPr';
    if (slideObject && slideObject._type === SLIDE_OBJECT_TYPES.text && slideObject.options._bodyProp) {
        // PPT-2019 EX: <a:bodyPr wrap="square" lIns="1270" tIns="1270" rIns="1270" bIns="1270" rtlCol="0" anchor="ctr"/>
        // A: Enable or disable textwrapping none or square
        bodyProperties += slideObject.options._bodyProp.wrap ? ' wrap="square"' : ' wrap="none"';
        // B: Textbox margins [padding]
        if (slideObject.options._bodyProp.lIns || slideObject.options._bodyProp.lIns === 0)
            bodyProperties += ' lIns="' + slideObject.options._bodyProp.lIns + '"';
        if (slideObject.options._bodyProp.tIns || slideObject.options._bodyProp.tIns === 0)
            bodyProperties += ' tIns="' + slideObject.options._bodyProp.tIns + '"';
        if (slideObject.options._bodyProp.rIns || slideObject.options._bodyProp.rIns === 0)
            bodyProperties += ' rIns="' + slideObject.options._bodyProp.rIns + '"';
        if (slideObject.options._bodyProp.bIns || slideObject.options._bodyProp.bIns === 0)
            bodyProperties += ' bIns="' + slideObject.options._bodyProp.bIns + '"';
        // C: Add rtl after margins
        bodyProperties += ' rtlCol="0"';
        // D: Add anchorPoints
        if (slideObject.options._bodyProp.anchor)
            bodyProperties += ' anchor="' + slideObject.options._bodyProp.anchor + '"'; // VALS: [t,ctr,b]
        if (slideObject.options._bodyProp.vert)
            bodyProperties += ' vert="' + slideObject.options._bodyProp.vert + '"'; // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
        // E: Close <a:bodyPr element
        bodyProperties += '>';
        /**
         * F: Text Fit/AutoFit/Shrink option
         * @see: http://officeopenxml.com/drwSp-text-bodyPr-fit.php
         * @see: http://www.datypic.com/sc/ooxml/g-a_EG_TextAutofit.html
         */
        if (slideObject.options.fit) {
            // NOTE: Use of '<a:noAutofit/>' instead of '' causes issues in PPT-2013!
            if (slideObject.options.fit === 'none')
                bodyProperties += '';
            // NOTE: Shrink does not work automatically - PowerPoint calculates the `fontScale` value dynamically upon resize
            //else if (slideObject.options.fit === 'shrink') bodyProperties += '<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>' // MS-PPT > Format shape > Text Options: "Shrink text on overflow"
            else if (slideObject.options.fit === 'shrink')
                bodyProperties += '<a:normAutofit/>';
            else if (slideObject.options.fit === 'resize')
                bodyProperties += '<a:spAutoFit/>';
        }
        //
        // DEPRECATED: below (@deprecated v3.3.0)
        if (slideObject.options.shrinkText)
            bodyProperties += '<a:normAutofit/>'; // MS-PPT > Format shape > Text Options: "Shrink text on overflow"
        /* DEPRECATED: below (@deprecated v3.3.0)
         * MS-PPT > Format shape > Text Options: "Resize shape to fit text" [spAutoFit]
         * NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
         */
        bodyProperties += slideObject.options._bodyProp.autoFit !== false ? '<a:spAutoFit/>' : '';
        // LAST: Close _bodyProp
        bodyProperties += '</a:bodyPr>';
    }
    else {
        // DEFAULT:
        bodyProperties += ' wrap="square" rtlCol="0">';
        bodyProperties += '</a:bodyPr>';
    }
    // LAST: Return Close _bodyProp
    return slideObject._type === SLIDE_OBJECT_TYPES.tablecell ? '<a:bodyPr/>' : bodyProperties;
}
/**
 * Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
 * @param {ISlideObject|TableCell} slideObj - slideObj or tableCell
 * @note PPT text lines [lines followed by line-breaks] are created using <p>-aragraph's
 * @note Bullets are a paragragh-level formatting device
 * @template
 *	<p:txBody>
 *		<a:bodyPr wrap="square" rtlCol="0">
 *			<a:spAutoFit/>
 *		</a:bodyPr>
 *		<a:lstStyle/>
 *		<a:p>
 *			<a:pPr algn="ctr"/>
 *			<a:r>
 *				<a:rPr lang="en-US" dirty="0" err="1"/>
 *				<a:t>textbox text</a:t>
 *			</a:r>
 *			<a:endParaRPr lang="en-US" dirty="0"/>
 *		</a:p>
 *	</p:txBody>
 * @returns XML containing the param object's text and formatting
 */
function genXmlTextBody(slideObj) {
    var opts = slideObj.options || {};
    var tmpTextObjects = [];
    var arrTextObjects = [];
    // FIRST: Shapes without text, etc. may be sent here during build, but have no text to render so return an empty string
    if (opts && slideObj._type !== SLIDE_OBJECT_TYPES.tablecell && (typeof slideObj.text === 'undefined' || slideObj.text === null))
        return '';
    // STEP 1: Start textBody
    var strSlideXml = slideObj._type === SLIDE_OBJECT_TYPES.tablecell ? '<a:txBody>' : '<p:txBody>';
    // STEP 2: Add bodyProperties
    {
        // A: 'bodyPr'
        strSlideXml += genXmlBodyProperties(slideObj);
        // B: 'lstStyle'
        // NOTE: shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
        // FIXME: LINE horiz-align doesnt work (text is always to the left inside line) (FYI: the PPT code diff is substantial!)
        if (opts.h === 0 && opts.line && opts.align)
            strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>';
        else if (slideObj._type === 'placeholder')
            strSlideXml += "<a:lstStyle>" + genXmlParagraphProperties(slideObj, true) + "</a:lstStyle>";
        else
            strSlideXml += '<a:lstStyle/>';
    }
    /* STEP 3: Modify slideObj.text to array
        CASES:
        addText( 'string' ) // string
        addText( 'line1\n line2' ) // string with lineBreak
        addText( {text:'word1'} ) // TextProps object
        addText( ['barry','allen'] ) // array of strings
        addText( [{text:'word1'}, {text:'word2'}] ) // TextProps object array
        addText( [{text:'line1\n line2'}, {text:'end word'}] ) // TextProps object array with lineBreak
    */
    if (typeof slideObj.text === 'string' || typeof slideObj.text === 'number') {
        // Handle cases 1,2
        tmpTextObjects.push({ text: slideObj.text.toString(), options: opts || {} });
    }
    else if (slideObj.text && !Array.isArray(slideObj.text) && typeof slideObj.text === 'object' && Object.keys(slideObj.text).indexOf('text') > -1) {
        //} else if (!Array.isArray(slideObj.text) && slideObj.text!.hasOwnProperty('text')) { // 20210706: replaced with below as ts compiler rejected it
        // Handle case 3
        tmpTextObjects.push({ text: slideObj.text || '', options: slideObj.options || {} });
    }
    else if (Array.isArray(slideObj.text)) {
        // Handle cases 4,5,6
        // NOTE: use cast as text is TextProps[]|TableCell[] and their `options` dont overlap (they share the same TextBaseProps though)
        tmpTextObjects = slideObj.text.map(function (item) { return ({ text: item.text, options: item.options }); });
    }
    // STEP 4: Iterate over text objects, set text/options, break into pieces if '\n'/breakLine found
    tmpTextObjects.forEach(function (itext, idx) {
        if (!itext.text)
            itext.text = '';
        // A: Set options
        itext.options = itext.options || opts || {};
        if (idx === 0 && itext.options && !itext.options.bullet && opts.bullet)
            itext.options.bullet = opts.bullet;
        // B: Cast to text-object and fix line-breaks (if needed)
        if (typeof itext.text === 'string' || typeof itext.text === 'number') {
            // 1: Convert "\n" or any variation into CRLF
            itext.text = itext.text.toString().replace(/\r*\n/g, CRLF);
        }
        // C: If text string has line-breaks, then create a separate text-object for each (much easier than dealing with split inside a loop below)
        // NOTE: Filter for trailing lineBreak prevents the creation of an empty textObj as the last item
        if (itext.text.indexOf(CRLF) > -1 && itext.text.match(/\n$/g) === null) {
            itext.text.split(CRLF).forEach(function (line) {
                itext.options.breakLine = true;
                arrTextObjects.push({ text: line, options: itext.options });
            });
        }
        else {
            arrTextObjects.push(itext);
        }
    });
    // STEP 5: Group textObj into lines by checking for lineBreak, bullets, alignment change, etc.
    var arrLines = [];
    var arrTexts = [];
    arrTextObjects.forEach(function (textObj, idx) {
        // A: Align or Bullet trigger new line
        if (arrTexts.length > 0 && (textObj.options.align || opts.align)) {
            // Only start a new paragraph when align *changes*
            if (textObj.options.align != arrTextObjects[idx - 1].options.align) {
                arrLines.push(arrTexts);
                arrTexts = [];
            }
        }
        else if (arrTexts.length > 0 && textObj.options.bullet && arrTexts.length > 0) {
            arrLines.push(arrTexts);
            arrTexts = [];
            textObj.options.breakLine = false; // For cases with both `bullet` and `brekaLine` - prevent double lineBreak
        }
        // B: Add this text to current line
        arrTexts.push(textObj);
        // C: BreakLine begins new line **after** adding current text
        if (arrTexts.length > 0 && textObj.options.breakLine) {
            // Avoid starting a para right as loop is exhausted
            if (idx + 1 < arrTextObjects.length) {
                arrLines.push(arrTexts);
                arrTexts = [];
            }
        }
        // D: Flush buffer
        if (idx + 1 === arrTextObjects.length)
            arrLines.push(arrTexts);
    });
    // STEP 6: Loop over each line and create paragraph props, text run, etc.
    arrLines.forEach(function (line) {
        var reqsClosingFontSize = false;
        // A: Start paragraph, add paraProps
        strSlideXml += '<a:p>';
        // NOTE: `rtlMode` is like other opts, its propagated up to each text:options, so just check the 1st one
        var paragraphPropXml = "<a:pPr " + (line[0].options && line[0].options.rtlMode ? ' rtl="1" ' : '');
        // B: Start paragraph, loop over lines and add text runs
        line.forEach(function (textObj, idx) {
            // A: Set line index
            textObj.options._lineIdx = idx;
            // A.1: Add soft break if not the first run of the line.
            if (idx > 0 && textObj.options.softBreakBefore) {
                strSlideXml += "<a:br/>";
            }
            // B: Inherit pPr-type options from parent shape's `options`
            textObj.options.align = textObj.options.align || opts.align;
            textObj.options.lineSpacing = textObj.options.lineSpacing || opts.lineSpacing;
            textObj.options.lineSpacingMultiple = textObj.options.lineSpacingMultiple || opts.lineSpacingMultiple;
            textObj.options.indentLevel = textObj.options.indentLevel || opts.indentLevel;
            textObj.options.paraSpaceBefore = textObj.options.paraSpaceBefore || opts.paraSpaceBefore;
            textObj.options.paraSpaceAfter = textObj.options.paraSpaceAfter || opts.paraSpaceAfter;
            paragraphPropXml = genXmlParagraphProperties(textObj, false);
            strSlideXml += paragraphPropXml;
            // C: Inherit any main options (color, fontSize, etc.)
            // NOTE: We only pass the text.options to genXmlTextRun (not the Slide.options),
            // so the run building function cant just fallback to Slide.color, therefore, we need to do that here before passing options below.
            Object.entries(opts).forEach(function (_a) {
                var key = _a[0], val = _a[1];
                // RULE: Hyperlinks should not inherit `color` from main options (let PPT default tolocal color, eg: blue on MacOS)
                if (textObj.options.hyperlink && key === 'color')
                    ;
                // NOTE: This loop will pick up unecessary keys (`x`, etc.), but it doesnt hurt anything
                else if (key !== 'bullet' && !textObj.options[key])
                    textObj.options[key] = val;
            });
            // D: Add formatted textrun
            strSlideXml += genXmlTextRun(textObj);
            // E: Flag close fontSize for empty [lineBreak] elements
            if ((!textObj.text && opts.fontSize) || textObj.options.fontSize) {
                reqsClosingFontSize = true;
                opts.fontSize = opts.fontSize || textObj.options.fontSize;
            }
        });
        /* C: Append 'endParaRPr' (when needed) and close current open paragraph
         * NOTE: (ISSUE#20, ISSUE#193): Add 'endParaRPr' with font/size props or PPT default (Arial/18pt en-us) is used making row "too tall"/not honoring options
         */
        if (slideObj._type === SLIDE_OBJECT_TYPES.tablecell && (opts.fontSize || opts.fontFace)) {
            if (opts.fontFace) {
                strSlideXml += "<a:endParaRPr lang=\"" + (opts.lang || 'en-US') + "\"" + (opts.fontSize ? " sz=\"" + Math.round(opts.fontSize * 100) + "\"" : '') + ' dirty="0">';
                strSlideXml += "<a:latin typeface=\"" + opts.fontFace + "\" charset=\"0\"/>";
                strSlideXml += "<a:ea typeface=\"" + opts.fontFace + "\" charset=\"0\"/>";
                strSlideXml += "<a:cs typeface=\"" + opts.fontFace + "\" charset=\"0\"/>";
                strSlideXml += '</a:endParaRPr>';
            }
            else {
                strSlideXml += "<a:endParaRPr lang=\"" + (opts.lang || 'en-US') + "\"" + (opts.fontSize ? " sz=\"" + Math.round(opts.fontSize * 100) + "\"" : '') + ' dirty="0"/>';
            }
        }
        else if (reqsClosingFontSize) {
            // Empty [lineBreak] lines should not contain runProp, however, they need to specify fontSize in `endParaRPr`
            strSlideXml += "<a:endParaRPr lang=\"" + (opts.lang || 'en-US') + "\"" + (opts.fontSize ? " sz=\"" + Math.round(opts.fontSize * 100) + "\"" : '') + ' dirty="0"/>';
        }
        else {
            strSlideXml += "<a:endParaRPr lang=\"" + (opts.lang || 'en-US') + "\" dirty=\"0\"/>"; // Added 20180101 to address PPT-2007 issues
        }
        // D: End paragraph
        strSlideXml += '</a:p>';
    });
    // STEP 7: Close the textBody
    strSlideXml += slideObj._type === SLIDE_OBJECT_TYPES.tablecell ? '</a:txBody>' : '</p:txBody>';
    // LAST: Return XML
    return strSlideXml;
}
/**
 * Generate an XML Placeholder
 * @param {ISlideObject} placeholderObj
 * @returns XML
 */
function genXmlPlaceholder(placeholderObj) {
    if (!placeholderObj)
        return '';
    var placeholderIdx = placeholderObj.options && placeholderObj.options._placeholderIdx ? placeholderObj.options._placeholderIdx : '';
    var placeholderType = placeholderObj.options && placeholderObj.options._placeholderType ? placeholderObj.options._placeholderType : '';
    return "<p:ph\n\t\t" + (placeholderIdx ? ' idx="' + placeholderIdx + '"' : '') + "\n\t\t" + (placeholderType && PLACEHOLDER_TYPES[placeholderType] ? ' type="' + PLACEHOLDER_TYPES[placeholderType] + '"' : '') + "\n\t\t" + (placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : '') + "\n\t\t/>";
}
// XML-GEN: First 6 functions create the base /ppt files
/**
 * Generate XML ContentType
 * @param {PresSlide[]} slides - slides
 * @param {SlideLayout[]} slideLayouts - slide layouts
 * @param {PresSlide} masterSlide - master slide
 * @returns XML
 */
function makeXmlContTypes(slides, slideLayouts, masterSlide) {
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
    strXml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
    strXml += '<Default Extension="xml" ContentType="application/xml"/>';
    strXml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
    strXml += '<Default Extension="jpeg" ContentType="image/jpeg"/>';
    strXml += '<Default Extension="jpg" ContentType="image/jpg"/>';
    // STEP 1: Add standard/any media types used in Presentation
    strXml += '<Default Extension="png" ContentType="image/png"/>';
    strXml += '<Default Extension="gif" ContentType="image/gif"/>';
    strXml += '<Default Extension="m4v" ContentType="video/mp4"/>'; // NOTE: Hard-Code this extension as it wont be created in loop below (as extn !== type)
    strXml += '<Default Extension="mp4" ContentType="video/mp4"/>'; // NOTE: Hard-Code this extension as it wont be created in loop below (as extn !== type)
    slides.forEach(function (slide) {
        (slide._relsMedia || []).forEach(function (rel) {
            if (rel.type !== 'image' && rel.type !== 'online' && rel.type !== 'chart' && rel.extn !== 'm4v' && strXml.indexOf(rel.type) === -1) {
                strXml += '<Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>';
            }
        });
    });
    strXml += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
    strXml += '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>';
    // STEP 2: Add presentation and slide master(s)/slide(s)
    strXml += '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>';
    strXml += '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>';
    slides.forEach(function (slide, idx) {
        strXml +=
            '<Override PartName="/ppt/slideMasters/slideMaster' +
                (idx + 1) +
                '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>';
        strXml += '<Override PartName="/ppt/slides/slide' + (idx + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
        // Add charts if any
        slide._relsChart.forEach(function (rel) {
            strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
        });
    });
    // STEP 3: Core PPT
    strXml += '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>';
    strXml += '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>';
    strXml += '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
    strXml += '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>';
    // STEP 4: Add Slide Layouts
    slideLayouts.forEach(function (layout, idx) {
        strXml +=
            '<Override PartName="/ppt/slideLayouts/slideLayout' +
                (idx + 1) +
                '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
        (layout._relsChart || []).forEach(function (rel) {
            strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
        });
    });
    // STEP 5: Add notes slide(s)
    slides.forEach(function (_slide, idx) {
        strXml +=
            ' <Override PartName="/ppt/notesSlides/notesSlide' +
                (idx + 1) +
                '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>';
    });
    // STEP 6: Add rels
    masterSlide._relsChart.forEach(function (rel) {
        strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
    });
    masterSlide._relsMedia.forEach(function (rel) {
        if (rel.type !== 'image' && rel.type !== 'online' && rel.type !== 'chart' && rel.extn !== 'm4v' && strXml.indexOf(rel.type) === -1)
            strXml += ' <Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>';
    });
    // LAST: Finish XML (Resume core)
    strXml += ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
    strXml += ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    strXml += '</Types>';
    return strXml;
}
/**
 * Creates `_rels/.rels`
 * @returns XML
 */
function makeXmlRootRels() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n\t\t<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\n\t\t<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\n\t\t<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>\n\t\t</Relationships>";
}
/**
 * Creates `docProps/app.xml`
 * @param {PresSlide[]} slides - Presenation Slides
 * @param {string} company - "Company" metadata
 * @returns XML
 */
function makeXmlApp(slides, company) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n\t<TotalTime>0</TotalTime>\n\t<Words>0</Words>\n\t<Application>Microsoft Office PowerPoint</Application>\n\t<PresentationFormat>On-screen Show (16:9)</PresentationFormat>\n\t<Paragraphs>0</Paragraphs>\n\t<Slides>" + slides.length + "</Slides>\n\t<Notes>" + slides.length + "</Notes>\n\t<HiddenSlides>0</HiddenSlides>\n\t<MMClips>0</MMClips>\n\t<ScaleCrop>false</ScaleCrop>\n\t<HeadingPairs>\n\t\t<vt:vector size=\"6\" baseType=\"variant\">\n\t\t\t<vt:variant><vt:lpstr>Fonts Used</vt:lpstr></vt:variant>\n\t\t\t<vt:variant><vt:i4>2</vt:i4></vt:variant>\n\t\t\t<vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>\n\t\t\t<vt:variant><vt:i4>1</vt:i4></vt:variant>\n\t\t\t<vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>\n\t\t\t<vt:variant><vt:i4>" + slides.length + "</vt:i4></vt:variant>\n\t\t</vt:vector>\n\t</HeadingPairs>\n\t<TitlesOfParts>\n\t\t<vt:vector size=\"" + (slides.length + 1 + 2) + "\" baseType=\"lpstr\">\n\t\t\t<vt:lpstr>Arial</vt:lpstr>\n\t\t\t<vt:lpstr>Calibri</vt:lpstr>\n\t\t\t<vt:lpstr>Office Theme</vt:lpstr>\n\t\t\t" + slides.map(function (_slideObj, idx) { return '<vt:lpstr>Slide ' + (idx + 1) + '</vt:lpstr>\n'; }).join('') + "\n\t\t</vt:vector>\n\t</TitlesOfParts>\n\t<Company>" + company + "</Company>\n\t<LinksUpToDate>false</LinksUpToDate>\n\t<SharedDoc>false</SharedDoc>\n\t<HyperlinksChanged>false</HyperlinksChanged>\n\t<AppVersion>16.0000</AppVersion>\n\t</Properties>";
}
/**
 * Creates `docProps/core.xml`
 * @param {string} title - metadata data
 * @param {string} company - metadata data
 * @param {string} author - metadata value
 * @param {string} revision - metadata value
 * @returns XML
 */
function makeXmlCore(title, subject, author, revision) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\t<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n\t\t<dc:title>" + encodeXmlEntities(title) + "</dc:title>\n\t\t<dc:subject>" + encodeXmlEntities(subject) + "</dc:subject>\n\t\t<dc:creator>" + encodeXmlEntities(author) + "</dc:creator>\n\t\t<cp:lastModifiedBy>" + encodeXmlEntities(author) + "</cp:lastModifiedBy>\n\t\t<cp:revision>" + revision + "</cp:revision>\n\t\t<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + new Date().toISOString().replace(/\.\d\d\dZ/, 'Z') + "</dcterms:created>\n\t\t<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + new Date().toISOString().replace(/\.\d\d\dZ/, 'Z') + "</dcterms:modified>\n\t</cp:coreProperties>";
}
/**
 * Creates `ppt/_rels/presentation.xml.rels`
 * @param {PresSlide[]} slides - Presenation Slides
 * @returns XML
 */
function makeXmlPresentationRels(slides) {
    var intRelNum = 1;
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
    strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    strXml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>';
    for (var idx = 1; idx <= slides.length; idx++) {
        strXml +=
            '<Relationship Id="rId' + ++intRelNum + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + idx + '.xml"/>';
    }
    intRelNum++;
    strXml +=
        '<Relationship Id="rId' +
            intRelNum +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>' +
            '<Relationship Id="rId' +
            (intRelNum + 1) +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>' +
            '<Relationship Id="rId' +
            (intRelNum + 2) +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>' +
            '<Relationship Id="rId' +
            (intRelNum + 3) +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
            '<Relationship Id="rId' +
            (intRelNum + 4) +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>' +
            '</Relationships>';
    return strXml;
}
// XML-GEN: Functions that run 1-N times (once for each Slide)
/**
 * Generates XML for the slide file (`ppt/slides/slide1.xml`)
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
function makeXmlSlide(slide) {
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF +
        "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
        "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"" +
        ((slide && slide.hidden ? ' show="0"' : '') + ">") +
        ("" + slideObjectToXml(slide)) +
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>");
}
/**
 * Get text content of Notes from Slide
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} notes text
 */
function getNotesFromSlide(slide) {
    var notesText = '';
    slide._slideObjects.forEach(function (data) {
        if (data._type === SLIDE_OBJECT_TYPES.notes)
            notesText += data.text && data.text[0] ? data.text[0].text : '';
    });
    return notesText.replace(/\r*\n/g, CRLF);
}
/**
 * Generate XML for Notes Master (notesMaster1.xml)
 * @returns {string} XML
 */
function makeXmlNotesMaster() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<p:notesMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Header Placeholder 1\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"hdr\" sz=\"quarter\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"2971800\" cy=\"458788\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"/><a:lstStyle><a:lvl1pPr algn=\"l\"><a:defRPr sz=\"1200\"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Date Placeholder 2\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"dt\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"3884613\" y=\"0\"/><a:ext cx=\"2971800\" cy=\"458788\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"/><a:lstStyle><a:lvl1pPr algn=\"r\"><a:defRPr sz=\"1200\"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id=\"{5282F153-3F37-0F45-9E97-73ACFA13230C}\" type=\"datetimeFigureOut\"><a:rPr lang=\"en-US\"/><a:t>7/23/19</a:t></a:fld><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"4\" name=\"Slide Image Placeholder 3\"/><p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"sldImg\" idx=\"2\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"685800\" y=\"1143000\"/><a:ext cx=\"5486400\" cy=\"3086100\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"12700\"><a:solidFill><a:prstClr val=\"black\"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"ctr\"/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"5\" name=\"Notes Placeholder 4\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"quarter\" idx=\"3\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"685800\" y=\"4400550\"/><a:ext cx=\"5486400\" cy=\"3600450\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"/><a:lstStyle/><a:p><a:pPr lvl=\"0\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl=\"1\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl=\"2\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl=\"3\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl=\"4\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Fifth level</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"6\" name=\"Footer Placeholder 5\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"ftr\" sz=\"quarter\" idx=\"4\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"8685213\"/><a:ext cx=\"2971800\" cy=\"458787\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"b\"/><a:lstStyle><a:lvl1pPr algn=\"l\"><a:defRPr sz=\"1200\"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"7\" name=\"Slide Number Placeholder 6\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"sldNum\" sz=\"quarter\" idx=\"5\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"3884613\" y=\"8685213\"/><a:ext cx=\"2971800\" cy=\"458787\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"b\"/><a:lstStyle><a:lvl1pPr algn=\"r\"><a:defRPr sz=\"1200\"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id=\"{CE5E9CC1-C706-0F49-92D6-E571CC5EEA8F}\" type=\"slidenum\"><a:rPr lang=\"en-US\"/><a:t>\u2039#\u203A</a:t></a:fld><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri=\"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}\"><p14:creationId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"1024086991\"/></p:ext></p:extLst></p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/><p:notesStyle><a:lvl1pPr marL=\"0\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL=\"457200\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL=\"914400\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL=\"1371600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL=\"1828800\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL=\"2286000\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL=\"2743200\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL=\"3200400\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL=\"3657600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl9pPr></p:notesStyle></p:notesMaster>";
}
/**
 * Creates Notes Slide (`ppt/notesSlides/notesSlide1.xml`)
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
function makeXmlNotesSlide(slide) {
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        CRLF +
        '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>' +
        '<p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>' +
        '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>' +
        '</a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>' +
        '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/>' +
        '</p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>' +
        '<p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>' +
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
        '<a:rPr lang="en-US" dirty="0"/><a:t>' +
        encodeXmlEntities(getNotesFromSlide(slide)) +
        '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody>' +
        '</p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>' +
        '<p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr>' +
        '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>' +
        '<a:fld id="' +
        SLDNUMFLDID +
        '" type="slidenum">' +
        '<a:rPr lang="en-US"/><a:t>' +
        slide._slideNum +
        '</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>' +
        '</p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">' +
        '<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991"/>' +
        '</p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>');
}
/**
 * Generates the XML layout resource from a layout object
 * @param {SlideLayout} layout - slide layout (master)
 * @return {string} XML
 */
function makeXmlLayout(layout) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\t\t<p:sldLayout xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" preserve=\"1\">\n\t\t" + slideObjectToXml(layout) + "\n\t\t<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>";
}
/**
 * Creates Slide Master 1 (`ppt/slideMasters/slideMaster1.xml`)
 * @param {PresSlide} slide - slide object that represents master slide layout
 * @param {SlideLayout[]} layouts - slide layouts
 * @return {string} XML
 */
function makeXmlMaster(slide, layouts) {
    // NOTE: Pass layouts as static rels because they are not referenced any time
    var layoutDefs = layouts.map(function (_layoutDef, idx) { return '<p:sldLayoutId id="' + (LAYOUT_IDX_SERIES_BASE + idx) + '" r:id="rId' + (slide._rels.length + idx + 1) + '"/>'; });
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
    strXml +=
        '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">';
    strXml += slideObjectToXml(slide);
    strXml +=
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
    strXml += '<p:sldLayoutIdLst>' + layoutDefs.join('') + '</p:sldLayoutIdLst>';
    strXml += '<p:hf sldNum="0" hdr="0" ftr="0" dt="0"/>';
    strXml +=
        '<p:txStyles>' +
            ' <p:titleStyle>' +
            '  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>' +
            ' </p:titleStyle>' +
            ' <p:bodyStyle>' +
            '  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
            '  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
            '  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
            '  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
            '  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="»"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
            '  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
            '  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
            '  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
            '  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
            ' </p:bodyStyle>' +
            ' <p:otherStyle>' +
            '  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' +
            '  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
            '  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
            '  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
            '  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
            '  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
            '  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
            '  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
            '  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
            '  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
            ' </p:otherStyle>' +
            '</p:txStyles>';
    strXml += '</p:sldMaster>';
    return strXml;
}
/**
 * Generates XML string for a slide layout relation file
 * @param {number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @param {SlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
function makeXmlSlideLayoutRel(layoutNumber, slideLayouts) {
    return slideObjectRelationsToXml(slideLayouts[layoutNumber - 1], [
        {
            target: '../slideMasters/slideMaster1.xml',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
        },
    ]);
}
/**
 * Creates `ppt/_rels/slide*.xml.rels`
 * @param {PresSlide[]} slides
 * @param {SlideLayout[]} slideLayouts - Slide Layout(s)
 * @param {number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
function makeXmlSlideRel(slides, slideLayouts, slideNumber) {
    return slideObjectRelationsToXml(slides[slideNumber - 1], [
        {
            target: '../slideLayouts/slideLayout' + getLayoutIdxForSlide(slides, slideLayouts, slideNumber) + '.xml',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
        },
        {
            target: '../notesSlides/notesSlide' + slideNumber + '.xml',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
        },
    ]);
}
/**
 * Generates XML string for a slide relation file.
 * @param {number} slideNumber - 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
function makeXmlNotesSlideRel(slideNumber) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\t\t<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n\t\t\t<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster\" Target=\"../notesMasters/notesMaster1.xml\"/>\n\t\t\t<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"../slides/slide" + slideNumber + ".xml\"/>\n\t\t</Relationships>";
}
/**
 * Creates `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @param {PresSlide} masterSlide - Slide object
 * @param {SlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
function makeXmlMasterRel(masterSlide, slideLayouts) {
    var defaultRels = slideLayouts.map(function (_layoutDef, idx) { return ({
        target: "../slideLayouts/slideLayout" + (idx + 1) + ".xml",
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
    }); });
    defaultRels.push({ target: '../theme/theme1.xml', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' });
    return slideObjectRelationsToXml(masterSlide, defaultRels);
}
/**
 * Creates `ppt/notesMasters/_rels/notesMaster1.xml.rels`
 * @return {string} XML
 */
function makeXmlNotesMasterRel() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n\t\t<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>\n\t\t</Relationships>";
}
/**
 * For the passed slide number, resolves name of a layout that is used for.
 * @param {PresSlide[]} slides - srray of slides
 * @param {SlideLayout[]} slideLayouts - array of slideLayouts
 * @param {number} slideNumber
 * @return {number} slide number
 */
function getLayoutIdxForSlide(slides, slideLayouts, slideNumber) {
    for (var i = 0; i < slideLayouts.length; i++) {
        if (slideLayouts[i]._name === slides[slideNumber - 1]._slideLayout._name) {
            return i + 1;
        }
    }
    // IMPORTANT: Return 1 (for `slideLayout1.xml`) when no def is found
    // So all objects are in Layout1 and every slide that references it uses this layout.
    return 1;
}
// XML-GEN: Last 5 functions create root /ppt files
/**
 * Creates `ppt/theme/theme1.xml`
 * @return {string} XML
 */
function makeXmlTheme() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2><a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2><a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4><a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6><a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"\u6E38\u30B4\u30B7\u30C3\u30AF Light\"/><a:font script=\"Hang\" typeface=\"\uB9D1\uC740 \uACE0\uB515\"/><a:font script=\"Hans\" typeface=\"\u7B49\u7EBF Light\"/><a:font script=\"Hant\" typeface=\"\u65B0\u7D30\u660E\u9AD4\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Angsana New\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"\u6E38\u30B4\u30B7\u30C3\u30AF\"/><a:font script=\"Hang\" typeface=\"\uB9D1\uC740 \uACE0\uB515\"/><a:font script=\"Hans\" typeface=\"\u7B49\u7EBF\"/><a:font script=\"Hant\" typeface=\"\u65B0\u7D30\u660E\u9AD4\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Cordia New\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"/></a:ext></a:extLst></a:theme>";
}
/**
 * Create presentation file (`ppt/presentation.xml`)
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
 * @param {IPresentationProps} pres - presentation
 * @return {string} XML
 */
function makeXmlPresentation(pres) {
    var strXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF +
        "<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
        ("xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " + (pres.rtlMode ? 'rtl="1"' : '') + " saveSubsetFonts=\"1\" autoCompressPictures=\"0\">");
    // STEP 1: Add slide master (SPEC: tag 1 under <presentation>)
    strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>';
    // STEP 2: Add all Slides (SPEC: tag 3 under <presentation>)
    strXml += '<p:sldIdLst>';
    pres.slides.forEach(function (slide) { return (strXml += "<p:sldId id=\"" + slide._slideId + "\" r:id=\"rId" + slide._rId + "\"/>"); });
    strXml += '</p:sldIdLst>';
    // STEP 3: Add Notes Master (SPEC: tag 2 under <presentation>)
    // (NOTE: length+2 is from `presentation.xml.rels` func (since we have to match this rId, we just use same logic))
    // IMPORTANT: In this order (matches PPT2019) PPT will give corruption message on open!
    // IMPORTANT: Placing this before `<p:sldIdLst>` causes warning in modern powerpoint!
    // IMPORTANT: Presentations open without warning Without this line, however, the pres isnt preview in Finder anymore or viewable in iOS!
    strXml += "<p:notesMasterIdLst><p:notesMasterId r:id=\"rId" + (pres.slides.length + 2) + "\"/></p:notesMasterIdLst>";
    // STEP 4: Add sizes
    strXml += "<p:sldSz cx=\"" + pres.presLayout.width + "\" cy=\"" + pres.presLayout.height + "\"/>";
    strXml += "<p:notesSz cx=\"" + pres.presLayout.height + "\" cy=\"" + pres.presLayout.width + "\"/>";
    // STEP 5: Add text styles
    strXml += '<p:defaultTextStyle>';
    for (var idy = 1; idy < 10; idy++) {
        strXml +=
            "<a:lvl" + idy + "pPr marL=\"" + (idy - 1) * 457200 + "\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">" +
                "<a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>" +
                ("</a:defRPr></a:lvl" + idy + "pPr>");
    }
    strXml += '</p:defaultTextStyle>';
    // STEP 6: Add Sections (if any)
    if (pres.sections && pres.sections.length > 0) {
        strXml += '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">';
        strXml += '<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">';
        pres.sections.forEach(function (sect) {
            strXml += "<p14:section name=\"" + encodeXmlEntities(sect.title) + "\" id=\"{" + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + "}\"><p14:sldIdLst>";
            sect._slides.forEach(function (slide) { return (strXml += "<p14:sldId id=\"" + slide._slideId + "\"/>"); });
            strXml += "</p14:sldIdLst></p14:section>";
        });
        strXml += '</p14:sectionLst></p:ext>';
        strXml += '<p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext>';
        strXml += '</p:extLst>';
    }
    // Done
    strXml += '</p:presentation>';
    return strXml;
}
/**
 * Create `ppt/presProps.xml`
 * @return {string} XML
 */
function makeXmlPresProps() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<p:presentationPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>";
}
/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
function makeXmlTableStyles() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<a:tblStyleLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" def=\"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}\"/>";
}
/**
 * Creates `ppt/viewProps.xml`
 * @return {string} XML
 */
function makeXmlViewProps() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<p:viewPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:normalViewPr horzBarState=\"maximized\"><p:restoredLeft sz=\"15611\"/><p:restoredTop sz=\"94610\"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid=\"0\" snapToObjects=\"1\"><p:cViewPr varScale=\"1\"><p:scale><a:sx n=\"136\" d=\"100\"/><a:sy n=\"136\" d=\"100\"/></p:scale><p:origin x=\"216\" y=\"312\"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n=\"1\" d=\"1\"/><a:sy n=\"1\" d=\"1\"/></p:scale><p:origin x=\"0\" y=\"0\"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx=\"76200\" cy=\"76200\"/></p:viewPr>";
}
/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {ShadowProps} ShadowProps - shadow options
 */
function correctShadowOptions(ShadowProps) {
    if (!ShadowProps || typeof ShadowProps !== 'object') {
        //console.warn("`shadow` options must be an object. Ex: `{shadow: {type:'none'}}`")
        return;
    }
    // OPT: `type`
    if (ShadowProps.type !== 'outer' && ShadowProps.type !== 'inner' && ShadowProps.type !== 'none') {
        console.warn('Warning: shadow.type options are `outer`, `inner` or `none`.');
        ShadowProps.type = 'outer';
    }
    // OPT: `angle`
    if (ShadowProps.angle) {
        // A: REALITY-CHECK
        if (isNaN(Number(ShadowProps.angle)) || ShadowProps.angle < 0 || ShadowProps.angle > 359) {
            console.warn('Warning: shadow.angle can only be 0-359');
            ShadowProps.angle = 270;
        }
        // B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
        ShadowProps.angle = Math.round(Number(ShadowProps.angle));
    }
    // OPT: `opacity`
    if (ShadowProps.opacity) {
        // A: REALITY-CHECK
        if (isNaN(Number(ShadowProps.opacity)) || ShadowProps.opacity < 0 || ShadowProps.opacity > 1) {
            console.warn('Warning: shadow.opacity can only be 0-1');
            ShadowProps.opacity = 0.75;
        }
        // B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
        ShadowProps.opacity = Number(ShadowProps.opacity);
    }
}

/**
 * PptxGenJS: Slide Object Generators
 */
/** counter for included charts (used for index in their filenames) */
var _chartCounter = 0;
/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * @param {SlideMasterProps} props - slide definition
 * @param {PresSlide|SlideLayout} target - empty slide object that should be updated by the passed definition
 */
function createSlideMaster(props, target) {
    // STEP 1: Add background if either the slide or layout has background props
    //	if (props.background || target.background) addBackgroundDefinition(props.background, target)
    if (props.bkgd)
        target.bkgd = props.bkgd; // DEPRECATED: (remove in v4.0.0)
    // STEP 2: Add all Slide Master objects in the order they were given
    if (props.objects && Array.isArray(props.objects) && props.objects.length > 0) {
        props.objects.forEach(function (object, idx) {
            var key = Object.keys(object)[0];
            var tgt = target;
            if (MASTER_OBJECTS[key] && key === 'chart')
                addChartDefinition(tgt, object[key].type, object[key].data, object[key].opts);
            else if (MASTER_OBJECTS[key] && key === 'image')
                addImageDefinition(tgt, object[key]);
            else if (MASTER_OBJECTS[key] && key === 'line')
                addShapeDefinition(tgt, SHAPE_TYPE.LINE, object[key]);
            else if (MASTER_OBJECTS[key] && key === 'rect')
                addShapeDefinition(tgt, SHAPE_TYPE.RECTANGLE, object[key]);
            else if (MASTER_OBJECTS[key] && key === 'text')
                addTextDefinition(tgt, [{ text: object[key].text }], object[key].options, false);
            else if (MASTER_OBJECTS[key] && key === 'placeholder') {
                // TODO: 20180820: Check for existing `name`?
                object[key].options.placeholder = object[key].options.name;
                delete object[key].options.name; // remap name for earier handling internally
                object[key].options._placeholderType = object[key].options.type;
                delete object[key].options.type; // remap name for earier handling internally
                object[key].options._placeholderIdx = 100 + idx;
                addTextDefinition(tgt, [{ text: object[key].text }], object[key].options, true);
                // TODO: ISSUE#599 - only text is suported now (add more below)
                //else if (object[key].image) addImageDefinition(tgt, object[key].image)
                /* 20200120: So... image placeholders go into the "slideLayoutN.xml" file and addImage doesnt do this yet...
                    <p:sp>
                  <p:nvSpPr>
                    <p:cNvPr id="7" name="Picture Placeholder 6">
                      <a:extLst>
                        <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                          <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{CE1AE45D-8641-0F4F-BDB5-080E69CCB034}"/>
                        </a:ext>
                      </a:extLst>
                    </p:cNvPr>
                    <p:cNvSpPr>
                */
            }
        });
    }
    // STEP 3: Add Slide Numbers (NOTE: Do this last so numbers are not covered by objects!)
    if (props.slideNumber && typeof props.slideNumber === 'object')
        target._slideNumberProps = props.slideNumber;
}
/**
 * Generate the chart based on input data.
 * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
 *
 * @param {CHART_NAME | IChartMulti[]} `type` should belong to: 'column', 'pie'
 * @param {[]} `data` a JSON object with follow the following format
 * @param {IChartOptsLib} `opt` chart options
 * @param {PresSlide} `target` slide object that the chart will be added to
 * @return {object} chart object
 * {
 *   title: 'eSurvey chart',
 *   data: [
 *		{
 *			name: 'Income',
 *			labels: ['2005', '2006', '2007', '2008', '2009'],
 *			values: [23.5, 26.2, 30.1, 29.5, 24.6]
 *		},
 *		{
 *			name: 'Expense',
 *			labels: ['2005', '2006', '2007', '2008', '2009'],
 *			values: [18.1, 22.8, 23.9, 25.1, 25]
 *		}
 *	 ]
 *	}
 */
function addChartDefinition(target, type, data, opt) {
    function correctGridLineOptions(glOpts) {
        if (!glOpts || glOpts.style === 'none')
            return;
        if (glOpts.size !== undefined && (isNaN(Number(glOpts.size)) || glOpts.size <= 0)) {
            console.warn('Warning: chart.gridLine.size must be greater than 0.');
            delete glOpts.size; // delete prop to used defaults
        }
        if (glOpts.style && ['solid', 'dash', 'dot'].indexOf(glOpts.style) < 0) {
            console.warn('Warning: chart.gridLine.style options: `solid`, `dash`, `dot`.');
            delete glOpts.style;
        }
    }
    var chartId = ++_chartCounter;
    var resultObject = {
        _type: null,
        text: null,
        options: null,
        chartRid: null,
    };
    // DESIGN: `type` can an object (ex: `pptx.charts.DOUGHNUT`) or an array of chart objects
    // EX: addChartDefinition([ { type:pptx.charts.BAR, data:{name:'', labels:[], values[]} }, {<etc>} ])
    // Multi-Type Charts
    var tmpOpt;
    var tmpData = [], options;
    if (Array.isArray(type)) {
        // For multi-type charts there needs to be data for each type,
        // as well as a single data source for non-series operations.
        // The data is indexed below to keep the data in order when segmented
        // into types.
        type.forEach(function (obj) {
            tmpData = tmpData.concat(obj.data);
        });
        tmpOpt = data || opt;
    }
    else {
        tmpData = data;
        tmpOpt = opt;
    }
    tmpData.forEach(function (item, i) {
        item.index = i;
    });
    options = tmpOpt && typeof tmpOpt === 'object' ? tmpOpt : {};
    // STEP 1: TODO: check for reqd fields, correct type, etc
    // `type` exists in CHART_TYPE
    // Array.isArray(data)
    /*
        if ( Array.isArray(rel.data) && rel.data.length > 0 && typeof rel.data[0] === 'object'
            && rel.data[0].labels && Array.isArray(rel.data[0].labels)
            && rel.data[0].values && Array.isArray(rel.data[0].values) ) {
            obj = rel.data[0];
        }
        else {
            console.warn("USAGE: addChart( 'pie', [ {name:'Sales', labels:['Jan','Feb'], values:[10,20]} ], {x:1, y:1} )");
            return;
        }
        */
    // STEP 2: Set default options/decode user options
    // A: Core
    options._type = type;
    options.x = typeof options.x !== 'undefined' && options.x != null && !isNaN(Number(options.x)) ? options.x : 1;
    options.y = typeof options.y !== 'undefined' && options.y != null && !isNaN(Number(options.y)) ? options.y : 1;
    options.w = options.w || '50%';
    options.h = options.h || '50%';
    // B: Options: misc
    if (['bar', 'col'].indexOf(options.barDir || '') < 0)
        options.barDir = 'col';
    // barGrouping: "21.2.3.17 ST_Grouping (Grouping)"
    // barGrouping must be handled before data label validation as it can affect valid label positioning
    if (options._type === CHART_TYPE.AREA) {
        if (['stacked', 'standard', 'percentStacked'].indexOf(options.barGrouping || '') < 0)
            options.barGrouping = 'standard';
    }
    if (options._type === CHART_TYPE.BAR) {
        if (['clustered', 'stacked', 'percentStacked'].indexOf(options.barGrouping || '') < 0)
            options.barGrouping = 'clustered';
    }
    if (options._type === CHART_TYPE.BAR3D) {
        if (['clustered', 'stacked', 'standard', 'percentStacked'].indexOf(options.barGrouping || '') < 0)
            options.barGrouping = 'standard';
    }
    if (options.barGrouping && options.barGrouping.indexOf('tacked') > -1) {
        if (!options.barGapWidthPct)
            options.barGapWidthPct = 50;
    }
    // Clean up and validate data label positions
    // REFERENCE: https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e2b1697c-7adc-463d-9081-3daef72f656f?redirectedfrom=MSDN
    if (options.dataLabelPosition) {
        if (options._type === CHART_TYPE.AREA || options._type === CHART_TYPE.BAR3D || options._type === CHART_TYPE.DOUGHNUT || options._type === CHART_TYPE.RADAR)
            delete options.dataLabelPosition;
        if (options._type === CHART_TYPE.PIE) {
            if (['bestFit', 'ctr', 'inEnd', 'outEnd'].indexOf(options.dataLabelPosition) < 0)
                delete options.dataLabelPosition;
        }
        if (options._type === CHART_TYPE.BUBBLE || options._type === CHART_TYPE.LINE || options._type === CHART_TYPE.SCATTER) {
            if (['b', 'ctr', 'l', 'r', 't'].indexOf(options.dataLabelPosition) < 0)
                delete options.dataLabelPosition;
        }
        if (options._type === CHART_TYPE.BAR) {
            if (['stacked', 'percentStacked'].indexOf(options.barGrouping || '') < 0) {
                if (['ctr', 'inBase', 'inEnd'].indexOf(options.dataLabelPosition) < 0)
                    delete options.dataLabelPosition;
            }
            if (['clustered'].indexOf(options.barGrouping || '') < 0) {
                if (['ctr', 'inBase', 'inEnd', 'outEnd'].indexOf(options.dataLabelPosition) < 0)
                    delete options.dataLabelPosition;
            }
        }
    }
    options.dataLabelBkgrdColors = options.dataLabelBkgrdColors === true || options.dataLabelBkgrdColors === false ? options.dataLabelBkgrdColors : false;
    if (['b', 'l', 'r', 't', 'tr'].indexOf(options.legendPos || '') < 0)
        options.legendPos = 'r';
    // 3D bar: ST_Shape
    if (['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax'].indexOf(options.bar3DShape || '') < 0)
        options.bar3DShape = 'box';
    // lineDataSymbol: http://www.datypic.com/sc/ooxml/a-val-32.html
    // Spec has [plus,star,x] however neither PPT2013 nor PPT-Online support them
    if (['circle', 'dash', 'diamond', 'dot', 'none', 'square', 'triangle'].indexOf(options.lineDataSymbol || '') < 0)
        options.lineDataSymbol = 'circle';
    if (['gap', 'span'].indexOf(options.displayBlanksAs || '') < 0)
        options.displayBlanksAs = 'span';
    if (['standard', 'marker', 'filled'].indexOf(options.radarStyle || '') < 0)
        options.radarStyle = 'standard';
    options.lineDataSymbolSize = options.lineDataSymbolSize && !isNaN(options.lineDataSymbolSize) ? options.lineDataSymbolSize : 6;
    options.lineDataSymbolLineSize = options.lineDataSymbolLineSize && !isNaN(options.lineDataSymbolLineSize) ? valToPts(options.lineDataSymbolLineSize) : valToPts(0.75);
    // `layout` allows the override of PPT defaults to maximize space
    if (options.layout) {
        ['x', 'y', 'w', 'h'].forEach(function (key) {
            var val = options.layout[key];
            if (isNaN(Number(val)) || val < 0 || val > 1) {
                console.warn('Warning: chart.layout.' + key + ' can only be 0-1');
                delete options.layout[key]; // remove invalid value so that default will be used
            }
        });
    }
    // Set gridline defaults
    options.catGridLine = options.catGridLine || (options._type === CHART_TYPE.SCATTER ? { color: 'D9D9D9', size: 1 } : { style: 'none' });
    options.valGridLine = options.valGridLine || (options._type === CHART_TYPE.SCATTER ? { color: 'D9D9D9', size: 1 } : {});
    options.serGridLine = options.serGridLine || (options._type === CHART_TYPE.SCATTER ? { color: 'D9D9D9', size: 1 } : { style: 'none' });
    correctGridLineOptions(options.catGridLine);
    correctGridLineOptions(options.valGridLine);
    correctGridLineOptions(options.serGridLine);
    correctShadowOptions(options.shadow);
    // C: Options: plotArea
    options.showDataTable = options.showDataTable === true || options.showDataTable === false ? options.showDataTable : false;
    options.showDataTableHorzBorder = options.showDataTableHorzBorder === true || options.showDataTableHorzBorder === false ? options.showDataTableHorzBorder : true;
    options.showDataTableVertBorder = options.showDataTableVertBorder === true || options.showDataTableVertBorder === false ? options.showDataTableVertBorder : true;
    options.showDataTableOutline = options.showDataTableOutline === true || options.showDataTableOutline === false ? options.showDataTableOutline : true;
    options.showDataTableKeys = options.showDataTableKeys === true || options.showDataTableKeys === false ? options.showDataTableKeys : true;
    options.showLabel = options.showLabel === true || options.showLabel === false ? options.showLabel : false;
    options.showLegend = options.showLegend === true || options.showLegend === false ? options.showLegend : false;
    options.showPercent = options.showPercent === true || options.showPercent === false ? options.showPercent : true;
    options.showTitle = options.showTitle === true || options.showTitle === false ? options.showTitle : false;
    options.showValue = options.showValue === true || options.showValue === false ? options.showValue : false;
    options.showLeaderLines = options.showLeaderLines === true || options.showLeaderLines === false ? options.showLeaderLines : false;
    options.catAxisLineShow = typeof options.catAxisLineShow !== 'undefined' ? options.catAxisLineShow : true;
    options.valAxisLineShow = typeof options.valAxisLineShow !== 'undefined' ? options.valAxisLineShow : true;
    options.serAxisLineShow = typeof options.serAxisLineShow !== 'undefined' ? options.serAxisLineShow : true;
    options.v3DRotX = !isNaN(options.v3DRotX) && options.v3DRotX >= -90 && options.v3DRotX <= 90 ? options.v3DRotX : 30;
    options.v3DRotY = !isNaN(options.v3DRotY) && options.v3DRotY >= 0 && options.v3DRotY <= 360 ? options.v3DRotY : 30;
    options.v3DRAngAx = options.v3DRAngAx === true || options.v3DRAngAx === false ? options.v3DRAngAx : true;
    options.v3DPerspective = !isNaN(options.v3DPerspective) && options.v3DPerspective >= 0 && options.v3DPerspective <= 240 ? options.v3DPerspective : 30;
    // D: Options: chart
    options.barGapWidthPct = !isNaN(options.barGapWidthPct) && options.barGapWidthPct >= 0 && options.barGapWidthPct <= 1000 ? options.barGapWidthPct : 150;
    options.barGapDepthPct = !isNaN(options.barGapDepthPct) && options.barGapDepthPct >= 0 && options.barGapDepthPct <= 1000 ? options.barGapDepthPct : 150;
    options.chartColors = Array.isArray(options.chartColors)
        ? options.chartColors
        : options._type === CHART_TYPE.PIE || options._type === CHART_TYPE.DOUGHNUT
            ? PIECHART_COLORS
            : BARCHART_COLORS;
    options.chartColorsOpacity = options.chartColorsOpacity && !isNaN(options.chartColorsOpacity) ? options.chartColorsOpacity : null;
    //
    options.border = options.border && typeof options.border === 'object' ? options.border : null;
    if (options.border && (!options.border.pt || isNaN(options.border.pt)))
        options.border.pt = 1;
    if (options.border && (!options.border.color || typeof options.border.color !== 'string' || options.border.color.length !== 6))
        options.border.color = '363636';
    //
    options.dataBorder = options.dataBorder && typeof options.dataBorder === 'object' ? options.dataBorder : null;
    if (options.dataBorder && (!options.dataBorder.pt || isNaN(options.dataBorder.pt)))
        options.dataBorder.pt = 0.75;
    if (options.dataBorder && (!options.dataBorder.color || typeof options.dataBorder.color !== 'string' || options.dataBorder.color.length !== 6))
        options.dataBorder.color = 'F9F9F9';
    //
    if (!options.dataLabelFormatCode && options._type === CHART_TYPE.SCATTER)
        options.dataLabelFormatCode = 'General';
    if (!options.dataLabelFormatCode && (options._type === CHART_TYPE.PIE || options._type === CHART_TYPE.DOUGHNUT))
        options.dataLabelFormatCode = options.showPercent ? '0%' : 'General';
    options.dataLabelFormatCode = options.dataLabelFormatCode && typeof options.dataLabelFormatCode === 'string' ? options.dataLabelFormatCode : '#,##0';
    //
    // Set default format for Scatter chart labels to custom string if not defined
    if (!options.dataLabelFormatScatter && options._type === CHART_TYPE.SCATTER)
        options.dataLabelFormatScatter = 'custom';
    //
    options.lineSize = typeof options.lineSize === 'number' ? options.lineSize : 2;
    options.valAxisMajorUnit = typeof options.valAxisMajorUnit === 'number' ? options.valAxisMajorUnit : null;
    options.valAxisCrossesAt = options.valAxisCrossesAt || 'autoZero';
    // STEP 4: Set props
    resultObject._type = 'chart';
    resultObject.options = options;
    resultObject.chartRid = getNewRelId(target);
    // STEP 5: Add this chart to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
    target._relsChart.push({
        rId: getNewRelId(target),
        data: tmpData,
        opts: options,
        type: options._type,
        globalId: chartId,
        fileName: 'chart' + chartId + '.xml',
        Target: '/ppt/charts/chart' + chartId + '.xml',
    });
    target._slideObjects.push(resultObject);
    return resultObject;
}
/**
 * Adds an image object to a slide definition.
 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
 * @param {ImageProps} `opt` - object containing `path`/`data`, `x`, `y`, etc.
 * @param {PresSlide} `target` - slide that the image should be added to (if not specified as the 2nd arg)
 * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
 * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
 */
function addImageDefinition(target, opt) {
    var newObject = {
        _type: null,
        text: null,
        options: null,
        image: null,
        imageRid: null,
        hyperlink: null,
    };
    // FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
    var intPosX = opt.x || 0;
    var intPosY = opt.y || 0;
    var intWidth = opt.w || 0;
    var intHeight = opt.h || 0;
    var sizing = opt.sizing || null;
    var objHyperlink = opt.hyperlink || '';
    var strImageData = opt.data || '';
    var strImagePath = opt.path || '';
    var imageRelId = getNewRelId(target);
    // REALITY-CHECK:
    if (!strImagePath && !strImageData) {
        console.error("ERROR: addImage() requires either 'data' or 'path' parameter!");
        return null;
    }
    else if (strImagePath && typeof strImagePath !== 'string') {
        console.error("ERROR: addImage() 'path' should be a string, ex: {path:'/img/sample.png'} - you sent " + strImagePath);
        return null;
    }
    else if (strImageData && typeof strImageData !== 'string') {
        console.error("ERROR: addImage() 'data' should be a string, ex: {data:'image/png;base64,NMP[...]'} - you sent " + strImageData);
        return null;
    }
    else if (strImageData && typeof strImageData === 'string' && strImageData.toLowerCase().indexOf('base64,') === -1) {
        console.error("ERROR: Image `data` value lacks a base64 header! Ex: 'image/png;base64,NMP[...]')");
        return null;
    }
    // STEP 1: Set extension
    // NOTE: Split to address URLs with params (eg: `path/brent.jpg?someParam=true`)
    var strImgExtn = strImagePath
        .substring(strImagePath.lastIndexOf('/') + 1)
        .split('?')[0]
        .split('.')
        .pop()
        .split('#')[0] || 'png';
    // However, pre-encoded images can be whatever mime-type they want (and good for them!)
    if (strImageData && /image\/(\w+);/.exec(strImageData) && /image\/(\w+);/.exec(strImageData).length > 0) {
        strImgExtn = /image\/(\w+);/.exec(strImageData)[1];
    }
    else if (strImageData && strImageData.toLowerCase().indexOf('image/svg+xml') > -1) {
        strImgExtn = 'svg';
    }
    // STEP 2: Set type/path
    newObject._type = SLIDE_OBJECT_TYPES.image;
    newObject.image = strImagePath || 'preencoded.png';
    // STEP 3: Set image properties & options
    // FIXME: Measure actual image when no intWidth/intHeight params passed
    // ....: This is an async process: we need to make getSizeFromImage use callback, then set H/W...
    // if ( !intWidth || !intHeight ) { var imgObj = getSizeFromImage(strImagePath);
    newObject.options = {
        x: intPosX || 0,
        y: intPosY || 0,
        w: intWidth || 1,
        h: intHeight || 1,
        altText: opt.altText || '',
        rounding: typeof opt.rounding === 'boolean' ? opt.rounding : false,
        sizing: sizing,
        placeholder: opt.placeholder,
        rotate: opt.rotate || 0,
        flipV: opt.flipV || false,
        flipH: opt.flipH || false,
    };
    // STEP 4: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
    if (strImgExtn === 'svg') {
        // SVG files consume *TWO* rId's: (a png version and the svg image)
        // <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
        // <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/>
        target._relsMedia.push({
            path: strImagePath || strImageData + 'png',
            type: 'image/png',
            extn: 'png',
            data: strImageData || '',
            rId: imageRelId,
            Target: '../media/image-' + target._slideNum + '-' + (target._relsMedia.length + 1) + '.png',
            isSvgPng: true,
            svgSize: { w: getSmartParseNumber(newObject.options.w, 'X', target._presLayout), h: getSmartParseNumber(newObject.options.h, 'Y', target._presLayout) },
        });
        newObject.imageRid = imageRelId;
        target._relsMedia.push({
            path: strImagePath || strImageData,
            type: 'image/svg+xml',
            extn: strImgExtn,
            data: strImageData || '',
            rId: imageRelId + 1,
            Target: '../media/image-' + target._slideNum + '-' + (target._relsMedia.length + 1) + '.' + strImgExtn,
        });
        newObject.imageRid = imageRelId + 1;
    }
    else {
        target._relsMedia.push({
            path: strImagePath || 'preencoded.' + strImgExtn,
            type: 'image/' + strImgExtn,
            extn: strImgExtn,
            data: strImageData || '',
            rId: imageRelId,
            Target: '../media/image-' + target._slideNum + '-' + (target._relsMedia.length + 1) + '.' + strImgExtn,
        });
        newObject.imageRid = imageRelId;
    }
    // STEP 5: Hyperlink support
    if (typeof objHyperlink === 'object') {
        if (!objHyperlink.url && !objHyperlink.slide)
            throw new Error('ERROR: `hyperlink` option requires either: `url` or `slide`');
        else {
            imageRelId++;
            target._rels.push({
                type: SLIDE_OBJECT_TYPES.hyperlink,
                data: objHyperlink.slide ? 'slide' : 'dummy',
                rId: imageRelId,
                Target: objHyperlink.url || objHyperlink.slide.toString(),
            });
            objHyperlink._rId = imageRelId;
            newObject.hyperlink = objHyperlink;
        }
    }
    // STEP 6: Add object to slide
    target._slideObjects.push(newObject);
}
/**
 * Adds a media object to a slide definition.
 * @param {PresSlide} `target` - slide object that the text will be added to
 * @param {MediaProps} `opt` - media options
 */
function addMediaDefinition(target, opt) {
    var intRels = target._relsMedia.length + 1;
    var intPosX = opt.x || 0;
    var intPosY = opt.y || 0;
    var intSizeX = opt.w || 2;
    var intSizeY = opt.h || 2;
    var strData = opt.data || '';
    var strLink = opt.link || '';
    var strPath = opt.path || '';
    var strType = opt.type || 'audio';
    var strExtn = 'mp3';
    var slideData = {
        _type: SLIDE_OBJECT_TYPES.media,
    };
    // STEP 1: REALITY-CHECK
    if (!strPath && !strData && strType !== 'online') {
        throw new Error("addMedia() error: either 'data' or 'path' are required!");
    }
    else if (strData && strData.toLowerCase().indexOf('base64,') === -1) {
        throw new Error("addMedia() error: `data` value lacks a base64 header! Ex: 'video/mpeg;base64,NMP[...]')");
    }
    // Online Video: requires `link`
    if (strType === 'online' && !strLink) {
        throw new Error('addMedia() error: online videos require `link` value');
    }
    // FIXME: 20190707
    //strType = strData ? strData.split(';')[0].split('/')[0] : strType
    strExtn = strData ? strData.split(';')[0].split('/')[1] : strPath.split('.').pop();
    // STEP 2: Set type, media
    slideData.mtype = strType;
    slideData.media = strPath || 'preencoded.mov';
    slideData.options = {};
    // STEP 3: Set media properties & options
    slideData.options.x = intPosX;
    slideData.options.y = intPosY;
    slideData.options.w = intSizeX;
    slideData.options.h = intSizeY;
    // STEP 4: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
    // NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
    if (strType === 'online') {
        // A: Add video
        target._relsMedia.push({
            path: strPath || 'preencoded' + strExtn,
            data: 'dummy',
            type: 'online',
            extn: strExtn,
            rId: intRels + 1,
            Target: strLink,
        });
        slideData.mediaRid = target._relsMedia[target._relsMedia.length - 1].rId;
        // B: Add preview/overlay image
        target._relsMedia.push({
            path: 'preencoded.png',
            data: IMG_PLAYBTN,
            type: 'image/png',
            extn: 'png',
            rId: intRels + 2,
            Target: '../media/image-' + target._slideNum + '-' + (target._relsMedia.length + 1) + '.png',
        });
    }
    else {
        /* NOTE: Audio/Video files consume *TWO* rId's:
         * <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
         * <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
         */
        // A: "relationships/video"
        target._relsMedia.push({
            path: strPath || 'preencoded' + strExtn,
            type: strType + '/' + strExtn,
            extn: strExtn,
            data: strData || '',
            rId: intRels + 0,
            Target: '../media/media-' + target._slideNum + '-' + (target._relsMedia.length + 1) + '.' + strExtn,
        });
        slideData.mediaRid = target._relsMedia[target._relsMedia.length - 1].rId;
        // B: "relationships/media"
        target._relsMedia.push({
            path: strPath || 'preencoded' + strExtn,
            type: strType + '/' + strExtn,
            extn: strExtn,
            data: strData || '',
            rId: intRels + 1,
            Target: '../media/media-' + target._slideNum + '-' + (target._relsMedia.length + 0) + '.' + strExtn,
        });
        // C: Add preview/overlay image
        target._relsMedia.push({
            data: IMG_PLAYBTN,
            path: 'preencoded.png',
            type: 'image/png',
            extn: 'png',
            rId: intRels + 2,
            Target: '../media/image-' + target._slideNum + '-' + (target._relsMedia.length + 1) + '.png',
        });
    }
    // LAST
    target._slideObjects.push(slideData);
}
/**
 * Adds Notes to a slide.
 * @param {PresSlide} `target` slide object
 * @param {string} `notes`
 * @since 2.3.0
 */
function addNotesDefinition(target, notes) {
    target._slideObjects.push({
        _type: SLIDE_OBJECT_TYPES.notes,
        text: [{ text: notes }],
    });
}
/**
 * Adds a shape object to a slide definition.
 * @param {PresSlide} target slide object that the shape should be added to
 * @param {SHAPE_NAME} shapeName shape name
 * @param {ShapeProps} opts shape options
 */
function addShapeDefinition(target, shapeName, opts) {
    var options = typeof opts === 'object' ? opts : {};
    options.line = options.line || { type: 'none' };
    var newObject = {
        _type: SLIDE_OBJECT_TYPES.text,
        shape: shapeName || SHAPE_TYPE.RECTANGLE,
        options: options,
        text: null,
    };
    // Reality check
    if (!shapeName)
        throw new Error('Missing/Invalid shape parameter! Example: `addShape(pptxgen.shapes.LINE, {x:1, y:1, w:1, h:1});`');
    // 1: ShapeLineProps defaults
    var newLineOpts = {
        type: options.line.type || 'solid',
        color: options.line.color || DEF_SHAPE_LINE_COLOR,
        transparency: options.line.transparency || 0,
        width: options.line.width || 1,
        dashType: options.line.dashType || 'solid',
        beginArrowType: options.line.beginArrowType || null,
        endArrowType: options.line.endArrowType || null,
    };
    if (typeof options.line === 'object' && options.line.type !== 'none')
        options.line = newLineOpts;
    // 2: Set options defaults
    options.x = options.x || (options.x === 0 ? 0 : 1);
    options.y = options.y || (options.y === 0 ? 0 : 1);
    options.w = options.w || (options.w === 0 ? 0 : 1);
    options.h = options.h || (options.h === 0 ? 0 : 1);
    // 3: Handle line (lots of deprecated opts)
    if (typeof options.line === 'string') {
        var tmpOpts = newLineOpts;
        tmpOpts.color = options.line + ''; // @deprecated `options.line` string (was line color)
        options.line = tmpOpts;
    }
    if (typeof options.lineSize === 'number')
        options.line.width = options.lineSize; // @deprecated (part of `ShapeLineProps` now)
    if (typeof options.lineDash === 'string')
        options.line.dashType = options.lineDash; // @deprecated (part of `ShapeLineProps` now)
    if (typeof options.lineHead === 'string')
        options.line.beginArrowType = options.lineHead; // @deprecated (part of `ShapeLineProps` now)
    if (typeof options.lineTail === 'string')
        options.line.endArrowType = options.lineTail; // @deprecated (part of `ShapeLineProps` now)
    // 4: Create hyperlink rels
    createHyperlinkRels(target, newObject);
    // LAST: Add object to slide
    target._slideObjects.push(newObject);
}
/**
 * Adds a table object to a slide definition.
 * @param {PresSlide} target - slide object that the table should be added to
 * @param {TableRow[]} tableRows - table data
 * @param {TableProps} options - table options
 * @param {SlideLayout} slideLayout - Slide layout
 * @param {PresLayout} presLayout - Presentation layout
 * @param {Function} addSlide - method
 * @param {Function} getSlide - method
 */
function addTableDefinition(target, tableRows, options, slideLayout, presLayout, addSlide, getSlide) {
    var opt = options && typeof options === 'object' ? options : {};
    var slides = [target]; // Create array of Slides as more may be added by auto-paging
    // STEP 1: REALITY-CHECK
    {
        // A: check for empty
        if (tableRows === null || tableRows.length === 0 || !Array.isArray(tableRows)) {
            throw new Error("addTable: Array expected! EX: 'slide.addTable( [rows], {options} );' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)");
        }
        // B: check for non-well-formatted array (ex: rows=['a','b'] instead of [['a','b']])
        if (!tableRows[0] || !Array.isArray(tableRows[0])) {
            throw new Error("addTable: 'rows' should be an array of cells! EX: 'slide.addTable( [ ['A'], ['B'], {text:'C',options:{align:'center'}} ] );' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)");
        }
        // TODO: FUTURE: This is wacky and wont function right (shows .w value when there is none from demo.js?!) 20191219
        /*
        if (opt.w && opt.colW) {
            console.warn('addTable: please use either `colW` or `w` - not both (table will use `colW` and ignore `w`)')
            console.log(`${opt.w} ${opt.colW}`)
        }
        */
    }
    // STEP 2: Transform `tableRows` into well-formatted TableCell's
    // tableRows can be object or plain text array: `[{text:'cell 1'}, {text:'cell 2', options:{color:'ff0000'}}]` | `["cell 1", "cell 2"]`
    var arrRows = [];
    tableRows.forEach(function (row) {
        var newRow = [];
        if (Array.isArray(row)) {
            row.forEach(function (cell) {
                // A:
                var newCell = {
                    _type: SLIDE_OBJECT_TYPES.tablecell,
                    text: '',
                    options: typeof cell === 'object' && cell.options ? cell.options : {},
                };
                // B:
                if (typeof cell === 'string' || typeof cell === 'number')
                    newCell.text = cell.toString();
                else if (cell.text) {
                    // Cell can contain complex text type, or string, or number
                    if (typeof cell.text === 'string' || typeof cell.text === 'number')
                        newCell.text = cell.text.toString();
                    else if (cell.text)
                        newCell.text = cell.text;
                    // Capture options
                    if (cell.options && typeof cell.options === 'object')
                        newCell.options = cell.options;
                }
                // C: Set cell borders
                newCell.options.border = newCell.options.border || opt.border || [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }];
                var cellBorder = newCell.options.border;
                // CASE 1: border interface is: BorderOptions | [BorderOptions, BorderOptions, BorderOptions, BorderOptions]
                if (!Array.isArray(cellBorder) && typeof cellBorder === 'object')
                    newCell.options.border = [cellBorder, cellBorder, cellBorder, cellBorder];
                // Handle: [null, null, {type:'solid'}, null]
                if (!newCell.options.border[0])
                    newCell.options.border[0] = { type: 'none' };
                if (!newCell.options.border[1])
                    newCell.options.border[1] = { type: 'none' };
                if (!newCell.options.border[2])
                    newCell.options.border[2] = { type: 'none' };
                if (!newCell.options.border[3])
                    newCell.options.border[3] = { type: 'none' };
                // set complete BorderOptions for all sides
                var arrSides = [0, 1, 2, 3];
                arrSides.forEach(function (idx) {
                    newCell.options.border[idx] = {
                        type: newCell.options.border[idx].type || DEF_CELL_BORDER.type,
                        color: newCell.options.border[idx].color || DEF_CELL_BORDER.color,
                        pt: typeof newCell.options.border[idx].pt === 'number' ? newCell.options.border[idx].pt : DEF_CELL_BORDER.pt,
                    };
                });
                // LAST:
                newRow.push(newCell);
            });
        }
        else {
            console.log('addTable: tableRows has a bad row. A row should be an array of cells. You provided:');
            console.log(row);
        }
        arrRows.push(newRow);
    });
    // STEP 3: Set options
    opt.x = getSmartParseNumber(opt.x || (opt.x === 0 ? 0 : EMU / 2), 'X', presLayout);
    opt.y = getSmartParseNumber(opt.y || (opt.y === 0 ? 0 : EMU / 2), 'Y', presLayout);
    if (opt.h)
        opt.h = getSmartParseNumber(opt.h, 'Y', presLayout); // NOTE: Dont set default `h` - leaving it null triggers auto-rowH in `makeXMLSlide()`
    opt.fontSize = opt.fontSize || DEF_FONT_SIZE;
    opt.margin = opt.margin === 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_PT;
    if (typeof opt.margin === 'number')
        opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)];
    if (!opt.color)
        opt.color = opt.color || DEF_FONT_COLOR; // Set default color if needed (table option > inherit from Slide > default to black)
    if (typeof opt.border === 'string') {
        console.warn("addTable `border` option must be an object. Ex: `{border: {type:'none'}}`");
        opt.border = null;
    }
    else if (Array.isArray(opt.border)) {
        [0, 1, 2, 3].forEach(function (idx) {
            opt.border[idx] = opt.border[idx]
                ? { type: opt.border[idx].type || DEF_CELL_BORDER.type, color: opt.border[idx].color || DEF_CELL_BORDER.color, pt: opt.border[idx].pt || DEF_CELL_BORDER.pt }
                : { type: 'none' };
        });
    }
    opt.autoPage = typeof opt.autoPage === 'boolean' ? opt.autoPage : false;
    opt.autoPageRepeatHeader = typeof opt.autoPageRepeatHeader === 'boolean' ? opt.autoPageRepeatHeader : false;
    opt.autoPageHeaderRows = typeof opt.autoPageHeaderRows !== 'undefined' && !isNaN(Number(opt.autoPageHeaderRows)) ? Number(opt.autoPageHeaderRows) : 1;
    opt.autoPageLineWeight = typeof opt.autoPageLineWeight !== 'undefined' && !isNaN(Number(opt.autoPageLineWeight)) ? Number(opt.autoPageLineWeight) : 0;
    if (opt.autoPageLineWeight) {
        if (opt.autoPageLineWeight > 1)
            opt.autoPageLineWeight = 1;
        else if (opt.autoPageLineWeight < -1)
            opt.autoPageLineWeight = -1;
    }
    // autoPage ^^^
    // Set/Calc table width
    // Get slide margins - start with default values, then adjust if master or slide margins exist
    var arrTableMargin = DEF_SLIDE_MARGIN_IN;
    // Case 1: Master margins
    if (slideLayout && typeof slideLayout._margin !== 'undefined') {
        if (Array.isArray(slideLayout._margin))
            arrTableMargin = slideLayout._margin;
        else if (!isNaN(Number(slideLayout._margin)))
            arrTableMargin = [Number(slideLayout._margin), Number(slideLayout._margin), Number(slideLayout._margin), Number(slideLayout._margin)];
    }
    // Case 2: Table margins
    /* FIXME: add `_margin` option to slide options
        else if ( addNewSlide._margin ) {
            if ( Array.isArray(addNewSlide._margin) ) arrTableMargin = addNewSlide._margin;
            else if ( !isNaN(Number(addNewSlide._margin)) ) arrTableMargin = [Number(addNewSlide._margin), Number(addNewSlide._margin), Number(addNewSlide._margin), Number(addNewSlide._margin)];
        }
    */
    /**
     * Calc table width depending upon what data we have - several scenarios exist (including bad data, eg: colW doesnt match col count)
     * The API does not require a `w` value, but XML generation does, hence, code to calc a width below using colW value(s)
     */
    if (opt.colW) {
        var firstRowColCnt = arrRows[0].reduce(function (totalLen, c) {
            if (c && c.options && c.options.colspan && typeof c.options.colspan === 'number') {
                totalLen += c.options.colspan;
            }
            else {
                totalLen += 1;
            }
            return totalLen;
        }, 0);
        // Ex: `colW = 3` or `colW = '3'`
        if (typeof opt.colW === 'string' || typeof opt.colW === 'number') {
            opt.w = Math.floor(Number(opt.colW) * firstRowColCnt);
            opt.colW = null; // IMPORTANT: Unset `colW` so table is created using `opt.w`, which will evenly divide cols
        }
        // Ex: `colW=[3]` but with >1 cols (same as above, user is saying "use this width for all")
        else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length === 1 && firstRowColCnt > 1) {
            opt.w = Math.floor(Number(opt.colW) * firstRowColCnt);
            opt.colW = null; // IMPORTANT: Unset `colW` so table is created using `opt.w`, which will evenly divide cols
        }
        // Err: Mismatched colW and cols count
        else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length !== firstRowColCnt) {
            console.warn('addTable: mismatch: (colW.length != data.length) Therefore, defaulting to evenly distributed col widths.');
            opt.colW = null;
        }
    }
    else if (opt.w) {
        opt.w = getSmartParseNumber(opt.w, 'X', presLayout);
    }
    else {
        opt.w = Math.floor(presLayout._sizeW / EMU - arrTableMargin[1] - arrTableMargin[3]);
    }
    // STEP 4: Convert units to EMU now (we use different logic in makeSlide->table - smartCalc is not used)
    if (opt.x && opt.x < 20)
        opt.x = inch2Emu(opt.x);
    if (opt.y && opt.y < 20)
        opt.y = inch2Emu(opt.y);
    if (opt.w && opt.w < 20)
        opt.w = inch2Emu(opt.w);
    if (opt.h && opt.h < 20)
        opt.h = inch2Emu(opt.h);
    // STEP 5: Loop over cells: transform each to ITableCell; check to see whether to skip autopaging while here
    arrRows.forEach(function (row) {
        row.forEach(function (cell, idy) {
            // A: Transform cell data if needed
            /* Table rows can be an object or plain text - transform into object when needed
                // EX:
                var arrTabRows1 = [
                    [ { text:'A1\nA2', options:{rowspan:2, fill:'99FFCC'} } ]
                    ,[ 'B2', 'C2', 'D2', 'E2' ]
                ]
            */
            if (typeof cell === 'number' || typeof cell === 'string') {
                // Grab table formatting `opts` to use here so text style/format inherits as it should
                row[idy] = { _type: SLIDE_OBJECT_TYPES.tablecell, text: row[idy].toString(), options: opt };
            }
            else if (typeof cell === 'object') {
                // ARG0: `text`
                if (typeof cell.text === 'number')
                    row[idy].text = row[idy].text.toString();
                else if (typeof cell.text === 'undefined' || cell.text === null)
                    row[idy].text = '';
                // ARG1: `options`: ensure options exists
                row[idy].options = cell.options || {};
                // Set type to tabelcell
                row[idy]._type = SLIDE_OBJECT_TYPES.tablecell;
            }
            // B: Check for fine-grained formatting, disable auto-page when found
            // Since genXmlTextBody already checks for text array ( text:[{},..{}] ) we're done!
            // Text in individual cells will be formatted as they are added by calls to genXmlTextBody within table builder
            if (cell.text && Array.isArray(cell.text))
                opt.autoPage = false;
        });
    });
    // STEP 6: Auto-Paging: (via {options} and used internally)
    // (used internally by `tableToSlides()` to not engage recursion - we've already paged the table data, just add this one)
    if (opt && opt.autoPage === false) {
        // Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
        createHyperlinkRels(target, arrRows);
        // Add slideObjects (NOTE: Use `extend` to avoid mutation)
        target._slideObjects.push({
            _type: SLIDE_OBJECT_TYPES.table,
            arrTabRows: arrRows,
            options: Object.assign({}, opt),
        });
    }
    else {
        if (opt.autoPageRepeatHeader)
            opt._arrObjTabHeadRows = arrRows.filter(function (_row, idx) { return idx < opt.autoPageHeaderRows; });
        // Loop over rows and create 1-N tables as needed (ISSUE#21)
        getSlidesForTableRows(arrRows, opt, presLayout, slideLayout).forEach(function (slide, idx) {
            // A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
            if (!getSlide(target._slideNum + idx))
                slides.push(addSlide(slideLayout ? slideLayout._name : null));
            // B: Reset opt.y to `option`/`margin` after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
            if (idx > 0)
                opt.y = inch2Emu(opt.autoPageSlideStartY || opt.newSlideStartY || arrTableMargin[0]);
            // C: Add this table to new Slide
            {
                var newSlide = getSlide(target._slideNum + idx);
                opt.autoPage = false;
                // Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
                createHyperlinkRels(newSlide, slide.rows);
                // Add rows to new slide
                newSlide.addTable(slide.rows, Object.assign({}, opt));
            }
        });
    }
}
/**
 * Adds a text object to a slide definition.
 * @param {PresSlide} target - slide object that the text should be added to
 * @param {string|TextProps[]} text text string or object
 * @param {TextPropsOptions} opts text options
 * @param {boolean} isPlaceholder whether this a placeholder object
 * @since: 1.0.0
 */
function addTextDefinition(target, text, opts, isPlaceholder) {
    var newObject = {
        _type: isPlaceholder ? SLIDE_OBJECT_TYPES.placeholder : SLIDE_OBJECT_TYPES.text,
        shape: (opts && opts.shape) || SHAPE_TYPE.RECTANGLE,
        text: !text || text.length === 0 ? [{ text: '', options: null }] : text,
        options: opts || {},
    };
    function cleanOpts(itemOpts) {
        // STEP 1: Set some options
        {
            // A.1: Color (placeholders should inherit their colors or override them, so don't default them)
            if (!itemOpts.placeholder) {
                itemOpts.color = itemOpts.color || newObject.options.color || target.color || DEF_FONT_COLOR;
            }
            // A.2: Placeholder should inherit their bullets or override them, so don't default them
            if (itemOpts.placeholder || isPlaceholder) {
                itemOpts.bullet = itemOpts.bullet || false;
            }
            // A.3: Text targeting a placeholder need to inherit the placeholders options (eg: margin, valign, etc.) (Issue #640)
            if (itemOpts.placeholder && target._slideLayout && target._slideLayout._slideObjects) {
                var placeHold = target._slideLayout._slideObjects.filter(function (item) { return item._type === 'placeholder' && item.options && item.options.placeholder && item.options.placeholder === itemOpts.placeholder; })[0];
                if (placeHold && placeHold.options)
                    itemOpts = __assign(__assign({}, itemOpts), placeHold.options);
            }
            // B:
            if (itemOpts.shape === SHAPE_TYPE.LINE) {
                // ShapeLineProps defaults
                var newLineOpts = {
                    type: itemOpts.line.type || 'solid',
                    color: itemOpts.line.color || DEF_SHAPE_LINE_COLOR,
                    transparency: itemOpts.line.transparency || 0,
                    width: itemOpts.line.width || 1,
                    dashType: itemOpts.line.dashType || 'solid',
                    beginArrowType: itemOpts.line.beginArrowType || null,
                    endArrowType: itemOpts.line.endArrowType || null,
                };
                if (typeof itemOpts.line === 'object')
                    itemOpts.line = newLineOpts;
                // 3: Handle line (lots of deprecated opts)
                if (typeof itemOpts.line === 'string') {
                    var tmpOpts = newLineOpts;
                    tmpOpts.color = itemOpts.line.toString(); // @deprecated `itemOpts.line` string (was line color)
                    itemOpts.line = tmpOpts;
                }
                if (typeof itemOpts.lineSize === 'number')
                    itemOpts.line.width = itemOpts.lineSize; // @deprecated (part of `ShapeLineProps` now)
                if (typeof itemOpts.lineDash === 'string')
                    itemOpts.line.dashType = itemOpts.lineDash; // @deprecated (part of `ShapeLineProps` now)
                if (typeof itemOpts.lineHead === 'string')
                    itemOpts.line.beginArrowType = itemOpts.lineHead; // @deprecated (part of `ShapeLineProps` now)
                if (typeof itemOpts.lineTail === 'string')
                    itemOpts.line.endArrowType = itemOpts.lineTail; // @deprecated (part of `ShapeLineProps` now)
            }
            // C: Line opts
            itemOpts.line = itemOpts.line || {};
            itemOpts.lineSpacing = itemOpts.lineSpacing && !isNaN(itemOpts.lineSpacing) ? itemOpts.lineSpacing : null;
            itemOpts.lineSpacingMultiple = itemOpts.lineSpacingMultiple && !isNaN(itemOpts.lineSpacingMultiple) ? itemOpts.lineSpacingMultiple : null;
            // D: Transform text options to bodyProperties as thats how we build XML
            itemOpts._bodyProp = itemOpts._bodyProp || {};
            itemOpts._bodyProp.autoFit = itemOpts.autoFit || false; // DEPRECATED: (3.3.0) If true, shape will collapse to text size (Fit To shape)
            itemOpts._bodyProp.anchor = !itemOpts.placeholder ? TEXT_VALIGN.ctr : null; // VALS: [t,ctr,b]
            itemOpts._bodyProp.vert = itemOpts.vert || null; // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
            itemOpts._bodyProp.wrap = typeof itemOpts.wrap === 'boolean' ? itemOpts.wrap : true;
            // E: Inset
            if ((itemOpts.inset && !isNaN(Number(itemOpts.inset))) || itemOpts.inset === 0) {
                itemOpts._bodyProp.lIns = inch2Emu(itemOpts.inset);
                itemOpts._bodyProp.rIns = inch2Emu(itemOpts.inset);
                itemOpts._bodyProp.tIns = inch2Emu(itemOpts.inset);
                itemOpts._bodyProp.bIns = inch2Emu(itemOpts.inset);
            }
            // F: Transform @deprecated props
            if (typeof itemOpts.underline === 'boolean' && itemOpts.underline === true)
                itemOpts.underline = { style: 'sng' };
        }
        // STEP 2: Transform `align`/`valign` to XML values, store in _bodyProp for XML gen
        {
            if ((itemOpts.align || '').toLowerCase().indexOf('c') === 0)
                itemOpts._bodyProp.align = TEXT_HALIGN.center;
            else if ((itemOpts.align || '').toLowerCase().indexOf('l') === 0)
                itemOpts._bodyProp.align = TEXT_HALIGN.left;
            else if ((itemOpts.align || '').toLowerCase().indexOf('r') === 0)
                itemOpts._bodyProp.align = TEXT_HALIGN.right;
            else if ((itemOpts.align || '').toLowerCase().indexOf('j') === 0)
                itemOpts._bodyProp.align = TEXT_HALIGN.justify;
            if ((itemOpts.valign || '').toLowerCase().indexOf('b') === 0)
                itemOpts._bodyProp.anchor = TEXT_VALIGN.b;
            else if ((itemOpts.valign || '').toLowerCase().indexOf('m') === 0)
                itemOpts._bodyProp.anchor = TEXT_VALIGN.ctr;
            else if ((itemOpts.valign || '').toLowerCase().indexOf('t') === 0)
                itemOpts._bodyProp.anchor = TEXT_VALIGN.t;
        }
        // STEP 3: ROBUST: Set rational values for some shadow props if needed
        correctShadowOptions(itemOpts.shadow);
        return itemOpts;
    }
    // STEP 1: Create/Clean object options
    newObject.options = cleanOpts(newObject.options);
    // STEP 2: Create/Clean text options
    newObject.text.forEach(function (item) { return (item.options = cleanOpts(item.options || {})); });
    // STEP 3: Create hyperlinks
    createHyperlinkRels(target, newObject.text || '');
    // LAST: Add object to Slide
    target._slideObjects.push(newObject);
}
/**
 * Adds placeholder objects to slide
 * @param {PresSlide} slide - slide object containing layouts
 */
function addPlaceholdersToSlideLayouts(slide) {
    (slide._slideLayout._slideObjects || []).forEach(function (slideLayoutObj) {
        if (slideLayoutObj._type === SLIDE_OBJECT_TYPES.placeholder) {
            // A: Search for this placeholder on Slide before we add
            // NOTE: Check to ensure a placeholder does not already exist on the Slide
            // They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
            if (slide._slideObjects.filter(function (slideObj) { return slideObj.options && slideObj.options.placeholder === slideLayoutObj.options.placeholder; }).length === 0) {
                addTextDefinition(slide, [{ text: '' }], slideLayoutObj.options, false);
            }
        }
    });
}
/* -------------------------------------------------------------------------------- */
/**
 * Adds a background image or color to a slide definition.
 * @param {BackgroundProps} props - color string or an object with image definition
 * @param {PresSlide} target - slide object that the background is set to
 */
function addBackgroundDefinition(props, target) {
    // A: @deprecated
    if (target.bkgd) {
        if (!target.background)
            target.background = {};
        if (typeof target.bkgd === 'string')
            target.background.color = target.bkgd;
        else {
            if (target.bkgd.data)
                target.background.data = target.bkgd.data;
            if (target.bkgd.path)
                target.background.path = target.bkgd.path;
            if (target.bkgd['src'])
                target.background.path = target.bkgd['src']; // @deprecated (drop in 4.x)
        }
    }
    if (target.background && target.background.fill)
        target.background.color = target.background.fill;
    // B: Handle media
    if (props && (props.path || props.data)) {
        // Allow the use of only the data key (`path` isnt reqd)
        props.path = props.path || 'preencoded.png';
        var strImgExtn = (props.path.split('.').pop() || 'png').split('?')[0]; // Handle "blah.jpg?width=540" etc.
        if (strImgExtn === 'jpg')
            strImgExtn = 'jpeg'; // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup
        target._relsMedia = target._relsMedia || [];
        var intRels = target._relsMedia.length + 1;
        // NOTE: `Target` cannot have spaces (eg:"Slide 1-image-1.jpg") or a "presentation is corrupt" warning comes up
        target._relsMedia.push({
            path: props.path,
            type: SLIDE_OBJECT_TYPES.image,
            extn: strImgExtn,
            data: props.data || null,
            rId: intRels,
            Target: "../media/" + (target._name || '').replace(/\s+/gi, '-') + "-image-" + (target._relsMedia.length + 1) + "." + strImgExtn,
        });
        target._bkgdImgRid = intRels;
    }
}
/**
 * Parses text/text-objects from `addText()` and `addTable()` methods; creates 'hyperlink'-type Slide Rels for each hyperlink found
 * @param {PresSlide} target - slide object that any hyperlinks will be be added to
 * @param {number | string | TextProps | TextProps[] | ITableCell[][]} text - text to parse
 */
function createHyperlinkRels(target, text) {
    var textObjs = [];
    // Only text objects can have hyperlinks, bail when text param is plain text
    if (typeof text === 'string' || typeof text === 'number')
        return;
    // IMPORTANT: "else if" Array.isArray must come before typeof===object! Otherwise, code will exhaust recursion!
    else if (Array.isArray(text))
        textObjs = text;
    else if (typeof text === 'object')
        textObjs = [text];
    textObjs.forEach(function (text) {
        // `text` can be an array of other `text` objects (table cell word-level formatting), continue parsing using recursion
        if (Array.isArray(text)) {
            createHyperlinkRels(target, text);
        }
        else if (Array.isArray(text.text)) {
            // this handles TableCells with hyperlinks
            createHyperlinkRels(target, text.text);
        }
        else if (text && typeof text === 'object' && text.options && text.options.hyperlink && !text.options.hyperlink._rId) {
            if (typeof text.options.hyperlink !== 'object')
                console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:'https://github.com'}` ");
            else if (!text.options.hyperlink.url && !text.options.hyperlink.slide)
                console.log("ERROR: 'hyperlink requires either: `url` or `slide`'");
            else {
                var relId = getNewRelId(target);
                target._rels.push({
                    type: SLIDE_OBJECT_TYPES.hyperlink,
                    data: text.options.hyperlink.slide ? 'slide' : 'dummy',
                    rId: relId,
                    Target: encodeXmlEntities(text.options.hyperlink.url) || text.options.hyperlink.slide.toString(),
                });
                text.options.hyperlink._rId = relId;
            }
        }
    });
}

/**
 * PptxGenJS: Slide Class
 */
var Slide = /** @class */ (function () {
    function Slide(params) {
        this.addSlide = params.addSlide;
        this.getSlide = params.getSlide;
        this._name = 'Slide ' + params.slideNumber;
        this._presLayout = params.presLayout;
        this._rId = params.slideRId;
        this._rels = [];
        this._relsChart = [];
        this._relsMedia = [];
        this._setSlideNum = params.setSlideNum;
        this._slideId = params.slideId;
        this._slideLayout = params.slideLayout || null;
        this._slideNum = params.slideNumber;
        this._slideObjects = [];
        /** NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
         * `defineSlideMaster` and `addNewSlide.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
         * so, lastly, add to the Slide now.
         */
        this._slideNumberProps = this._slideLayout && this._slideLayout._slideNumberProps ? this._slideLayout._slideNumberProps : null;
    }
    Object.defineProperty(Slide.prototype, "bkgd", {
        get: function () {
            return this._bkgd;
        },
        set: function (value) {
            this._bkgd = value;
            if (!this._background || !this._background.color) {
                if (!this._background)
                    this._background = {};
                if (typeof value === 'string')
                    this._background.color = value;
            }
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Slide.prototype, "background", {
        get: function () {
            return this._background;
        },
        set: function (props) {
            this._background = props;
            // Add background (image data/path must be captured before `exportPresentation()` is called)
            if (props)
                addBackgroundDefinition(props, this);
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Slide.prototype, "color", {
        get: function () {
            return this._color;
        },
        set: function (value) {
            this._color = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Slide.prototype, "hidden", {
        get: function () {
            return this._hidden;
        },
        set: function (value) {
            this._hidden = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Slide.prototype, "slideNumber", {
        get: function () {
            return this._slideNumberProps;
        },
        /**
         * @type {SlideNumberProps}
         */
        set: function (value) {
            // NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
            this._slideNumberProps = value;
            this._setSlideNum(value);
        },
        enumerable: false,
        configurable: true
    });
    /**
     * Add chart to Slide
     * @param {CHART_NAME|IChartMulti[]} type - chart type
     * @param {object[]} data - data object
     * @param {IChartOpts} options - chart options
     * @return {Slide} this Slide
     */
    Slide.prototype.addChart = function (type, data, options) {
        // FUTURE: TODO-VERSION-4: Remove first arg - only take data and opts, with "type" required on opts
        // Set `_type` on IChartOptsLib as its what is used as object is passed around
        var optionsWithType = options || {};
        optionsWithType._type = type;
        addChartDefinition(this, type, data, options);
        return this;
    };
    /**
     * Add image to Slide
     * @param {ImageProps} options - image options
     * @return {Slide} this Slide
     */
    Slide.prototype.addImage = function (options) {
        addImageDefinition(this, options);
        return this;
    };
    /**
     * Add media (audio/video) to Slide
     * @param {MediaProps} options - media options
     * @return {Slide} this Slide
     */
    Slide.prototype.addMedia = function (options) {
        addMediaDefinition(this, options);
        return this;
    };
    /**
     * Add speaker notes to Slide
     * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
     * @param {string} notes - notes to add to slide
     * @return {Slide} this Slide
     */
    Slide.prototype.addNotes = function (notes) {
        addNotesDefinition(this, notes);
        return this;
    };
    /**
     * Add shape to Slide
     * @param {SHAPE_NAME} shapeName - shape name
     * @param {ShapeProps} options - shape options
     * @return {Slide} this Slide
     */
    Slide.prototype.addShape = function (shapeName, options) {
        // NOTE: As of v3.1.0, <script> users are passing the old shape object from the shapes file (orig to the project)
        // But React/TypeScript users are passing the shapeName from an enum, which is a simple string, so lets cast
        // <script./> => `pptx.shapes.RECTANGLE` [string] "rect" ... shapeName['name'] = 'rect'
        // TypeScript => `pptxgen.shapes.RECTANGLE` [string] "rect" ... shapeName = 'rect'
        //let shapeNameDecode = typeof shapeName === 'object' && shapeName['name'] ? shapeName['name'] : shapeName
        addShapeDefinition(this, shapeName, options);
        return this;
    };
    /**
     * Add table to Slide
     * @param {TableRow[]} tableRows - table rows
     * @param {TableProps} options - table options
     * @return {Slide} this Slide
     */
    Slide.prototype.addTable = function (tableRows, options) {
        // FUTURE: we pass `this` - we dont need to pass layouts - they can be read from this!
        addTableDefinition(this, tableRows, options, this._slideLayout, this._presLayout, this.addSlide, this.getSlide);
        return this;
    };
    /**
     * Add text to Slide
     * @param {string|TextProps[]} text - text string or complex object
     * @param {TextPropsOptions} options - text options
     * @return {Slide} this Slide
     */
    Slide.prototype.addText = function (text, options) {
        var textParam = typeof text === 'string' || typeof text === 'number' ? [{ text: text, options: options }] : text;
        addTextDefinition(this, textParam, options, false);
        return this;
    };
    return Slide;
}());

/**
 * PptxGenJS: Chart Generation
 */
/**
 * Based on passed data, creates Excel Worksheet that is used as a data source for a chart.
 * @param {ISlideRelChart} chartObject - chart object
 * @param {JSZip} zip - file that the resulting XLSX should be added to
 * @return {Promise} promise of generating the XLSX file
 */
function createExcelWorksheet(chartObject, zip) {
    var data = chartObject.data;
    return new Promise(function (resolve, reject) {
        var zipExcel = new JSZip();
        var intBubbleCols = (data.length - 1) * 2 + 1; // 1 for "X-Values", then 2 for every Y-Axis
        // A: Add folders
        zipExcel.folder('_rels');
        zipExcel.folder('docProps');
        zipExcel.folder('xl/_rels');
        zipExcel.folder('xl/tables');
        zipExcel.folder('xl/theme');
        zipExcel.folder('xl/worksheets');
        zipExcel.folder('xl/worksheets/_rels');
        // B: Add core contents
        {
            zipExcel.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                '  <Default Extension="xml" ContentType="application/xml"/>' +
                '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                //+ '  <Default Extension="jpeg" ContentType="image/jpg"/><Default Extension="png" ContentType="image/png"/>'
                //+ '  <Default Extension="bmp" ContentType="image/bmp"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="tif" ContentType="image/tif"/><Default Extension="pdf" ContentType="application/pdf"/><Default Extension="mov" ContentType="application/movie"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
                //+ '  <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
                '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                '  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
                '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
                '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
                '  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
                '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
                '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
                '</Types>\n');
            zipExcel.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                '</Relationships>\n');
            zipExcel.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
                '<Application>Microsoft Excel</Application>' +
                '<DocSecurity>0</DocSecurity>' +
                '<ScaleCrop>false</ScaleCrop>' +
                '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>' +
                '</Properties>\n');
            zipExcel.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
                '<dc:creator>PptxGenJS</dc:creator>' +
                '<cp:lastModifiedBy>Ely, Brent</cp:lastModifiedBy>' +
                '<dcterms:created xsi:type="dcterms:W3CDTF">' +
                new Date().toISOString() +
                '</dcterms:created>' +
                '<dcterms:modified xsi:type="dcterms:W3CDTF">' +
                new Date().toISOString() +
                '</dcterms:modified>' +
                '</cp:coreProperties>\n');
            zipExcel.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
                '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
                '</Relationships>\n');
            zipExcel.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="0" formatCode="General"/></numFmts><fonts count="4"><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="10"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="18"/><color indexed="8"/>' +
                '<name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><dxfs count="0"/><tableStyles count="0"/><colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ffff0000"/><rgbColor rgb="ff00ff00"/><rgbColor rgb="ff0000ff"/>' +
                '<rgbColor rgb="ffffff00"/><rgbColor rgb="ffff00ff"/><rgbColor rgb="ff00ffff"/><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ff878787"/><rgbColor rgb="fff9f9f9"/></indexedColors></colors></styleSheet>\n');
            zipExcel.file('xl/theme/theme1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>');
            zipExcel.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?>' +
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
                '<fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/>' +
                '<workbookPr />' +
                '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="15960" windowHeight="18080"/></bookViews>' +
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" /></sheets>' +
                '<calcPr calcId="171026" concurrentCalc="0"/>' +
                '</workbook>\n');
            zipExcel.file('xl/worksheets/_rels/sheet1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
                '</Relationships>\n');
        }
        // sharedStrings.xml
        {
            // A: Start XML
            var strSharedStrings_1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            if (chartObject.opts._type === CHART_TYPE.BUBBLE) {
                strSharedStrings_1 +=
                    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (intBubbleCols + 1) + '" uniqueCount="' + (intBubbleCols + 1) + '">';
            }
            else if (chartObject.opts._type === CHART_TYPE.SCATTER) {
                strSharedStrings_1 +=
                    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (data.length + 1) + '" uniqueCount="' + (data.length + 1) + '">';
            }
            else {
                strSharedStrings_1 +=
                    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' +
                        (data[0].labels.length + data.length + 1) +
                        '" uniqueCount="' +
                        (data[0].labels.length + data.length + 1) +
                        '">';
                // B: Add 'blank' for A1
                strSharedStrings_1 += '<si><t xml:space="preserve"></t></si>';
            }
            // C: Add `name`/Series
            if (chartObject.opts._type === CHART_TYPE.BUBBLE) {
                data.forEach(function (objData, idx) {
                    if (idx === 0)
                        strSharedStrings_1 += '<si><t>X-Axis</t></si>';
                    else {
                        strSharedStrings_1 += '<si><t>' + encodeXmlEntities(objData.name || ' ') + '</t></si>';
                        strSharedStrings_1 += '<si><t>' + encodeXmlEntities('Size ' + idx) + '</t></si>';
                    }
                });
            }
            else {
                data.forEach(function (objData) {
                    strSharedStrings_1 += '<si><t>' + encodeXmlEntities((objData.name || ' ').replace('X-Axis', 'X-Values')) + '</t></si>';
                });
            }
            // D: Add `labels`/Categories
            if (chartObject.opts._type !== CHART_TYPE.BUBBLE && chartObject.opts._type !== CHART_TYPE.SCATTER) {
                data[0].labels.forEach(function (label) {
                    strSharedStrings_1 += '<si><t>' + encodeXmlEntities(label) + '</t></si>';
                });
            }
            strSharedStrings_1 += '</sst>\n';
            zipExcel.file('xl/sharedStrings.xml', strSharedStrings_1);
        }
        // tables/table1.xml
        {
            var strTableXml_1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            if (chartObject.opts._type === CHART_TYPE.BUBBLE) ;
            else if (chartObject.opts._type === CHART_TYPE.SCATTER) {
                strTableXml_1 +=
                    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:' +
                        LETTERS[data.length - 1] +
                        (data[0].values.length + 1) +
                        '" totalsRowShown="0">';
                strTableXml_1 += '<tableColumns count="' + data.length + '">';
                data.forEach(function (_obj, idx) {
                    strTableXml_1 += '<tableColumn id="' + (idx + 1) + '" name="' + (idx === 0 ? 'X-Values' : 'Y-Value ' + idx) + '" />';
                });
            }
            else {
                strTableXml_1 +=
                    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:' +
                        LETTERS[data.length] +
                        (data[0].labels.length + 1) +
                        '" totalsRowShown="0">';
                strTableXml_1 += '<tableColumns count="' + (data.length + 1) + '">';
                strTableXml_1 += '<tableColumn id="1" name=" " />';
                data.forEach(function (obj, idx) {
                    strTableXml_1 += '<tableColumn id="' + (idx + 2) + '" name="' + encodeXmlEntities(obj.name) + '" />';
                });
            }
            strTableXml_1 += '</tableColumns>';
            strTableXml_1 += '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0" />';
            strTableXml_1 += '</table>';
            zipExcel.file('xl/tables/table1.xml', strTableXml_1);
        }
        // worksheets/sheet1.xml
        {
            var strSheetXml_1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            strSheetXml_1 +=
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
            if (chartObject.opts._type === CHART_TYPE.BUBBLE) {
                strSheetXml_1 += '<dimension ref="A1:' + LETTERS[intBubbleCols - 1] + (data[0].values.length + 1) + '" />';
            }
            else if (chartObject.opts._type === CHART_TYPE.SCATTER) {
                strSheetXml_1 += '<dimension ref="A1:' + LETTERS[data.length - 1] + (data[0].values.length + 1) + '" />';
            }
            else {
                strSheetXml_1 += '<dimension ref="A1:' + LETTERS[data.length] + (data[0].labels.length + 1) + '" />';
            }
            strSheetXml_1 += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1" /></sheetView></sheetViews>';
            strSheetXml_1 += '<sheetFormatPr baseColWidth="10" defaultColWidth="11.5" defaultRowHeight="12" />';
            if (chartObject.opts._type === CHART_TYPE.BUBBLE) {
                strSheetXml_1 += '<cols>';
                strSheetXml_1 += '<col min="1" max="' + data.length + '" width="11" customWidth="1" />';
                strSheetXml_1 += '</cols>';
                /* EX: INPUT: `data`
                [
                    { name:'X-Axis'  , values:[10,11,12,13,14,15,16,17,18,19,20] },
                    { name:'Y-Axis 1', values:[ 1, 6, 7, 8, 9], sizes:[ 4, 5, 6, 7, 8] },
                    { name:'Y-Axis 2', values:[33,32,42,53,63], sizes:[11,12,13,14,15] }
                ];
                */
                /* EX: OUTPUT: bubbleChart Worksheet:
                    -|----A-----|------B-----|------C-----|------D-----|------E-----|
                    1| X-Values | Y-Values 1 | Y-Sizes 1  | Y-Values 2 | Y-Sizes 2  |
                    2|    11    |     22     |      4     |     33     |      8     |
                    -|----------|------------|------------|------------|------------|
                */
                strSheetXml_1 += '<sheetData>';
                // A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
                strSheetXml_1 += '<row r="1" spans="1:' + intBubbleCols + '">';
                strSheetXml_1 += '<c r="A1" t="s"><v>0</v></c>';
                for (var idxBc = 1; idxBc < intBubbleCols; idxBc++) {
                    strSheetXml_1 += '<c r="' + (idxBc < 26 ? LETTERS[idxBc] : 'A' + LETTERS[idxBc % LETTERS.length]) + '1" t="s">'; // NOTE: use `t="s"` for label cols!
                    strSheetXml_1 += '<v>' + idxBc + '</v>';
                    strSheetXml_1 += '</c>';
                }
                strSheetXml_1 += '</row>';
                // B: Add row for each X-Axis value (Y-Axis* value is optional)
                data[0].values.forEach(function (val, idx) {
                    // Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
                    strSheetXml_1 += '<row r="' + (idx + 2) + '" spans="1:' + intBubbleCols + '">';
                    strSheetXml_1 += '<c r="A' + (idx + 2) + '"><v>' + val + '</v></c>';
                    // Add Y-Axis 1->N (idy=0 = Xaxis)
                    var idxColLtr = 1;
                    for (var idy = 1; idy < data.length; idy++) {
                        // y-value
                        strSheetXml_1 += '<c r="' + (idxColLtr < 26 ? LETTERS[idxColLtr] : 'A' + LETTERS[idxColLtr % LETTERS.length]) + '' + (idx + 2) + '">';
                        strSheetXml_1 += '<v>' + (data[idy].values[idx] || '') + '</v>';
                        strSheetXml_1 += '</c>';
                        idxColLtr++;
                        // y-size
                        strSheetXml_1 += '<c r="' + (idxColLtr < 26 ? LETTERS[idxColLtr] : 'A' + LETTERS[idxColLtr % LETTERS.length]) + '' + (idx + 2) + '">';
                        strSheetXml_1 += '<v>' + (data[idy].sizes[idx] || '') + '</v>';
                        strSheetXml_1 += '</c>';
                        idxColLtr++;
                    }
                    strSheetXml_1 += '</row>';
                });
            }
            else if (chartObject.opts._type === CHART_TYPE.SCATTER) {
                strSheetXml_1 += '<cols>';
                strSheetXml_1 += '<col min="1" max="' + data.length + '" width="11" customWidth="1" />';
                //data.forEach((obj,idx)=>{ strSheetXml += '<col min="'+(idx+1)+'" max="'+(idx+1)+'" width="11" customWidth="1" />' });
                strSheetXml_1 += '</cols>';
                /* EX: INPUT: `data`
                [
                    { name:'X-Axis'  , values:[10,11,12,13,14,15,16,17,18,19,20] },
                    { name:'Y-Axis 1', values:[ 1, 6, 7, 8, 9] },
                    { name:'Y-Axis 2', values:[33,32,42,53,63] }
                ];
                */
                /* EX: OUTPUT: scatterChart Worksheet:
                    -|----A-----|------B-----|
                    1| X-Values | Y-Values 1 |
                    2|    11    |     22     |
                    -|----------|------------|
                */
                strSheetXml_1 += '<sheetData>';
                // A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
                strSheetXml_1 += '<row r="1" spans="1:' + data.length + '">';
                strSheetXml_1 += '<c r="A1" t="s"><v>0</v></c>';
                for (var idxSd = 1; idxSd < data.length; idxSd++) {
                    strSheetXml_1 += '<c r="' + (idxSd < 26 ? LETTERS[idxSd] : 'A' + LETTERS[idxSd % LETTERS.length]) + '1" t="s">'; // NOTE: use `t="s"` for label cols!
                    strSheetXml_1 += '<v>' + idxSd + '</v>';
                    strSheetXml_1 += '</c>';
                }
                strSheetXml_1 += '</row>';
                // B: Add row for each X-Axis value (Y-Axis* value is optional)
                data[0].values.forEach(function (val, idx) {
                    // Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
                    strSheetXml_1 += '<row r="' + (idx + 2) + '" spans="1:' + data.length + '">';
                    strSheetXml_1 += '<c r="A' + (idx + 2) + '"><v>' + val + '</v></c>';
                    // Add Y-Axis 1->N
                    for (var idy = 1; idy < data.length; idy++) {
                        strSheetXml_1 += '<c r="' + (idy < 26 ? LETTERS[idy] : 'A' + LETTERS[idy % LETTERS.length]) + '' + (idx + 2) + '">';
                        strSheetXml_1 += '<v>' + (data[idy].values[idx] || data[idy].values[idx] === 0 ? data[idy].values[idx] : '') + '</v>';
                        strSheetXml_1 += '</c>';
                    }
                    strSheetXml_1 += '</row>';
                });
            }
            else {
                strSheetXml_1 += '<cols>';
                strSheetXml_1 += '<col min="1" max="1" width="11" customWidth="1" />';
                //data.forEach(function(){ strSheetXml += '<col min="10" max="100" width="10" customWidth="1" />' });
                strSheetXml_1 += '</cols>';
                strSheetXml_1 += '<sheetData>';
                /* EX: INPUT: `data`
                [
                    { name:'Red', labels:['Jan..May-17'], values:[11,13,14,15,16] },
                    { name:'Amb', labels:['Jan..May-17'], values:[22, 6, 7, 8, 9] },
                    { name:'Grn', labels:['Jan..May-17'], values:[33,32,42,53,63] }
                ];
                */
                /* EX: OUTPUT: lineChart Worksheet:
                    -|---A---|--B--|--C--|--D--|
                    1|       | Red | Amb | Grn |
                    2|Jan-17 |   11|   22|   33|
                    3|Feb-17 |   55|   43|   70|
                    4|Mar-17 |   56|  143|   99|
                    5|Apr-17 |   65|    3|  120|
                    6|May-17 |   75|   93|  170|
                    -|-------|-----|-----|-----|
                */
                // A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
                strSheetXml_1 += '<row r="1" spans="1:' + (data.length + 1) + '">';
                strSheetXml_1 += '<c r="A1" t="s"><v>0</v></c>';
                for (var idx = 1; idx <= data.length; idx++) {
                    // FIXME: Max cols is 52
                    strSheetXml_1 += '<c r="' + (idx < 26 ? LETTERS[idx] : 'A' + LETTERS[idx % LETTERS.length]) + '1" t="s">'; // NOTE: use `t="s"` for label cols!
                    strSheetXml_1 += '<v>' + idx + '</v>';
                    strSheetXml_1 += '</c>';
                }
                strSheetXml_1 += '</row>';
                // B: Add data row(s) for each category
                data[0].labels.forEach(function (_cat, idx) {
                    // Leading col is reserved for the label, so hard-code it, then loop over col values
                    strSheetXml_1 += '<row r="' + (idx + 2) + '" spans="1:' + (data.length + 1) + '">';
                    strSheetXml_1 += '<c r="A' + (idx + 2) + '" t="s">';
                    strSheetXml_1 += '<v>' + (data.length + idx + 1) + '</v>';
                    strSheetXml_1 += '</c>';
                    for (var idy = 0; idy < data.length; idy++) {
                        strSheetXml_1 += '<c r="' + (idy + 1 < 26 ? LETTERS[idy + 1] : 'A' + LETTERS[(idy + 1) % LETTERS.length]) + '' + (idx + 2) + '">';
                        strSheetXml_1 += '<v>' + (data[idy].values[idx] || '') + '</v>';
                        strSheetXml_1 += '</c>';
                    }
                    strSheetXml_1 += '</row>';
                });
            }
            strSheetXml_1 += '</sheetData>';
            strSheetXml_1 += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />';
            // Link the `table1.xml` file to define an actual Table in Excel
            // NOTE: This only works with scatter charts - all others give a "cannot find linked file" error
            // ....: Since we dont need the table anyway (chart data can be edited/range selected, etc.), just dont use this
            // ....: Leaving this so nobody foolishly attempts to add this in the future
            // strSheetXml += '<tableParts count="1"><tablePart r:id="rId1" /></tableParts>';
            strSheetXml_1 += '</worksheet>\n';
            zipExcel.file('xl/worksheets/sheet1.xml', strSheetXml_1);
        }
        // C: Add XLSX to PPTX export
        zipExcel
            .generateAsync({ type: 'base64' })
            .then(function (content) {
            // 1: Create the embedded Excel worksheet with labels and data
            zip.file('ppt/embeddings/Microsoft_Excel_Worksheet' + chartObject.globalId + '.xlsx', content, { base64: true });
            // 2: Create the chart.xml and rel files
            zip.file('ppt/charts/_rels/' + chartObject.fileName + '.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet' +
                chartObject.globalId +
                '.xlsx"/>' +
                '</Relationships>');
            zip.file('ppt/charts/' + chartObject.fileName, makeXmlCharts(chartObject));
            // 3: Done
            resolve(null);
        })
            .catch(function (strErr) {
            reject(strErr);
        });
    });
}
/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 * @param {ISlideRelChart} rel - chart object
 * @return {string} XML
 */
function makeXmlCharts(rel) {
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    var usesSecondaryValAxis = false;
    // STEP 1: Create chart
    {
        // CHARTSPACE: BEGIN vvv
        strXml +=
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        strXml += '<c:date1904 val="0"/>'; // ppt defaults to 1904 dates, excel to 1900
        strXml += '<c:chart>';
        // OPTION: Title
        if (rel.opts.showTitle) {
            strXml += genXmlTitle({
                title: rel.opts.title || 'Chart Title',
                color: rel.opts.titleColor,
                fontFace: rel.opts.titleFontFace,
                fontSize: rel.opts.titleFontSize || DEF_FONT_TITLE_SIZE,
                titleAlign: rel.opts.titleAlign,
                titleBold: rel.opts.titleBold,
                titlePos: rel.opts.titlePos,
                titleRotate: rel.opts.titleRotate,
            });
            strXml += '<c:autoTitleDeleted val="0"/>';
        }
        else {
            // NOTE: Add autoTitleDeleted tag in else to prevent default creation of chart title even when showTitle is set to false
            strXml += '<c:autoTitleDeleted val="1"/>';
        }
        /** Add 3D view tag
         * @see: https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_perspective_topic_ID0E6BUQB.html
         */
        if (rel.opts._type === CHART_TYPE.BAR3D) {
            strXml += '<c:view3D>';
            strXml += ' <c:rotX val="' + rel.opts.v3DRotX + '"/>';
            strXml += ' <c:rotY val="' + rel.opts.v3DRotY + '"/>';
            strXml += ' <c:rAngAx val="' + (rel.opts.v3DRAngAx === false ? 0 : 1) + '"/>';
            strXml += ' <c:perspective val="' + rel.opts.v3DPerspective + '"/>';
            strXml += '</c:view3D>';
        }
        strXml += '<c:plotArea>';
        // IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
        if (rel.opts.layout) {
            strXml += '<c:layout>';
            strXml += ' <c:manualLayout>';
            strXml += '  <c:layoutTarget val="inner" />';
            strXml += '  <c:xMode val="edge" />';
            strXml += '  <c:yMode val="edge" />';
            strXml += '  <c:x val="' + (rel.opts.layout.x || 0) + '" />';
            strXml += '  <c:y val="' + (rel.opts.layout.y || 0) + '" />';
            strXml += '  <c:w val="' + (rel.opts.layout.w || 1) + '" />';
            strXml += '  <c:h val="' + (rel.opts.layout.h || 1) + '" />';
            strXml += ' </c:manualLayout>';
            strXml += '</c:layout>';
        }
        else {
            strXml += '<c:layout/>';
        }
    }
    // A: Create Chart XML -----------------------------------------------------------
    if (Array.isArray(rel.opts._type)) {
        rel.opts._type.forEach(function (type) {
            // TODO: FIXME: theres `options` on chart rels??
            var options = getMix(rel.opts, type.options);
            //let options: IChartOptsLib = { type: type.type, }
            var valAxisId = options['secondaryValAxis'] ? AXIS_ID_VALUE_SECONDARY : AXIS_ID_VALUE_PRIMARY;
            var catAxisId = options['secondaryCatAxis'] ? AXIS_ID_CATEGORY_SECONDARY : AXIS_ID_CATEGORY_PRIMARY;
            usesSecondaryValAxis = usesSecondaryValAxis || options.secondaryValAxis;
            strXml += makeChartType(type.type, type.data, options, valAxisId, catAxisId);
        });
    }
    else {
        strXml += makeChartType(rel.opts._type, rel.data, rel.opts, AXIS_ID_VALUE_PRIMARY, AXIS_ID_CATEGORY_PRIMARY);
    }
    // B: Axes -----------------------------------------------------------
    if (rel.opts._type !== CHART_TYPE.PIE && rel.opts._type !== CHART_TYPE.DOUGHNUT) {
        // Param check
        if (rel.opts.valAxes && rel.opts.valAxes.length > 1 && !usesSecondaryValAxis) {
            throw new Error('Secondary axis must be used by one of the multiple charts');
        }
        if (rel.opts.catAxes) {
            if (!rel.opts.valAxes || rel.opts.valAxes.length !== rel.opts.catAxes.length) {
                throw new Error('There must be the same number of value and category axes.');
            }
            strXml += makeCatAxis(getMix(rel.opts, rel.opts.catAxes[0]), AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY);
            if (rel.opts.catAxes[1]) {
                strXml += makeCatAxis(getMix(rel.opts, rel.opts.catAxes[1]), AXIS_ID_CATEGORY_SECONDARY, AXIS_ID_VALUE_PRIMARY);
            }
        }
        else {
            strXml += makeCatAxis(rel.opts, AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY);
        }
        if (rel.opts.valAxes) {
            strXml += makeValAxis(getMix(rel.opts, rel.opts.valAxes[0]), AXIS_ID_VALUE_PRIMARY);
            if (rel.opts.valAxes[1]) {
                strXml += makeValAxis(getMix(rel.opts, rel.opts.valAxes[1]), AXIS_ID_VALUE_SECONDARY);
            }
        }
        else {
            strXml += makeValAxis(rel.opts, AXIS_ID_VALUE_PRIMARY);
            // Add series axis for 3D bar
            if (rel.opts._type === CHART_TYPE.BAR3D) {
                strXml += makeSerAxis(rel.opts, AXIS_ID_SERIES_PRIMARY, AXIS_ID_VALUE_PRIMARY);
            }
        }
    }
    // C: Chart Properties and plotArea Options: Border, Data Table, Fill, Legend
    {
        // NOTE: DataTable goes between '</c:valAx>' and '<c:spPr>'
        if (rel.opts.showDataTable) {
            strXml += '<c:dTable>';
            strXml += '  <c:showHorzBorder val="' + (rel.opts.showDataTableHorzBorder === false ? 0 : 1) + '"/>';
            strXml += '  <c:showVertBorder val="' + (rel.opts.showDataTableVertBorder === false ? 0 : 1) + '"/>';
            strXml += '  <c:showOutline    val="' + (rel.opts.showDataTableOutline === false ? 0 : 1) + '"/>';
            strXml += '  <c:showKeys       val="' + (rel.opts.showDataTableKeys === false ? 0 : 1) + '"/>';
            strXml += '  <c:spPr>';
            strXml += '    <a:noFill/>';
            strXml +=
                '    <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>';
            strXml += '    <a:effectLst/>';
            strXml += '  </c:spPr>';
            strXml += '  <c:txPr>';
            strXml += '	  <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>';
            strXml += '	  <a:lstStyle/>';
            strXml += '	  <a:p>';
            strXml += '		<a:pPr rtl="0">';
            strXml += "       <a:defRPr sz=\"" + Math.round((rel.opts.dataTableFontSize || DEF_FONT_SIZE) * 100) + "\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">";
            strXml += '			<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>';
            strXml += '			<a:latin typeface="+mn-lt"/>';
            strXml += '			<a:ea typeface="+mn-ea"/>';
            strXml += '			<a:cs typeface="+mn-cs"/>';
            strXml += '		  </a:defRPr>';
            strXml += '		</a:pPr>';
            strXml += '		<a:endParaRPr lang="en-US"/>';
            strXml += '	  </a:p>';
            strXml += '	</c:txPr>';
            strXml += '</c:dTable>';
        }
        strXml += '  <c:spPr>';
        // OPTION: Fill
        strXml += rel.opts.fill ? genXmlColorSelection(rel.opts.fill) : '<a:noFill/>';
        // OPTION: Border
        strXml += rel.opts.border ? "<a:ln w=\"" + valToPts(rel.opts.border.pt) + "\" cap=\"flat\">" + genXmlColorSelection(rel.opts.border.color) + "</a:ln>" : '<a:ln><a:noFill/></a:ln>';
        // Close shapeProp/plotArea before Legend
        strXml += '    <a:effectLst/>';
        strXml += '  </c:spPr>';
        strXml += '</c:plotArea>';
        // OPTION: Legend
        // IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
        if (rel.opts.showLegend) {
            strXml += '<c:legend>';
            strXml += '<c:legendPos val="' + rel.opts.legendPos + '"/>';
            //strXml += '<c:layout/>'
            strXml += '<c:overlay val="0"/>';
            if (rel.opts.legendFontFace || rel.opts.legendFontSize || rel.opts.legendColor) {
                strXml += '<c:txPr>';
                strXml += '  <a:bodyPr/>';
                strXml += '  <a:lstStyle/>';
                strXml += '  <a:p>';
                strXml += '    <a:pPr>';
                strXml += rel.opts.legendFontSize ? '<a:defRPr sz="' + Math.round(Number(rel.opts.legendFontSize) * 100) + '">' : '<a:defRPr>';
                if (rel.opts.legendColor)
                    strXml += genXmlColorSelection(rel.opts.legendColor);
                if (rel.opts.legendFontFace)
                    strXml += '<a:latin typeface="' + rel.opts.legendFontFace + '"/>';
                if (rel.opts.legendFontFace)
                    strXml += '<a:cs    typeface="' + rel.opts.legendFontFace + '"/>';
                strXml += '      </a:defRPr>';
                strXml += '    </a:pPr>';
                strXml += '    <a:endParaRPr lang="en-US"/>';
                strXml += '  </a:p>';
                strXml += '</c:txPr>';
            }
            strXml += '</c:legend>';
        }
    }
    strXml += '  <c:plotVisOnly val="1"/>';
    strXml += '  <c:dispBlanksAs val="' + rel.opts.displayBlanksAs + '"/>';
    if (rel.opts._type === CHART_TYPE.SCATTER)
        strXml += '<c:showDLblsOverMax val="1"/>';
    strXml += '</c:chart>';
    // D: CHARTSPACE SHAPE PROPS
    strXml += '<c:spPr>';
    strXml += '  <a:noFill/>';
    strXml += '  <a:ln w="12700" cap="flat"><a:noFill/><a:miter lim="400000"/></a:ln>';
    strXml += '  <a:effectLst/>';
    strXml += '</c:spPr>';
    // E: DATA (Add relID)
    strXml += '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>';
    // LAST: chartSpace end
    strXml += '</c:chartSpace>';
    return strXml;
}
/**
 * Create XML string for any given chart type
 * @param {CHART_NAME} `chartType` chart type name
 * @param {OptsChartData[]} `data` chart data
 * @param {IChartOptsLib} `opts` chart options
 * @param {string} `valAxisId`
 * @param {string} `catAxisId`
 * @param {boolean} `isMultiTypeChart`
 * @example '<c:bubbleChart>'
 * @example '<c:lineChart>'
 * @return {string} XML
 */
function makeChartType(chartType, data, opts, valAxisId, catAxisId, isMultiTypeChart) {
    // NOTE: "Chart Range" (as shown in "select Chart Area dialog") is calculated.
    // ....: Ensure each X/Y Axis/Col has same row height (esp. applicable to XY Scatter where X can often be larger than Y's)
    var strXml = '';
    switch (chartType) {
        case CHART_TYPE.AREA:
        case CHART_TYPE.BAR:
        case CHART_TYPE.BAR3D:
        case CHART_TYPE.LINE:
        case CHART_TYPE.RADAR:
            // 1: Start Chart
            strXml += '<c:' + chartType + 'Chart>';
            if (chartType === CHART_TYPE.AREA && opts.barGrouping === 'stacked') {
                strXml += '<c:grouping val="' + opts.barGrouping + '"/>';
            }
            if (chartType === CHART_TYPE.BAR || chartType === CHART_TYPE.BAR3D) {
                strXml += '<c:barDir val="' + opts.barDir + '"/>';
                strXml += '<c:grouping val="' + opts.barGrouping + '"/>';
            }
            if (chartType === CHART_TYPE.RADAR) {
                strXml += '<c:radarStyle val="' + opts.radarStyle + '"/>';
            }
            strXml += '<c:varyColors val="0"/>';
            // 2: "Series" block for every data row
            /* EX:
                data: [
                 {
                   name: 'Region 1',
                   labels: ['April', 'May', 'June', 'July'],
                   values: [17, 26, 53, 96]
                 },
                 {
                   name: 'Region 2',
                   labels: ['April', 'May', 'June', 'July'],
                   values: [55, 43, 70, 58]
                 }
                ]
            */
            var colorIndex_1 = -1; // Maintain the color index by region
            data.forEach(function (obj) {
                colorIndex_1++;
                var idx = obj.index;
                strXml += '<c:ser>';
                strXml += '  <c:idx val="' + idx + '"/>';
                strXml += '  <c:order val="' + idx + '"/>';
                strXml += '  <c:tx>';
                strXml += '    <c:strRef>';
                strXml += '      <c:f>Sheet1!$' + getExcelColName(idx + 1) + '$1</c:f>';
                strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt></c:strCache>';
                strXml += '    </c:strRef>';
                strXml += '  </c:tx>';
                strXml += '  <c:invertIfNegative val="0"/>';
                // Fill and Border
                // TODO: CURRENT: Pull#727
                // WIP: let seriesColor = obj.color ? obj.color : opts.chartColors ? opts.chartColors[colorIndex % opts.chartColors.length] : null
                var seriesColor = opts.chartColors ? opts.chartColors[colorIndex_1 % opts.chartColors.length] : null;
                strXml += '  <c:spPr>';
                if (seriesColor === 'transparent') {
                    strXml += '<a:noFill/>';
                }
                else if (opts.chartColorsOpacity) {
                    strXml += '<a:solidFill>' + createColorElement(seriesColor, "<a:alpha val=\"" + Math.round(opts.chartColorsOpacity * 1000) + "\"/>") + '</a:solidFill>';
                }
                else {
                    strXml += '<a:solidFill>' + createColorElement(seriesColor) + '</a:solidFill>';
                }
                if (chartType === CHART_TYPE.LINE) {
                    if (opts.lineSize === 0) {
                        strXml += '<a:ln><a:noFill/></a:ln>';
                    }
                    else {
                        strXml += '<a:ln w="' + valToPts(opts.lineSize) + '" cap="flat"><a:solidFill>' + createColorElement(seriesColor) + '</a:solidFill>';
                        strXml += '<a:prstDash val="' + (opts.lineDash || 'solid') + '"/><a:round/></a:ln>';
                    }
                }
                else if (opts.dataBorder) {
                    strXml +=
                        '<a:ln w="' +
                            valToPts(opts.dataBorder.pt) +
                            '" cap="flat"><a:solidFill>' +
                            createColorElement(opts.dataBorder.color) +
                            '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
                }
                strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                strXml += '  </c:spPr>';
                // Data Labels per series
                // [20190117] NOTE: Adding these to RADAR chart causes unrecoverable corruption!
                if (chartType !== CHART_TYPE.RADAR) {
                    strXml += '  <c:dLbls>';
                    strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
                    if (opts.dataLabelBkgrdColors) {
                        strXml += '    <c:spPr>';
                        strXml += '       <a:solidFill>' + createColorElement(seriesColor) + '</a:solidFill>';
                        strXml += '    </c:spPr>';
                    }
                    strXml += '    <c:txPr>';
                    strXml += '      <a:bodyPr/>';
                    strXml += '      <a:lstStyle/>';
                    strXml += '      <a:p><a:pPr>';
                    strXml += '        <a:defRPr b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="' + (opts.dataLabelFontItalic ? 1 : 0) + '" strike="noStrike" sz="' + Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100) + '" u="none">';
                    strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
                    strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
                    strXml += '        </a:defRPr>';
                    strXml += '      </a:pPr></a:p>';
                    strXml += '    </c:txPr>';
                    if (opts.dataLabelPosition)
                        strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>';
                    strXml += '    <c:showLegendKey val="0"/>';
                    strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
                    strXml += '    <c:showCatName val="0"/>';
                    strXml += '    <c:showSerName val="0"/>';
                    strXml += '    <c:showPercent val="0"/>';
                    strXml += '    <c:showBubbleSize val="0"/>';
                    strXml += "    <c:showLeaderLines val=\"" + (opts.showLeaderLines ? '1' : '0') + "\"/>";
                    strXml += '  </c:dLbls>';
                }
                // 'c:marker' tag: `lineDataSymbol`
                if (chartType === CHART_TYPE.LINE || chartType === CHART_TYPE.RADAR) {
                    strXml += '<c:marker>';
                    strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>';
                    if (opts.lineDataSymbolSize) {
                        // Defaults to "auto" otherwise (but this is usually too small, so there is a default)
                        strXml += '  <c:size val="' + opts.lineDataSymbolSize + '"/>';
                    }
                    strXml += '  <c:spPr>';
                    strXml +=
                        '    <a:solidFill>' +
                            createColorElement(opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx]) +
                            '</a:solidFill>';
                    strXml +=
                        '    <a:ln w="' +
                            opts.lineDataSymbolLineSize +
                            '" cap="flat"><a:solidFill>' +
                            createColorElement(opts.lineDataSymbolLineColor || seriesColor) +
                            '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
                    strXml += '    <a:effectLst/>';
                    strXml += '  </c:spPr>';
                    strXml += '</c:marker>';
                }
                // Allow users with a single data set to pass their own array of colors (check for this using != ours)
                // Color chart bars various colors when >1 color
                // NOTE: `<c:dPt>` created with various colors will change PPT legend by design so each dataPt/color is an legend item!
                if ((chartType === CHART_TYPE.BAR || chartType === CHART_TYPE.BAR3D) &&
                    data.length === 1 &&
                    opts.chartColors !== BARCHART_COLORS &&
                    opts.chartColors.length > 1) {
                    // Series Data Point colors
                    obj.values.forEach(function (value, index) {
                        var arrColors = value < 0 ? opts.invertedColors || opts.chartColors || BARCHART_COLORS : opts.chartColors || [];
                        strXml += '  <c:dPt>';
                        strXml += '    <c:idx val="' + index + '"/>';
                        strXml += '      <c:invertIfNegative val="0"/>';
                        strXml += '    <c:bubble3D val="0"/>';
                        strXml += '    <c:spPr>';
                        if (opts.lineSize === 0) {
                            strXml += '<a:ln><a:noFill/></a:ln>';
                        }
                        else if (chartType === CHART_TYPE.BAR) {
                            strXml += '<a:solidFill>';
                            strXml += '  <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
                            strXml += '</a:solidFill>';
                        }
                        else {
                            strXml += '<a:ln>';
                            strXml += '  <a:solidFill>';
                            strXml += '   <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
                            strXml += '  </a:solidFill>';
                            strXml += '</a:ln>';
                        }
                        strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                        strXml += '    </c:spPr>';
                        strXml += '  </c:dPt>';
                    });
                }
                // 2: "Categories"
                {
                    strXml += '<c:cat>';
                    if (opts.catLabelFormatCode) {
                        // Use 'numRef' as catLabelFormatCode implies that we are expecting numbers here
                        strXml += '  <c:numRef>';
                        strXml += '    <c:f>Sheet1!$A$2:$A$' + (obj.labels.length + 1) + '</c:f>';
                        strXml += '    <c:numCache>';
                        strXml += '      <c:formatCode>' + (opts.catLabelFormatCode || 'General') + '</c:formatCode>';
                        strXml += '      <c:ptCount val="' + obj.labels.length + '"/>';
                        obj.labels.forEach(function (label, idx) {
                            strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>';
                        });
                        strXml += '    </c:numCache>';
                        strXml += '  </c:numRef>';
                    }
                    else {
                        strXml += '  <c:strRef>';
                        strXml += '    <c:f>Sheet1!$A$2:$A$' + (obj.labels.length + 1) + '</c:f>';
                        strXml += '    <c:strCache>';
                        strXml += '	     <c:ptCount val="' + obj.labels.length + '"/>';
                        obj.labels.forEach(function (label, idx) {
                            strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>';
                        });
                        strXml += '    </c:strCache>';
                        strXml += '  </c:strRef>';
                    }
                    strXml += '</c:cat>';
                }
                // 3: "Values"
                {
                    strXml += '<c:val>';
                    strXml += '  <c:numRef>';
                    strXml += '    <c:f>Sheet1!$' + getExcelColName(idx + 1) + '$2:$' + getExcelColName(idx + 1) + '$' + (obj.labels.length + 1) + '</c:f>';
                    strXml += '    <c:numCache>';
                    strXml += '      <c:formatCode>' + (opts.valLabelFormatCode || opts.dataTableFormatCode || 'General') + '</c:formatCode>';
                    strXml += '      <c:ptCount val="' + obj.labels.length + '"/>';
                    obj.values.forEach(function (value, idx) {
                        strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value === 0 ? value : '') + '</c:v></c:pt>';
                    });
                    strXml += '    </c:numCache>';
                    strXml += '  </c:numRef>';
                    strXml += '</c:val>';
                }
                // Option: `smooth`
                if (chartType === CHART_TYPE.LINE)
                    strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>';
                // 4: Close "SERIES"
                strXml += '</c:ser>';
            });
            // 3: "Data Labels"
            {
                strXml += '  <c:dLbls>';
                strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
                strXml += '    <c:txPr>';
                strXml += '      <a:bodyPr/>';
                strXml += '      <a:lstStyle/>';
                strXml += '      <a:p><a:pPr>';
                strXml +=
                    '        <a:defRPr b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="' + (opts.dataLabelFontItalic ? 1 : 0) + '" strike="noStrike" sz="' + Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100) + '" u="none">';
                strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
                strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
                strXml += '        </a:defRPr>';
                strXml += '      </a:pPr></a:p>';
                strXml += '    </c:txPr>';
                if (opts.dataLabelPosition)
                    strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>';
                strXml += '    <c:showLegendKey val="0"/>';
                strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
                strXml += '    <c:showCatName val="0"/>';
                strXml += '    <c:showSerName val="0"/>';
                strXml += '    <c:showPercent val="0"/>';
                strXml += '    <c:showBubbleSize val="0"/>';
                strXml += "    <c:showLeaderLines val=\"" + (opts.showLeaderLines ? '1' : '0') + "\"/>";
                strXml += '  </c:dLbls>';
            }
            // 4: Add more chart options (gapWidth, line Marker, etc.)
            if (chartType === CHART_TYPE.BAR) {
                strXml += '  <c:gapWidth val="' + opts.barGapWidthPct + '"/>';
                strXml += '  <c:overlap val="' + ((opts.barGrouping || '').indexOf('tacked') > -1 ? 100 : 0) + '"/>';
            }
            else if (chartType === CHART_TYPE.BAR3D) {
                strXml += '  <c:gapWidth val="' + opts.barGapWidthPct + '"/>';
                strXml += '  <c:gapDepth val="' + opts.barGapDepthPct + '"/>';
                strXml += '  <c:shape val="' + opts.bar3DShape + '"/>';
            }
            else if (chartType === CHART_TYPE.LINE) {
                strXml += '  <c:marker val="1"/>';
            }
            // 5: Add axisId (NOTE: order matters! (category comes first))
            strXml += '  <c:axId val="' + catAxisId + '"/>';
            strXml += '  <c:axId val="' + valAxisId + '"/>';
            strXml += '  <c:axId val="' + AXIS_ID_SERIES_PRIMARY + '"/>';
            // 6: Close Chart tag
            strXml += '</c:' + chartType + 'Chart>';
            // end switch
            break;
        case CHART_TYPE.SCATTER:
            /*
                `data` = [
                    { name:'X-Axis',    values:[1,2,3,4,5,6,7,8,9,10,11,12] },
                    { name:'Y-Value 1', values:[13, 20, 21, 25] },
                    { name:'Y-Value 2', values:[ 1,  2,  5,  9] }
                ];
            */
            // 1: Start Chart
            strXml += '<c:' + chartType + 'Chart>';
            strXml += '<c:scatterStyle val="lineMarker"/>';
            strXml += '<c:varyColors val="0"/>';
            // 2: Series: (One for each Y-Axis)
            colorIndex_1 = -1;
            data.filter(function (_obj, idx) { return idx > 0; }).forEach(function (obj, idx) {
                colorIndex_1++;
                strXml += '<c:ser>';
                strXml += '  <c:idx val="' + idx + '"/>';
                strXml += '  <c:order val="' + idx + '"/>';
                strXml += '  <c:tx>';
                strXml += '    <c:strRef>';
                strXml += '      <c:f>Sheet1!$' + LETTERS[idx + 1] + '$1</c:f>';
                strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + obj.name + '</c:v></c:pt></c:strCache>';
                strXml += '    </c:strRef>';
                strXml += '  </c:tx>';
                // 'c:spPr': Fill, Border, Line, LineStyle (dash, etc.), Shadow
                strXml += '  <c:spPr>';
                {
                    var tmpSerColor = opts.chartColors[colorIndex_1 % opts.chartColors.length];
                    if (tmpSerColor === 'transparent') {
                        strXml += '<a:noFill/>';
                    }
                    else if (opts.chartColorsOpacity) {
                        strXml += '<a:solidFill>' + createColorElement(tmpSerColor, '<a:alpha val="' + Math.round(opts.chartColorsOpacity * 1000) + '"/>') + '</a:solidFill>';
                    }
                    else {
                        strXml += '<a:solidFill>' + createColorElement(tmpSerColor) + '</a:solidFill>';
                    }
                    if (opts.lineSize === 0) {
                        strXml += '<a:ln><a:noFill/></a:ln>';
                    }
                    else {
                        strXml += '<a:ln w="' + valToPts(opts.lineSize) + '" cap="flat"><a:solidFill>' + createColorElement(tmpSerColor) + '</a:solidFill>';
                        strXml += '<a:prstDash val="' + (opts.lineDash || 'solid') + '"/><a:round/></a:ln>';
                    }
                    // Shadow
                    strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                }
                strXml += '  </c:spPr>';
                // 'c:marker' tag: `lineDataSymbol`
                {
                    strXml += '<c:marker>';
                    strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>';
                    if (opts.lineDataSymbolSize) {
                        // Defaults to "auto" otherwise (but this is usually too small, so there is a default)
                        strXml += '  <c:size val="' + opts.lineDataSymbolSize + '"/>';
                    }
                    strXml += '  <c:spPr>';
                    strXml +=
                        '    <a:solidFill>' +
                            createColorElement(opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx]) +
                            '</a:solidFill>';
                    strXml +=
                        '    <a:ln w="' +
                            opts.lineDataSymbolLineSize +
                            '" cap="flat"><a:solidFill>' +
                            createColorElement(opts.lineDataSymbolLineColor || opts.chartColors[colorIndex_1 % opts.chartColors.length]) +
                            '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
                    strXml += '    <a:effectLst/>';
                    strXml += '  </c:spPr>';
                    strXml += '</c:marker>';
                }
                // Option: scatter data point labels
                if (opts.showLabel) {
                    var chartUuid_1 = getUuid('-xxxx-xxxx-xxxx-xxxxxxxxxxxx');
                    if (obj.labels && (opts.dataLabelFormatScatter === 'custom' || opts.dataLabelFormatScatter === 'customXY')) {
                        strXml += '<c:dLbls>';
                        obj.labels.forEach(function (label, idx) {
                            if (opts.dataLabelFormatScatter === 'custom' || opts.dataLabelFormatScatter === 'customXY') {
                                strXml += '  <c:dLbl>';
                                strXml += '    <c:idx val="' + idx + '"/>';
                                strXml += '    <c:tx>';
                                strXml += '      <c:rich>';
                                strXml += '			<a:bodyPr>';
                                strXml += '				<a:spAutoFit/>';
                                strXml += '			</a:bodyPr>';
                                strXml += '        	<a:lstStyle/>';
                                strXml += '        	<a:p>';
                                strXml += '				<a:pPr>';
                                strXml += '					<a:defRPr/>';
                                strXml += '				</a:pPr>';
                                strXml += '          	<a:r>';
                                strXml += '            		<a:rPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>';
                                strXml += '            		<a:t>' + encodeXmlEntities(label) + '</a:t>';
                                strXml += '          	</a:r>';
                                // Apply XY values at end of custom label
                                // Do not apply the values if the label was empty or just spaces
                                // This allows for selective labelling where required
                                if (opts.dataLabelFormatScatter === 'customXY' && !/^ *$/.test(label)) {
                                    strXml += '          	<a:r>';
                                    strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>';
                                    strXml += '          		<a:t> (</a:t>';
                                    strXml += '          	</a:r>';
                                    strXml += '          	<a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="XVALUE">';
                                    strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>';
                                    strXml += '          		<a:pPr>';
                                    strXml += '          			<a:defRPr/>';
                                    strXml += '          		</a:pPr>';
                                    strXml += '          		<a:t>[' + encodeXmlEntities(obj.name) + '</a:t>';
                                    strXml += '          	</a:fld>';
                                    strXml += '          	<a:r>';
                                    strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>';
                                    strXml += '          		<a:t>, </a:t>';
                                    strXml += '          	</a:r>';
                                    strXml += '          	<a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="YVALUE">';
                                    strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>';
                                    strXml += '          		<a:pPr>';
                                    strXml += '          			<a:defRPr/>';
                                    strXml += '          		</a:pPr>';
                                    strXml += '          		<a:t>[' + encodeXmlEntities(obj.name) + ']</a:t>';
                                    strXml += '          	</a:fld>';
                                    strXml += '          	<a:r>';
                                    strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>';
                                    strXml += '          		<a:t>)</a:t>';
                                    strXml += '          	</a:r>';
                                    strXml += '          	<a:endParaRPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>';
                                }
                                strXml += '        	</a:p>';
                                strXml += '      </c:rich>';
                                strXml += '    </c:tx>';
                                strXml += '    <c:spPr>';
                                strXml += '    	<a:noFill/>';
                                strXml += '    	<a:ln>';
                                strXml += '    		<a:noFill/>';
                                strXml += '    	</a:ln>';
                                strXml += '    	<a:effectLst/>';
                                strXml += '    </c:spPr>';
                                if (opts.dataLabelPosition)
                                    strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>';
                                strXml += '    <c:showLegendKey val="0"/>';
                                strXml += '    <c:showVal val="0"/>';
                                strXml += '    <c:showCatName val="0"/>';
                                strXml += '    <c:showSerName val="0"/>';
                                strXml += '    <c:showPercent val="0"/>';
                                strXml += '    <c:showBubbleSize val="0"/>';
                                strXml += '	   <c:showLeaderLines val="1"/>';
                                strXml += '    <c:extLst>';
                                strXml += '      <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"/>';
                                strXml += '      <c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">';
                                strXml += '			<c16:uniqueId val="{' + '00000000'.substring(0, 8 - (idx + 1).toString().length).toString() + (idx + 1) + chartUuid_1 + '}"/>';
                                strXml += '      </c:ext>';
                                strXml += '		</c:extLst>';
                                strXml += '</c:dLbl>';
                            }
                        });
                        strXml += '</c:dLbls>';
                    }
                    if (opts.dataLabelFormatScatter === 'XY') {
                        strXml += '<c:dLbls>';
                        strXml += '	<c:spPr>';
                        strXml += '		<a:noFill/>';
                        strXml += '		<a:ln>';
                        strXml += '			<a:noFill/>';
                        strXml += '		</a:ln>';
                        strXml += '	  	<a:effectLst/>';
                        strXml += '	</c:spPr>';
                        strXml += '	<c:txPr>';
                        strXml += '		<a:bodyPr>';
                        strXml += '			<a:spAutoFit/>';
                        strXml += '		</a:bodyPr>';
                        strXml += '		<a:lstStyle/>';
                        strXml += '		<a:p>';
                        strXml += '	    	<a:pPr>';
                        strXml += '        		<a:defRPr/>';
                        strXml += '	    	</a:pPr>';
                        strXml += '	    	<a:endParaRPr lang="en-US"/>';
                        strXml += '		</a:p>';
                        strXml += '	</c:txPr>';
                        if (opts.dataLabelPosition)
                            strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>';
                        strXml += '	<c:showLegendKey val="0"/>';
                        strXml += " <c:showVal val=\"" + (opts.showLabel ? '1' : '0') + "\"/>";
                        strXml += " <c:showCatName val=\"" + (opts.showLabel ? '1' : '0') + "\"/>";
                        strXml += '	<c:showSerName val="0"/>';
                        strXml += '	<c:showPercent val="0"/>';
                        strXml += '	<c:showBubbleSize val="0"/>';
                        strXml += '	<c:extLst>';
                        strXml += '		<c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">';
                        strXml += '			<c15:showLeaderLines val="1"/>';
                        strXml += '		</c:ext>';
                        strXml += '	</c:extLst>';
                        strXml += '</c:dLbls>';
                    }
                }
                // Color bar chart bars various colors
                // Allow users with a single data set to pass their own array of colors (check for this using != ours)
                if (data.length === 1 && opts.chartColors !== BARCHART_COLORS) {
                    // Series Data Point colors
                    obj.values.forEach(function (value, index) {
                        var arrColors = value < 0 ? opts.invertedColors || opts.chartColors || BARCHART_COLORS : opts.chartColors || [];
                        strXml += '  <c:dPt>';
                        strXml += '    <c:idx val="' + index + '"/>';
                        strXml += '      <c:invertIfNegative val="0"/>';
                        strXml += '    <c:bubble3D val="0"/>';
                        strXml += '    <c:spPr>';
                        if (opts.lineSize === 0) {
                            strXml += '<a:ln><a:noFill/></a:ln>';
                        }
                        else {
                            strXml += '<a:solidFill>';
                            strXml += ' <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
                            strXml += '</a:solidFill>';
                        }
                        strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                        strXml += '    </c:spPr>';
                        strXml += '  </c:dPt>';
                    });
                }
                // 3: "Values": Scatter Chart has 2: `xVal` and `yVal`
                {
                    // X-Axis is always the same
                    strXml += '<c:xVal>';
                    strXml += '  <c:numRef>';
                    strXml += '    <c:f>Sheet1!$A$2:$A$' + (data[0].values.length + 1) + '</c:f>';
                    strXml += '    <c:numCache>';
                    strXml += '      <c:formatCode>General</c:formatCode>';
                    strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
                    data[0].values.forEach(function (value, idx) {
                        strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value === 0 ? value : '') + '</c:v></c:pt>';
                    });
                    strXml += '    </c:numCache>';
                    strXml += '  </c:numRef>';
                    strXml += '</c:xVal>';
                    // Y-Axis vals are this object's `values`
                    strXml += '<c:yVal>';
                    strXml += '  <c:numRef>';
                    strXml += '    <c:f>Sheet1!$' + getExcelColName(idx + 1) + '$2:$' + getExcelColName(idx + 1) + '$' + (data[0].values.length + 1) + '</c:f>';
                    strXml += '    <c:numCache>';
                    strXml += '      <c:formatCode>General</c:formatCode>';
                    // NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
                    strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
                    data[0].values.forEach(function (_value, idx) {
                        strXml += '<c:pt idx="' + idx + '"><c:v>' + (obj.values[idx] || obj.values[idx] === 0 ? obj.values[idx] : '') + '</c:v></c:pt>';
                    });
                    strXml += '    </c:numCache>';
                    strXml += '  </c:numRef>';
                    strXml += '</c:yVal>';
                }
                // Option: `smooth`
                strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>';
                // 4: Close "SERIES"
                strXml += '</c:ser>';
            });
            // 3: Data Labels
            {
                strXml += '  <c:dLbls>';
                strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
                strXml += '    <c:txPr>';
                strXml += '      <a:bodyPr/>';
                strXml += '      <a:lstStyle/>';
                strXml += '      <a:p><a:pPr>';
                strXml += '        <a:defRPr b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="' + (opts.dataLabelFontItalic ? 1 : 0) + '" strike="noStrike" sz="' + Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100) + '" u="none">';
                strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
                strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
                strXml += '        </a:defRPr>';
                strXml += '      </a:pPr></a:p>';
                strXml += '    </c:txPr>';
                if (opts.dataLabelPosition)
                    strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>';
                strXml += '    <c:showLegendKey val="0"/>';
                strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
                strXml += '    <c:showCatName val="0"/>';
                strXml += '    <c:showSerName val="0"/>';
                strXml += '    <c:showPercent val="0"/>';
                strXml += '    <c:showBubbleSize val="0"/>';
                strXml += '  </c:dLbls>';
            }
            // 4: Add axisId (NOTE: order matters! (category comes first))
            strXml += '  <c:axId val="' + catAxisId + '"/>';
            strXml += '  <c:axId val="' + valAxisId + '"/>';
            // 5: Close Chart tag
            strXml += '</c:' + chartType + 'Chart>';
            // end switch
            break;
        case CHART_TYPE.BUBBLE:
            /*
                `data` = [
                    { name:'X-Axis',     values:[1,2,3,4,5,6,7,8,9,10,11,12] },
                    { name:'Y-Values 1', values:[13, 20, 21, 25], sizes:[10, 5, 20, 15] },
                    { name:'Y-Values 2', values:[ 1,  2,  5,  9], sizes:[ 5, 3,  9,  3] }
                ];
            */
            // 1: Start Chart
            strXml += '<c:' + chartType + 'Chart>';
            strXml += '<c:varyColors val="0"/>';
            // 2: Series: (One for each Y-Axis)
            colorIndex_1 = -1;
            var idxColLtr_1 = 1;
            data.filter(function (_obj, idx) { return idx > 0; }).forEach(function (obj, idx) {
                colorIndex_1++;
                strXml += '<c:ser>';
                strXml += '  <c:idx val="' + idx + '"/>';
                strXml += '  <c:order val="' + idx + '"/>';
                // A: `<c:tx>`
                strXml += '  <c:tx>';
                strXml += '    <c:strRef>';
                strXml += '      <c:f>Sheet1!$' + LETTERS[idxColLtr_1] + '$1</c:f>';
                strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + obj.name + '</c:v></c:pt></c:strCache>';
                strXml += '    </c:strRef>';
                strXml += '  </c:tx>';
                // B: '<c:spPr>': Fill, Border, Line, LineStyle (dash, etc.), Shadow
                {
                    strXml += '<c:spPr>';
                    var tmpSerColor = opts.chartColors[colorIndex_1 % opts.chartColors.length];
                    if (tmpSerColor === 'transparent') {
                        strXml += '<a:noFill/>';
                    }
                    else if (opts.chartColorsOpacity) {
                        strXml += '<a:solidFill>' + createColorElement(tmpSerColor, '<a:alpha val="' + Math.round(opts.chartColorsOpacity * 1000) + '"/>') + '</a:solidFill>';
                    }
                    else {
                        strXml += '<a:solidFill>' + createColorElement(tmpSerColor) + '</a:solidFill>';
                    }
                    if (opts.lineSize === 0) {
                        strXml += '<a:ln><a:noFill/></a:ln>';
                    }
                    else if (opts.dataBorder) {
                        strXml +=
                            '<a:ln w="' +
                                valToPts(opts.dataBorder.pt) +
                                '" cap="flat"><a:solidFill>' +
                                createColorElement(opts.dataBorder.color) +
                                '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
                    }
                    else {
                        strXml += '<a:ln w="' + valToPts(opts.lineSize) + '" cap="flat"><a:solidFill>' + createColorElement(tmpSerColor) + '</a:solidFill>';
                        strXml += '<a:prstDash val="' + (opts.lineDash || 'solid') + '"/><a:round/></a:ln>';
                    }
                    // Shadow
                    strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                    strXml += '</c:spPr>';
                }
                // C: '<c:dLbls>' "Data Labels"
                // Let it be defaulted for now
                // D: '<c:xVal>'/'<c:yVal>' "Values": Scatter Chart has 2: `xVal` and `yVal`
                {
                    // X-Axis is always the same
                    strXml += '<c:xVal>';
                    strXml += '  <c:numRef>';
                    strXml += '    <c:f>Sheet1!$A$2:$A$' + (data[0].values.length + 1) + '</c:f>';
                    strXml += '    <c:numCache>';
                    strXml += '      <c:formatCode>General</c:formatCode>';
                    strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
                    data[0].values.forEach(function (value, idx) {
                        strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value === 0 ? value : '') + '</c:v></c:pt>';
                    });
                    strXml += '    </c:numCache>';
                    strXml += '  </c:numRef>';
                    strXml += '</c:xVal>';
                    // Y-Axis vals are this object's `values`
                    strXml += '<c:yVal>';
                    strXml += '  <c:numRef>';
                    strXml += '    <c:f>Sheet1!$' + getExcelColName(idxColLtr_1) + '$2:$' + getExcelColName(idxColLtr_1) + '$' + (data[0].values.length + 1) + '</c:f>';
                    idxColLtr_1++;
                    strXml += '    <c:numCache>';
                    strXml += '      <c:formatCode>General</c:formatCode>';
                    // NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
                    strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
                    data[0].values.forEach(function (_value, idx) {
                        strXml += '<c:pt idx="' + idx + '"><c:v>' + (obj.values[idx] || obj.values[idx] === 0 ? obj.values[idx] : '') + '</c:v></c:pt>';
                    });
                    strXml += '    </c:numCache>';
                    strXml += '  </c:numRef>';
                    strXml += '</c:yVal>';
                }
                // E: '<c:bubbleSize>'
                strXml += '  <c:bubbleSize>';
                strXml += '    <c:numRef>';
                strXml += '      <c:f>Sheet1!$' + getExcelColName(idxColLtr_1) + '$2:$' + getExcelColName(idx + 2) + '$' + (obj.sizes.length + 1) + '</c:f>';
                idxColLtr_1++;
                strXml += '      <c:numCache>';
                strXml += '        <c:formatCode>General</c:formatCode>';
                strXml += '	       <c:ptCount val="' + obj.sizes.length + '"/>';
                obj.sizes.forEach(function (value, idx) {
                    strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || '') + '</c:v></c:pt>';
                });
                strXml += '      </c:numCache>';
                strXml += '    </c:numRef>';
                strXml += '  </c:bubbleSize>';
                strXml += '  <c:bubble3D val="0"/>';
                // F: Close "SERIES"
                strXml += '</c:ser>';
            });
            // 3: Data Labels
            {
                strXml += '  <c:dLbls>';
                strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
                strXml += '    <c:txPr>';
                strXml += '      <a:bodyPr/>';
                strXml += '      <a:lstStyle/>';
                strXml += '      <a:p><a:pPr>';
                strXml += '        <a:defRPr b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="' + (opts.dataLabelFontItalic ? 1 : 0) + '" strike="noStrike" sz="' + Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100) + '" u="none">';
                strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
                strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
                strXml += '        </a:defRPr>';
                strXml += '      </a:pPr></a:p>';
                strXml += '    </c:txPr>';
                if (opts.dataLabelPosition)
                    strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>';
                strXml += '    <c:showLegendKey val="0"/>';
                strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
                strXml += '    <c:showCatName val="0"/>';
                strXml += '    <c:showSerName val="0"/>';
                strXml += '    <c:showPercent val="0"/>';
                strXml += '    <c:showBubbleSize val="0"/>';
                strXml += '  </c:dLbls>';
            }
            // 4: Add bubble options
            //strXml += '  <c:bubbleScale val="100"/>';
            //strXml += '  <c:showNegBubbles val="0"/>';
            // Commented out to let it default to PPT until we create options
            // 5: Add axisId (NOTE: order matters! (category comes first))
            strXml += '  <c:axId val="' + catAxisId + '"/>';
            strXml += '  <c:axId val="' + valAxisId + '"/>';
            // 6: Close Chart tag
            strXml += '</c:' + chartType + 'Chart>';
            // end switch
            break;
        case CHART_TYPE.DOUGHNUT:
        case CHART_TYPE.PIE:
            // Use the same let name so code blocks from barChart are interchangeable
            var obj = data[0];
            /* EX:
                data: [
                 {
                   name: 'Project Status',
                   labels: ['Red', 'Amber', 'Green', 'Unknown'],
                   values: [10, 20, 38, 2]
                 }
                ]
            */
            // 1: Start Chart
            strXml += '<c:' + chartType + 'Chart>';
            strXml += '  <c:varyColors val="0"/>';
            strXml += '<c:ser>';
            strXml += '  <c:idx val="0"/>';
            strXml += '  <c:order val="0"/>';
            strXml += '  <c:tx>';
            strXml += '    <c:strRef>';
            strXml += '      <c:f>Sheet1!$B$1</c:f>';
            strXml += '      <c:strCache>';
            strXml += '        <c:ptCount val="1"/>';
            strXml += '        <c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt>';
            strXml += '      </c:strCache>';
            strXml += '    </c:strRef>';
            strXml += '  </c:tx>';
            strXml += '  <c:spPr>';
            strXml += '    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>';
            strXml += '    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
            if (opts.dataNoEffects) {
                strXml += '<a:effectLst/>';
            }
            else {
                strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
            }
            strXml += '  </c:spPr>';
            //strXml += '<c:explosion val="0"/>'
            // 2: "Data Point" block for every data row
            obj.labels.forEach(function (_label, idx) {
                strXml += '<c:dPt>';
                strXml += " <c:idx val=\"" + idx + "\"/>";
                strXml += ' <c:bubble3D val="0"/>';
                strXml += ' <c:spPr>';
                strXml += "<a:solidFill>" + createColorElement(opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx]) + "</a:solidFill>";
                if (opts.dataBorder) {
                    strXml += "<a:ln w=\"" + valToPts(opts.dataBorder.pt) + "\" cap=\"flat\"><a:solidFill>" + createColorElement(opts.dataBorder.color) + "</a:solidFill><a:prstDash val=\"solid\"/><a:round/></a:ln>";
                }
                strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                strXml += '  </c:spPr>';
                strXml += '</c:dPt>';
            });
            // 3: "Data Label" block for every data Label
            strXml += '<c:dLbls>';
            obj.labels.forEach(function (_label, idx) {
                strXml += '<c:dLbl>';
                strXml += " <c:idx val=\"" + idx + "\"/>";
                strXml += "  <c:numFmt formatCode=\"" + (opts.dataLabelFormatCode || 'General') + "\" sourceLinked=\"0\"/>";
                strXml += '  <c:spPr/><c:txPr>';
                strXml += '   <a:bodyPr/><a:lstStyle/>';
                strXml += '   <a:p><a:pPr>';
                strXml += "   <a:defRPr sz=\"" + Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100) + "\" b=\"" + (opts.dataLabelFontBold ? 1 : 0) + "\" i=\"" + (opts.dataLabelFontItalic ? 1 : 0) + "\" u=\"none\" strike=\"noStrike\">";
                strXml += '    <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
                strXml += "    <a:latin typeface=\"" + (opts.dataLabelFontFace || 'Arial') + "\"/>";
                strXml += '   </a:defRPr>';
                strXml += '      </a:pPr></a:p>';
                strXml += '    </c:txPr>';
                if (chartType === CHART_TYPE.PIE && opts.dataLabelPosition)
                    strXml += "    <c:dLblPos val=\"" + opts.dataLabelPosition + "\"/>";
                strXml += '    <c:showLegendKey val="0"/>';
                strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
                strXml += '    <c:showCatName val="' + (opts.showLabel ? '1' : '0') + '"/>';
                strXml += '    <c:showSerName val="0"/>';
                strXml += '    <c:showPercent val="' + (opts.showPercent ? '1' : '0') + '"/>';
                strXml += '    <c:showBubbleSize val="0"/>';
                strXml += '  </c:dLbl>';
            });
            strXml += " <c:numFmt formatCode=\"" + (opts.dataLabelFormatCode || 'General') + "\" sourceLinked=\"0\"/>";
            strXml += '	<c:txPr>';
            strXml += '	  <a:bodyPr/>';
            strXml += '	  <a:lstStyle/>';
            strXml += '	  <a:p>';
            strXml += '		<a:pPr>';
            strXml += '		  <a:defRPr sz="1800" b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="' + (opts.dataLabelFontItalic ? 1 : 0) + '" u="none" strike="noStrike">';
            strXml += '			<a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/>';
            strXml += '		  </a:defRPr>';
            strXml += '		</a:pPr>';
            strXml += '	  </a:p>';
            strXml += '	</c:txPr>';
            strXml += chartType === CHART_TYPE.PIE ? '<c:dLblPos val="ctr"/>' : '';
            strXml += '	<c:showLegendKey val="0"/>';
            strXml += '	<c:showVal val="0"/>';
            strXml += '	<c:showCatName val="1"/>';
            strXml += '	<c:showSerName val="0"/>';
            strXml += '	<c:showPercent val="1"/>';
            strXml += '	<c:showBubbleSize val="0"/>';
            strXml += " <c:showLeaderLines val=\"" + (opts.showLeaderLines ? '1' : '0') + "\"/>";
            strXml += '</c:dLbls>';
            // 2: "Categories"
            strXml += '<c:cat>';
            strXml += '  <c:strRef>';
            strXml += '    <c:f>Sheet1!$A$2:$A$' + (obj.labels.length + 1) + '</c:f>';
            strXml += '    <c:strCache>';
            strXml += '	     <c:ptCount val="' + obj.labels.length + '"/>';
            obj.labels.forEach(function (label, idx) {
                strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>';
            });
            strXml += '    </c:strCache>';
            strXml += '  </c:strRef>';
            strXml += '</c:cat>';
            // 3: Create vals
            strXml += '  <c:val>';
            strXml += '    <c:numRef>';
            strXml += '      <c:f>Sheet1!$B$2:$B$' + (obj.labels.length + 1) + '</c:f>';
            strXml += '      <c:numCache>';
            strXml += '	       <c:ptCount val="' + obj.labels.length + '"/>';
            obj.values.forEach(function (value, idx) {
                strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value === 0 ? value : '') + '</c:v></c:pt>';
            });
            strXml += '      </c:numCache>';
            strXml += '    </c:numRef>';
            strXml += '  </c:val>';
            // 4: Close "SERIES"
            strXml += '  </c:ser>';
            strXml += "  <c:firstSliceAng val=\"" + (opts.firstSliceAng ? Math.round(opts.firstSliceAng) : 0) + "\"/>";
            if (chartType === CHART_TYPE.DOUGHNUT)
                strXml += '  <c:holeSize val="' + (opts.holeSize || 50) + '"/>';
            strXml += '</c:' + chartType + 'Chart>';
            // Done with Doughnut/Pie
            break;
        default:
            strXml += '';
            break;
    }
    return strXml;
}
/**
 * Create Category axis
 * @param {IChartOptsLib} opts - chart options
 * @param {string} axisId - value
 * @param {string} valAxisId - value
 * @return {string} XML
 */
function makeCatAxis(opts, axisId, valAxisId) {
    var strXml = '';
    // Build cat axis tag
    // NOTE: Scatter and Bubble chart need two Val axises as they display numbers on x axis
    if (opts._type === CHART_TYPE.SCATTER || opts._type === CHART_TYPE.BUBBLE) {
        strXml += '<c:valAx>';
    }
    else {
        strXml += '<c:' + (opts.catLabelFormatCode ? 'dateAx' : 'catAx') + '>';
    }
    strXml += '  <c:axId val="' + axisId + '"/>';
    strXml += '  <c:scaling>';
    strXml += '<c:orientation val="' + (opts.catAxisOrientation || (opts.barDir === 'col' ? 'minMax' : 'minMax')) + '"/>';
    if (opts.catAxisMaxVal || opts.catAxisMaxVal === 0)
        strXml += '<c:max val="' + opts.catAxisMaxVal + '"/>';
    if (opts.catAxisMinVal || opts.catAxisMinVal === 0)
        strXml += '<c:min val="' + opts.catAxisMinVal + '"/>';
    strXml += '</c:scaling>';
    strXml += '  <c:delete val="' + (opts.catAxisHidden ? 1 : 0) + '"/>';
    strXml += '  <c:axPos val="' + (opts.barDir === 'col' ? 'b' : 'l') + '"/>';
    strXml += opts.catGridLine.style !== 'none' ? createGridLineElement(opts.catGridLine) : '';
    // '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
    if (opts.showCatAxisTitle) {
        strXml += genXmlTitle({
            color: opts.catAxisTitleColor,
            fontFace: opts.catAxisTitleFontFace,
            fontSize: opts.catAxisTitleFontSize,
            titleRotate: opts.catAxisTitleRotate,
            title: opts.catAxisTitle || 'Axis Title',
        });
    }
    // NOTE: Adding Val Axis Formatting if scatter or bubble charts
    if (opts._type === CHART_TYPE.SCATTER || opts._type === CHART_TYPE.BUBBLE) {
        strXml += '  <c:numFmt formatCode="' + (opts.valAxisLabelFormatCode ? opts.valAxisLabelFormatCode : 'General') + '" sourceLinked="0"/>';
    }
    else {
        strXml += '  <c:numFmt formatCode="' + (opts.catLabelFormatCode || 'General') + '" sourceLinked="0"/>';
    }
    if (opts._type === CHART_TYPE.SCATTER) {
        strXml += '  <c:majorTickMark val="none"/>';
        strXml += '  <c:minorTickMark val="none"/>';
        strXml += '  <c:tickLblPos val="nextTo"/>';
    }
    else {
        strXml += '  <c:majorTickMark val="' + (opts.catAxisMajorTickMark || 'out') + '"/>';
        strXml += '  <c:minorTickMark val="' + (opts.catAxisMinorTickMark || 'none') + '"/>';
        strXml += '  <c:tickLblPos val="' + (opts.catAxisLabelPos || (opts.barDir === 'col' ? 'low' : 'nextTo')) + '"/>';
    }
    strXml += '  <c:spPr>';
    strXml += '    <a:ln w="' + (opts.catAxisLineSize ? valToPts(opts.catAxisLineSize) : ONEPT) + '" cap="flat">';
    strXml += opts.catAxisLineShow === false ? '<a:noFill/>' : '<a:solidFill>' + createColorElement(opts.catAxisLineColor || DEF_CHART_GRIDLINE.color) + '</a:solidFill>';
    strXml += '      <a:prstDash val="' + (opts.catAxisLineStyle || 'solid') + '"/>';
    strXml += '      <a:round/>';
    strXml += '    </a:ln>';
    strXml += '  </c:spPr>';
    strXml += '  <c:txPr>';
    strXml += '    <a:bodyPr' + (opts.catAxisLabelRotate ? ' rot="' + convertRotationDegrees(opts.catAxisLabelRotate) + '"' : '') + '/>'; // don't specify rot 0 so we get the auto behavior
    strXml += '    <a:lstStyle/>';
    strXml += '    <a:p>';
    strXml += '    <a:pPr>';
    strXml +=
        '    <a:defRPr sz="' +
            Math.round((opts.catAxisLabelFontSize || DEF_FONT_SIZE) * 100) +
            '" b="' + (opts.catAxisLabelFontBold ? 1 : 0) + '" i="' + (opts.catAxisLabelFontItalic ? 1 : 0) + '" u="none" strike="noStrike">';
    strXml += '      <a:solidFill>' + createColorElement(opts.catAxisLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
    strXml += '      <a:latin typeface="' + (opts.catAxisLabelFontFace || 'Arial') + '"/>';
    strXml += '   </a:defRPr>';
    strXml += '  </a:pPr>';
    strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>';
    strXml += '  </a:p>';
    strXml += ' </c:txPr>';
    strXml += ' <c:crossAx val="' + valAxisId + '"/>';
    strXml += ' <c:' + (typeof opts.valAxisCrossesAt === 'number' ? 'crossesAt' : 'crosses') + ' val="' + opts.valAxisCrossesAt + '"/>';
    strXml += ' <c:auto val="1"/>';
    strXml += ' <c:lblAlgn val="ctr"/>';
    strXml += ' <c:noMultiLvlLbl val="1"/>';
    if (opts.catAxisLabelFrequency)
        strXml += ' <c:tickLblSkip val="' + opts.catAxisLabelFrequency + '"/>';
    // Issue#149: PPT will auto-adjust these as needed after calcing the date bounds, so we only include them when specified by user
    // Allow major and minor units to be set for double value axis charts
    if (opts.catLabelFormatCode || opts._type === CHART_TYPE.SCATTER || opts._type === CHART_TYPE.BUBBLE) {
        if (opts.catLabelFormatCode) {
            ['catAxisBaseTimeUnit', 'catAxisMajorTimeUnit', 'catAxisMinorTimeUnit'].forEach(function (opt) {
                // Validate input as poorly chosen/garbage options will cause chart corruption and it wont render at all!
                if (opts[opt] && (typeof opts[opt] !== 'string' || ['days', 'months', 'years'].indexOf(opts[opt].toLowerCase()) === -1)) {
                    console.warn('`' + opt + "` must be one of: 'days','months','years' !");
                    opts[opt] = null;
                }
            });
            if (opts.catAxisBaseTimeUnit)
                strXml += '<c:baseTimeUnit val="' + opts.catAxisBaseTimeUnit.toLowerCase() + '"/>';
            if (opts.catAxisMajorTimeUnit)
                strXml += '<c:majorTimeUnit val="' + opts.catAxisMajorTimeUnit.toLowerCase() + '"/>';
            if (opts.catAxisMinorTimeUnit)
                strXml += '<c:minorTimeUnit val="' + opts.catAxisMinorTimeUnit.toLowerCase() + '"/>';
        }
        if (opts.catAxisMajorUnit)
            strXml += '<c:majorUnit val="' + opts.catAxisMajorUnit + '"/>';
        if (opts.catAxisMinorUnit)
            strXml += '<c:minorUnit val="' + opts.catAxisMinorUnit + '"/>';
    }
    // Close cat axis tag
    // NOTE: Added closing tag of val or cat axis based on chart type
    if (opts._type === CHART_TYPE.SCATTER || opts._type === CHART_TYPE.BUBBLE) {
        strXml += '</c:valAx>';
    }
    else {
        strXml += '</c:' + (opts.catLabelFormatCode ? 'dateAx' : 'catAx') + '>';
    }
    return strXml;
}
/**
 * Create Value Axis (Used by `bar3D`)
 * @param {IChartOptsLib} opts - chart options
 * @param {string} valAxisId - value
 * @return {string} XML
 */
function makeValAxis(opts, valAxisId) {
    var axisPos = valAxisId === AXIS_ID_VALUE_PRIMARY ? (opts.barDir === 'col' ? 'l' : 'b') : opts.barDir !== 'col' ? 'r' : 't';
    var strXml = '';
    var isRight = axisPos === 'r' || axisPos === 't';
    var crosses = isRight ? 'max' : 'autoZero';
    var crossAxId = valAxisId === AXIS_ID_VALUE_PRIMARY ? AXIS_ID_CATEGORY_PRIMARY : AXIS_ID_CATEGORY_SECONDARY;
    strXml += '<c:valAx>';
    strXml += '  <c:axId val="' + valAxisId + '"/>';
    strXml += '  <c:scaling>';
    if (opts.valAxisLogScaleBase)
        strXml += "    <c:logBase val=\"" + opts.valAxisLogScaleBase + "\"/>";
    strXml += '    <c:orientation val="' + (opts.valAxisOrientation || (opts.barDir === 'col' ? 'minMax' : 'minMax')) + '"/>';
    if (opts.valAxisMaxVal || opts.valAxisMaxVal === 0)
        strXml += '<c:max val="' + opts.valAxisMaxVal + '"/>';
    if (opts.valAxisMinVal || opts.valAxisMinVal === 0)
        strXml += '<c:min val="' + opts.valAxisMinVal + '"/>';
    strXml += '  </c:scaling>';
    strXml += '  <c:delete val="' + (opts.valAxisHidden ? 1 : 0) + '"/>';
    strXml += '  <c:axPos val="' + axisPos + '"/>';
    if (opts.valGridLine.style !== 'none')
        strXml += createGridLineElement(opts.valGridLine);
    // '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
    if (opts.showValAxisTitle) {
        strXml += genXmlTitle({
            color: opts.valAxisTitleColor,
            fontFace: opts.valAxisTitleFontFace,
            fontSize: opts.valAxisTitleFontSize,
            titleRotate: opts.valAxisTitleRotate,
            title: opts.valAxisTitle || 'Axis Title',
        });
    }
    strXml += "<c:numFmt formatCode='" + (opts.valAxisLabelFormatCode ? opts.valAxisLabelFormatCode : 'General') + "' sourceLinked=\"0\"/>";
    if (opts._type === CHART_TYPE.SCATTER) {
        strXml += '  <c:majorTickMark val="none"/>';
        strXml += '  <c:minorTickMark val="none"/>';
        strXml += '  <c:tickLblPos val="nextTo"/>';
    }
    else {
        strXml += ' <c:majorTickMark val="' + (opts.valAxisMajorTickMark || 'out') + '"/>';
        strXml += ' <c:minorTickMark val="' + (opts.valAxisMinorTickMark || 'none') + '"/>';
        strXml += ' <c:tickLblPos val="' + (opts.valAxisLabelPos || (opts.barDir === 'col' ? 'nextTo' : 'low')) + '"/>';
    }
    strXml += ' <c:spPr>';
    strXml += '   <a:ln w="' + (opts.valAxisLineSize ? valToPts(opts.valAxisLineSize) : ONEPT) + '" cap="flat">';
    strXml += opts.valAxisLineShow === false ? '<a:noFill/>' : '<a:solidFill>' + createColorElement(opts.valAxisLineColor || DEF_CHART_GRIDLINE.color) + '</a:solidFill>';
    strXml += '     <a:prstDash val="' + (opts.valAxisLineStyle || 'solid') + '"/>';
    strXml += '     <a:round/>';
    strXml += '   </a:ln>';
    strXml += ' </c:spPr>';
    strXml += ' <c:txPr>';
    strXml += '  <a:bodyPr ' + (opts.valAxisLabelRotate ? 'rot="' + convertRotationDegrees(opts.valAxisLabelRotate) + '"' : '') + '/>'; // don't specify rot 0 so we get the auto behavior
    strXml += '  <a:lstStyle/>';
    strXml += '  <a:p>';
    strXml += '    <a:pPr>';
    strXml +=
        '      <a:defRPr sz="' + Math.round((opts.valAxisLabelFontSize || DEF_FONT_SIZE) * 100) + '" b="' + (opts.valAxisLabelFontBold ? 1 : 0) + '" i="' + (opts.valAxisLabelFontItalic ? 1 : 0) + '" u="none" strike="noStrike">';
    strXml += '        <a:solidFill>' + createColorElement(opts.valAxisLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
    strXml += '        <a:latin typeface="' + (opts.valAxisLabelFontFace || 'Arial') + '"/>';
    strXml += '      </a:defRPr>';
    strXml += '    </a:pPr>';
    strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>';
    strXml += '  </a:p>';
    strXml += ' </c:txPr>';
    strXml += ' <c:crossAx val="' + crossAxId + '"/>';
    strXml += ' <c:crosses val="' + crosses + '"/>';
    strXml +=
        ' <c:crossBetween val="' +
            (opts._type === CHART_TYPE.SCATTER || (Array.isArray(opts._type) && opts._type.filter(function (type) { return type.type === CHART_TYPE.AREA; }).length > 0 ? true : false)
                ? 'midCat'
                : 'between') +
            '"/>';
    if (opts.valAxisMajorUnit)
        strXml += ' <c:majorUnit val="' + opts.valAxisMajorUnit + '"/>';
    if (opts.valAxisDisplayUnit)
        strXml += "<c:dispUnits><c:builtInUnit val=\"" + opts.valAxisDisplayUnit + "\"/>" + (opts.valAxisDisplayUnitLabel ? '<c:dispUnitsLbl/>' : '') + "</c:dispUnits>";
    strXml += '</c:valAx>';
    return strXml;
}
/**
 * Create Series Axis (Used by `bar3D`)
 * @param {IChartOptsLib} opts - chart options
 * @param {string} axisId - axis ID
 * @param {string} valAxisId - value
 * @return {string} XML
 */
function makeSerAxis(opts, axisId, valAxisId) {
    var strXml = '';
    // Build ser axis tag
    strXml += '<c:serAx>';
    strXml += '  <c:axId val="' + axisId + '"/>';
    strXml += '  <c:scaling><c:orientation val="' + (opts.serAxisOrientation || (opts.barDir === 'col' ? 'minMax' : 'minMax')) + '"/></c:scaling>';
    strXml += '  <c:delete val="' + (opts.serAxisHidden ? 1 : 0) + '"/>';
    strXml += '  <c:axPos val="' + (opts.barDir === 'col' ? 'b' : 'l') + '"/>';
    strXml += opts.serGridLine.style !== 'none' ? createGridLineElement(opts.serGridLine) : '';
    // '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
    if (opts.showSerAxisTitle) {
        strXml += genXmlTitle({
            color: opts.serAxisTitleColor,
            fontFace: opts.serAxisTitleFontFace,
            fontSize: opts.serAxisTitleFontSize,
            titleRotate: opts.serAxisTitleRotate,
            title: opts.serAxisTitle || 'Axis Title',
        });
    }
    strXml += '  <c:numFmt formatCode="' + (opts.serLabelFormatCode || 'General') + '" sourceLinked="0"/>';
    strXml += '  <c:majorTickMark val="out"/>';
    strXml += '  <c:minorTickMark val="none"/>';
    strXml += '  <c:tickLblPos val="' + (opts.serAxisLabelPos || opts.barDir === 'col' ? 'low' : 'nextTo') + '"/>';
    strXml += '  <c:spPr>';
    strXml += '    <a:ln w="12700" cap="flat">';
    strXml += opts.serAxisLineShow === false ? '<a:noFill/>' : '<a:solidFill>' + createColorElement(opts.serAxisLineColor || DEF_CHART_GRIDLINE.color) + '</a:solidFill>';
    strXml += '      <a:prstDash val="solid"/>';
    strXml += '      <a:round/>';
    strXml += '    </a:ln>';
    strXml += '  </c:spPr>';
    strXml += '  <c:txPr>';
    strXml += '    <a:bodyPr/>'; // don't specify rot 0 so we get the auto behavior
    strXml += '    <a:lstStyle/>';
    strXml += '    <a:p>';
    strXml += '    <a:pPr>';
    strXml += "    <a:defRPr sz=\"" + Math.round((opts.serAxisLabelFontSize || DEF_FONT_SIZE) * 100) + "\" b=\"" + (opts.serAxisLabelFontBold || 0) + "\" i=\"" + (opts.serAxisLabelFontItalic || 0) + "\" u=\"none\" strike=\"noStrike\">";
    strXml += '      <a:solidFill>' + createColorElement(opts.serAxisLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
    strXml += '      <a:latin typeface="' + (opts.serAxisLabelFontFace || 'Arial') + '"/>';
    strXml += '   </a:defRPr>';
    strXml += '  </a:pPr>';
    strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>';
    strXml += '  </a:p>';
    strXml += ' </c:txPr>';
    strXml += ' <c:crossAx val="' + valAxisId + '"/>';
    strXml += ' <c:crosses val="autoZero"/>';
    if (opts.serAxisLabelFrequency)
        strXml += ' <c:tickLblSkip val="' + opts.serAxisLabelFrequency + '"/>';
    // Issue#149: PPT will auto-adjust these as needed after calcing the date bounds, so we only include them when specified by user
    if (opts.serLabelFormatCode) {
        ['serAxisBaseTimeUnit', 'serAxisMajorTimeUnit', 'serAxisMinorTimeUnit'].forEach(function (opt) {
            // Validate input as poorly chosen/garbage options will cause chart corruption and it wont render at all!
            if (opts[opt] && (typeof opts[opt] !== 'string' || ['days', 'months', 'years'].indexOf(opt.toLowerCase()) === -1)) {
                console.warn('`' + opt + "` must be one of: 'days','months','years' !");
                opts[opt] = null;
            }
        });
        if (opts.serAxisBaseTimeUnit)
            strXml += ' <c:baseTimeUnit  val="' + opts.serAxisBaseTimeUnit.toLowerCase() + '"/>';
        if (opts.serAxisMajorTimeUnit)
            strXml += ' <c:majorTimeUnit val="' + opts.serAxisMajorTimeUnit.toLowerCase() + '"/>';
        if (opts.serAxisMinorTimeUnit)
            strXml += ' <c:minorTimeUnit val="' + opts.serAxisMinorTimeUnit.toLowerCase() + '"/>';
        if (opts.serAxisMajorUnit)
            strXml += ' <c:majorUnit     val="' + opts.serAxisMajorUnit + '"/>';
        if (opts.serAxisMinorUnit)
            strXml += ' <c:minorUnit     val="' + opts.serAxisMinorUnit + '"/>';
    }
    // Close ser axis tag
    strXml += '</c:serAx>';
    return strXml;
}
/**
 * Create char title elements
 * @param {IChartPropsTitle} opts - options
 * @return {string} XML `<c:title>`
 */
function genXmlTitle(opts) {
    var align = opts.titleAlign === 'left' || opts.titleAlign === 'right' ? "<a:pPr algn=\"" + opts.titleAlign.substring(0, 1) + "\">" : "<a:pPr>";
    var rotate = opts.titleRotate ? "<a:bodyPr rot=\"" + convertRotationDegrees(opts.titleRotate) + "\"/>" : "<a:bodyPr/>"; // don't specify rotation to get default (ex. vertical for cat axis)
    var sizeAttr = opts.fontSize ? 'sz="' + Math.round(opts.fontSize * 100) + '"' : ''; // only set the font size if specified.  Powerpoint will handle the default size
    var titleBold = opts.titleBold === true ? 1 : 0;
    var layout = opts.titlePos && opts.titlePos.x && opts.titlePos.y
        ? "<c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"" + opts.titlePos.x + "\"/><c:y val=\"" + opts.titlePos.y + "\"/></c:manualLayout></c:layout>"
        : "<c:layout/>";
    return "<c:title>\n\t  <c:tx>\n\t    <c:rich>\n\t      " + rotate + "\n\t      <a:lstStyle/>\n\t      <a:p>\n\t        " + align + "\n\t        <a:defRPr " + sizeAttr + " b=\"" + titleBold + "\" i=\"0\" u=\"none\" strike=\"noStrike\">\n\t          <a:solidFill>" + createColorElement(opts.color || DEF_FONT_COLOR) + "</a:solidFill>\n\t          <a:latin typeface=\"" + (opts.fontFace || 'Arial') + "\"/>\n\t        </a:defRPr>\n\t      </a:pPr>\n\t      <a:r>\n\t        <a:rPr " + sizeAttr + " b=\"" + titleBold + "\" i=\"0\" u=\"none\" strike=\"noStrike\">\n\t          <a:solidFill>" + createColorElement(opts.color || DEF_FONT_COLOR) + "</a:solidFill>\n\t          <a:latin typeface=\"" + (opts.fontFace || 'Arial') + "\"/>\n\t        </a:rPr>\n\t        <a:t>" + (encodeXmlEntities(opts.title) || '') + "</a:t>\n\t      </a:r>\n\t    </a:p>\n\t    </c:rich>\n\t  </c:tx>\n\t  " + layout + "\n\t  <c:overlay val=\"0\"/>\n\t</c:title>";
}
/**
 * Calc and return excel column name for a given column length
 * @param {number} length - col length
 * @return {string} column name (ex: 'A2')
 */
function getExcelColName(length) {
    var strName = '';
    if (length <= 26) {
        strName = LETTERS[length];
    }
    else {
        strName += LETTERS[Math.floor(length / LETTERS.length) - 1];
        strName += LETTERS[length % LETTERS.length];
    }
    return strName;
}
/**
 * Creates `a:innerShdw` or `a:outerShdw` depending on pass options `opts`.
 * @param {Object} opts optional shadow properties
 * @param {Object} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 * @example { type: 'outer', blur: 3, offset: (23000 / 12700), angle: 90, color: '000000', opacity: 0.35, rotateWithShape: true };
 * @return {string} XML
 */
function createShadowElement(options, defaults) {
    if (!options) {
        return '<a:effectLst/>';
    }
    else if (typeof options !== 'object') {
        console.warn("`shadow` options must be an object. Ex: `{shadow: {type:'none'}}`");
        return '<a:effectLst/>';
    }
    var strXml = '<a:effectLst>', opts = getMix(defaults, options), type = opts['type'] || 'outer', blur = valToPts(opts['blur']), offset = valToPts(opts['offset']), angle = Math.round(opts['angle'] * 60000), color = opts['color'], opacity = Math.round(opts['opacity'] * 100000), rotateWithShape = opts['rotateWithShape'] ? 1 : 0;
    strXml += '<a:' + type + 'Shdw sx="100000" sy="100000" kx="0" ky="0"  algn="bl" blurRad="' + blur + '" ';
    strXml += 'rotWithShape="' + +rotateWithShape + '"';
    strXml += ' dist="' + offset + '" dir="' + angle + '">';
    strXml += '<a:srgbClr val="' + color + '">';
    strXml += '<a:alpha val="' + opacity + '"/></a:srgbClr>';
    strXml += '</a:' + type + 'Shdw>';
    strXml += '</a:effectLst>';
    return strXml;
}
/**
 * Create Grid Line Element
 * @param {OptsChartGridLine} glOpts {size, color, style}
 * @return {string} XML
 */
function createGridLineElement(glOpts) {
    var strXml = '<c:majorGridlines>';
    strXml += ' <c:spPr>';
    strXml += '  <a:ln w="' + valToPts(glOpts.size || DEF_CHART_GRIDLINE.size) + '" cap="flat">';
    strXml += '  <a:solidFill><a:srgbClr val="' + (glOpts.color || DEF_CHART_GRIDLINE.color) + '"/></a:solidFill>'; // should accept scheme colors as implemented in [Pull #135]
    strXml += '   <a:prstDash val="' + (glOpts.style || DEF_CHART_GRIDLINE.style) + '"/><a:round/>';
    strXml += '  </a:ln>';
    strXml += ' </c:spPr>';
    strXml += '</c:majorGridlines>';
    return strXml;
}

/**
 * PptxGenJS: Media Methods
 */
/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlide | SlideLayout} layout - slide layout
 * @return {Promise} promise
 */
function encodeSlideMediaRels(layout) {
    var fs = typeof require !== 'undefined' && typeof window === 'undefined' ? require('fs') : null; // NodeJS
    var https = typeof require !== 'undefined' && typeof window === 'undefined' ? require('https') : null; // NodeJS
    var imageProms = [];
    // A: Read/Encode each audio/image/video thats not already encoded (eg: base64 provided by user)
    layout._relsMedia
        .filter(function (rel) { return rel.type !== 'online' && !rel.data && (!rel.path || (rel.path && rel.path.indexOf('preencoded') === -1)); })
        .forEach(function (rel) {
        imageProms.push(new Promise(function (resolve, reject) {
            if (fs && rel.path.indexOf('http') !== 0) {
                // DESIGN: Node local-file encoding is syncronous, so we can load all images here, then call export with a callback (if any)
                try {
                    var bitmap = fs.readFileSync(rel.path);
                    rel.data = Buffer.from(bitmap).toString('base64');
                    resolve('done');
                }
                catch (ex) {
                    rel.data = IMG_BROKEN;
                    reject('ERROR: Unable to read media: "' + rel.path + '"\n' + ex.toString());
                }
            }
            else if (fs && https && rel.path.indexOf('http') === 0) {
                https.get(rel.path, function (res) {
                    var rawData = '';
                    res.setEncoding('binary'); // IMPORTANT: Only binary encoding works
                    res.on('data', function (chunk) { return (rawData += chunk); });
                    res.on('end', function () {
                        rel.data = Buffer.from(rawData, 'binary').toString('base64');
                        resolve('done');
                    });
                    res.on('error', function (ex) {
                        rel.data = IMG_BROKEN;
                        reject("ERROR! Unable to load image (https.get): " + rel.path);
                    });
                });
            }
            else {
                // A: Declare XHR and onload/onerror handlers
                // DESIGN: `XMLHttpRequest()` plus `FileReader()` = Ablity to read any file into base64!
                var xhr_1 = new XMLHttpRequest();
                xhr_1.onload = function () {
                    var reader = new FileReader();
                    reader.onloadend = function () {
                        rel.data = reader.result;
                        if (!rel.isSvgPng) {
                            resolve('done');
                        }
                        else {
                            createSvgPngPreview(rel)
                                .then(function () {
                                resolve('done');
                            })
                                .catch(function (ex) {
                                reject(ex);
                            });
                        }
                    };
                    reader.readAsDataURL(xhr_1.response);
                };
                xhr_1.onerror = function (ex) {
                    rel.data = IMG_BROKEN;
                    reject("ERROR! Unable to load image (xhr.onerror): " + rel.path);
                };
                // B: Execute request
                xhr_1.open('GET', rel.path);
                xhr_1.responseType = 'blob';
                xhr_1.send();
            }
        }));
    });
    // B: SVG: base64 data still requires a png to be generated (`isSvgPng` flag this as the preview image, not the SVG itself)
    layout._relsMedia
        .filter(function (rel) { return rel.isSvgPng && rel.data; })
        .forEach(function (rel) {
        if (fs) {
            //console.log('Sorry, SVG is not supported in Node (more info: https://github.com/gitbrent/PptxGenJS/issues/401)')
            rel.data = IMG_BROKEN;
            imageProms.push(Promise.resolve().then(function () { return 'done'; }));
        }
        else {
            imageProms.push(createSvgPngPreview(rel));
        }
    });
    return imageProms;
}
/**
 * Create SVG preview image
 * @param {ISlideRelMedia} rel - slide rel
 * @return {Promise} promise
 */
function createSvgPngPreview(rel) {
    return new Promise(function (resolve, reject) {
        // A: Create
        var image = new Image();
        // B: Set onload event
        image.onload = function () {
            // First: Check for any errors: This is the best method (try/catch wont work, etc.)
            if (image.width + image.height === 0) {
                image.onerror('h/w=0');
            }
            var canvas = document.createElement('CANVAS');
            var ctx = canvas.getContext('2d');
            canvas.width = image.width;
            canvas.height = image.height;
            ctx.drawImage(image, 0, 0);
            // Users running on local machine will get the following error:
            // "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
            // when the canvas.toDataURL call executes below.
            try {
                rel.data = canvas.toDataURL(rel.type);
                resolve('done');
            }
            catch (ex) {
                image.onerror(ex);
            }
            canvas = null;
        };
        image.onerror = function (ex) {
            rel.data = IMG_BROKEN;
            reject("ERROR! Unable to load image (image.onerror): " + rel.path);
        };
        // C: Load image
        image.src = typeof rel.data === 'string' ? rel.data : IMG_BROKEN;
    });
}

/**
 *  :: pptxgen.ts ::
 *
 *  JavaScript framework that creates PowerPoint (pptx) presentations
 *  https://github.com/gitbrent/PptxGenJS
 *
 *  This framework is released under the MIT Public License (MIT)
 *
 *  PptxGenJS (C) 2015-present Brent Ely -- https://github.com/gitbrent
 *
 *  Some code derived from the OfficeGen project:
 *  github.com/Ziv-Barber/officegen/ (Copyright 2013 Ziv Barber)
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in all
 *  copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 *  SOFTWARE.
 */
var VERSION = '3.7.1';
var PptxGenJS = /** @class */ (function () {
    function PptxGenJS() {
        var _this = this;
        /**
         * PptxGenJS Library Version
         */
        this._version = VERSION;
        // Exposed class props
        this._alignH = AlignH;
        this._alignV = AlignV;
        this._chartType = ChartType;
        this._outputType = OutputType;
        this._schemeColor = SchemeColor;
        this._shapeType = ShapeType;
        /**
         * @depricated use `ChartType`
         */
        this._charts = CHART_TYPE;
        /**
         * @depricated use `SchemeColor`
         */
        this._colors = SCHEME_COLOR_NAMES;
        /**
         * @depricated use `ShapeType`
         */
        this._shapes = SHAPE_TYPE;
        /**
         * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
         * @param {string} masterName - slide master name
         * @return {PresSlide} new Slide
         */
        this.addNewSlide = function (masterName) {
            // Continue using sections if the first slide using auto-paging has a Section
            var sectAlreadyInUse = _this.sections.length > 0 &&
                _this.sections[_this.sections.length - 1]._slides.filter(function (slide) { return slide._slideNum === _this.slides[_this.slides.length - 1]._slideNum; }).length > 0;
            return _this.addSlide({
                masterName: masterName,
                sectionTitle: sectAlreadyInUse ? _this.sections[_this.sections.length - 1].title : null,
            });
        };
        /**
         * Provides an API for `addTableDefinition` to get slide reference by number
         * @param {number} slideNum - slide number
         * @return {PresSlide} Slide
         * @since 3.0.0
         */
        this.getSlide = function (slideNum) { return _this.slides.filter(function (slide) { return slide._slideNum === slideNum; })[0]; };
        /**
         * Enables the `Slide` class to set PptxGenJS [Presentation] master/layout slidenumbers
         * @param {SlideNumberProps} slideNum - slide number config
         */
        this.setSlideNumber = function (slideNum) {
            // 1: Add slideNumber to slideMaster1.xml
            _this.masterSlide._slideNumberProps = slideNum;
            // 2: Add slideNumber to DEF_PRES_LAYOUT_NAME layout
            _this.slideLayouts.filter(function (layout) { return layout._name === DEF_PRES_LAYOUT_NAME; })[0]._slideNumberProps = slideNum;
        };
        /**
         * Create all chart and media rels for this Presentation
         * @param {PresSlide | SlideLayout} slide - slide with rels
         * @param {JSZip} zip - JSZip instance
         * @param {Promise<any>[]} chartPromises - promise array
         */
        this.createChartMediaRels = function (slide, zip, chartPromises) {
            slide._relsChart.forEach(function (rel) { return chartPromises.push(createExcelWorksheet(rel, zip)); });
            slide._relsMedia.forEach(function (rel) {
                if (rel.type !== 'online' && rel.type !== 'hyperlink') {
                    // A: Loop vars
                    var data = rel.data && typeof rel.data === 'string' ? rel.data : '';
                    // B: Users will undoubtedly pass various string formats, so correct prefixes as needed
                    if (data.indexOf(',') === -1 && data.indexOf(';') === -1)
                        data = 'image/png;base64,' + data;
                    else if (data.indexOf(',') === -1)
                        data = 'image/png;base64,' + data;
                    else if (data.indexOf(';') === -1)
                        data = 'image/png;' + data;
                    // C: Add media
                    zip.file(rel.Target.replace('..', 'ppt'), data.split(',').pop(), { base64: true });
                }
            });
        };
        /**
         * Create and export the .pptx file
         * @param {string} exportName - output file type
         * @param {Blob} blobContent - Blob content
         * @return {Promise<string>} Promise with file name
         */
        this.writeFileToBrowser = function (exportName, blobContent) {
            // STEP 1: Create element
            var eleLink = document.createElement('a');
            eleLink.setAttribute('style', 'display:none;');
            eleLink.dataset.interception = 'off'; // @see https://docs.microsoft.com/en-us/sharepoint/dev/spfx/hyperlinking
            document.body.appendChild(eleLink);
            // STEP 2: Download file to browser
            // DESIGN: Use `createObjectURL()` (or MS-specific func for IE11) to D/L files in client browsers (FYI: synchronously executed)
            if (window.navigator.msSaveOrOpenBlob) {
                // @see https://docs.microsoft.com/en-us/microsoft-edge/dev-guide/html5/file-api/blob
                var blob_1 = new Blob([blobContent], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
                eleLink.onclick = function () {
                    window.navigator.msSaveOrOpenBlob(blob_1, exportName);
                };
                eleLink.click();
                // Clean-up
                document.body.removeChild(eleLink);
                // Done
                return Promise.resolve(exportName);
            }
            else if (window.URL.createObjectURL) {
                var url_1 = window.URL.createObjectURL(new Blob([blobContent], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
                eleLink.href = url_1;
                eleLink.download = exportName;
                eleLink.click();
                // Clean-up (NOTE: Add a slight delay before removing to avoid 'blob:null' error in Firefox Issue#81)
                setTimeout(function () {
                    window.URL.revokeObjectURL(url_1);
                    document.body.removeChild(eleLink);
                }, 100);
                // Done
                return Promise.resolve(exportName);
            }
        };
        /**
         * Create and export the .pptx file
         * @param {WRITE_OUTPUT_TYPE} outputType - output file type
         * @return {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} Promise with data or stream (node) or filename (browser)
         */
        this.exportPresentation = function (props) {
            var arrChartPromises = [];
            var arrMediaPromises = [];
            var zip = new JSZip();
            // STEP 1: Read/Encode all Media before zip as base64 content, etc. is required
            _this.slides.forEach(function (slide) {
                arrMediaPromises = arrMediaPromises.concat(encodeSlideMediaRels(slide));
            });
            _this.slideLayouts.forEach(function (layout) {
                arrMediaPromises = arrMediaPromises.concat(encodeSlideMediaRels(layout));
            });
            arrMediaPromises = arrMediaPromises.concat(encodeSlideMediaRels(_this.masterSlide));
            // STEP 2: Wait for Promises (if any) then generate the PPTX file
            return Promise.all(arrMediaPromises).then(function () {
                // A: Add empty placeholder objects to slides that don't already have them
                _this.slides.forEach(function (slide) {
                    if (slide._slideLayout)
                        addPlaceholdersToSlideLayouts(slide);
                });
                // B: Add all required folders and files
                zip.folder('_rels');
                zip.folder('docProps');
                zip.folder('ppt').folder('_rels');
                zip.folder('ppt/charts').folder('_rels');
                zip.folder('ppt/embeddings');
                zip.folder('ppt/media');
                zip.folder('ppt/slideLayouts').folder('_rels');
                zip.folder('ppt/slideMasters').folder('_rels');
                zip.folder('ppt/slides').folder('_rels');
                zip.folder('ppt/theme');
                zip.folder('ppt/notesMasters').folder('_rels');
                zip.folder('ppt/notesSlides').folder('_rels');
                zip.file('[Content_Types].xml', makeXmlContTypes(_this.slides, _this.slideLayouts, _this.masterSlide)); // TODO: pass only `this` like below! 20200206
                zip.file('_rels/.rels', makeXmlRootRels());
                zip.file('docProps/app.xml', makeXmlApp(_this.slides, _this.company)); // TODO: pass only `this` like below! 20200206
                zip.file('docProps/core.xml', makeXmlCore(_this.title, _this.subject, _this.author, _this.revision)); // TODO: pass only `this` like below! 20200206
                zip.file('ppt/_rels/presentation.xml.rels', makeXmlPresentationRels(_this.slides));
                zip.file('ppt/theme/theme1.xml', makeXmlTheme());
                zip.file('ppt/presentation.xml', makeXmlPresentation(_this));
                zip.file('ppt/presProps.xml', makeXmlPresProps());
                zip.file('ppt/tableStyles.xml', makeXmlTableStyles());
                zip.file('ppt/viewProps.xml', makeXmlViewProps());
                // C: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
                _this.slideLayouts.forEach(function (layout, idx) {
                    zip.file('ppt/slideLayouts/slideLayout' + (idx + 1) + '.xml', makeXmlLayout(layout));
                    zip.file('ppt/slideLayouts/_rels/slideLayout' + (idx + 1) + '.xml.rels', makeXmlSlideLayoutRel(idx + 1, _this.slideLayouts));
                });
                _this.slides.forEach(function (slide, idx) {
                    zip.file('ppt/slides/slide' + (idx + 1) + '.xml', makeXmlSlide(slide));
                    zip.file('ppt/slides/_rels/slide' + (idx + 1) + '.xml.rels', makeXmlSlideRel(_this.slides, _this.slideLayouts, idx + 1));
                    // Create all slide notes related items. Notes of empty strings are created for slides which do not have notes specified, to keep track of _rels.
                    zip.file('ppt/notesSlides/notesSlide' + (idx + 1) + '.xml', makeXmlNotesSlide(slide));
                    zip.file('ppt/notesSlides/_rels/notesSlide' + (idx + 1) + '.xml.rels', makeXmlNotesSlideRel(idx + 1));
                });
                zip.file('ppt/slideMasters/slideMaster1.xml', makeXmlMaster(_this.masterSlide, _this.slideLayouts));
                zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', makeXmlMasterRel(_this.masterSlide, _this.slideLayouts));
                zip.file('ppt/notesMasters/notesMaster1.xml', makeXmlNotesMaster());
                zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', makeXmlNotesMasterRel());
                // D: Create all Rels (images, media, chart data)
                _this.slideLayouts.forEach(function (layout) {
                    _this.createChartMediaRels(layout, zip, arrChartPromises);
                });
                _this.slides.forEach(function (slide) {
                    _this.createChartMediaRels(slide, zip, arrChartPromises);
                });
                _this.createChartMediaRels(_this.masterSlide, zip, arrChartPromises);
                // E: Wait for Promises (if any) then generate the PPTX file
                return Promise.all(arrChartPromises).then(function () {
                    if (props.outputType === 'STREAM') {
                        // A: stream file
                        return zip.generateAsync({ type: 'nodebuffer', compression: props.compression ? 'DEFLATE' : 'STORE' });
                    }
                    else if (props.outputType) {
                        // B: Node [fs]: Output type user option or default
                        return zip.generateAsync({ type: props.outputType });
                    }
                    else {
                        // C: Browser: Output blob as app/ms-pptx
                        return zip.generateAsync({ type: 'blob', compression: props.compression ? 'DEFLATE' : 'STORE' });
                    }
                });
            });
        };
        // Set available layouts
        this.LAYOUTS = {
            LAYOUT_4x3: { name: 'screen4x3', width: 9144000, height: 6858000 },
            LAYOUT_16x9: { name: 'screen16x9', width: 9144000, height: 5143500 },
            LAYOUT_16x10: { name: 'screen16x10', width: 9144000, height: 5715000 },
            LAYOUT_WIDE: { name: 'custom', width: 12192000, height: 6858000 },
        };
        // Core
        this._author = 'PptxGenJS';
        this._company = 'PptxGenJS';
        this._revision = '1'; // Note: Must be a whole number
        this._subject = 'PptxGenJS Presentation';
        this._title = 'PptxGenJS Presentation';
        // PptxGenJS props
        this._presLayout = {
            name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
            _sizeW: this.LAYOUTS[DEF_PRES_LAYOUT].width,
            _sizeH: this.LAYOUTS[DEF_PRES_LAYOUT].height,
            width: this.LAYOUTS[DEF_PRES_LAYOUT].width,
            height: this.LAYOUTS[DEF_PRES_LAYOUT].height,
        };
        this._rtlMode = false;
        //
        this._slideLayouts = [
            {
                _margin: DEF_SLIDE_MARGIN_IN,
                _name: DEF_PRES_LAYOUT_NAME,
                _presLayout: this._presLayout,
                _rels: [],
                _relsChart: [],
                _relsMedia: [],
                _slide: null,
                _slideNum: 1000,
                _slideNumberProps: null,
                _slideObjects: [],
            },
        ];
        this._slides = [];
        this._sections = [];
        this._masterSlide = {
            addChart: null,
            addImage: null,
            addMedia: null,
            addNotes: null,
            addShape: null,
            addTable: null,
            addText: null,
            //
            _name: null,
            _presLayout: this._presLayout,
            _rId: null,
            _rels: [],
            _relsChart: [],
            _relsMedia: [],
            _slideId: null,
            _slideLayout: null,
            _slideNum: null,
            _slideNumberProps: null,
            _slideObjects: [],
        };
    }
    Object.defineProperty(PptxGenJS.prototype, "layout", {
        get: function () {
            return this._layout;
        },
        set: function (value) {
            var newLayout = this.LAYOUTS[value];
            if (newLayout) {
                this._layout = value;
                this._presLayout = newLayout;
            }
            else {
                throw new Error('UNKNOWN-LAYOUT');
            }
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "version", {
        get: function () {
            return this._version;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "author", {
        get: function () {
            return this._author;
        },
        set: function (value) {
            this._author = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "company", {
        get: function () {
            return this._company;
        },
        set: function (value) {
            this._company = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "revision", {
        get: function () {
            return this._revision;
        },
        set: function (value) {
            this._revision = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "subject", {
        get: function () {
            return this._subject;
        },
        set: function (value) {
            this._subject = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "title", {
        get: function () {
            return this._title;
        },
        set: function (value) {
            this._title = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "rtlMode", {
        get: function () {
            return this._rtlMode;
        },
        set: function (value) {
            this._rtlMode = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "masterSlide", {
        get: function () {
            return this._masterSlide;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "slides", {
        get: function () {
            return this._slides;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "sections", {
        get: function () {
            return this._sections;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "slideLayouts", {
        get: function () {
            return this._slideLayouts;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "AlignH", {
        get: function () {
            return this._alignH;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "AlignV", {
        get: function () {
            return this._alignV;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "ChartType", {
        get: function () {
            return this._chartType;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "OutputType", {
        get: function () {
            return this._outputType;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "presLayout", {
        get: function () {
            return this._presLayout;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "SchemeColor", {
        get: function () {
            return this._schemeColor;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "ShapeType", {
        get: function () {
            return this._shapeType;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "charts", {
        get: function () {
            return this._charts;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "colors", {
        get: function () {
            return this._colors;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "shapes", {
        get: function () {
            return this._shapes;
        },
        enumerable: false,
        configurable: true
    });
    // EXPORT METHODS
    /**
     * Export the current Presentation to stream
     * @param {WriteBaseProps} props - output properties
     * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file stream
     */
    PptxGenJS.prototype.stream = function (props) {
        var propsCompress = typeof props === 'object' && props.hasOwnProperty('compression') ? props.compression : false;
        return this.exportPresentation({
            compression: propsCompress,
            outputType: 'STREAM',
        });
    };
    /**
     * Export the current Presentation as JSZip content with the selected type
     * @param {WriteProps} props - output properties
     * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file content in selected type
     */
    PptxGenJS.prototype.write = function (props) {
        // DEPRECATED: @deprecated v3.5.0 - outputType - [[remove in v4.0.0]]
        var propsOutpType = typeof props === 'object' && props.hasOwnProperty('outputType') ? props.outputType : props ? props : null;
        var propsCompress = typeof props === 'object' && props.hasOwnProperty('compression') ? props.compression : false;
        return this.exportPresentation({
            compression: propsCompress,
            outputType: propsOutpType,
        });
    };
    /**
     * Export the current Presentation. Writes file to local file system if `fs` exists, otherwise, initiates download in browsers
     * @param {WriteFileProps} props - output file properties
     * @returns {Promise<string>} the presentation name
     */
    PptxGenJS.prototype.writeFile = function (props) {
        var _this = this;
        var fs = typeof require !== 'undefined' && typeof window === 'undefined' ? require('fs') : null; // NodeJS
        // DEPRECATED: @deprecated v3.5.0 - fileName - [[remove in v4.0.0]]
        if (typeof props === 'string')
            console.log('Warning: `writeFile(filename)` is deprecated - please use `WriteFileProps` argument (v3.5.0)');
        var propsExpName = typeof props === 'object' && props.hasOwnProperty('fileName') ? props.fileName : typeof props === 'string' ? props : '';
        var propsCompress = typeof props === 'object' && props.hasOwnProperty('compression') ? props.compression : false;
        var fileName = propsExpName ? (propsExpName.toString().toLowerCase().endsWith('.pptx') ? propsExpName : propsExpName + '.pptx') : 'Presentation.pptx';
        return this.exportPresentation({
            compression: propsCompress,
            outputType: fs ? 'nodebuffer' : null,
        }).then(function (content) {
            if (fs) {
                // Node: Output
                return new Promise(function (resolve, reject) {
                    fs.writeFile(fileName, content, function (err) {
                        if (err) {
                            reject(err);
                        }
                        else {
                            resolve(fileName);
                        }
                    });
                });
            }
            else {
                // Browser: Output blob as app/ms-pptx
                return _this.writeFileToBrowser(fileName, content);
            }
        });
    };
    // PRESENTATION METHODS
    /**
     * Add a new Section to Presentation
     * @param {ISectionProps} section - section properties
     * @example pptx.addSection({ title:'Charts' });
     */
    PptxGenJS.prototype.addSection = function (section) {
        if (!section)
            console.warn('addSection requires an argument');
        else if (!section.title)
            console.warn('addSection requires a title');
        var newSection = {
            _type: 'user',
            _slides: [],
            title: section.title,
        };
        if (section.order)
            this.sections.splice(section.order, 0, newSection);
        else
            this._sections.push(newSection);
    };
    /**
     * Add a new Slide to Presentation
     * @param {AddSlideProps} options - slide options
     * @returns {PresSlide} the new Slide
     */
    PptxGenJS.prototype.addSlide = function (options) {
        // TODO: DEPRECATED: arg0 string "masterSlideName" dep as of 3.2.0
        var masterSlideName = typeof options === 'string' ? options : options && options.masterName ? options.masterName : '';
        var slideLayout = {
            _name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
            _presLayout: this.presLayout,
            _rels: [],
            _relsChart: [],
            _relsMedia: [],
            _slideNum: this.slides.length + 1,
        };
        if (masterSlideName) {
            var tmpLayout = this.slideLayouts.filter(function (layout) { return layout._name === masterSlideName; })[0];
            if (tmpLayout)
                slideLayout = tmpLayout;
        }
        var newSlide = new Slide({
            addSlide: this.addNewSlide,
            getSlide: this.getSlide,
            presLayout: this.presLayout,
            setSlideNum: this.setSlideNumber,
            slideId: this.slides.length + 256,
            slideRId: this.slides.length + 2,
            slideNumber: this.slides.length + 1,
            slideLayout: slideLayout,
        });
        // A: Add slide to pres
        this._slides.push(newSlide);
        // B: Sections
        // B-1: Add slide to section (if any provided)
        if (options && options.sectionTitle) {
            var sect = this.sections.filter(function (section) { return section.title === options.sectionTitle; })[0];
            if (!sect)
                console.warn("addSlide: unable to find section with title: \"" + options.sectionTitle + "\"");
            else
                sect._slides.push(newSlide);
        }
        // B-2: Handle slides without a section when sections are already is use ("loose" slides arent allowed, they all need a section)
        else if (this.sections && this.sections.length > 0 && (!options || !options.sectionTitle)) {
            var lastSect = this._sections[this.sections.length - 1];
            // CASE 1: The latest section is a default type - just add this one
            if (lastSect._type === 'default')
                lastSect._slides.push(newSlide);
            // CASE 2: There latest section is NOT a default type - create the defualt, add this slide
            else
                this._sections.push({
                    title: "Default-" + (this.sections.filter(function (sect) { return sect._type === 'default'; }).length + 1),
                    _type: 'default',
                    _slides: [newSlide],
                });
        }
        return newSlide;
    };
    /**
     * Create a custom Slide Layout in any size
     * @param {PresLayout} layout - layout properties
     * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
     */
    PptxGenJS.prototype.defineLayout = function (layout) {
        // @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
        if (!layout)
            console.warn('defineLayout requires `{name, width, height}`');
        else if (!layout.name)
            console.warn('defineLayout requires `name`');
        else if (!layout.width)
            console.warn('defineLayout requires `width`');
        else if (!layout.height)
            console.warn('defineLayout requires `height`');
        else if (typeof layout.height !== 'number')
            console.warn('defineLayout `height` should be a number (inches)');
        else if (typeof layout.width !== 'number')
            console.warn('defineLayout `width` should be a number (inches)');
        this.LAYOUTS[layout.name] = {
            name: layout.name,
            _sizeW: Math.round(Number(layout.width) * EMU),
            _sizeH: Math.round(Number(layout.height) * EMU),
            width: Math.round(Number(layout.width) * EMU),
            height: Math.round(Number(layout.height) * EMU),
        };
    };
    /**
     * Create a new slide master [layout] for the Presentation
     * @param {SlideMasterProps} props - layout properties
     */
    PptxGenJS.prototype.defineSlideMaster = function (props) {
        if (!props.title)
            throw new Error('defineSlideMaster() object argument requires a `title` value. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)');
        var newLayout = {
            _margin: props.margin || DEF_SLIDE_MARGIN_IN,
            _name: props.title,
            _presLayout: this.presLayout,
            _rels: [],
            _relsChart: [],
            _relsMedia: [],
            _slide: null,
            _slideNum: 1000 + this.slideLayouts.length + 1,
            _slideNumberProps: props.slideNumber || null,
            _slideObjects: [],
            background: props.background || null,
            bkgd: props.bkgd || null,
        };
        // STEP 1: Create the Slide Master/Layout
        createSlideMaster(props, newLayout);
        // STEP 2: Add it to layout defs
        this.slideLayouts.push(newLayout);
        // STEP 3: Add background (image data/path must be captured before `exportPresentation()` is called)
        if (props.background || props.bkgd)
            addBackgroundDefinition(props.background, newLayout);
        // STEP 4: Add slideNumber to master slide (if any)
        if (newLayout._slideNumberProps && !this.masterSlide._slideNumberProps)
            this.masterSlide._slideNumberProps = newLayout._slideNumberProps;
    };
    // HTML-TO-SLIDES METHODS
    /**
     * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
     * @param {string} eleId - table HTML element ID
     * @param {TableToSlidesProps} options - generation options
     */
    PptxGenJS.prototype.tableToSlides = function (eleId, options) {
        if (options === void 0) { options = {}; }
        // @note `verbose` option is undocumented; used for verbose output of layout process
        genTableToSlides(this, eleId, options, options && options.masterSlideName ? this.slideLayouts.filter(function (layout) { return layout._name === options.masterSlideName; })[0] : null);
    };
    return PptxGenJS;
}());

export default PptxGenJS;
