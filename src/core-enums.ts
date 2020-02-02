/**
 * PptxGenJS Enums
 * NOTE: `enum` wont work for objects, so use `Object.freeze`
 */

import { IBorderOptions } from './core-interfaces'

// CONST
export const EMU: number = 914400 // One (1) inch (OfficeXML measures in EMU (English Metric Units))
export const ONEPT: number = 12700 // One (1) point (pt)
export const CRLF: string = '\r\n' // AKA: Chr(13) & Chr(10)
export const LAYOUT_IDX_SERIES_BASE: number = 2147483649
export const REGEX_HEX_COLOR: RegExp = /^[0-9a-fA-F]{6}$/
export const LINEH_MODIFIER = 1.67 // AKA: Golden Ratio Typography

export const DEF_CELL_BORDER: IBorderOptions = { color: '666666' }
export const DEF_CELL_MARGIN_PT: [number, number, number, number] = [3, 3, 3, 3] // TRBL-style
export const DEF_CHART_GRIDLINE = { color: '888888', style: 'solid', size: 1 }
export const DEF_FONT_COLOR: string = '000000'
export const DEF_FONT_SIZE: number = 12
export const DEF_FONT_TITLE_SIZE: number = 18
export const DEF_PRES_LAYOUT = 'LAYOUT_16x9'
export const DEF_PRES_LAYOUT_NAME = 'DEFAULT'
export const DEF_SLIDE_BKGD = 'FFFFFF'
export const DEF_SLIDE_MARGIN_IN: [number, number, number, number] = [0.5, 0.5, 0.5, 0.5] // TRBL-style
export const DEF_SHAPE_SHADOW = { type: 'outer', blur: 3, offset: 23000 / 12700, angle: 90, color: '000000', opacity: 0.35, rotateWithShape: true }
export const DEF_TEXT_SHADOW = { type: 'outer', blur: 8, offset: 4, angle: 270, color: '000000', opacity: 0.75 }
export const DEF_TEXT_GLOW = { size: 8, color: 'FFFFFF', opacity: 0.75 }

export const AXIS_ID_VALUE_PRIMARY: string = '2094734552'
export const AXIS_ID_VALUE_SECONDARY: string = '2094734553'
export const AXIS_ID_CATEGORY_PRIMARY: string = '2094734554'
export const AXIS_ID_CATEGORY_SECONDARY: string = '2094734555'
export const AXIS_ID_SERIES_PRIMARY: string = '2094734556'

export type JSZIP_OUTPUT_TYPE = 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
export type WRITE_OUTPUT_TYPE = JSZIP_OUTPUT_TYPE | 'STREAM'
export type CHART_NAME = 'area' | 'bar' | 'bar3D' | 'bubble' | 'doughnut' | 'line' | 'pie' | 'radar' | 'scatter'

export const LETTERS: Array<string> = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')
export const BARCHART_COLORS: Array<string> = [
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
]
export const PIECHART_COLORS: Array<string> = [
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
]
export enum TEXT_HALIGN {
	'left' = 'left',
	'center' = 'center',
	'right' = 'right',
	'justify' = 'justify',
}
export enum TEXT_VALIGN {
	'b' = 'b',
	'ctr' = 'ctr',
	't' = 't',
}

export const SLDNUMFLDID: string = '{F7021451-1387-4CA6-816F-3879F97B5CBC}'

// ENUM
// TODO: 3.5 or v4.0: rationalize ts-def exported enum names/case!
// NOTE: First tsdef enum named correctly (shapes -> 'Shape', colors -> 'Color'), etc.
export enum OutputType {
	'arraybuffer' = 'arraybuffer',
	'base64' = 'base64',
	'binarystring' = 'binarystring',
	'blob' = 'blob',
	'nodebuffer' = 'nodebuffer',
	'uint8array' = 'uint8array',
}
export enum ChartType {
	'area' = 'area',
	'bar' = 'bar',
	'bar3d' = 'bar3D',
	'bubble' = 'bubble',
	'doughnut' = 'doughnut',
	'line' = 'line',
	'pie' = 'pie',
	'radar' = 'radar',
	'scatter' = 'scatter',
}
export enum ShapeType {
	'accentBorderCallout1' = 'accentBorderCallout1',
	'accentBorderCallout2' = 'accentBorderCallout2',
	'accentBorderCallout3' = 'accentBorderCallout3',
	'accentCallout1' = 'accentCallout1',
	'accentCallout2' = 'accentCallout2',
	'accentCallout3' = 'accentCallout3',
	'actionButtonBackPrevious' = 'actionButtonBackPrevious',
	'actionButtonBeginning' = 'actionButtonBeginning',
	'actionButtonBlank' = 'actionButtonBlank',
	'actionButtonDocument' = 'actionButtonDocument',
	'actionButtonEnd' = 'actionButtonEnd',
	'actionButtonForwardNext' = 'actionButtonForwardNext',
	'actionButtonHelp' = 'actionButtonHelp',
	'actionButtonHome' = 'actionButtonHome',
	'actionButtonInformation' = 'actionButtonInformation',
	'actionButtonMovie' = 'actionButtonMovie',
	'actionButtonReturn' = 'actionButtonReturn',
	'actionButtonSound' = 'actionButtonSound',
	'arc' = 'arc',
	'bentArrow' = 'bentArrow',
	'bentUpArrow' = 'bentUpArrow',
	'bevel' = 'bevel',
	'blockArc' = 'blockArc',
	'borderCallout1' = 'borderCallout1',
	'borderCallout2' = 'borderCallout2',
	'borderCallout3' = 'borderCallout3',
	'bracePair' = 'bracePair',
	'bracketPair' = 'bracketPair',
	'callout1' = 'callout1',
	'callout2' = 'callout2',
	'callout3' = 'callout3',
	'can' = 'can',
	'chartPlus' = 'chartPlus',
	'chartStar' = 'chartStar',
	'chartX' = 'chartX',
	'chevron' = 'chevron',
	'chord' = 'chord',
	'circularArrow' = 'circularArrow',
	'cloud' = 'cloud',
	'cloudCallout' = 'cloudCallout',
	'corner' = 'corner',
	'cornerTabs' = 'cornerTabs',
	'cube' = 'cube',
	'curvedDownArrow' = 'curvedDownArrow',
	'curvedLeftArrow' = 'curvedLeftArrow',
	'curvedRightArrow' = 'curvedRightArrow',
	'curvedUpArrow' = 'curvedUpArrow',
	'decagon' = 'decagon',
	'diagStripe' = 'diagStripe',
	'diamond' = 'diamond',
	'dodecagon' = 'dodecagon',
	'donut' = 'donut',
	'doubleWave' = 'doubleWave',
	'downArrow' = 'downArrow',
	'downArrowCallout' = 'downArrowCallout',
	'ellipse' = 'ellipse',
	'ellipseRibbon' = 'ellipseRibbon',
	'ellipseRibbon2' = 'ellipseRibbon2',
	'flowChartAlternateProcess' = 'flowChartAlternateProcess',
	'flowChartCollate' = 'flowChartCollate',
	'flowChartConnector' = 'flowChartConnector',
	'flowChartDecision' = 'flowChartDecision',
	'flowChartDelay' = 'flowChartDelay',
	'flowChartDisplay' = 'flowChartDisplay',
	'flowChartDocument' = 'flowChartDocument',
	'flowChartExtract' = 'flowChartExtract',
	'flowChartInputOutput' = 'flowChartInputOutput',
	'flowChartInternalStorage' = 'flowChartInternalStorage',
	'flowChartMagneticDisk' = 'flowChartMagneticDisk',
	'flowChartMagneticDrum' = 'flowChartMagneticDrum',
	'flowChartMagneticTape' = 'flowChartMagneticTape',
	'flowChartManualInput' = 'flowChartManualInput',
	'flowChartManualOperation' = 'flowChartManualOperation',
	'flowChartMerge' = 'flowChartMerge',
	'flowChartMultidocument' = 'flowChartMultidocument',
	'flowChartOfflineStorage' = 'flowChartOfflineStorage',
	'flowChartOffpageConnector' = 'flowChartOffpageConnector',
	'flowChartOnlineStorage' = 'flowChartOnlineStorage',
	'flowChartOr' = 'flowChartOr',
	'flowChartPredefinedProcess' = 'flowChartPredefinedProcess',
	'flowChartPreparation' = 'flowChartPreparation',
	'flowChartProcess' = 'flowChartProcess',
	'flowChartPunchedCard' = 'flowChartPunchedCard',
	'flowChartPunchedTape' = 'flowChartPunchedTape',
	'flowChartSort' = 'flowChartSort',
	'flowChartSummingJunction' = 'flowChartSummingJunction',
	'flowChartTerminator' = 'flowChartTerminator',
	'folderCorner' = 'folderCorner',
	'frame' = 'frame',
	'funnel' = 'funnel',
	'gear6' = 'gear6',
	'gear9' = 'gear9',
	'halfFrame' = 'halfFrame',
	'heart' = 'heart',
	'heptagon' = 'heptagon',
	'hexagon' = 'hexagon',
	'homePlate' = 'homePlate',
	'horizontalScroll' = 'horizontalScroll',
	'irregularSeal1' = 'irregularSeal1',
	'irregularSeal2' = 'irregularSeal2',
	'leftArrow' = 'leftArrow',
	'leftArrowCallout' = 'leftArrowCallout',
	'leftBrace' = 'leftBrace',
	'leftBracket' = 'leftBracket',
	'leftCircularArrow' = 'leftCircularArrow',
	'leftRightArrow' = 'leftRightArrow',
	'leftRightArrowCallout' = 'leftRightArrowCallout',
	'leftRightCircularArrow' = 'leftRightCircularArrow',
	'leftRightRibbon' = 'leftRightRibbon',
	'leftRightUpArrow' = 'leftRightUpArrow',
	'leftUpArrow' = 'leftUpArrow',
	'lightningBolt' = 'lightningBolt',
	'line' = 'line',
	'lineInv' = 'lineInv',
	'mathDivide' = 'mathDivide',
	'mathEqual' = 'mathEqual',
	'mathMinus' = 'mathMinus',
	'mathMultiply' = 'mathMultiply',
	'mathNotEqual' = 'mathNotEqual',
	'mathPlus' = 'mathPlus',
	'moon' = 'moon',
	'nonIsoscelesTrapezoid' = 'nonIsoscelesTrapezoid',
	'noSmoking' = 'noSmoking',
	'notchedRightArrow' = 'notchedRightArrow',
	'octagon' = 'octagon',
	'parallelogram' = 'parallelogram',
	'pentagon' = 'pentagon',
	'pie' = 'pie',
	'pieWedge' = 'pieWedge',
	'plaque' = 'plaque',
	'plaqueTabs' = 'plaqueTabs',
	'plus' = 'plus',
	'quadArrow' = 'quadArrow',
	'quadArrowCallout' = 'quadArrowCallout',
	'rect' = 'rect',
	'ribbon' = 'ribbon',
	'ribbon2' = 'ribbon2',
	'rightArrow' = 'rightArrow',
	'rightArrowCallout' = 'rightArrowCallout',
	'rightBrace' = 'rightBrace',
	'rightBracket' = 'rightBracket',
	'round1Rect' = 'round1Rect',
	'round2DiagRect' = 'round2DiagRect',
	'round2SameRect' = 'round2SameRect',
	'roundRect' = 'roundRect',
	'rtTriangle' = 'rtTriangle',
	'smileyFace' = 'smileyFace',
	'snip1Rect' = 'snip1Rect',
	'snip2DiagRect' = 'snip2DiagRect',
	'snip2SameRect' = 'snip2SameRect',
	'snipRoundRect' = 'snipRoundRect',
	'squareTabs' = 'squareTabs',
	'star10' = 'star10',
	'star12' = 'star12',
	'star16' = 'star16',
	'star24' = 'star24',
	'star32' = 'star32',
	'star4' = 'star4',
	'star5' = 'star5',
	'star6' = 'star6',
	'star7' = 'star7',
	'star8' = 'star8',
	'stripedRightArrow' = 'stripedRightArrow',
	'sun' = 'sun',
	'swooshArrow' = 'swooshArrow',
	'teardrop' = 'teardrop',
	'trapezoid' = 'trapezoid',
	'triangle' = 'triangle',
	'upArrow' = 'upArrow',
	'upArrowCallout' = 'upArrowCallout',
	'upDownArrow' = 'upDownArrow',
	'upDownArrowCallout' = 'upDownArrowCallout',
	'uturnArrow' = 'uturnArrow',
	'verticalScroll' = 'verticalScroll',
	'wave' = 'wave',
	'wedgeEllipseCallout' = 'wedgeEllipseCallout',
	'wedgeRectCallout' = 'wedgeRectCallout',
	'wedgeRoundRectCallout' = 'wedgeRoundRectCallout',
}
export enum SchemeColor {
	'text1' = 'tx1',
	'text2' = 'tx2',
	'background1' = 'bg1',
	'background2' = 'bg2',
	'accent1' = 'accent1',
	'accent2' = 'accent2',
	'accent3' = 'accent3',
	'accent4' = 'accent4',
	'accent5' = 'accent5',
	'accent6' = 'accent6',
}
export enum AlignH {
	'left' = 'left',
	'center' = 'center',
	'right' = 'right',
	'justify' = 'justify',
}
export enum AlignV {
	'top' = 'top',
	'middle' = 'middle',
	'bottom' = 'bottom',
}

export enum SHAPE_TYPE {
	ACTION_BUTTON_BACK_OR_PREVIOUS = 'actionButtonBackPrevious',
	ACTION_BUTTON_BEGINNING = 'actionButtonBeginning',
	ACTION_BUTTON_CUSTOM = 'actionButtonBlank',
	ACTION_BUTTON_DOCUMENT = 'actionButtonDocument',
	ACTION_BUTTON_END = 'actionButtonEnd',
	ACTION_BUTTON_FORWARD_OR_NEXT = 'actionButtonForwardNext',
	ACTION_BUTTON_HELP = 'actionButtonHelp',
	ACTION_BUTTON_HOME = 'actionButtonHome',
	ACTION_BUTTON_INFORMATION = 'actionButtonInformation',
	ACTION_BUTTON_MOVIE = 'actionButtonMovie',
	ACTION_BUTTON_RETURN = 'actionButtonReturn',
	ACTION_BUTTON_SOUND = 'actionButtonSound',
	ARC = 'arc',
	BALLOON = 'wedgeRoundRectCallout',
	BENT_ARROW = 'bentArrow',
	BENT_UP_ARROW = 'bentUpArrow',
	BEVEL = 'bevel',
	BLOCK_ARC = 'blockArc',
	CAN = 'can',
	CHART_PLUS = 'chartPlus',
	CHART_STAR = 'chartStar',
	CHART_X = 'chartX',
	CHEVRON = 'chevron',
	CHORD = 'chord',
	CIRCULAR_ARROW = 'circularArrow',
	CLOUD = 'cloud',
	CLOUD_CALLOUT = 'cloudCallout',
	CORNER = 'corner',
	CORNER_TABS = 'cornerTabs',
	CROSS = 'plus',
	CUBE = 'cube',
	CURVED_DOWN_ARROW = 'curvedDownArrow',
	CURVED_DOWN_RIBBON = 'ellipseRibbon',
	CURVED_LEFT_ARROW = 'curvedLeftArrow',
	CURVED_RIGHT_ARROW = 'curvedRightArrow',
	CURVED_UP_ARROW = 'curvedUpArrow',
	CURVED_UP_RIBBON = 'ellipseRibbon2',
	DECAGON = 'decagon',
	DIAGONAL_STRIPE = 'diagStripe',
	DIAMOND = 'diamond',
	DODECAGON = 'dodecagon',
	DONUT = 'donut',
	DOUBLE_BRACE = 'bracePair',
	DOUBLE_BRACKET = 'bracketPair',
	DOUBLE_WAVE = 'doubleWave',
	DOWN_ARROW = 'downArrow',
	DOWN_ARROW_CALLOUT = 'downArrowCallout',
	DOWN_RIBBON = 'ribbon',
	EXPLOSION1 = 'irregularSeal1',
	EXPLOSION2 = 'irregularSeal2',
	FLOWCHART_ALTERNATE_PROCESS = 'flowChartAlternateProcess',
	FLOWCHART_CARD = 'flowChartPunchedCard',
	FLOWCHART_COLLATE = 'flowChartCollate',
	FLOWCHART_CONNECTOR = 'flowChartConnector',
	FLOWCHART_DATA = 'flowChartInputOutput',
	FLOWCHART_DECISION = 'flowChartDecision',
	FLOWCHART_DELAY = 'flowChartDelay',
	FLOWCHART_DIRECT_ACCESS_STORAGE = 'flowChartMagneticDrum',
	FLOWCHART_DISPLAY = 'flowChartDisplay',
	FLOWCHART_DOCUMENT = 'flowChartDocument',
	FLOWCHART_EXTRACT = 'flowChartExtract',
	FLOWCHART_INTERNAL_STORAGE = 'flowChartInternalStorage',
	FLOWCHART_MAGNETIC_DISK = 'flowChartMagneticDisk',
	FLOWCHART_MANUAL_INPUT = 'flowChartManualInput',
	FLOWCHART_MANUAL_OPERATION = 'flowChartManualOperation',
	FLOWCHART_MERGE = 'flowChartMerge',
	FLOWCHART_MULTIDOCUMENT = 'flowChartMultidocument',
	FLOWCHART_OFFLINE_STORAGE = 'flowChartOfflineStorage',
	FLOWCHART_OFFPAGE_CONNECTOR = 'flowChartOffpageConnector',
	FLOWCHART_OR = 'flowChartOr',
	FLOWCHART_PREDEFINED_PROCESS = 'flowChartPredefinedProcess',
	FLOWCHART_PREPARATION = 'flowChartPreparation',
	FLOWCHART_PROCESS = 'flowChartProcess',
	FLOWCHART_PUNCHED_TAPE = 'flowChartPunchedTape',
	FLOWCHART_SEQUENTIAL_ACCESS_STORAGE = 'flowChartMagneticTape',
	FLOWCHART_SORT = 'flowChartSort',
	FLOWCHART_STORED_DATA = 'flowChartOnlineStorage',
	FLOWCHART_SUMMING_JUNCTION = 'flowChartSummingJunction',
	FLOWCHART_TERMINATOR = 'flowChartTerminator',
	FOLDED_CORNER = 'folderCorner',
	FRAME = 'frame',
	FUNNEL = 'funnel',
	GEAR_6 = 'gear6',
	GEAR_9 = 'gear9',
	HALF_FRAME = 'halfFrame',
	HEART = 'heart',
	HEPTAGON = 'heptagon',
	HEXAGON = 'hexagon',
	HORIZONTAL_SCROLL = 'horizontalScroll',
	ISOSCELES_TRIANGLE = 'triangle',
	LEFT_ARROW = 'leftArrow',
	LEFT_ARROW_CALLOUT = 'leftArrowCallout',
	LEFT_BRACE = 'leftBrace',
	LEFT_BRACKET = 'leftBracket',
	LEFT_CIRCULAR_ARROW = 'leftCircularArrow',
	LEFT_RIGHT_ARROW = 'leftRightArrow',
	LEFT_RIGHT_ARROW_CALLOUT = 'leftRightArrowCallout',
	LEFT_RIGHT_CIRCULAR_ARROW = 'leftRightCircularArrow',
	LEFT_RIGHT_RIBBON = 'leftRightRibbon',
	LEFT_RIGHT_UP_ARROW = 'leftRightUpArrow',
	LEFT_UP_ARROW = 'leftUpArrow',
	LIGHTNING_BOLT = 'lightningBolt',
	LINE_CALLOUT_1 = 'borderCallout1',
	LINE_CALLOUT_1_ACCENT_BAR = 'accentCallout1',
	LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR = 'accentBorderCallout1',
	LINE_CALLOUT_1_NO_BORDER = 'callout1',
	LINE_CALLOUT_2 = 'borderCallout2',
	LINE_CALLOUT_2_ACCENT_BAR = 'accentCallout2',
	LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR = 'accentBorderCallout2',
	LINE_CALLOUT_2_NO_BORDER = 'callout2',
	LINE_CALLOUT_3 = 'borderCallout3',
	LINE_CALLOUT_3_ACCENT_BAR = 'accentCallout3',
	LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR = 'accentBorderCallout3',
	LINE_CALLOUT_3_NO_BORDER = 'callout3',
	LINE_CALLOUT_4 = 'borderCallout3',
	LINE_CALLOUT_4_ACCENT_BAR = 'accentCallout3',
	LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR = 'accentBorderCallout3',
	LINE_CALLOUT_4_NO_BORDER = 'callout3',
	LINE = 'line',
	LINE_INVERSE = 'lineInv',
	MATH_DIVIDE = 'mathDivide',
	MATH_EQUAL = 'mathEqual',
	MATH_MINUS = 'mathMinus',
	MATH_MULTIPLY = 'mathMultiply',
	MATH_NOT_EQUAL = 'mathNotEqual',
	MATH_PLUS = 'mathPlus',
	MOON = 'moon',
	NON_ISOSCELES_TRAPEZOID = 'nonIsoscelesTrapezoid',
	NOTCHED_RIGHT_ARROW = 'notchedRightArrow',
	NO_SYMBOL = 'noSmoking',
	OCTAGON = 'octagon',
	OVAL = 'ellipse',
	OVAL_CALLOUT = 'wedgeEllipseCallout',
	PARALLELOGRAM = 'parallelogram',
	PENTAGON = 'homePlate',
	PIE = 'pie',
	PIE_WEDGE = 'pieWedge',
	PLAQUE = 'plaque',
	PLAQUE_TABS = 'plaqueTabs',
	QUAD_ARROW = 'quadArrow',
	QUAD_ARROW_CALLOUT = 'quadArrowCallout',
	RECTANGLE = 'rect',
	RECTANGULAR_CALLOUT = 'wedgeRectCallout',
	REGULAR_PENTAGON = 'pentagon',
	RIGHT_ARROW = 'rightArrow',
	RIGHT_ARROW_CALLOUT = 'rightArrowCallout',
	RIGHT_BRACE = 'rightBrace',
	RIGHT_BRACKET = 'rightBracket',
	RIGHT_TRIANGLE = 'rtTriangle',
	ROUNDED_RECTANGLE = 'roundRect',
	ROUNDED_RECTANGULAR_CALLOUT = 'wedgeRoundRectCallout',
	ROUND_1_RECTANGLE = 'round1Rect',
	ROUND_2_DIAG_RECTANGLE = 'round2DiagRect',
	ROUND_2_SAME_RECTANGLE = 'round2SameRect',
	SMILEY_FACE = 'smileyFace',
	SNIP_1_RECTANGLE = 'snip1Rect',
	SNIP_2_DIAG_RECTANGLE = 'snip2DiagRect',
	SNIP_2_SAME_RECTANGLE = 'snip2SameRect',
	SNIP_ROUND_RECTANGLE = 'snipRoundRect',
	SQUARE_TABS = 'squareTabs',
	STAR_10_POINT = 'star10',
	STAR_12_POINT = 'star12',
	STAR_16_POINT = 'star16',
	STAR_24_POINT = 'star24',
	STAR_32_POINT = 'star32',
	STAR_4_POINT = 'star4',
	STAR_5_POINT = 'star5',
	STAR_6_POINT = 'star6',
	STAR_7_POINT = 'star7',
	STAR_8_POINT = 'star8',
	STRIPED_RIGHT_ARROW = 'stripedRightArrow',
	SUN = 'sun',
	SWOOSH_ARROW = 'swooshArrow',
	TEAR = 'teardrop',
	TRAPEZOID = 'trapezoid',
	UP_ARROW = 'upArrow',
	UP_ARROW_CALLOUT = 'upArrowCallout',
	UP_DOWN_ARROW = 'upDownArrow',
	UP_DOWN_ARROW_CALLOUT = 'upDownArrowCallout',
	UP_RIBBON = 'ribbon2',
	U_TURN_ARROW = 'uturnArrow',
	VERTICAL_SCROLL = 'verticalScroll',
	WAVE = 'wave',
}
export type SHAPE_NAME =
	| 'actionButtonBackPrevious'
	| 'actionButtonBeginning'
	| 'actionButtonBlank'
	| 'actionButtonDocument'
	| 'actionButtonEnd'
	| 'actionButtonForwardNext'
	| 'actionButtonHelp'
	| 'actionButtonHome'
	| 'actionButtonInformation'
	| 'actionButtonMovie'
	| 'actionButtonReturn'
	| 'actionButtonSound'
	| 'arc'
	| 'wedgeRoundRectCallout'
	| 'bentArrow'
	| 'bentUpArrow'
	| 'bevel'
	| 'blockArc'
	| 'can'
	| 'chartPlus'
	| 'chartStar'
	| 'chartX'
	| 'chevron'
	| 'chord'
	| 'circularArrow'
	| 'cloud'
	| 'cloudCallout'
	| 'corner'
	| 'cornerTabs'
	| 'plus'
	| 'cube'
	| 'curvedDownArrow'
	| 'ellipseRibbon'
	| 'curvedLeftArrow'
	| 'curvedRightArrow'
	| 'curvedUpArrow'
	| 'ellipseRibbon2'
	| 'decagon'
	| 'diagStripe'
	| 'diamond'
	| 'dodecagon'
	| 'donut'
	| 'bracePair'
	| 'bracketPair'
	| 'doubleWave'
	| 'downArrow'
	| 'downArrowCallout'
	| 'ribbon'
	| 'irregularSeal1'
	| 'irregularSeal2'
	| 'flowChartAlternateProcess'
	| 'flowChartPunchedCard'
	| 'flowChartCollate'
	| 'flowChartConnector'
	| 'flowChartInputOutput'
	| 'flowChartDecision'
	| 'flowChartDelay'
	| 'flowChartMagneticDrum'
	| 'flowChartDisplay'
	| 'flowChartDocument'
	| 'flowChartExtract'
	| 'flowChartInternalStorage'
	| 'flowChartMagneticDisk'
	| 'flowChartManualInput'
	| 'flowChartManualOperation'
	| 'flowChartMerge'
	| 'flowChartMultidocument'
	| 'flowChartOfflineStorage'
	| 'flowChartOffpageConnector'
	| 'flowChartOr'
	| 'flowChartPredefinedProcess'
	| 'flowChartPreparation'
	| 'flowChartProcess'
	| 'flowChartPunchedTape'
	| 'flowChartMagneticTape'
	| 'flowChartSort'
	| 'flowChartOnlineStorage'
	| 'flowChartSummingJunction'
	| 'flowChartTerminator'
	| 'folderCorner'
	| 'frame'
	| 'funnel'
	| 'gear6'
	| 'gear9'
	| 'halfFrame'
	| 'heart'
	| 'heptagon'
	| 'hexagon'
	| 'horizontalScroll'
	| 'triangle'
	| 'leftArrow'
	| 'leftArrowCallout'
	| 'leftBrace'
	| 'leftBracket'
	| 'leftCircularArrow'
	| 'leftRightArrow'
	| 'leftRightArrowCallout'
	| 'leftRightCircularArrow'
	| 'leftRightRibbon'
	| 'leftRightUpArrow'
	| 'leftUpArrow'
	| 'lightningBolt'
	| 'borderCallout1'
	| 'accentCallout1'
	| 'accentBorderCallout1'
	| 'callout1'
	| 'borderCallout2'
	| 'accentCallout2'
	| 'accentBorderCallout2'
	| 'callout2'
	| 'borderCallout3'
	| 'accentCallout3'
	| 'accentBorderCallout3'
	| 'callout3'
	| 'borderCallout3'
	| 'accentCallout3'
	| 'accentBorderCallout3'
	| 'callout3'
	| 'line'
	| 'lineInv'
	| 'mathDivide'
	| 'mathEqual'
	| 'mathMinus'
	| 'mathMultiply'
	| 'mathNotEqual'
	| 'mathPlus'
	| 'moon'
	| 'nonIsoscelesTrapezoid'
	| 'notchedRightArrow'
	| 'noSmoking'
	| 'octagon'
	| 'ellipse'
	| 'wedgeEllipseCallout'
	| 'parallelogram'
	| 'homePlate'
	| 'pie'
	| 'pieWedge'
	| 'plaque'
	| 'plaqueTabs'
	| 'quadArrow'
	| 'quadArrowCallout'
	| 'rect'
	| 'wedgeRectCallout'
	| 'pentagon'
	| 'rightArrow'
	| 'rightArrowCallout'
	| 'rightBrace'
	| 'rightBracket'
	| 'rtTriangle'
	| 'roundRect'
	| 'wedgeRoundRectCallout'
	| 'round1Rect'
	| 'round2DiagRect'
	| 'round2SameRect'
	| 'smileyFace'
	| 'snip1Rect'
	| 'snip2DiagRect'
	| 'snip2SameRect'
	| 'snipRoundRect'
	| 'squareTabs'
	| 'star10'
	| 'star12'
	| 'star16'
	| 'star24'
	| 'star32'
	| 'star4'
	| 'star5'
	| 'star6'
	| 'star7'
	| 'star8'
	| 'stripedRightArrow'
	| 'sun'
	| 'swooshArrow'
	| 'teardrop'
	| 'trapezoid'
	| 'upArrow'
	| 'upArrowCallout'
	| 'upDownArrow'
	| 'upDownArrowCallout'
	| 'ribbon2'
	| 'uturnArrow'
	| 'verticalScroll'
	| 'wave'

export enum CHART_TYPE {
	'AREA' = 'area',
	'BAR' = 'bar',
	'BAR3D' = 'bar3D',
	'BUBBLE' = 'bubble',
	'DOUGHNUT' = 'doughnut',
	'LINE' = 'line',
	'PIE' = 'pie',
	'RADAR' = 'radar',
	'SCATTER' = 'scatter',
}

export enum SCHEME_COLOR_NAMES {
	'TEXT1' = 'tx1',
	'TEXT2' = 'tx2',
	'BACKGROUND1' = 'bg1',
	'BACKGROUND2' = 'bg2',
	'ACCENT1' = 'accent1',
	'ACCENT2' = 'accent2',
	'ACCENT3' = 'accent3',
	'ACCENT4' = 'accent4',
	'ACCENT5' = 'accent5',
	'ACCENT6' = 'accent6',
}

export enum MASTER_OBJECTS {
	'chart' = 'chart',
	'image' = 'image',
	'line' = 'line',
	'rect' = 'rect',
	'text' = 'text',
	'placeholder' = 'placeholder',
}

export enum SLIDE_OBJECT_TYPES {
	'chart' = 'chart',
	'hyperlink' = 'hyperlink',
	'image' = 'image',
	'media' = 'media',
	'online' = 'online',
	'placeholder' = 'placeholder',
	'table' = 'table',
	'tablecell' = 'tablecell',
	'text' = 'text',
	'notes' = 'notes',
}
export enum PLACEHOLDER_TYPES {
	'title' = 'title',
	'body' = 'body',
	'image' = 'pic',
	'chart' = 'chart',
	'table' = 'tbl',
	'media' = 'media',
}

/**
 * NOTE: 20170304: BULLET_TYPES: Only default is used so far. I'd like to combine the two pieces of code that use these before implementing these as options
 * Since we close <p> within the text object bullets, its slightly more difficult than combining into a func and calling to get the paraProp
 * and i'm not sure if anyone will even use these... so, skipping for now.
 */
export enum BULLET_TYPES {
	'DEFAULT' = '&#x2022;',
	'CHECK' = '&#x2713;',
	'STAR' = '&#x2605;',
	'TRIANGLE' = '&#x25B6;',
}

// IMAGES (base64)
export const IMG_BROKEN =
	'data:image/png;base64,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'

export const IMG_PLAYBTN =
	'data:image/png;base64,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'
