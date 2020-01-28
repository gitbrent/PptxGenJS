// Type definitions for pptxgenjs 3.1.0
// Project: https://gitbrent.github.io/PptxGenJS/
// Definitions by: Brent Ely <https://github.com/gitbrent/>
//                 Michael Beaumont <https://github.com/michaelbeaumont>
//                 Nicholas Tietz-Sokolsky <https://github.com/ntietz>
//                 David Adams <https://github.com/iota-pi>
//                 Stephen Cronin <https://github.com/cronin4392>
// TypeScript Version: 3.x

export as namespace PptxGenJS

export = PptxGenJS

declare class PptxGenJS {
	/**
	 * PptxGenJS Library Version
	 */
	readonly version: string

	// Presentation Props

	/**
	 * Presentation layout name
	 * Standard layouts:
	 * - 'LAYOUT_4x3'   (10" x 7.5")
	 * - 'LAYOUT_16x9'  (10" x 5.625")
	 * - 'LAYOUT_16x10' (10" x 6.25")
	 * - 'LAYOUT_WIDE'  (13.33" x 7.5")
	 * Custom layouts:
	 * Use `pptx.defineLayout()` to create custom layouts (e.g.: 'A4')
	 * @type {string}
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 */
	layout: string
	/**
	 * Whether Right-to-Left (RTL) mode is enabled
	 */
	rtlMode: boolean

	// Presentation Metadata
	author: string
	company: string
	/**
	 * @type {string}
	 * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PPT will throw errors upon opening!)
	 */
	revision: string
	subject: string
	title: string

	// Methods

	/**
	 * Add a Slide to Presenation
	 * @param {string} masterSlideName - Master Slide name
	 * @returns {ISlide} the new Slide
	 */
	addSlide(masterSlideName?: string): PptxGenJS.Slide
	/**
	 * Define a custom Slide Layout
	 * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 })
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 * @param {IUserLayout} layout - an object with user-defined w/h
	 */
	defineLayout(layout: PptxGenJS.IUserLayout): void
	/**
	 * Adds a new slide master [layout] to the Presentation
	 * @param {ISlideMasterOptions} slideMasterOpts - layout definition
	 */
	defineSlideMaster(slideMasterOpts: PptxGenJS.ISlideMasterOptions): void
	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @note `verbose` option is undocumented used for verbose output of layout process
	 * @param {string} tabEleId - HTMLElementID of the table
	 * @param {ITableToSlidesOpts} inOpts - array of options (e.g.: tabsize)
	 */
	tableToSlides(tableElementId: string, opts?: PptxGenJS.ITableToSlidesOpts): void

	// Export

	/**
	 * Export the current Presenation to stream
	 * @since 3.0.0
	 * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file stream
	 */
	stream(): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>
	/**
	 * Export the current Presenation to selected/default type
	 * @since 3.0.0
	 * @param {JSZIP_OUTPUT_TYPE} outputType - 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
	 * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file
	 */
	write(outputType: PptxGenJS.JSZIP_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>
	/**
	 * Export the current Presenation to local file (initiates download in browsers)
	 * @since 3.0.0
	 * @param {string} exportName - file name
	 * @returns {Promise<string>} file name
	 */
	writeFile(exportName?: string): Promise<string>
}

declare namespace PptxGenJS {
	// JSZIP
	export type JSZIP_OUTPUT_TYPE = 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'

	// `core-interfaces.d.ts`
	// import { CHART_NAME, SLIDE_OBJECT_TYPES, TEXT_HALIGN, TEXT_VALIGN, PLACEHOLDER_TYPES, SHAPE_NAME } from './core-enums'
	export type CHART_NAME = 'area' | 'bar' | 'bar3D' | 'bubble' | 'doughnut' | 'line' | 'pie' | 'radar' | 'scatter'
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
		'notes' = 'notes'
	}
	export enum TEXT_HALIGN {
		'left' = 'left',
		'center' = 'center',
		'right' = 'right',
		'justify' = 'justify'
	}
	export enum TEXT_VALIGN {
		'b' = 'b',
		'ctr' = 'ctr',
		't' = 't'
	}
	export enum PLACEHOLDER_TYPES {
		'title' = 'title',
		'body' = 'body',
		'image' = 'pic',
		'chart' = 'chart',
		'table' = 'tbl',
		'media' = 'media'
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
	// ======

	// charts and shapes for `pptxgen.charts.` `pptxgen.shapes.`
	export enum charts {
		'AREA' = 'area',
		'BAR' = 'bar',
		'BAR3D' = 'bar3D',
		'BUBBLE' = 'bubble',
		'DOUGHNUT' = 'doughnut',
		'LINE' = 'line',
		'PIE' = 'pie',
		'RADAR' = 'radar',
		'SCATTER' = 'scatter'
	}
	export enum shapes {
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
		WAVE = 'wave'
	}

	/**
	 * Coordinate (string is in the form of 'N%')
	 */
	export type HexColor = string
	export type ThemeColor = 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
	export type Color = HexColor | ThemeColor
	export type Coord = number | string
	export type Margin = number | [number, number, number, number]
	export type HAlign = 'left' | 'center' | 'right' | 'justify'
	export type VAlign = 'top' | 'middle' | 'bottom'
	export type ChartAxisTickMark = 'none' | 'inside' | 'outside' | 'cross'
	export type HyperLink = {
		rId: number
		slide?: number
		tooltip?: string
		url?: string
	}
	export type ShapeFill =
		| Color
		| {
				type: string
				color: Color
				alpha?: number
		  }
	export type BkgdOpts = {
		src?: string
		path?: string
		data?: string
	}
	type MediaType = 'audio' | 'online' | 'video'

	export interface FontOptions {
		fontFace?: string
		fontSize?: number
	}
	export interface PositionOptions {
		x?: Coord
		y?: Coord
		w?: Coord
		h?: Coord
	}
	export interface OptsDataOrPath {
		data?: string
		path?: string
	}
	export interface OptsChartData {
		index?: number
		name?: string
		labels?: string[]
		values?: number[]
		sizes?: number[]
	}
	export interface OptsChartGridLine {
		size?: number
		color?: string
		style?: 'solid' | 'dash' | 'dot' | 'none'
	}
	export interface IBorderOptions {
		color?: HexColor
		pt?: number
		type?: string
	}
	export interface IShadowOptions {
		type: 'outer' | 'inner' | 'none'
		angle: number
		opacity: number
		blur?: number
		offset?: number
		color?: string
	}
	export interface IGlowOptions {
		size: number
		opacity: number
		color?: string
	}
	export interface IChartOpts extends PositionOptions, OptsChartGridLine {
		_type?: CHART_NAME | IChartMulti[]
		axisPos?: string
		bar3DShape?: string
		barDir?: string
		barGapDepthPct?: number
		barGapWidthPct?: number
		barGrouping?: string
		border?: IBorderOptions
		catAxes?: number[]
		catAxisBaseTimeUnit?: string
		catAxisHidden?: boolean
		catAxisLabelColor?: string
		catAxisLabelFontBold?: boolean
		catAxisLabelFontFace?: string
		catAxisLabelFontSize?: number
		catAxisLabelFrequency?: string
		catAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		catAxisLabelRotate?: number
		catAxisLineShow?: boolean
		catAxisMajorTickMark?: ChartAxisTickMark
		catAxisMajorTimeUnit?: string
		catAxisMajorUnit?: number
		catAxisMaxVal?: number
		catAxisMinorTickMark?: ChartAxisTickMark
		catAxisMinorTimeUnit?: string
		catAxisMinorUnit?: string
		catAxisMinVal?: number
		catAxisOrientation?: 'minMax' | 'minMax'
		catAxisTitle?: string
		catAxisTitleColor?: string
		catAxisTitleFontFace?: string
		catAxisTitleFontSize?: number
		catAxisTitleRotate?: number
		catGridLine?: OptsChartGridLine
		catLabelFormatCode?: string
		chartColors?: string[]
		chartColorsOpacity?: number
		dataBorder?: IBorderOptions
		dataLabelBkgrdColors?: boolean
		dataLabelColor?: string
		dataLabelFontBold?: boolean
		dataLabelFontFace?: string
		dataLabelFontSize?: number
		dataLabelFormatCode?: string
		dataLabelFormatScatter?: 'custom' | 'customXY' | 'XY'
		dataLabelPosition?: 'b' | 'bestFit' | 'ctr' | 'l' | 'r' | 't' | 'inEnd' | 'outEnd' | 'bestFit'
		dataNoEffects?: string
		dataTableFontSize?: number
		displayBlanksAs?: string
		fill?: string
		hasArea?: boolean
		holeSize?: number
		invertedColors?: string
		lang?: string
		layout?: PositionOptions
		legendColor?: string
		legendFontFace?: string
		legendFontSize?: number
		legendPos?: string
		lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'
		lineDataSymbol?: 'circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'square' | 'triangle'
		lineDataSymbolLineColor?: string
		lineDataSymbolLineSize?: number
		lineDataSymbolSize?: number
		lineSize?: number
		lineSmooth?: boolean
		radarStyle?: 'standard' | 'marker' | 'filled'
		serAxisBaseTimeUnit?: string
		serAxisHidden?: boolean
		serAxisLabelColor?: string
		serAxisLabelFontFace?: string
		serAxisLabelFontSize?: string
		serAxisLabelFrequency?: string
		serAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		serAxisLineShow?: boolean
		serAxisMajorTimeUnit?: string
		serAxisMajorUnit?: number
		serAxisMinorTimeUnit?: string
		serAxisMinorUnit?: number
		serAxisOrientation?: string
		serAxisTitle?: string
		serAxisTitleColor?: string
		serAxisTitleFontFace?: string
		serAxisTitleFontSize?: number
		serAxisTitleRotate?: number
		serGridLine?: OptsChartGridLine
		serLabelFormatCode?: string
		shadow?: IShadowOptions
		showCatAxisTitle?: boolean
		showDataTable?: boolean
		showDataTableHorzBorder?: boolean
		showDataTableKeys?: boolean
		showDataTableOutline?: boolean
		showDataTableVertBorder?: boolean
		showLabel?: boolean
		showLeaderLines?: boolean
		showLegend?: boolean
		showPercent?: boolean
		showSerAxisTitle?: boolean
		showTitle?: boolean
		showValAxisTitle?: boolean
		showValue?: boolean
		title?: string
		titleAlign?: string
		titleColor?: string
		titleFontFace?: string
		titleFontSize?: number
		titlePos?: {
			x: number
			y: number
		}
		titleRotate?: number
		v3DPerspective?: number
		v3DRAngAx?: boolean
		v3DRotX?: number
		v3DRotY?: number
		valAxes?: number[]
		valAxisCrossesAt?: string | number
		valAxisDisplayUnit?: 'billions' | 'hundredMillions' | 'hundreds' | 'hundredThousands' | 'millions' | 'tenMillions' | 'tenThousands' | 'thousands' | 'trillions'
		valAxisHidden?: boolean
		valAxisLabelColor?: string
		valAxisLabelFontBold?: boolean
		valAxisLabelFontFace?: string
		valAxisLabelFontSize?: number
		valAxisLabelFormatCode?: string
		valAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		valAxisLabelRotate?: number
		valAxisLineShow?: boolean
		valAxisMajorTickMark?: ChartAxisTickMark
		valAxisMajorUnit?: number
		valAxisMaxVal?: number
		valAxisMinorTickMark?: ChartAxisTickMark
		valAxisMinVal?: number
		valAxisOrientation?: 'minMax' | 'minMax'
		valAxisTitle?: string
		valAxisTitleColor?: string
		valAxisTitleFontFace?: string
		valAxisTitleFontSize?: number
		valAxisTitleRotate?: number
		valGridLine?: OptsChartGridLine
		valueBarColors?: string[]
	}
	export interface IImageOpts extends PositionOptions, OptsDataOrPath {
		type?: 'audio' | 'online' | 'video'
		sizing?: {
			type: 'crop' | 'contain' | 'cover'
			w: number
			h: number
			x?: number
			y?: number
		}
		hyperlink?: HyperLink
		rounding?: boolean
		placeholder?: any
		rotate?: number
	}
	export interface IMediaOpts extends PositionOptions, OptsDataOrPath {
		link: string
		onlineVideoLink?: string
		type?: MediaType
	}
	export interface IShapeOptions extends PositionOptions {
		align?: HAlign
		fill?: ShapeFill
		flipH?: boolean
		flipV?: boolean
		lineSize?: number
		lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'
		lineHead?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle'
		lineTail?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle'
		line?: Color
		rectRadius?: number
		rotate?: number
		shadow?: IShadowOptions
	}
	export interface IChartTitleOpts extends FontOptions {
		title: string
		color?: String
		rotate?: number
		titleAlign?: string
		titlePos?: {
			x: number
			y: number
		}
	}
	export interface IChartMulti {
		type: CHART_NAME
		data: []
		options: {}
	}
	export interface ITableToSlidesOpts extends ITableOptions {
		addImage?: {
			url: string
			x: number
			y: number
			w?: number
			h?: number
		}
		addShape?: {
			shape: any
			opts: {}
		}
		addTable?: {
			rows: any[]
			opts: {}
		}
		addText?: {
			text: any[]
			opts: {}
		}
		_arrObjTabHeadRows?: [ITableToSlidesCell[]?]
		addHeaderToEach?: boolean
		autoPage?: boolean
		autoPageCharWeight?: number
		autoPageLineWeight?: number
		colW?: number | number[]
		masterSlideName?: string
		masterSlide?: ISlideLayout
		newSlideStartY?: number
		slideMargin?: Margin
		verbose?: boolean
	}
	export interface ITableCellOpts extends FontOptions {
		autoPageCharWeight?: number
		autoPageLineWeight?: number
		align?: HAlign
		bold?: boolean
		border?: IBorderOptions | [IBorderOptions, IBorderOptions, IBorderOptions, IBorderOptions]
		color?: Color
		colspan?: number
		fill?: ShapeFill
		margin?: Margin
		rowspan?: number
		valign?: VAlign
	}
	export interface ITableOptions extends PositionOptions, FontOptions {
		align?: HAlign
		autoPage?: boolean
		autoPageCharWeight?: number
		autoPageLineWeight?: number
		border?: IBorderOptions | [IBorderOptions, IBorderOptions, IBorderOptions, IBorderOptions]
		color?: Color
		colspan?: number
		colW?: number | number[]
		fill?: Color
		margin?: Margin
		newSlideStartY?: number
		rowW?: number | number[]
		rowspan?: number
		valign?: VAlign
	}
	export interface ITableToSlidesCell {
		type: SLIDE_OBJECT_TYPES.tablecell
		text?: string
		options?: ITableCellOpts
	}
	export interface ITableCell {
		type: SLIDE_OBJECT_TYPES.tablecell
		text?: string
		options?: ITableCellOpts
		lines?: string[]
		lineHeight?: number
		hmerge?: boolean
		vmerge?: boolean
		optImp?: any
	}
	export type TableRow = number[] | string[] | ITableCell[]
	export type ITableRow = ITableCell[]
	export interface TableRowSlide {
		rows: ITableRow[]
	}
	export interface ITextOpts extends PositionOptions, OptsDataOrPath, FontOptions {
		align?: HAlign
		autoFit?: boolean
		bodyProp?: {
			autoFit?: boolean
			align?: TEXT_HALIGN
			anchor?: TEXT_VALIGN
			lIns?: number
			rIns?: number
			tIns?: number
			bIns?: number
			vert?: 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl'
			wrap?: boolean
		}
		bold?: boolean
		breakLine?: boolean
		bullet?:
			| boolean
			| {
					type?: string
					code?: string
					style?: string
					startAt?: number
			  }
		charSpacing?: number
		color?: string
		fill?: ShapeFill
		glow?: IGlowOptions
		hyperlink?: HyperLink
		indentLevel?: number
		inset?: number
		isTextBox?: boolean
		italic?: boolean
		lang?: string
		line?: Color
		lineIdx?: number
		lineSize?: number
		lineSpacing?: number
		margin?: Margin
		outline?: {
			color: Color
			size: number
		}
		paraSpaceAfter?: number
		paraSpaceBefore?: number
		placeholder?: string
		rotate?: number
		rtlMode?: boolean
		shadow?: IShadowOptions
		shape?: SHAPE_NAME
		shrinkText?: boolean
		strike?: boolean
		subscript?: boolean
		superscript?: boolean
		underline?: boolean
		valign?: VAlign
		vert?: 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl'
		wrap?: boolean
	}
	export interface IText {
		text: string
		options?: ITextOpts
	}
	/**
	 * The Presenation Layout (ex: 'LAYOUT_WIDE')
	 */
	export interface ILayout {
		name: string
		width?: number
		height?: number
	}
	export interface IUserLayout {
		name: string
		width: number
		height: number
	}
	export interface ISlideNumber extends PositionOptions, FontOptions {
		color?: string
	}
	export interface ISlideMasterOptions {
		title: string
		height?: number
		width?: number
		margin?: Margin
		bkgd?: string | BkgdOpts
		objects?: (
			| {
					chart: {}
			  }
			| {
					image: {}
			  }
			| {
					line: {}
			  }
			| {
					rect: {}
			  }
			| {
					text: {
						options: ITextOpts
					}
			  }
			| {
					placeholder: {
						options: ISlideMstrObjPlchldrOpts
						text?: string
					}
			  })[]
		slideNumber?: ISlideNumber
	}
	export interface ISlideMstrObjPlchldrOpts {
		name: string
		type: PLACEHOLDER_TYPES
		x: Coord
		y: Coord
		w: Coord
		h: Coord
	}
	export interface ISlideRelChart extends OptsChartData {
		type: CHART_NAME | IChartMulti[]
		opts: IChartOpts
		data: OptsChartData[]
		rId: number
		Target: string
		globalId: number
		fileName: string
	}
	export interface ISlideRel {
		type: SLIDE_OBJECT_TYPES
		Target: string
		fileName?: string
		data: any[] | string
		opts?: IChartOpts
		path?: string
		extn?: string
		globalId?: number
		rId: number
	}
	export interface ISlideRelMedia {
		type: string
		opts?: IMediaOpts
		path?: string
		extn?: string
		data?: string | ArrayBuffer
		isSvgPng?: boolean
		svgSize?: {
			w: number
			h: number
		}
		rId: number
		Target: string
	}
	export interface IObjectOptions extends IShapeOptions, ITableCellOpts, ITextOpts {
		x?: Coord
		y?: Coord
		cx?: Coord
		cy?: Coord
		w?: Coord
		h?: Coord
		margin?: Margin
		colW?: number | number[]
		rowH?: number | number[]
		sizing?: {
			type?: string
			x?: number
			y?: number
			w?: number
			h?: number
		}
		rounding?: string
		placeholderIdx?: number
		placeholderType?: PLACEHOLDER_TYPES
	}
	export interface ISlideObject {
		type: SLIDE_OBJECT_TYPES
		options?: IObjectOptions
		text?: string | IText[]
		arrTabRows?: [ITableCell[]?]
		chartRid?: number
		image?: string
		imageRid?: number
		hyperlink?: HyperLink
		media?: string
		mtype?: MediaType
		mediaRid?: number
		shape?: SHAPE_NAME
	}
	export interface ISlideLayout {
		presLayout: ILayout
		name: string
		number: number
		bkgd?: string
		bkgdImgRid?: number
		slide?: {
			back: string
			bkgdImgRid?: number
			color: string
			hidden?: boolean
		}
		data: ISlideObject[]
		rels: ISlideRel[]
		relsChart: ISlideRelChart[]
		relsMedia: ISlideRelMedia[]
		margin?: Margin
		slideNumberObj?: ISlideNumber
	}
	export interface ISlide {
		addChart: Function
		addImage: Function
		addMedia: Function
		addNotes: Function
		addShape: Function
		addTable: Function
		addText: Function
		bkgd?: string
		bkgdImgRid?: number
		color?: string
		data?: ISlideObject[]
		hidden?: boolean
		margin?: Margin
		name?: string
		number: number
		presLayout: ILayout
		rels: ISlideRel[]
		relsChart: ISlideRelChart[]
		relsMedia: ISlideRelMedia[]
		slideLayout: ISlideLayout
		slideNumberObj?: ISlideNumber
	}
	export interface IPresentation {
		author: string
		company: string
		revision: string
		subject: string
		title: string
		isBrowser: boolean
		fileName: string
		fileExtn: string
		pptLayout: ILayout
		rtlMode: boolean
		saveCallback?: null
		masterSlide?: ISlide
		chartCounter: number
		imageCounter: number
		slides?: ISlide[]
		slideLayouts?: ISlideLayout[]
	}

	/**
	 * `slide.d.ts`
	 */
	export class Slide {
		bkgd: string
		color: string
		hidden: boolean
		slideNumber: ISlideNumber
		/**
		 * Generate the chart based on input data.
		 * @see OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
		 * @param {CHART_TYPE_NAMES|IChartMulti[]} `type` - chart type
		 * @param {object[]} data - a JSON object with follow the following format
		 * @param {IChartOpts} options - chart options
		 * @example
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
		 * }
		 * @return {Slide} this class
		 */
		addChart(type: CHART_NAME | IChartMulti[], data: [], options?: IChartOpts): Slide
		/**
		 * Add Image object
		 * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
		 * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
		 * @param {IImageOpts} options - image options
		 * @return {Slide} this class
		 */
		addImage(options: IImageOpts): Slide
		/**
		 * Add Media (audio/video) object
		 * @param {IMediaOpts} options - media options
		 * @return {Slide} this class
		 */
		addMedia(options: IMediaOpts): Slide
		/**
		 * Add Speaker Notes to Slide
		 * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
		 * @param {string} notes - notes to add to slide
		 * @return {Slide} this class
		 */
		addNotes(notes: string): Slide
		/**
		 * Add shape object to Slide
		 * @param {SHAPE_NAME} shape - shape object
		 * @param {IShapeOptions} options - shape options
		 * @return {Slide} this class
		 */
		addShape(shape: SHAPE_NAME, options?: IShapeOptions): Slide
		/**
		 * Add shape object to Slide
		 * @note can be recursive
		 * @param {TableRow[]} arrTabRows - table rows
		 * @param {ITableOptions} options - table options
		 * @return {Slide} this class
		 */
		addTable(arrTabRows: TableRow[], options?: ITableOptions): Slide
		/**
		 * Add text object to Slide
		 * @param {string|IText[]} text - text string or complex object
		 * @param {ITextOpts} options - text options
		 * @return {Slide} this class
		 * @since: 1.0.0
		 */
		addText(text: string | IText[], options?: ITextOpts): Slide
	}
}
