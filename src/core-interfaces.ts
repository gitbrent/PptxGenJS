/**
 * PptxGenJS Interfaces
 */

import { CHART_NAME, PLACEHOLDER_TYPES, SHAPE_NAME, SLIDE_OBJECT_TYPES, TEXT_HALIGN, TEXT_VALIGN } from './core-enums'

// Common
// ======

/**
 * Coordinate number - either:
 * - Inches
 * - Percentage
 *
 * @example 10.25
 * coordinate in inches
 * @example '75%'
 * coordinate as percentage of slide size
 */
export type Coord = number | string
/**
 * Color in Hex format
 * @example 'FF3399'
 */
export type HexColor = string
export type ThemeColor = 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
export type Color = HexColor | ThemeColor
export type Margin = number | [number, number, number, number]
export type HAlign = 'left' | 'center' | 'right' | 'justify'
export type VAlign = 'top' | 'middle' | 'bottom'
export type MediaType = 'audio' | 'online' | 'video'
export type ChartAxisTickMark = 'none' | 'inside' | 'outside' | 'cross'
export type ShapeFill = Color | { type: string; color: Color; alpha?: number }
export type HyperLink = {
	slide?: number
	url?: string
	tooltip?: string
}
/**
 * Either `data` or `path` is required
 */
export type OptsDataOrPath = {
	/**
	 * URL or relative path
	 *
	 * @example 'https://onedrives.com/myimg.png`
	 * retrieve image via URL
	 * @example '/home/gitbrent/images/myimg.png`
	 * retrieve image via local path
	 */
	path?: string
	/**
	 * base64-encoded string
	 * - Useful for avoiding potential path/server issues
	 *
	 * @example 'image/png;base64,iVtDafDrBF[...]='
	 * adds a pre-encoded image
	 */
	data?: string
}
export interface IHyperLink extends HyperLink {
	rId: number
}
export interface BkgdOpts extends OptsDataOrPath {
	/**
	 * Color (hex format)
	 * @example 'FF3399'
	 */
	fill?: HexColor
}
export type TextOptions = {
	/**
	 * text alignment
	 * @default 'left'
	 */
	align?: HAlign
	bold?: boolean
	breakLine?: boolean
	bullet?: boolean | { type?: string; code?: string; marginPt: number; style?: string; startAt?: number }
	color?: Color
	fontFace?: string
	fontSize?: number
	italic?: boolean
	/**
	 * language
	 * - ISO 639-1 standard language code
	 * @default 'en-US' // english US
	 * @example 'fr-CA' // french Canadian
	 */
	lang?: string
	valign?: VAlign
}
export type PositionOptions = {
	/**
	 * Horizontal position
	 * - inches or percentage
	 * @example 10.25
	 * position in inches
	 * @example '75%'
	 * position as percentage of slide size
	 */
	x?: Coord
	/**
	 * Vertical position
	 * - inches or percentage
	 * @example 10.25
	 * position in inches
	 * @example '75%'
	 * position as percentage of slide size
	 */
	y?: Coord
	/**
	 * Height
	 * - inches or percentage
	 * @example 10.25
	 * height in inches
	 * @example '75%'
	 * height as percentage of slide size
	 */
	h?: Coord
	/**
	 * Width
	 * - inches or percentage
	 * @example 10.25
	 * width in inches
	 * @example '75%'
	 * width as percentage of slide size
	 */
	w?: Coord
}
export type OptsChartData = {
	index?: number
	labels?: string[]
	name?: string
	sizes?: number[]
	values?: number[]
}
export type OptsChartGridLine = {
	/**
	 * line color (hex format)
	 * @example 'FF3399'
	 */
	color?: HexColor
	/**
	 * line size (points)
	 */
	size?: number
	/**
	 * line style
	 */
	style?: 'solid' | 'dash' | 'dot' | 'none'
}

// FUTURE: BREAKING-CHANGE: (soln: use `OptsDataLabelPosition|string` until 3.5/4.0)
/*
export interface OptsDataLabelPosition {
	pie: 'ctr' | 'inEnd' | 'outEnd' | 'bestFit'
	scatter: 'b' | 'ctr' | 'l' | 'r' | 't'
	// TODO: add all othere chart types
}
*/

// Opts
// ====
export interface IBorderOptions {
	/**
	 * Border color (hex format)
	 * @example 'FF3399'
	 */
	color?: HexColor
	/**
	 * Border size (points)
	 */
	pt?: number
	/**
	 * Border type
	 */
	type?: 'none' | 'dash' | 'solid'
}
export interface IShadowOptions {
	/**
	 * shadow type
	 */
	type: 'outer' | 'inner' | 'none'
	/**
	 * opacity (0.0 - 1.0)
	 * @example 0.5
	 * 50% opaque
	 */
	opacity: number // TODO: "Transparency (0-100%)" in PPT // TODO: deprecate and add `transparency`
	/**
	 * blue (points)
	 * - range: 0-100
	 */
	blur?: number
	/**
	 * angle (degrees)
	 * - range: 0-359
	 */
	angle: number
	/**
	 * shadow offset (points)
	 * - range: 0-200
	 */
	offset?: number // TODO: "Distance" in PPT
	/**
	 * shadow color (hex format)
	 * @example 'FF3399'
	 */
	color?: HexColor
}
export interface IGlowOptions {
	/**
	 * Border color (hex format)
	 * @example 'FF3399'
	 */
	color?: HexColor
	/**
	 * opacity (0.0 - 1.0)
	 * @example 0.5
	 * 50% opaque
	 */
	opacity: number
	/**
	 * size (points)
	 */
	size: number
}

// charts =========================================================================================

export interface IChartTitleOpts extends TextOptions {
	color?: Color
	rotate?: number
	title: string
	titleAlign?: string
	titlePos?: { x: number; y: number }
}
export interface IChartMulti {
	type: CHART_NAME
	data: any[]
	options: {}
}
export interface IChartPropsBase {
	axisPos?: string
	border?: IBorderOptions
	chartColors?: string[]
	chartColorsOpacity?: number
	dataBorder?: IBorderOptions
	displayBlanksAs?: string
	fill?: string
	invertedColors?: string
	lang?: string
	layout?: PositionOptions
	shadow?: IShadowOptions
	showLabel?: boolean
	showLeaderLines?: boolean
	showLegend?: boolean
	showPercent?: boolean
	showTitle?: boolean
	showValue?: boolean
	v3DPerspective?: number
	v3DRAngAx?: boolean
	v3DRotX?: number
	v3DRotY?: number
}
export interface IChartPropsAxisCat {
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
	showCatAxisTitle?: boolean
}
export interface IChartPropsAxisSer {
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
	showSerAxisTitle?: boolean
}
export interface IChartPropsAxisVal {
	showValAxisTitle?: boolean
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
}
export interface IChartPropsChartBar {
	bar3DShape?: string
	barDir?: string
	barGapDepthPct?: number
	barGapWidthPct?: number
	barGrouping?: string
	valueBarColors?: string[]
}
export interface IChartPropsChartDoughnut {
	dataNoEffects?: boolean
	holeSize?: number
}
export interface IChartPropsChartLine {
	lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'
	lineDataSymbol?: 'circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'square' | 'triangle'
	lineDataSymbolLineColor?: string
	lineDataSymbolLineSize?: number
	lineDataSymbolSize?: number
	lineSize?: number
	lineSmooth?: boolean
}
export interface IChartPropsChartPie {
	dataNoEffects?: boolean
}
export interface IChartPropsChartRadar {
	radarStyle?: 'standard' | 'marker' | 'filled'
}
export interface IChartPropsDataLabel {
	dataLabelBkgrdColors?: boolean
	dataLabelColor?: string
	dataLabelFontBold?: boolean
	dataLabelFontFace?: string
	dataLabelFontSize?: number
	dataLabelFormatCode?: string
	dataLabelFormatScatter?: 'custom' | 'customXY' | 'XY'
	dataLabelPosition?: 'b' | 'bestFit' | 'ctr' | 'l' | 'r' | 't' | 'inEnd' | 'outEnd' | 'bestFit'
}
export interface IChartPropsDataTable {
	dataTableFontSize?: number
	showDataTable?: boolean
	showDataTableHorzBorder?: boolean
	showDataTableKeys?: boolean
	showDataTableOutline?: boolean
	showDataTableVertBorder?: boolean
}
export interface IChartPropsLegend {
	legendColor?: string
	legendFontFace?: string
	legendFontSize?: number
	legendPos?: 'b' | 'l' | 'r' | 't' | 'tr'
}
export interface IChartPropsTitle {
	title?: string
	titleAlign?: string
	titleColor?: string
	titleFontFace?: string
	titleFontSize?: number
	titlePos?: { x: number; y: number }
	titleRotate?: number
}
export interface IChartOpts
	extends IChartPropsAxisCat,
		IChartPropsAxisSer,
		IChartPropsAxisVal,
		IChartPropsBase,
		IChartPropsChartBar,
		IChartPropsChartDoughnut,
		IChartPropsChartLine,
		IChartPropsChartPie,
		IChartPropsChartRadar,
		IChartPropsDataLabel,
		IChartPropsDataTable,
		IChartPropsLegend,
		IChartPropsTitle,
		OptsChartGridLine,
		PositionOptions {}
export interface IChartOptsLib extends IChartOpts {
	_type?: CHART_NAME | IChartMulti[]
}

// image / media ==================================================================================

export interface ImageOpts extends PositionOptions, OptsDataOrPath {
	hyperlink?: IHyperLink
	/**
	 * image rotation (degrees)
	 * - range: -360 to 360
	 * @example 180 // rotate image 180 degrees
	 */
	rotate?: number
	rounding?: boolean
	sizing?: {
		type: 'contain' | 'cover' | 'crop'
		w: number
		h: number
		x?: number
		y?: number
	}
}
export interface IImageOpts extends ImageOpts {
	type?: 'audio' | 'online' | 'video'
	placeholder?: any
}
/**
 * Add media (audio/video) to slide
 * @requires either `link` or `path`
 */
export interface IMediaOpts extends PositionOptions, OptsDataOrPath {
	/**
	 * Media type
	 * - Use 'online' to embed a YouTube video (only supported in recent versions of PowerPoint)
	 */
	type: MediaType
	/**
	 * video embed link
	 * - works with YouTube
	 * - other sites may not show correctly in PowerPoint
	 * @example 'https://www.youtube.com/embed/Dph6ynRVyUc' // embed a youtube video
	 */
	link?: string
	/**
	 * full or local path
	 * @example 'https://freesounds/simpsons/bart.mp3' // embed mp3 audio clip from server
	 * @example '/sounds/simpsons_haha.mp3' // embed mp3 audio clip from local directory
	 */
	path?: string
}

// shapes =========================================================================================

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

// tables =========================================================================================

// TODO: create TableToSlidesOpts (incl. verbose)
export interface ITableToSlidesOpts extends TableOptions {
	addImage?: { url: string; x: number; y: number; w?: number; h?: number }
	addShape?: { shape: any; options: {} }
	addTable?: { rows: any[]; options: {} }
	addText?: { text: any[]; options: {} }
	//
	_arrObjTabHeadRows?: TableRow[]
	addHeaderToEach?: boolean // TODO: 20200528: rename to "autoPageRepeatHeader"
	autoPage?: boolean
	autoPageCharWeight?: number // -1.0 to 1.0
	autoPageLineWeight?: number // -1.0 to 1.0
	colW?: number | number[]
	masterSlideName?: string
	masterSlide?: ISlideLayout
	newSlideStartY?: number
	slideMargin?: Margin
	verbose?: boolean // Undocumented; shows verbose output
}
export interface ITableCellOpts extends TextOptions {
	autoPageCharWeight?: number
	autoPageLineWeight?: number
	border?: IBorderOptions | [IBorderOptions, IBorderOptions, IBorderOptions, IBorderOptions]
	colspan?: number
	fill?: ShapeFill
	margin?: Margin
	rowspan?: number
	valign?: VAlign
}
export interface TableOptions extends PositionOptions, TextOptions {
	/**
	 * Whether to create new slides as table rows overflow each slide
	 * @default false
	 */
	autoPage?: boolean
	/**
	 * Character weight - affects line length before wrapping begins
	 * @type float (-1.0 to 1.0)
	 * @default 0
	 */
	autoPageCharWeight?: number
	/**
	 * Number of rows that comprise table headers.
	 * Required when `autoPageRepeatHeader` is set to true.
	 * @example 2 - repeats the first two table rows on each new slide created
	 * @default 1
	 * @since v3.3.0
	 */
	autoPageHeaderRows?: number
	/**
	 * Line weight - affects line height before paging begins
	 * @type float (-1.0 to 1.0)
	 * @default 0
	 */
	autoPageLineWeight?: number
	/**
	 * Whether table header row(s) should be repeated on each new slide creating by autoPage.
	 * Use `autoPageHeaderRows` to designate how many rows comprise the table header (1+).
	 * @default false
	 * @since v3.3.0
	 */
	autoPageRepeatHeader?: boolean
	/**
	 * Table border
	 * - single value is applied to all 4 sides
	 * - array of values in TRBL order for individual sides
	 */
	border?: IBorderOptions | [IBorderOptions, IBorderOptions, IBorderOptions, IBorderOptions]
	/**
	 * Width of table columns
	 * - single value is applied to every column equally based upon `w`
	 * - array of values in applied to each column in order
	 * @default columns of equal width based upon `w`
	 */
	colW?: number | number[]
	/**
	 * Cell background color
	 */
	fill?: Color
	/**
	 * Cell margin
	 * - affects all table cells, is superceded by cell options
	 */
	margin?: Margin
	/**
	 * Starting `y` location on additional slides created by autoPage=true
	 * @default `y` value from table options
	 */
	newSlideStartY?: number
	/**
	 * Height of table rows
	 * - single value is applied to every row equally based upon `h`
	 * - array of values in applied to each row in order
	 * @default rows of equal height based upon `h`
	 */
	rowH?: number | number[]
}
export interface ITableOptions extends TableOptions {
	_arrObjTabHeadRows?: TableRow[]
}
export interface TableCell {
	text?: string | TableCell[]
	options?: ITableCellOpts
}
// TODO: replace this with `ITableCell`
export interface ITableToSlidesCell {
	type: SLIDE_OBJECT_TYPES.tablecell
	text?: string | TableCell[]
	options?: ITableCellOpts
}
export interface ITableCell extends TableCell {
	type: SLIDE_OBJECT_TYPES.tablecell
	lines?: string[]
	lineHeight?: number
	hmerge?: boolean
	vmerge?: boolean
	optImp?: any
}
export type ITableRow = ITableCell[]
// TODO: 20200523: Consistency: Remove `number[]` as Cell/IText only take strings
export type TableRow = number[] | string[] | TableCell[]
export interface TableRowSlide {
	rows: ITableRow[]
}

// text ===========================================================================================

export interface ITextOpts extends PositionOptions, OptsDataOrPath, TextOptions {
	autoFit?: boolean
	bodyProp?: {
		// Note: Many of these duplicated as user options are transformed to bodyProp options for XML processing
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
	charSpacing?: number
	fill?: ShapeFill
	glow?: IGlowOptions
	hyperlink?: IHyperLink
	indentLevel?: number
	inset?: number
	isTextBox?: boolean
	line?: Color
	lineIdx?: number
	lineSize?: number
	lineSpacing?: number
	margin?: Margin
	outline?: { color: Color; size: number }
	paraSpaceAfter?: number
	paraSpaceBefore?: number
	placeholder?: string
	rotate?: number // (degree * 60,000)
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

// Core
// ====
/**
 * Section options
 */
export interface ISectionProps {
	title: string
	order?: number // 1-n
}
export interface ISection {
	type: 'user' | 'default'
	title: string
	slides: ISlideLib[]
}
/**
 * The Presentation Layout (ex: 'LAYOUT_WIDE')
 */
export interface ILayout {
	name: string
	width?: number
	height?: number
}
export interface ILayoutProps {
	name: string
	width: number
	height: number
}
export interface ISlideNumber extends PositionOptions, TextOptions {
	align?: HAlign
	color?: string
}
export interface ISlideMasterOptions {
	title: string
	height?: number
	width?: number
	margin?: Margin
	background?: BkgdOpts
	bkgd?: string | BkgdOpts // @deprecated v3.3.0
	objects?: (
		| { chart: {} }
		| { image: {} }
		| { line: {} }
		| { rect: {} }
		| { text: { options: ITextOpts; text?: string } }
		| { placeholder: { options: ISlideMstrObjPlchldrOpts; text?: string } }
	)[]
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
	opts: IChartOptsLib
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
	svgSize?: { w: number; h: number }
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
	// table
	colW?: number | number[]
	rowH?: number | number[]
	// image:
	sizing?: {
		type?: string
		x?: number
		y?: number
		w?: number
		h?: number
	}
	rounding?: string
	// placeholder
	placeholderIdx?: number
	placeholderType?: PLACEHOLDER_TYPES
}
export interface ISlideObject {
	type: SLIDE_OBJECT_TYPES
	options?: IObjectOptions
	// text
	text?: string | IText[]
	// table
	arrTabRows?: ITableCell[][]
	// chart
	chartRid?: number
	// image:
	image?: string
	imageRid?: number
	hyperlink?: IHyperLink
	// media
	media?: string
	mtype?: MediaType
	mediaRid?: number
	shape?: SHAPE_NAME
}

export interface ISlideLayout {
	presLayout: ILayout
	name: string
	number: number
	background?: BkgdOpts
	bkgd?: string // @deprecated v3.3.0
	bkgdImgRid?: number
	slide?: {
		back: string
		bkgdImgRid?: number
		color: string
		hidden?: boolean
	}
	data: ISlideObject[]
	rels: ISlideRel[]
	relsChart: ISlideRelChart[] // needed as we use args:"ISlide|ISlideLayout" often
	relsMedia: ISlideRelMedia[] // needed as we use args:"ISlide|ISlideLayout" often
	margin?: Margin
	slideNumberObj?: ISlideNumber
}
export interface IAddSlideOptions {
	masterName?: string // TODO: 20200528: rename to "masterTitle" (createMaster uses `title` so lets be consistent)
	sectionTitle?: string
}
export interface ISlide {
	addChart: Function
	addImage: Function
	addMedia: Function
	addNotes: Function
	addShape: Function
	addTable: Function
	addText: Function
	background?: BkgdOpts
	bkgd?: string // @deprecated v3.3.0
	/**
	 * Default text color (hex format)
	 * @example 'FF3399'
	 */
	color?: HexColor
	hidden?: boolean
	slideNumber?: ISlideNumber
}
export interface ISlideLib extends ISlide {
	bkgdImgRid?: number // FIXME rename
	data?: ISlideObject[]
	id: number
	margin?: Margin
	name?: string
	number: number
	presLayout: ILayout
	rels: ISlideRel[]
	relsChart: ISlideRelChart[]
	relsMedia: ISlideRelMedia[]
	rId: number
	slideLayout: ISlideLayout
	slideNumberObj?: ISlideNumber // FIXME rename
}
export interface IPresentation {
	author: string
	company: string
	layout: string
	masterSlide: ISlide
	presLayout: ILayout
	revision: string
	rtlMode: boolean
	sections: ISection[]
	slideLayouts: ISlideLayout[]
	slides: ISlideLib[]
	subject: string
	title: string
}
