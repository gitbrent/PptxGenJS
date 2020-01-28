/**
 * PptxGenJS Interfaces
 */

import { CHART_NAME, SLIDE_OBJECT_TYPES, TEXT_HALIGN, TEXT_VALIGN, PLACEHOLDER_TYPES, SHAPE_NAME } from './core-enums'

// Common
// ======

/**
 * Coordinate (string is in the form of 'N%')
 */
export type HexColor = string // should match /^[0-9a-fA-F]{6}$/
export type ThemeColor = 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
export type Color = HexColor | ThemeColor
export type Coord = number | string // string is in form 'n%'
export type Margin = number | [number, number, number, number]
export type HAlign = 'left' | 'center' | 'right' | 'justify'
export type VAlign = 'top' | 'middle' | 'bottom'
export type ChartAxisTickMark = 'none' | 'inside' | 'outside' | 'cross'
export type HyperLink = { rId: number; slide?: number; tooltip?: string; url?: string }
export type ShapeFill = Color | { type: string; color: Color; alpha?: number }
export type BkgdOpts = { src?: string; path?: string; data?: string }
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
	data?: string // one option is required
	path?: string // one option is required
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

// TODO: FUTURE: BREAKING-CHANGE: (soln: use `OptsDataLabelPosition|string` until 3.5/4.0)
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
	color?: HexColor
	pt?: number
	type?: string // TODO: specify values, eg: 'none'
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

// TODO: divide chart opts into sections, then:
// `export type ChartOptions = ChartBaseOptions | ChartAxesOptions | ChartBarDataLineOptions | Chart3DBarOptions;`
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
	titlePos?: { x: number; y: number }
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
	sizing?: { type: 'crop' | 'contain' | 'cover'; w: number; h: number; x?: number; y?: number }
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
	titlePos?: { x: number; y: number }
}
export interface IChartMulti {
	type: CHART_NAME
	data: []
	options: {}
}

export interface ITableToSlidesOpts extends ITableOptions {
	addImage?: { url: string; x: number; y: number; w?: number; h?: number }
	addShape?: { shape: any; opts: {} }
	addTable?: { rows: any[]; opts: {} }
	addText?: { text: any[]; opts: {} }
	//
	_arrObjTabHeadRows?: [ITableToSlidesCell[]?]
	addHeaderToEach?: boolean
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
// TODO: replace this with `ITableCell`
export interface ITableToSlidesCell {
	type: SLIDE_OBJECT_TYPES.tablecell
	text?: string
	options?: ITableCellOpts
}
export interface ITableCell {
	type: SLIDE_OBJECT_TYPES.tablecell
	text?: string
	options?: ITableCellOpts
	// internal fields below
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
	bold?: boolean
	breakLine?: boolean
	bullet?: boolean | { type?: string; code?: string; style?: string; startAt?: number }
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
		| { chart: {} }
		| { image: {} }
		| { line: {} }
		| { rect: {} }
		| { text: { options: ITextOpts } }
		| { placeholder: { options: ISlideMstrObjPlchldrOpts; text?: string } })[]
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
	arrTabRows?: [ITableCell[]?]
	// chart
	chartRid?: number
	// image:
	image?: string
	imageRid?: number
	hyperlink?: HyperLink
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
	relsChart: ISlideRelChart[] // needed as we use args:"ISlide|ISlideLayout" often
	relsMedia: ISlideRelMedia[] // needed as we use args:"ISlide|ISlideLayout" often
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
	bkgdImgRid?: number // FIXME rename
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
	slideNumberObj?: ISlideNumber // FIXME rename
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
