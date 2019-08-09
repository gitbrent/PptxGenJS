/**
 * PptxGenJS Interfaces
 */

import { CHART_TYPE_NAMES, SLIDE_OBJECT_TYPES, TEXT_HALIGN, TEXT_VALIGN, PLACEHOLDER_TYPES } from './core-enums'

// FIXME: WONT FORMAT ANYMORE!!! ("TSTypeAliasDeclaration" error)

// Common
// ======

/**
 * Coordinate (string is in the form of 'N%')
 */
export type HexColor = string // should match /^[0-9a-fA-F]{6}$/
export type ThemeColor = 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
export type Color = HexColor | ThemeColor
export type Coord = number | string // string is in form 'n%'
export type HAlign = 'left' | 'center' | 'right' | 'justify'
export type VAlign = 'top' | 'middle' | 'bottom'
export type HyperLink = { rId: number; slide?: number; tooltip?: string; url?: string }

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
	labels?: Array<string>
	values?: Array<number>
	sizes?: Array<number>
}
export interface OptsChartGridLine {
	size?: number
	color?: string
	style?: 'solid' | 'dash' | 'dot' | 'none'
}
type MediaType = 'audio' | 'online' | 'video'

// Opts
// ====
export interface BorderOptions {
	color?: HexColor
	pt?: number
	type?: string // TODO: specify values, eg: 'none'
}
export interface ShadowOpts {
	type: string
	angle: number
	opacity: number
	blur?: number
	offset?: number
	color?: string
}

// TODO:
//   export type ChartOptions = ChartBaseOptions | ChartAxesOptions | ChartBarDataLineOptions | Chart3DBarOptions;

export interface IChartOpts extends PositionOptions, OptsChartGridLine {
	type: CHART_TYPE_NAMES | IChartMulti[]
	layout?: PositionOptions
	barDir?: string
	barGrouping?: string
	barGapWidthPct?: number
	barGapDepthPct?: number
	bar3DShape?: string
	catAxisLineShow?: boolean
	catAxisMaxVal?: number
	catAxisMinVal?: number
	catAxisHidden?: boolean
	catAxisOrientation?: 'minMax' | 'minMax'
	catAxisLabelRotate?: number
	catAxisLabelFontBold?: boolean
	catAxisTitleColor?: string
	catAxisTitleFontFace?: string
	catAxisTitleFontSize?: number
	catAxisTitleRotate?: number
	catAxisTitle?: string
	catAxisLabelFontSize?: number
	catAxisLabelColor?: string
	catAxisLabelFontFace?: string
	catAxisLabelFrequency?: string
	catAxisBaseTimeUnit?: string
	catAxisMajorTimeUnit?: string
	catAxisMinorTimeUnit?: string
	catAxisMajorUnit?: string
	catAxisMinorUnit?: string
	catGridLine?: OptsChartGridLine
	valGridLine?: OptsChartGridLine
	chartColors?: Array<string>
	chartColorsOpacity?: number
	showLabel?: boolean
	lang?: string
	dataNoEffects?: string
	dataLabelFormatScatter?: string
	dataLabelFormatCode?: string
	dataLabelBkgrdColors?: boolean
	dataLabelFontSize?: number
	dataLabelColor?: string
	dataLabelFontFace?: string
	dataLabelPosition?: string
	displayBlanksAs?: string
	fill?: string
	border?: BorderOptions
	hasArea?: boolean
	catAxes?: Array<number>
	valAxes?: Array<number>
	lineDataSymbol?: string
	lineDataSymbolSize?: number
	lineDataSymbolLineColor?: string
	lineDataSymbolLineSize?: number
	showLegend?: boolean
	showCatAxisTitle?: boolean
	legendPos?: string
	legendFontFace?: string
	legendFontSize?: number
	legendColor?: string
	lineSmooth?: boolean
	invertedColors?: string
	serAxisOrientation?: string
	serAxisHidden?: boolean
	serGridLine?: OptsChartGridLine
	showSerAxisTitle?: boolean
	serLabelFormatCode?: string
	serAxisLabelPos?: string
	serAxisLineShow?: boolean
	serAxisLabelFontSize?: string
	serAxisLabelColor?: string
	serAxisLabelFontFace?: string
	serAxisLabelFrequency?: string
	serAxisBaseTimeUnit?: string
	serAxisMajorTimeUnit?: string
	serAxisMinorTimeUnit?: string
	serAxisMajorUnit?: number
	serAxisMinorUnit?: number
	serAxisTitleColor?: string
	serAxisTitleFontFace?: string
	serAxisTitleFontSize?: number
	serAxisTitleRotate?: number
	serAxisTitle?: string
	showDataTable?: boolean
	showDataTableHorzBorder?: boolean
	showDataTableVertBorder?: boolean
	showDataTableOutline?: boolean
	showDataTableKeys?: boolean
	title?: string
	titleFontSize?: number
	titleColor?: string
	titleFontFace?: string
	titleRotate?: number
	titleAlign?: string
	titlePos?: { x: number; y: number }
	dataLabelFontBold?: boolean
	valueBarColors?: Array<string>
	holeSize?: number
	showTitle?: boolean
	showValue?: boolean
	showPercent?: boolean
	catLabelFormatCode?: string
	dataBorder?: BorderOptions
	lineSize?: number
	lineDash?: string
	radarStyle?: string
	shadow?: ShadowOpts
	catAxisLabelPos?: string
	valAxisOrientation?: 'minMax' | 'minMax'
	valAxisCrossesAt?: string | number
	valAxisMaxVal?: number
	valAxisMinVal?: number
	valAxisHidden?: boolean
	valAxisTitleColor?: string
	valAxisTitleFontFace?: string
	valAxisTitleFontSize?: number
	valAxisTitleRotate?: number
	valAxisTitle?: string
	valAxisLabelFormatCode?: string
	valAxisLineShow?: boolean
	valAxisLabelRotate?: number
	valAxisLabelFontSize?: number
	valAxisLabelFontBold?: boolean
	valAxisLabelColor?: string
	valAxisLabelFontFace?: string
	valAxisMajorUnit?: number
	showValAxisTitle?: boolean
	axisPos?: string
	v3DRotX?: number
	v3DRotY?: number
	v3DRAngAx?: boolean
	v3DPerspective?: number
}
export interface IImageOpts extends PositionOptions, OptsDataOrPath {
	type?: 'audio' | 'online' | 'video'
	sizing?: { type: 'crop' | 'contain' | 'cover'; w: number; h: number; x?: number; y?: number }
	hyperlink?: HyperLink
	rounding?: boolean
	placeholder?: any
}
export interface IMediaOpts extends PositionOptions, OptsDataOrPath {
	link: string
	onlineVideoLink?: string
	type?: MediaType
}

export interface Shape {
	displayName: string
	name: string
	avLst: { [key: string]: number }
}
export type ShapeFill = Color | { type: string; color: Color; alpha?: number }
export interface ShapeOptions extends PositionOptions {
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
	shadow?: ShadowOpts
}

export interface IChartTitleOpts {
	title: string
	color?: String
	fontSize?: number
	fontFace?: string
	rotate?: number
	titleAlign?: string
	titlePos?: { x: number; y: number }
}
export interface IChartMulti {
	type: CHART_TYPE_NAMES
	data: []
	options: {}
}

export interface ITableToSlidesOpts extends PositionOptions {
	addImage?: { url: string; x: number; y: number; w?: number; h?: number }
	addShape?: { shape: any; opts: {} }
	addTable?: { rows: any[]; opts: {} }
	addText?: { text: any[]; opts: {} }
	//
	_arrObjTabHeadRows?: any[]
	addHeaderToEach?: boolean
	autoPage?: boolean
	colW?: number | number[]
	debug?: boolean
	lineWeight?: number
	masterSlideName?: string
	masterSlide?: ISlideLayout
	newSlideStartY?: number
	slideMargin?: number | [number, number, number, number]
}
export interface ITableCellOpts {
	align?: HAlign
	bold?: boolean
	border?: BorderOptions | [BorderOptions, BorderOptions, BorderOptions, BorderOptions]
	color?: string
	colspan?: number
	fill?: ShapeFill
	fontSize?: number
	lineWeight?: number
	margin?: number | [number, number, number, number]
	rowspan?: number
	valign?: VAlign
}
export interface ITableToSlidesCell {
	type: SLIDE_OBJECT_TYPES.tablecell
	text?: string
	options?: ITableCellOpts
}
export interface TableCell {
	type: SLIDE_OBJECT_TYPES.tablecell
	text?: string
	hmerge?: boolean
	vmerge?: boolean
	optImp?: any
	options?: ITableCellOpts
}
export type TableRow = [TableCell[]]
export interface TableOptions extends PositionOptions {
	align?: HAlign
	autoPage?: boolean
	border?: BorderOptions | [BorderOptions, BorderOptions, BorderOptions, BorderOptions]
	color?: Color
	colspan?: number
	colW?: number | number[]
	fill?: Color
	fontSize?: number
	lineWeight?: number
	margin?: number | [number, number, number, number]
	newSlideStartY?: number
	rowW?: number | number[]
	rowspan?: number
	valign?: VAlign
}

export interface ITextOpts extends PositionOptions, OptsDataOrPath {
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
	bullet?: boolean | { type?: string; code?: string }
	charSpacing?: number
	color?: string
	fontFace?: string
	fontSize?: number
	hyperlink?: HyperLink
	indentLevel?: number
	inset?: number
	isTextBox?: boolean
	italic?: boolean
	lang?: string
	line?: Color
	lineIdx?:number
	lineSize?: number
	lineSpacing?: number
	outline?: { color:Color, size:number }
	paraSpaceAfter?: number
	paraSpaceBefore?: number
	placeholder?: string
	rotate?: number // (degree * 60,000)
	rtlMode?: boolean
	shadow?: ShadowOpts
	shape?: Shape
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
export interface ISlideNumber extends PositionOptions {
	fontFace?: string
	fontSize?: number
	color?: string
}
export interface ISlideMasterDef {
	title: string
	height?: number
	width?: number
	margin?: number | [number, number, number, number]
	bkgd?: string
	objects?: [{}]
	slideNumber?: ISlideNumber
}
export interface ISlideRelChart extends OptsChartData {
	type: CHART_TYPE_NAMES | IChartMulti[]
	opts: IChartOpts
	data: Array<OptsChartData>
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

export interface ObjectOptions extends ShapeOptions, ITableCellOpts, ITextOpts {
	x?: Coord
	y?: Coord
	cx?: Coord
	cy?: Coord
	w?: Coord
	h?: Coord
	margin?: number | [number, number, number, number]
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
	options?: ObjectOptions
	// text
	text?: string | IText[]
	// table
	arrTabRows?: [TableCell[]?]
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
	shape?: Shape
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
	data: Array<ISlideObject>
	rels: Array<ISlideRel>
	relsChart: Array<ISlideRelChart> // needed as we use args:"ISlide|ISlideLayout" often
	relsMedia: Array<ISlideRelMedia> // needed as we use args:"ISlide|ISlideLayout" often
	margin?: number | [number, number, number, number]
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
	margin?: number | [number, number, number, number]
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
