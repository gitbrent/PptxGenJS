/**
 * PptxGenJS Interfaces
 */

import { CHART_TYPE_NAMES, SLIDE_OBJECT_TYPES } from './core-enums'

// FIXME: WONT FORMAT ANYMORE!!! ("TSTypeAliasDeclaration" error)

// Common
// ======

/**
 * Coordinate (string is in the form of 'N%')
 */
type Coord = number | string
export interface OptsCoords {
	x?: Coord
	y?: Coord
	w?: Coord
	h?: Coord
}
/**
 * `data`/`path` options (one option is required)
 */
export interface OptsDataOrPath {
	data?: string
	path?: string
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
export interface IBorderOpts {
	color?: string // '#696969'
	pt?: number
	type?: string // TODO: specify values
}
export interface IShadowOpts {
	type: string
	angle: number
	opacity: number
	blur?: number
	offset?: number
	color?: string
}
export interface IChartOpts extends OptsCoords, OptsChartGridLine {
	type: CHART_TYPE_NAMES|IChartMulti[]
	layout?: OptsCoords
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
	border?: IBorderOpts
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
	dataBorder?: IBorderOpts
	lineSize?: number
	lineDash?: string
	radarStyle?: string
	shadow?: IShadowOpts
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
export interface IImageOpts extends OptsCoords, OptsDataOrPath {
	type?: 'audio' | 'online' | 'video'
	sizing?: { type:'crop' | 'contain' | 'cover', w:number, h:number, x?:number, y?:number }
	hyperlink?: any
	rounding?: boolean
	placeholder?: any
}
export interface IMediaOpts extends OptsCoords, OptsDataOrPath {
	link: string
	onlineVideoLink?: string
	type?: MediaType
}
export interface ITextOpts extends OptsCoords, OptsDataOrPath {
	align?: string // "left" | "center" | "right"
	autoFit?: boolean
	color?: string
	fontSize?: number
	inset?: number
	lineSpacing?: number
	line?: string // color
	lineSize?: number
	placeholder?: string
	rotate?: number // VALS: degree * 60,000
	shadow?: IShadowOpts
	shape?: { name: string }
	vert?: 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl'
	valign?: string //"top" | "middle" | "bottom"
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
export interface ITableToSlidesOpts {
	addImage?: {url:string, x:number, y:number, w?:number, h?:number}
	addShape?: {shape:any, opts:{}}
	addTable?: {rows:any[], opts:{}}
	addText?: {text:any[], opts:{}}
	//
	x?: number
	y?: number
	w?: number
	h?: number
	_arrObjTabHeadRows?: any[]
	addHeaderToEach?: boolean
	autoPage?: boolean
	colW?: Array<number>
	debug?: boolean
	lineWeight?: number
	masterSlideName?: string
	masterSlide?: ISlideLayout
	newSlideStartY?: number
	slideMargin?: Array<number> | number
}
export interface ITableCellOpts {
	align?:'left'|'center'|'right'
	bold?:boolean
	border?:IBorderOpts
	color?:string
	colspan?:number
	fill?:string
	fontSize?:number
	isTableCell?: boolean;
	lineWeight?:number
	margin?:number|number[]
	rowspan?:number
	valign?:'top'|'middle'|'bottom'
}
export interface ITableToSlidesCell {
	text?: string
	options?: ITableCellOpts
}
export interface ITableCell {
	text?: string
	hmerge?: boolean
	vmerge?: boolean
	optImp?: any
	options?: ITableCellOpts
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
export interface ISlideNumber extends OptsCoords {
	fontFace?: string
	fontSize?: number
	color?: string
}
export interface ISlideMasterDef {
	title: string
	height?: number
	width?: number
	margin?: Array<number> | number
	bkgd?: string
	objects?: [{}]
	slideNumber?: ISlideNumber
}
export interface ISlideRelChart extends OptsChartData {
	type: CHART_TYPE_NAMES|IChartMulti[]
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
	opts?:IChartOpts
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
	svgSize?: {w:number, h:number}
	rId: number
	Target: string
}

export interface ObjectOptions {
	x?: Coord
	y?: Coord
	cx?: Coord
	cy?: Coord
	w?: number
	h?: number
	placeholder?: string
	shape?: object
	bodyProp?: {
		lIns?: number
		rIns?: number
		bIns?: number
		tIns?: number
	}
	isTextBox?: boolean
	line?: string
	margin?: number
	rectRadius?: number
	fill?: string
	shadow?: IShadowOpts
	colW?: number
	rowH?: number
	flipH?: boolean
	flipV?: boolean
	rotate?: number
	lineDash?: string
	lineSize?: number
	lineHead?: string
	lineTail?: string
	// image:
	sizing?: {
		type?: string
		x?: number
		y?: number
		w?: number
		h?: number
	}
	rounding?: string
}
export interface ISlideObject {
	type: SLIDE_OBJECT_TYPES
	options?: ObjectOptions
	// text
	text?: string
	// table
	arrTabRows?: [ITableCell[]?]
	// chart
	chartRid?: number
	// image:
	image?: string
	imageRid?: number
	hyperlink?: { rId: number; slide?: number; tooltip?: string; url?: string }
	// media
	media?: string
	mtype?: MediaType
	mediaRid?: number
}

export interface ISlideLayout {
	presLayout: ILayout
	name: string
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
	number?: number // needed as we use args:"ISlide|ISlideLayout" often
	relsChart: Array<ISlideRelChart> // needed as we use args:"ISlide|ISlideLayout" often
	relsMedia: Array<ISlideRelMedia> // needed as we use args:"ISlide|ISlideLayout" often
	margin?: Array<number> | number
	slideNumberObj?: ISlideNumber
	width: number
	height: number
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
	margin?: object
	name?: string
	number: number
	presLayout: ILayout
	rels: ISlideRel[]
	relsChart: ISlideRelChart[]
	relsMedia: ISlideRelMedia[]
	slideLayout: ISlideLayout
	slideNumberObj?: ISlideNumber // FIXME rename
	totalMediaRels:number
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
