// Type definitions for pptxgenjs 3.0.0
// Project: https://gitbrent.github.io/PptxGenJS/
// Definitions by: Brent Ely <https://github.com/gitbrent/>
//                 Michael Beaumont <https://github.com/michaelbeaumont>
//                 Nicholas Tietz-Sokolsky <https://github.com/ntietz>
//                 David Adams <https://github.com/iota-pi>
// TypeScript Version: 3.x

export as namespace PptxGenJS

export = PptxGenJS

// TESTING:
//let newppt = new PptxGenJS()
//newppt.addSlide().

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
     * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
     * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
     * @param {IUserLayout} layout - an object with user-defined w/h
     */
    defineLayout(layout: PptxGenJS.IUserLayout): void;
	/**
	 * Adds a new slide master [layout] to the Presentation
	 * @param {ISlideMasterOptions} slideMasterOpts - layout definition
	 */
	defineSlideMaster(slideMasterOpts: PptxGenJS.ISlideMasterOptions): void
	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @note `verbose` option is undocumented; used for verbose output of layout process
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
	// ENUMS
	export type JSZIP_OUTPUT_TYPE = 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
	// import { CHART_TYPE_NAMES, SLIDE_OBJECT_TYPES, TEXT_HALIGN, TEXT_VALIGN, PLACEHOLDER_TYPES } from './core-enums'
	export type CHART_TYPE_NAMES = 'area' | 'bar' | 'bar3D' | 'bubble' | 'doughnut' | 'line' | 'pie' | 'radar' | 'scatter'
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
	type MediaType = 'audio' | 'online' | 'video'

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
		catAxisMajorTickMark?: ChartAxisTickMark
		catAxisMinorTickMark?: ChartAxisTickMark
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
		border?: IBorderOptions
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
		serAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
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
		dataBorder?: IBorderOptions
		lineSize?: number
		lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'
		radarStyle?: string
		shadow?: IShadowOptions
		catAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
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
		valAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		valAxisLabelRotate?: number
		valAxisLabelFontSize?: number
		valAxisLabelFontBold?: boolean
		valAxisLabelColor?: string
		valAxisLabelFontFace?: string
		valAxisMajorUnit?: number
		valAxisMajorTickMark?: ChartAxisTickMark
		valAxisMinorTickMark?: ChartAxisTickMark
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
		rotate?: number
	}
	export interface IMediaOpts extends PositionOptions, OptsDataOrPath {
		link: string
		onlineVideoLink?: string
		type?: MediaType
	}

	export interface IShape {
		displayName: string
		name: string
		avLst: { [key: string]: number }
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
	export interface ITableCellOpts {
		autoPageCharWeight?: number
		autoPageLineWeight?: number
		align?: HAlign
		bold?: boolean
		border?: IBorderOptions | [IBorderOptions, IBorderOptions, IBorderOptions, IBorderOptions]
		color?: Color
		colspan?: number
		fill?: ShapeFill
		fontFace?: string
		fontSize?: number
		margin?: Margin
		rowspan?: number
		valign?: VAlign
	}
	export interface ITableOptions extends PositionOptions {
		align?: HAlign
		autoPage?: boolean
		autoPageCharWeight?: number
		autoPageLineWeight?: number
		border?: IBorderOptions | [IBorderOptions, IBorderOptions, IBorderOptions, IBorderOptions]
		color?: Color
		colspan?: number
		colW?: number | number[]
		fill?: Color
		fontSize?: number
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
		bullet?: boolean | { type?: string; code?: string; style?: string; startAt?: number }
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
		lineIdx?: number
		lineSize?: number
		lineSpacing?: number
		margin?: number
		outline?: { color: Color; size: number }
		paraSpaceAfter?: number
		paraSpaceBefore?: number
		placeholder?: string
		rotate?: number // (degree * 60,000)
		rtlMode?: boolean
		shadow?: IShadowOptions
		shape?: IShape
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
	    name: string;
	    width: number;
	    height: number;
	}
	export interface ISlideNumber extends PositionOptions {
		fontFace?: string
		fontSize?: number
		color?: string
	}
	export interface ISlideMasterOptions {
		title: string
		height?: number
		width?: number
		margin?: Margin
		bkgd?: string
		objects?: [

				| { chart: {} }
				| { image: {} }
				| { line: {} }
				| { rect: {} }
				| { text: { options: ITextOpts } }
				| { placeholder: { options: ISlideMstrObjPlchldrOpts; text?: string } }
		]
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
		shape?: IShape
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

	export class Slide {
		bkgd: string
		color: string
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
		addChart(type: CHART_TYPE_NAMES | IChartMulti[], data: [], options?: IChartOpts): Slide
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
		 * @param {IShape} shape - shape object
		 * @param {IShapeOptions} options - shape options
		 * @return {Slide} this class
		 */
		addShape(shape: IShape, options?: IShapeOptions): Slide
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
