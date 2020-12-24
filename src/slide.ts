/**
 * PptxGenJS: Slide Class
 */

import { CHART_NAME, SHAPE_NAME } from './core-enums'
import {
	BackgroundProps,
	HexColor,
	IChartMulti,
	IChartOpts,
	IChartOptsLib,
	ImageProps,
	PresLayout,
	SlideLayout,
	SlideNumberProps,
	ISlideObject,
	ISlideRel,
	ISlideRelChart,
	ISlideRelMedia,
	TableProps,
	TextProps,
	TextPropsOptions,
	MediaProps,
	ShapeProps,
	TableRow,
} from './core-interfaces'
import * as genObj from './gen-objects'

export default class Slide {
	private _setSlideNum: Function

	public addSlide: Function
	public getSlide: Function
	public _name: string
	public _presLayout: PresLayout
	public _rels: ISlideRel[]
	public _relsChart: ISlideRelChart[]
	public _relsMedia: ISlideRelMedia[]
	public _rId: number
	public _slideId: number
	public _slideLayout: SlideLayout
	public _slideNum: number
	public _slideNumberProps: SlideNumberProps
	public _slideObjects: ISlideObject[]

	constructor(params: {
		addSlide: Function
		getSlide: Function
		presLayout: PresLayout
		setSlideNum: Function
		slideId: number
		slideRId: number
		slideNumber: number
		slideLayout?: SlideLayout
	}) {
		this.addSlide = params.addSlide
		this.getSlide = params.getSlide
		this._name = 'Slide ' + params.slideNumber
		this._presLayout = params.presLayout
		this._rId = params.slideRId
		this._rels = []
		this._relsChart = []
		this._relsMedia = []
		this._setSlideNum = params.setSlideNum
		this._slideId = params.slideId
		this._slideLayout = params.slideLayout || null
		this._slideNum = params.slideNumber
		this._slideObjects = []
		/** NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		 * `defineSlideMaster` and `addNewSlide.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
		 * so, lastly, add to the Slide now.
		 */
		this._slideNumberProps = this._slideLayout && this._slideLayout._slideNumberProps ? this._slideLayout._slideNumberProps : null
	}

	/**
	 * Background color
	 * @type {string}
	 * @deprecated in v3.3.0 - use `background` instead
	 */
	private _bkgd: string
	public set bkgd(value: string) {
		this._bkgd = value
	}
	public get bkgd(): string {
		return this._bkgd
	}

	/**
	 * Background color or image
	 * @type {BackgroundProps}
	 * @example solid color `background: {fill:'FF0000'}
	 * @example base64 `background: {data:'image/png;base64,ABC[...]123'}`
	 * @example url  `background: {path:'https://some.url/image.jpg'}`
	 * @since v3.3.0
	 */
	private _background: BackgroundProps
	public set background(value: BackgroundProps) {
		genObj.addBackgroundDefinition(value, this)
	}
	public get background(): BackgroundProps {
		return this._background
	}

	/**
	 * Default font color
	 * @type {HexColor}
	 */
	private _color: HexColor
	public set color(value: HexColor) {
		this._color = value
	}
	public get color(): HexColor {
		return this._color
	}

	/**
	 * @type {boolean}
	 */
	private _hidden: boolean
	public set hidden(value: boolean) {
		this._hidden = value
	}
	public get hidden(): boolean {
		return this._hidden
	}

	/**
	 * @type {SlideNumberProps}
	 */
	public set slideNumber(value: SlideNumberProps) {
		// NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		this._slideNumberProps = value
		this._setSlideNum(value)
	}
	public get slideNumber(): SlideNumberProps {
		return this._slideNumberProps
	}

	/**
	 * Add chart to Slide
	 * @param {CHART_NAME|IChartMulti[]} type - chart type
	 * @param {object[]} data - data object
	 * @param {IChartOpts} options - chart options
	 * @return {Slide} this Slide
	 */
	addChart(type: CHART_NAME | IChartMulti[], data: any[], options?: IChartOpts): Slide {
		// FUTURE: TODO-VERSION-4: Remove first arg - only take data and opts, with "type" required on opts
		// Set `_type` on IChartOptsLib as its what is used as object is passed around
		let optionsWithType: IChartOptsLib = options || {}
		optionsWithType._type = type
		genObj.addChartDefinition(this, type, data, options)
		return this
	}

	/**
	 * Add image to Slide
	 * @param {ImageProps} options - image options
	 * @return {Slide} this Slide
	 */
	addImage(options: ImageProps): Slide {
		genObj.addImageDefinition(this, options)
		return this
	}

	/**
	 * Add media (audio/video) to Slide
	 * @param {MediaProps} options - media options
	 * @return {Slide} this Slide
	 */
	addMedia(options: MediaProps): Slide {
		genObj.addMediaDefinition(this, options)
		return this
	}

	/**
	 * Add speaker notes to Slide
	 * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
	 * @param {string} notes - notes to add to slide
	 * @return {Slide} this Slide
	 */
	addNotes(notes: string): Slide {
		genObj.addNotesDefinition(this, notes)
		return this
	}

	/**
	 * Add shape to Slide
	 * @param {SHAPE_NAME} shapeName - shape name
	 * @param {ShapeProps} options - shape options
	 * @return {Slide} this Slide
	 */
	addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Slide {
		// NOTE: As of v3.1.0, <script> users are passing the old shape object from the shapes file (orig to the project)
		// But React/TypeScript users are passing the shapeName from an enum, which is a simple string, so lets cast
		// <script./> => `pptx.shapes.RECTANGLE` [string] "rect" ... shapeName['name'] = 'rect'
		// TypeScript => `pptxgen.shapes.RECTANGLE` [string] "rect" ... shapeName = 'rect'
		//let shapeNameDecode = typeof shapeName === 'object' && shapeName['name'] ? shapeName['name'] : shapeName
		genObj.addShapeDefinition(this, shapeName, options)
		return this
	}

	/**
	 * Add table to Slide
	 * @param {TableRow[]} tableRows - table rows
	 * @param {TableProps} options - table options
	 * @return {Slide} this Slide
	 */
	addTable(tableRows: TableRow[], options?: TableProps): Slide {
		// FUTURE: we pass `this` - we dont need to pass layouts - they can be read from this!
		genObj.addTableDefinition(this, tableRows, options, this._slideLayout, this._presLayout, this.addSlide, this.getSlide)
		return this
	}

	/**
	 * Add text to Slide
	 * @param {string|TextProps[]} text - text string or complex object
	 * @param {TextPropsOptions} options - text options
	 * @return {Slide} this Slide
	 */
	addText(text: string | TextProps[], options?: TextPropsOptions): Slide {
		genObj.addTextDefinition(this, text, options, false)
		return this
	}
}
