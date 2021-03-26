import { CHART_NAME, SHAPE_NAME, SLIDE_OBJECT_TYPES } from './core-enums'
import {
	IChartMulti,
	IChartOpts,
	IChartOptsLib,
	ImageProps,
	ISlideObject,
	MediaProps,
	PresSlide,
	ShapeProps,
	TableProps,
	TableRow,
	TextProps,
	TextPropsOptions,
} from './core-interfaces'
import * as genObj from './gen-objects'

export class Group {
	public _slideObjects: ISlideObject[]
	public _slide: PresSlide
	public addSlide: Function
	public getSlide: Function

	constructor(params: { slide: PresSlide; addSlide: Function; getSlide }) {
		this._slideObjects = []
		this._slide = params.slide
		this.addSlide = params.addSlide
		this.getSlide = params.getSlide
	}

	/**
	 * Add chart to Group
	 * @param {CHART_NAME|IChartMulti[]} type - chart type
	 * @param {object[]} data - data object
	 * @param {IChartOpts} options - chart options
	 * @return {Group} this Group
	 */
	addChart(type: CHART_NAME | IChartMulti[], data: any[], options?: IChartOpts): Group {
		// FUTURE: TODO-VERSION-4: Remove first arg - only take data and opts, with "type" required on opts
		// Set `_type` on IChartOptsLib as its what is used as object is passed around
		let optionsWithType: IChartOptsLib = options || {}
		optionsWithType._type = type
		genObj.addChartDefinition(this, this._slide, type, data, options)
		return this
	}

	/**
	 * Add image to Group
	 * @param {ImageProps} options - image options
	 * @return {Group} this Group
	 */
	addImage(options: ImageProps): Group {
		genObj.addImageDefinition(this, this._slide, options)
		return this
	}

	/**
	 * Add media (audio/video) to Group
	 * @param {MediaProps} options - media options
	 * @return {Group} this Group
	 */
	addMedia(options: MediaProps): Group {
		genObj.addMediaDefinition(this, this._slide, options)
		return this
	}

	/**
	 * Add shape to Group
	 * @param {SHAPE_NAME} shapeName - shape name
	 * @param {ShapeProps} options - shape options
	 * @return {Group} this Group
	 */
	addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Group {
		// NOTE: As of v3.1.0, <script> users are passing the old shape object from the shapes file (orig to the project)
		// But React/TypeScript users are passing the shapeName from an enum, which is a simple string, so lets cast
		// <script./> => `pptx.shapes.RECTANGLE` [string] "rect" ... shapeName['name'] = 'rect'
		// TypeScript => `pptxgen.shapes.RECTANGLE` [string] "rect" ... shapeName = 'rect'
		//let shapeNameDecode = typeof shapeName === 'object' && shapeName['name'] ? shapeName['name'] : shapeName
		genObj.addShapeDefinition(this, this._slide, shapeName, options)
		return this
	}

	/**
	 * Add table to Group
	 * @param {TableRow[]} tableRows - table rows
	 * @param {TableProps} options - table options
	 * @return {Group} this Group
	 */
	addTable(tableRows: TableRow[], options?: TableProps): Group {
		// FUTURE: we pass `this` - we dont need to pass layouts - they can be read from this!
		genObj.addTableDefinition(this, this._slide, tableRows, options, this._slide._slideLayout, this._slide._presLayout, this.addSlide, this.getSlide)
		return this
	}

	/**
	 * Add text to Group
	 * @param {string|TextProps[]} text - text string or complex object
	 * @param {TextPropsOptions} options - text options
	 * @return {Group} this Group
	 */
	addText(text: string | TextProps[], options?: TextPropsOptions): Group {
		let textParam = typeof text === 'string' || typeof text === 'number' ? [{ text: text, options: options } as TextProps] : text
		genObj.addTextDefinition(this, this._slide, textParam, options, false)
		return this
	}

	addGroup(): Group {
		const group = new Group({
			slide: this._slide,
			addSlide: this.addSlide,
			getSlide: this.getSlide,
		})
		this._slideObjects.push({
			_type: SLIDE_OBJECT_TYPES.group,
			group,
		})
		return group
	}
}
