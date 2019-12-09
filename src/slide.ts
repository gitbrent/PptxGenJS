/**
 * PptxGenJS Slide Class
 */

import { CHART_TYPE_NAMES, SLIDE_OBJECT_TYPES } from './core-enums'
import {
	IChartMulti,
	IChartOpts,
	IImageOpts,
	ILayout,
	IMediaOpts,
	ISlideLayout,
	ISlideNumber,
	ISlideRel,
	ISlideRelChart,
	ISlideRelMedia,
	ISlideObject,
	IShape,
	IShapeOptions,
	ITableOptions,
	IText,
	ITextOpts,
	TableRow,
} from './core-interfaces'

import * as genObj from './gen-objects'
import TextElement from './elements/text'
import ShapeElement from './elements/simple-shape'

export default class Slide {
	private _bkgd: string
	private _color: string
	private _setSlideNum: Function
	private _slideNumber: ISlideNumber

	public addSlide: Function
	public getSlide: Function
	public presLayout: ILayout
	public name: string
	public number: number
	public data: ISlideObject[]
	public rels: ISlideRel[]
	public relsChart: ISlideRelChart[]
	public relsMedia: ISlideRelMedia[]
	public slideLayout: ISlideLayout
	public slideNumberObj: ISlideNumber

	constructor(params: { addSlide: Function; getSlide: Function; presLayout: ILayout; setSlideNum: Function; slideNumber: number; slideLayout?: ISlideLayout }) {
		this.addSlide = params.addSlide
		this.getSlide = params.getSlide
		this.presLayout = params.presLayout
		this._setSlideNum = params.setSlideNum
		this.name = 'Slide ' + params.slideNumber
		this.number = params.slideNumber
		this.data = []
		this.rels = []
		this.relsChart = []
		this.relsMedia = []
		this.slideLayout = params.slideLayout || null
		// NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		// `defineSlideMaster` and `addNewSlide.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
		// so, lastly, add to the Slide now.
		this.slideNumberObj = this.slideLayout && this.slideLayout.slideNumberObj ? this.slideLayout.slideNumberObj : null
	}

	private _registerLink(data, target) {
		const relId = this.rels.length + target.relsChart.length + target.relsMedia.length + 1
		this.rels.push({
			type: SLIDE_OBJECT_TYPES.hyperlink,
			data,
			rId: relId,
			Target: target,
		})

		return relId
	}

	// TODO: add comments (also add to index.d.ts)
	public set bkgd(value: string) {
		this._bkgd = value
	}
	public get bkgd(): string {
		return this._bkgd
	}

	// TODO: add comments (also add to index.d.ts)
	public set color(value: string) {
		this._color = value
	}
	public get color(): string {
		return this._color
	}

	// TODO: add comments (also add to index.d.ts)
	public set slideNumber(value: ISlideNumber) {
		// NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		this.slideNumberObj = value
		this._slideNumber = value
		this._setSlideNum(value)
	}
	public get slideNumber(): ISlideNumber {
		return this._slideNumber
	}

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
	addChart(type: CHART_TYPE_NAMES | IChartMulti[], data: [], options?: IChartOpts): Slide {
		genObj.addChartDefinition(this, type, data, options)
		return this
	}

	/**
	 * Add Image object
	 * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
	 * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
	 * @param {IImageOpts} options - image options
	 * @return {Slide} this class
	 */
	addImage(options: IImageOpts): Slide {
		genObj.addImageDefinition(this, options)
		return this
	}

	/**
	 * Add Media (audio/video) object
	 * @param {IMediaOpts} options - media options
	 * @return {Slide} this class
	 */
	addMedia(options: IMediaOpts): Slide {
		genObj.addMediaDefinition(this, options)
		return this
	}

	/**
	 * Add Speaker Notes to Slide
	 * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
	 * @param {string} notes - notes to add to slide
	 * @return {Slide} this class
	 */
	addNotes(notes: string): Slide {
		genObj.addNotesDefinition(this, notes)
		return this
	}

	/**
	 * Add shape object to Slide
	 * @param {IShape} shape - shape object
	 * @param {IShapeOptions} options - shape options
	 * @return {Slide} this class
	 */
	addShape(shape: IShape, options?: IShapeOptions): Slide {
		this.data.push(new ShapeElement(shape, options))
		return this
	}

	/**
	 * Add shape object to Slide
	 * @note can be recursive
	 * @param {TableRow[]} arrTabRows - table rows
	 * @param {ITableOptions} options - table options
	 * @return {Slide} this class
	 */
	addTable(arrTabRows: TableRow[], options?: ITableOptions): Slide {
		// FIXME: TODO-3: we pass `this` - we dont need to pass layouts - they can be read from this!
		genObj.addTableDefinition(this, arrTabRows, options, this.slideLayout, this.presLayout, this.addSlide, this.getSlide)
		return this
	}

	/**
	 * Add text object to Slide
	 * @param {string|IText[]} text - text string or complex object
	 * @param {ITextOpts} options - text options
	 * @return {Slide} this class
	 * @since: 1.0.0
	 */
	addText(text: string | IText[], options?: ITextOpts): Slide {
		this.data.push(new TextElement(text, options, this._registerLink.bind(this)))
		return this
		//genObj.addTextDefinition(this, text, options, false)
		//return this
	}
}
