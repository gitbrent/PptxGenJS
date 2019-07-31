/**
 * Slide Class
 */

import { CHART_TYPE_NAMES } from './core-enums'
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
	Shape,
	ShapeOptions,
    TableOptions,
	IText,
	ITextOpts,
} from './core-interfaces'
import * as genObj from './gen-objects'

export default class Slide {
	private _bkgd: string
	private _color: string
	private _setSlideNum: Function
	private _slideNumber: ISlideNumber

	public presLayout: ILayout
	public name: string
	public number: number
	public data: ISlideObject[]
	public rels: ISlideRel[]
	public relsChart: ISlideRelChart[]
	public relsMedia: ISlideRelMedia[]
	public slideLayout: ISlideLayout

	constructor(params: { presLayout: ILayout; setSlideNum: Function; slideNumber: number; slideLayout?: ISlideLayout }) {
		this.presLayout = params.presLayout
		this._setSlideNum = params.setSlideNum
		this.name = 'Slide ' + params.slideNumber
		this.number = params.slideNumber
		this.data = []
		this.rels = []
		this.relsChart = []
		this.relsMedia = []
		this.slideNumber = null
		this.slideLayout = params.slideLayout || null

		// NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		// `defineSlideMaster` and `addNewSlide.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
		// so, lastly, add to the Slide now.
		if (this.slideLayout && this.slideLayout.slideNumberObj && !this._slideNumber) this.slideNumber = this.slideLayout.slideNumberObj
	}

	// ==========================================================================
	// PUBLIC METHODS:
	// ==========================================================================

	public set bkgd(value: string) {
		this._bkgd = value
	}
	public get bkgd(): string {
		return this._bkgd
	}

	public set color(value: string) {
		this._color = value
	}
	public get color(): string {
		return this._color
	}

	public set slideNumber(value: ISlideNumber) {
		// NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		this._slideNumber = value
		this._setSlideNum(value)
	}
	public get slideNumber(): ISlideNumber {
		return this._slideNumber
	}

	/**
	 * Generate the chart based on input data.
	 * @see OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
	 *
	 * @param {CHART_TYPE_NAMES|IChartMulti[]} `type` - chart type
	 * @param {object[]} `data` - a JSON object with follow the following format
	 * @param {IChartOpts} `opt` - options
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
	 */
	addChart(type: CHART_TYPE_NAMES | IChartMulti[], data: [], opt?: IChartOpts) {
		genObj.addChartDefinition(type, data, opt, this)
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
		genObj.addImageDefinition(options, this)
		return this
	}

	addMedia(opt: IMediaOpts) {
		genObj.addMediaDefinition(this, opt)
		return this
	}

	addNotes(notes, opt) {
		genObj.addNotesDefinition(notes, opt, this)
		return this
	}

	/**
	 * Add shape object to Slide
	 * @param {Shape} shape - shape object
	 * @param {ShapeOptions} options - shape options
	 * @return {Slide} this class
	 */
	addShape(shape: Shape, opt: ShapeOptions): Slide {
		genObj.addShapeDefinition(shape, opt, this)
		return this
	}

	// RECURSIVE: (sometimes)
	// FIXME: TODO: dont forget to update the "this.color" refs below to "target.slide.color"!!!
	addTable(arrTabRows, inOpt:TableOptions): Slide {
		// FIXME: TODO-3: we pass `this` - we dont need to pass layouts - they can be read from this!
		genObj.addTableDefinition(this, arrTabRows, inOpt, this.slideLayout, this.presLayout)
		return this
	}

	/**
	 * Add text object to Slide
	 * @param {string|IText[]} text - text string or complex object
	 * @param {ITextOpts} options - text options
	 * @return {Slide} this class
	 * @since: 1.0.0
	 */
	addText(text: string | IText[], options: ITextOpts): Slide {
		genObj.addTextDefinition(text, options, this, false)
		return this
	}
}
