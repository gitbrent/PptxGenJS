/**
 * Slide Class
 */

import { CHART_TYPE_NAMES } from './enums'
import { IMediaOpts, ISlideNumber, ISlideLayout, ILayout, IChartMulti, IChartOpts } from './interfaces'
import * as genObj from './gen-objects'

export default class Slide {
	private _bkgd: string
	private _color: string
	private _slideNumber: ISlideNumber
	private _presLayout: ILayout

	public name: string
	public number: number
	public data: []
	public rels: any[]
	public relsChart: any[]
	public relsMedia: any[]
	public layoutName: any
	public slideLayout: ISlideLayout

	// TODO: slide.title (they're all "PowerPoint Presenation" now!) 20190712

	constructor(params: { presLayout: ILayout; slideNumber: number; slideLayout?: ISlideLayout }) {
		this._presLayout = params.presLayout
		this.name = 'Slide ' + params.slideNumber
		this.number = params.slideNumber
		this.data = []
		this.rels = []
		this.relsChart = []
		this.relsMedia = []
		this.slideNumber = null
		this.layoutName = params.slideLayout.name || '[ default ]'
		this.slideLayout = params.slideLayout || null

		// NOTE: Slide Numbers: In order for Slide Numbers to work normally, they need to be in all 3 files: master/layout/slide
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

	public set slideNumber(inObj: ISlideNumber) {
		// A:
		this._slideNumber = inObj

		// TODO: commented these as we need to reach up to do these (create some setters on PptxGen??)
		// B: Add slideNumber to slideMaster1.xml
		//if (!this.masterSlide.slideNumber) this.masterSlide.slideNumber = inObj

		// C: Add slideNumber to `BLANK` (default) layout
		//if (!this.slideLayouts[0].slideNumber) this.slideLayouts[0].slideNumber = inObj
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
	 * NOTE: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
	 * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
	 */
	addImage(objImage) {
		// TODO-3: create `IImageOpts` (name,path,w,rotate,etc.)
		genObj.addImageDefinition(objImage, this)
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

	addShape(shape, opt) {
		genObj.addShapeDefinition(shape, opt, this)
		return this
	}

	// RECURSIVE: (sometimes)
	// TODO: dont forget to update the "this.color" refs below to "target.slide.color"!!!
	addTable(arrTabRows, inOpt) {
		genObj.addTableDefinition(this, arrTabRows, inOpt, this.slideLayout, this._presLayout)
		return this
	}

	/**
	 * Add text object to Slide
	 *
	 * @param {object|string} `text` - text string or complex object
	 * @param {object} `options` - text options
	 * @since: 1.0.0
	 */
	addText(text, options) {
		genObj.addTextDefinition(text, options, this, false)
		return this
	}
}
