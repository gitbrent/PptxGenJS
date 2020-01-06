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
    TableRow
} from './core-interfaces'

import * as genObj from './gen-objects'
import { createImageConfig } from './gen-utils'
import Relations from './relations'
import TextElement from './elements/text'
import ShapeElement from './elements/simple-shape'
import ImageElement from './elements/image'
import ChartElement from './elements/chart'
import SlideNumberElement from './elements/slide-number'
import TableElement from './elements/table'

export default class Slide {
    private _bkgd: string
    private _color: string
    private relations

    public addSlide: Function
    public getSlide: Function
    public presLayout: ILayout
    public name: string
    public number: number
    public data: ISlideObject[]
    public slideLayout: ISlideLayout

    constructor(params: {
        addSlide: Function
        getSlide: Function
        presLayout: ILayout
        slideNumber: number
        slideLayout?: ISlideLayout
    }) {
        this.addSlide = params.addSlide
        this.getSlide = params.getSlide
        this.presLayout = params.presLayout
        this.name = 'Slide ' + params.slideNumber
        this.number = params.slideNumber
        this.data = []
        this.slideLayout = params.slideLayout || null
        this.relations = new Relations()
    }

    public get rels() {
        return this.relations.rels
    }

    public get relsChart() {
        return this.relations.relsChart
    }

    public get relsMedia() {
        return this.relations.relsMedia
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

    slideNumber(value) {
        return this.addSlideNumber(value)
    }

    addSlideNumber(value) {
        this.data.push(new SlideNumberElement(value))
        return this
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
    addChart(
        type: CHART_TYPE_NAMES | IChartMulti[],
        data: [],
        options?: IChartOpts
    ): Slide {
        this.data.push(new ChartElement(type, data, options, this.relations))
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
        if (!options.path && !options.data) {
            console.error(
                "ERROR: `addImage()` requires either 'data' or 'path' parameter!"
            )
            return null
        } else if (
            options.data &&
            options.data.toLowerCase().indexOf('base64,') === -1
        ) {
            console.error(
                "ERROR: Image `data` value lacks a base64 header! Ex: 'image/png;base64,NMP[...]')"
            )
            return null
        }

        this.data.push(new ImageElement(options, this.relations))
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

        this.data.push(
            new TableElement(
                this,
                arrTabRows,
                options,
                this.slideLayout,
                this.presLayout,
                this.addSlide,
                this.getSlide
            )
        )
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
        this.data.push(new TextElement(text, options, this.relations))
        return this
        //genObj.addTextDefinition(this, text, options, false)
        //return this
    }
}
