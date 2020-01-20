/**
 * PptxGenJS Slide Class
 */
import { CHART_TYPE_NAMES } from './core-enums';
import { IChartMulti, IChartOpts, IImageOpts, ILayout, IMediaOpts, IShape, IShapeOptions, ITableOptions, IText, ITextOpts, TableRow } from './core-interfaces';
import Relations from './relations';
import { Master } from './slideLayouts';
import GroupElement from './elements/group';
import ElementInterface from './elements/element-interface';
export default class Slide {
    private _bkgd;
    private _color;
    relations: Relations;
    addSlide: Function;
    getSlide: Function;
    presLayout?: ILayout;
    name: string;
    number: number;
    data: ElementInterface[];
    slideLayout: Master;
    notes: string[];
    hidden: boolean;
    constructor(params: {
        addSlide?: Function;
        getSlide?: Function;
        presLayout: ILayout;
        slideNumber: number;
        slideLayout?: Master;
    });
    readonly rels: any[];
    readonly relsChart: any[];
    readonly relsMedia: any[];
    bkgd: string;
    color: string;
    slideNumber(value: any): this;
    addSlideNumber(value: any): this;
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
    addChart(type: CHART_TYPE_NAMES | IChartMulti[], data: [], options?: IChartOpts): Slide;
    /**
     * Add Image object
     * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
     * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
     * @param {IImageOpts} options - image options
     * @return {Slide} this class
     */
    addImage(options: IImageOpts): Slide;
    /**
     * Add Media (audio/video) object
     * @param {IMediaOpts} options - media options
     * @return {Slide} this class
     */
    addMedia(options: IMediaOpts): Slide;
    /**
     * Add Speaker Notes to Slide
     * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
     * @param {string} notes - notes to add to slide
     * @return {Slide} this class
     */
    addNotes(notes: string): Slide;
    /**
     * Add shape object to Slide
     * @param {IShape} shape - shape object
     * @param {IShapeOptions} options - shape options
     * @return {Slide} this class
     */
    addShape(shape: IShape, options?: IShapeOptions): Slide;
    /**
     * Add shape object to Slide
     * @note can be recursive
     * @param {TableRow[]} arrTabRows - table rows
     * @param {ITableOptions} options - table options
     * @return {Slide} this class
     */
    addTable(arrTabRows: TableRow[], options?: ITableOptions): Slide;
    /**
     * Add text object to Slide
     * @param {string|IText[]} text - text string or complex object
     * @param {ITextOpts} options - text options
     * @return {Slide} this class
     * @since: 1.0.0
     */
    addText(text: string | IText[], options?: ITextOpts): Slide;
    newGroup(): GroupElement;
}
