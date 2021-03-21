/**
 * PptxGenJS: Slide Class
 */
import { CHART_NAME, SHAPE_NAME } from './core-enums';
import { BackgroundProps, HexColor, IChartMulti, IChartOpts, ImageProps, PresLayout, SlideLayout, SlideNumberProps, ISlideObject, ISlideRel, ISlideRelChart, ISlideRelMedia, TableProps, TextProps, TextPropsOptions, MediaProps, ShapeProps, TableRow } from './core-interfaces';
export default class Slide {
    private _setSlideNum;
    addSlide: Function;
    getSlide: Function;
    _name: string;
    _presLayout: PresLayout;
    _rels: ISlideRel[];
    _relsChart: ISlideRelChart[];
    _relsMedia: ISlideRelMedia[];
    _rId: number;
    _slideId: number;
    _slideLayout: SlideLayout;
    _slideNum: number;
    _slideNumberProps: SlideNumberProps;
    _slideObjects: ISlideObject[];
    constructor(params: {
        addSlide: Function;
        getSlide: Function;
        presLayout: PresLayout;
        setSlideNum: Function;
        slideId: number;
        slideRId: number;
        slideNumber: number;
        slideLayout?: SlideLayout;
    });
    /**
     * Background color
     * @type {string}
     * @deprecated in v3.3.0 - use `background` instead
     */
    private _bkgd;
    set bkgd(value: string);
    get bkgd(): string;
    /**
     * Background color or image
     * @type {BackgroundProps}
     * @example solid color `background: {fill:'FF0000'}
     * @example base64 `background: {data:'image/png;base64,ABC[...]123'}`
     * @example url  `background: {path:'https://some.url/image.jpg'}`
     * @since v3.3.0
     */
    private _background;
    set background(value: BackgroundProps);
    get background(): BackgroundProps;
    /**
     * Default font color
     * @type {HexColor}
     */
    private _color;
    set color(value: HexColor);
    get color(): HexColor;
    /**
     * @type {boolean}
     */
    private _hidden;
    set hidden(value: boolean);
    get hidden(): boolean;
    /**
     * @type {SlideNumberProps}
     */
    set slideNumber(value: SlideNumberProps);
    get slideNumber(): SlideNumberProps;
    /**
     * Add chart to Slide
     * @param {CHART_NAME|IChartMulti[]} type - chart type
     * @param {object[]} data - data object
     * @param {IChartOpts} options - chart options
     * @return {Slide} this Slide
     */
    addChart(type: CHART_NAME | IChartMulti[], data: any[], options?: IChartOpts): Slide;
    /**
     * Add image to Slide
     * @param {ImageProps} options - image options
     * @return {Slide} this Slide
     */
    addImage(options: ImageProps): Slide;
    /**
     * Add media (audio/video) to Slide
     * @param {MediaProps} options - media options
     * @return {Slide} this Slide
     */
    addMedia(options: MediaProps): Slide;
    /**
     * Add speaker notes to Slide
     * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
     * @param {string} notes - notes to add to slide
     * @return {Slide} this Slide
     */
    addNotes(notes: string): Slide;
    /**
     * Add shape to Slide
     * @param {SHAPE_NAME} shapeName - shape name
     * @param {ShapeProps} options - shape options
     * @return {Slide} this Slide
     */
    addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Slide;
    /**
     * Add table to Slide
     * @param {TableRow[]} tableRows - table rows
     * @param {TableProps} options - table options
     * @return {Slide} this Slide
     */
    addTable(tableRows: TableRow[], options?: TableProps): Slide;
    /**
     * Add text to Slide
     * @param {string|TextProps[]} text - text string or complex object
     * @param {TextPropsOptions} options - text options
     * @return {Slide} this Slide
     */
    addText(text: string | TextProps[], options?: TextPropsOptions): Slide;
}
