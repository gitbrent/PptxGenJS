/**
 * PptxGenJS: Slide Object Generators
 */
import { CHART_NAME, SHAPE_NAME } from './core-enums';
import { AddSlideProps, BackgroundProps, GroupProps, IChartMulti, IChartOptsLib, IOptsChartData, ImageProps, MediaProps, PresLayout, PresSlide, ShapeProps, SlideLayout, SlideMasterProps, TableProps, TableRow, TextProps, TextPropsOptions } from './core-interfaces';
/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * @param {SlideMasterProps} props - slide definition
 * @param {PresSlide|SlideLayout} target - empty slide object that should be updated by the passed definition
 */
export declare function createSlideMaster(props: SlideMasterProps, target: SlideLayout): void;
/**
 * Generate the chart based on input data.
 * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
 *
 * @param {CHART_NAME | IChartMulti[]} `type` should belong to: 'column', 'pie'
 * @param {[]} `data` a JSON object with follow the following format
 * @param {IChartOptsLib} `opt` chart options
 * @param {PresSlide} `target` slide object that the chart will be added to
 * @return {object} chart object
 * {
 *    title: 'eSurvey chart',
 *    data: [
 *        {
 *            name: 'Income',
 *            labels: ['2005', '2006', '2007', '2008', '2009'],
 *            values: [23.5, 26.2, 30.1, 29.5, 24.6]
 *        },
 *        {
 *            name: 'Expense',
 *            labels: ['2005', '2006', '2007', '2008', '2009'],
 *            values: [18.1, 22.8, 23.9, 25.1, 25]
 *        }
 *    ]
 * }
 */
export declare function addChartDefinition(target: PresSlide, type: CHART_NAME | IChartMulti[], data: IOptsChartData[], opt: IChartOptsLib): object;
/**
 * Adds an image object to a slide definition.
 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
 * @param {ImageProps} `opt` - object containing `path`/`data`, `x`, `y`, etc.
 * @param {PresSlide} `target` - slide that the image should be added to (if not specified as the 2nd arg)
 * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
 * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
 */
export declare function addImageDefinition(target: PresSlide, opt: ImageProps): void;
/**
 * Adds a media object to a slide definition.
 * @param {PresSlide} `target` - slide object that the media will be added to
 * @param {MediaProps} `opt` - media options
 */
export declare function addMediaDefinition(target: PresSlide, opt: MediaProps): void;
/**
 * Adds Notes to a slide.
 * @param {PresSlide} `target` slide object
 * @param {string} `notes`
 * @since 2.3.0
 */
export declare function addNotesDefinition(target: PresSlide, notes: string): void;
/**
 * Adds a shape object to a slide definition.
 * @param {PresSlide} target slide object that the shape should be added to
 * @param {SHAPE_NAME} shapeName shape name
 * @param {ShapeProps} opts shape options
 */
export declare function addShapeDefinition(target: PresSlide, shapeName: SHAPE_NAME, opts: ShapeProps): void;
/**
 * Adds a table object to a slide definition.
 * @param {PresSlide} target - slide object that the table should be added to
 * @param {TableRow[]} tableRows - table data
 * @param {TableProps} options - table options
 * @param {SlideLayout} slideLayout - Slide layout
 * @param {PresLayout} presLayout - Presentation layout
 * @param {Function} addSlide - method
 * @param {Function} getSlide - method
 */
export declare function addTableDefinition(target: PresSlide, tableRows: TableRow[], options: TableProps, slideLayout: SlideLayout, presLayout: PresLayout, addSlide: (options?: AddSlideProps) => PresSlide, getSlide: (slideNumber: number) => PresSlide): PresSlide[];
/**
 * Adds a picture placeholder object to a slide/layout definition.
 * Creates a native <p:pic> element with empty <p:blipFill/> that acts as an image placeholder.
 * @param {PresSlide} target - slide/layout object that the placeholder should be added to
 * @param {object} props - picture placeholder properties (x, y, w, h, name, idx)
 * @param {number} placeholderIdx - the placeholder index
 * @since: 1.0.6
 */
export declare function addPicturePlaceholderDefinition(target: PresSlide, props: {
    x: number;
    y: number;
    w: number;
    h: number;
    name?: string;
    idx?: number;
}, placeholderIdx: number): void;
/**
 * Adds a text object to a slide definition.
 * @param {PresSlide} target - slide object that the text should be added to
 * @param {string|TextProps[]} text text string or object
 * @param {TextPropsOptions} opts text options
 * @param {boolean} isPlaceholder whether this a placeholder object
 * @since: 1.0.0
 */
export declare function addTextDefinition(target: PresSlide, text: TextProps[], opts: TextPropsOptions, isPlaceholder: boolean): void;
/**
 * Adds placeholder objects to slide
 * @param {PresSlide} slide - slide object containing layouts
 */
export declare function addPlaceholdersToSlideLayouts(slide: PresSlide): void;
/**
 * Adds a background image or color to a slide definition.
 * @param {BackgroundProps} props - color string or an object with image definition
 * @param {PresSlide} target - slide object that the background is set to
 */
export declare function addBackgroundDefinition(props: BackgroundProps, target: SlideLayout): void;
/**
 * Adds a group object to the slide
 * @param {PresSlide} target - slide to add group to
 * @param {GroupProps} opts - group options including children
 */
export declare function addGroupDefinition(target: PresSlide, opts: GroupProps): void;
