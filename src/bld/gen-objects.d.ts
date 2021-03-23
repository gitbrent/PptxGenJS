/**
 * PptxGenJS: Slide Object Generators
 */
import { CHART_NAME, SHAPE_NAME } from './core-enums';
import { BackgroundProps, IChartMulti, IChartOptsLib, ImageProps, MediaProps, PresLayout, PresSlide, ShapeProps, SlideLayout, SlideMasterProps, TableProps, TableRow, TextProps, TextPropsOptions } from './core-interfaces';
/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * @param {SlideMasterProps} slideDef - slide definition
 * @param {PresSlide|SlideLayout} target - empty slide object that should be updated by the passed definition
 */
export declare function createSlideObject(slideDef: SlideMasterProps, target: SlideLayout): void;
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
 *	}
 */
export declare function addChartDefinition(target: PresSlide, type: CHART_NAME | IChartMulti[], data: any[], opt: IChartOptsLib): object;
/**
 * Adds an image object to a slide definition.
 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
 * @param {ImageProps} `opt` - object containing `path`/`data`, `x`, `y`, etc.
 * @param {PresSlide} `target` - slide that the image should be added to (if not specified as the 2nd arg)
 * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
 * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
 */
export declare function addImageDefinition(target: PresSlide, opt: ImageProps): any;
/**
 * Adds a media object to a slide definition.
 * @param {PresSlide} `target` - slide object that the text will be added to
 * @param {MediaProps} `opt` - media options
 */
export declare function addMediaDefinition(target: PresSlide, opt: MediaProps): void;
/**
 * Adds Notes to a slide.
 * @param {String} `notes`
 * @param {Object} opt (*unused*)
 * @param {PresSlide} `target` slide object
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
export declare function addTableDefinition(target: PresSlide, tableRows: TableRow[], options: TableProps, slideLayout: SlideLayout, presLayout: PresLayout, addSlide: Function, getSlide: Function): void;
/**
 * Adds a text object to a slide definition.
 * @param {PresSlide} target - slide object that the text should be added to
 * @param {string|TextProps[]} text text string or object
 * @param {TextPropsOptions} opts text options
 * @param {boolean} isPlaceholder` is this a placeholder object
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
 * @param {BackgroundProps} bkg - color string or an object with image definition
 * @param {PresSlide} target - slide object that the background is set to
 */
export declare function addBackgroundDefinition(bkg: BackgroundProps, target: SlideLayout): void;
