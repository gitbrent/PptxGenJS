/**
 * PPTX Units are "DXA" (except for font sizing)
 * ....: There are 1440 DXA per inch. 1 inch is 72 points. 1 DXA is 1/20th's of a point (20 DXA is 1 point).
 * ....: There is also something called EMU's (914400 EMUs is 1 inch, 12700 EMUs is 1pt).
 * SEE: https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
 *
 * OBJECT LAYOUTS: 16x9 (10" x 5.625"), 16x10 (10" x 6.25"), 4x3 (10" x 7.5"), Wide (13.33" x 7.5") and Custom (any size)
 *
 * REFERENCES:
 * @see [Structure of a PresentationML document (Open XML SDK)](https://msdn.microsoft.com/en-us/library/office/gg278335.aspx)
 * @see [TableStyleId enumeration](https://msdn.microsoft.com/en-us/library/office/hh273476(v=office.14).aspx)
 */
import { CHART_TYPES, JSZIP_OUTPUT_TYPE, SCHEME_COLOR_NAMES, WRITE_OUTPUT_TYPE } from './core-enums';
import { ILayout, ISlideMasterOptions, ITableToSlidesOpts } from './core-interfaces';
import { PowerPointShapes } from './core-shapes';
import Slide from './slide';
import * as JSZip from 'jszip';
import { Master } from './slideLayouts';
export default class PptxGenJS {
    /**
     * Presentation layout name
     * Available Layouts:
     * 'LAYOUT_4x3'   (10" x 7.5")
     * 'LAYOUT_16x9'  (10" x 5.625")
     * 'LAYOUT_16x10' (10" x 6.25")
     * 'LAYOUT_WIDE'  (13.33" x 7.5")
     * 'LAYOUT_USER'  (user specified, can be any size)
     * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
     */
    private _layout;
    layout: string;
    /**
     * Library Version
     */
    private _version;
    readonly version: string;
    private _author;
    author: string;
    private _company;
    company: string;
    private _theme;
    configureTheme(fontFamily: any, colorScheme: any): void;
    /**
     * Sets the Presentation's Revision
     * PowerPoint requires `revision` be a number only (without "." or ",") (otherwise, PPT will throw errors upon opening Presentation!)
     */
    private _revision;
    revision: string;
    private _subject;
    subject: string;
    private _title;
    title: string;
    /**
     * Whether Right-to-Left (RTL) mode is enabled
     */
    private _rtlMode;
    rtlMode: boolean;
    /**
     * `isBrowser` Presentation Option:
     * Target: Angular/React/Webpack, etc. This setting affects how files are saved: using `fs` for Node.js or browser libs
     */
    private _isBrowser;
    isBrowser: boolean;
    private _colorScheme;
    setColorScheme(colorScheme: any): void;
    /** master slide layout object */
    private masterSlide;
    /** this Presentation's Slide objects */
    private slides;
    /** slide layout definition objects, used for generating slide layout files */
    private slideLayouts;
    private LAYOUTS;
    private _charts;
    readonly charts: typeof CHART_TYPES;
    private _colors;
    readonly colors: typeof SCHEME_COLOR_NAMES;
    private _shapes;
    readonly shapes: typeof PowerPointShapes;
    private _presLayout;
    readonly presLayout: ILayout;
    constructor();
    /**
     * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
     * @param {string} masterName - slide master name
     * @return {Slide} new Slide
     */
    addNewSlide: (masterName: string) => Slide;
    /**
     * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
     * @since 3.0.0
     * @param {number} slideNum - slide number
     * @return {Slide} Slide
     */
    getSlide: (slideNum: number) => Slide;
    /**
     * Create all chart and media rels for this Presenation
     * @param {Slide | Master} slide - slide with rels
     * @param {JSZIP} zip - JSZip instance
     * @param {Promise<any>[]} chartPromises - promise array
     */
    createChartMediaRels: (slide: Slide | Master, zip: JSZip, chartPromises: Promise<any>[]) => void;
    /**
     * Create and export the .pptx file
     * @param {string} exportName - output file type
     * @param {Blob} blobContent - Blob content
     * @return {Promise<string>} Promise with file name
     */
    writeFileToBrowser: (exportName: string, blobContent: Blob) => Promise<string>;
    /**
     * Create and export the .pptx file
     * @param {WRITE_OUTPUT_TYPE} outputType - output file type
     * @return {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} Promise with data or stream (node) or filename (browser)
     */
    exportPresentation: (outputType?: WRITE_OUTPUT_TYPE) => Promise<string | Blob | ArrayBuffer | Buffer | Uint8Array>;
    /**
     * Export the current Presenation to stream
     * @since 3.0.0
     * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file stream
     */
    stream(): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>;
    /**
     * Export the current Presenation as JSZip content with the selected type
     * @since 3.0.0
     * @param {JSZIP_OUTPUT_TYPE} outputType - 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
     * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file content in selected type
     */
    write(outputType: JSZIP_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>;
    /**
     * Export the current Presenation. Writes file to local file system if `fs` exists, otherwise, initiates download in browsers
     * @since 3.0.0
     * @param {string} exportName - file name
     * @returns {Promise<string>} the presentation name
     */
    writeFile(exportName?: string): Promise<string>;
    /**
     * Add a Slide to Presenation
     * @param {string} masterSlideName - Master Slide name
     * @returns {Slide} the new Slide
     */
    addSlide(masterSlideName?: string): Slide;
    /**
     * Adds a new slide master [layout] to the Presentation
     * @param {ISlideMasterOptions} slideMasterOpts - layout definition
     */
    defineSlideMaster(slideMasterOpts: ISlideMasterOptions): Master;
    /**
     * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
     * @note `verbose` option is undocumented; used for verbose output of layout process
     * @param {string} tabEleId - HTMLElementID of the table
     * @param {ITableToSlidesOpts} inOpts - array of options (e.g.: tabsize)
     */
    tableToSlides(tableElementId: string, opts?: ITableToSlidesOpts): void;
}
