import * as JSZip from 'jszip';
import { JSZIP_OUTPUT_TYPE } from './enums';
import { ISlide, ILayout, IAddNewSlide, ISlideRelMedia, ITableCell, ISlideMasterDef } from './interfaces';
export default class PptxGenJS {
    private _version;
    readonly version: string;
    private _author;
    author: string;
    private _company;
    company: string;
    /**
     * DESC: Sets the Presentation's Revision
     * NOTE: PowerPoint requires `revision` be: number only (without "." or ",") otherwise, PPT will throw errors upon opening Presentation.
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
     * Presentation Layout: 'screen4x3', 'screen16x9', 'widescreen', etc.
     * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
     */
    private _layout;
    layout: ILayout;
    /**
     * `isBrowser` Presentation Option:
     * Target: Angular/React/Webpack, etc. This setting affects how files are saved: using `fs` for Node.js or browser libs
     */
    private _isBrowser;
    isBrowser: boolean;
    private fileName;
    private fileExtn;
    /** master slide layout object */
    private masterSlide;
    /** this Presentation's Slide objects */
    private slides;
    /** slide layout definition objects, used for generating slide layout files */
    private slideLayouts;
    private saveCallback;
    private NODEJS;
    private LAYOUTS;
    private Charts;
    private _imageCounter;
    private fs;
    private https;
    private sizeOf;
    constructor();
    /**
     * DESC: Export the .pptx file
     */
    doExportPresentation: (outputType?: JSZIP_OUTPUT_TYPE) => void;
    writeFileToBrowser: (strExportName: string, content: any) => void;
    createMediaFiles: (layout: ISlide, zip: JSZip, chartPromises: Promise<any>[]) => void;
    addPlaceholdersToSlides: (slide: ISlide) => void;
    getSizeFromImage: (inImgUrl: string) => {
        width: any;
        height: any;
    };
    encodeSlideMediaRels: (layout: any, arrRelsDone: any) => number;
    convertImgToDataURL: (slideRel: ISlideRelMedia) => void;
    /**
     * Node equivalent of `convertImgToDataURL()`: Use https to fetch, then use Buffer to encode to base64
     * @param {ISlideRelMedia} `slideRel` - slide rel
     */
    convertRemoteMediaToDataURL: (slideRel: ISlideRelMedia) => void;
    /**
     * (Browser Only): Convert SVG-base64 data to PNG-base64
     * @param {ISlideRelMedia} `slideRel` - slide rel
     */
    convertSvgToPngViaCanvas: (slideRel: ISlideRelMedia) => void;
    callbackImgToDataURLDone: (base64Data: string | ArrayBuffer, slideRel: ISlideRelMedia) => void;
    /**
     * Magic happens here
     */
    parseTextToLines: (cell: ITableCell, inWidth: number) => string[];
    /**
     * Magic happens here
     */
    getSlidesForTableRows: (inArrRows: any, opts: any) => any[];
    /**
     * Save (export) the Presentation .pptx file
     * @param {string} `inStrExportName` - Filename to use for the export
     * @param {Function} `funcCallback` - Callback function to be called when export is complete
     * @param {JSZIP_OUTPUT_TYPE} `outputType` - JSZip output type
     */
    save(inStrExportName: string, funcCallback?: Function, outputType?: JSZIP_OUTPUT_TYPE): void;
    /**
     * Add a new Slide to the Presentation
     * @param {string} inMasterName - name of Master Slide
     * @returns {IAddNewSlide} slideObj - new Slide object
     */
    addNewSlide(inMasterName?: string): IAddNewSlide;
    /**
     * Adds a new slide master [layout] to the presentation.
     * @param {ISlideMasterDef} inObjMasterDef - layout definition
     */
    defineSlideMaster(inObjMasterDef: ISlideMasterDef): void;
    /**
     * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
     * "Auto-Paging is the future!" --Elon Musk
     *
     * @param {string} `tabEleId` - HTMLElementID of the table
     * @param {object} `inOpts` - array of options (e.g.: tabsize)
     */
    addSlidesForTable(tabEleId: string, inOpts: any): void;
}
