import { JSZIP_OUTPUT_TYPE } from './enums';
import { ISlide, ILayout, IAddNewSlide, ISlideRelMedia } from './interfaces';
import { inch2Emu, rgbToHex } from './utils';
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
     * Set Right-to-Left (RTL) mode for users whose language requires this setting
     */
    private _rtlMode;
    rtlMode: boolean;
    /**
     * Sets the Presentation's Slide Layout {object}: [screen4x3, screen16x9, widescreen]
     * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
     * @param {string} inLayout - a const name from LAYOUTS variable
     * @param {object} inLayout - an object with user-defined `w` and `h`
     */
    private _pptLayout;
    pptLayout: ILayout;
    /**
     * Sets the Presentation Option: `isBrowser`
     * Target: Angular/React/Webpack, etc.
     * This setting affects how files are saved: using `fs` for Node.js or browser libs
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
    private fs;
    private https;
    private JSZip;
    private sizeOf;
    constructor();
    /**
     * DESC: Export the .pptx file
     */
    doExportPresentation: (outputType?: JSZIP_OUTPUT_TYPE) => void;
    writeFileToBrowser: (strExportName: any, content: any) => void;
    createMediaFiles: (layout: ISlide, zip: any, chartPromises: Promise<any>[]) => void;
    addPlaceholdersToSlides: (slide: any) => void;
    getSizeFromImage: (inImgUrl: any) => {
        width: any;
        height: any;
    };
    encodeSlideMediaRels: (layout: any, arrRelsDone: any) => number;
    convertImgToDataURL: (slideRel: ISlideRelMedia) => void;
    convertRemoteMediaToDataURL: (slideRel: ISlideRelMedia) => void;
    convertSvgToPngViaCanvas: (slideRel: any) => void;
    callbackImgToDataURLDone: (base64Data: string | ArrayBuffer, slideRel: ISlideRelMedia) => void;
    /**
     * Magic happens here
     */
    parseTextToLines: (cell: any, inWidth: any) => string[];
    /**
     * Magic happens here
     */
    getSlidesForTableRows: (inArrRows: any, opts: any) => any[];
    /**
     * Expose a couple private helper functions from above
     */
    inch2Emu: () => typeof inch2Emu;
    rgbToHex: () => typeof rgbToHex;
    /**
     * Save (export) the Presentation .pptx file
     * @param {string} `inStrExportName` - Filename to use for the export
     * @param {Function} `funcCallback` - Callback function to be called when export is complete
     * @param {JSZIP_OUTPUT_TYPE} `outputType` - JSZip output type
     */
    save(inStrExportName: string, funcCallback?: Function, outputType?: JSZIP_OUTPUT_TYPE): void;
    /**
     * Add a new Slide to the Presentation
     * @param {string} inMasterName - Name of Master Slide
     * @returns {ISlide[]} slideObj - The new Slide object
     */
    addNewSlide(inMasterName?: string): IAddNewSlide;
    /**
     * Adds a new slide master [layout] to the presentation.
     * @param {ILayout} inObjMasterDef - layout definition
     * @return {ISlide} this slide
     */
    defineSlideMaster(inObjMasterDef: any): this;
    /**
     * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
     * "Auto-Paging is the future!" --Elon Musk
     *
     * @param {string} tabEleId - The HTML Element ID of the table
     * @param {array} inOpts - An array of options (e.g.: tabsize)
     */
    addSlidesForTable(tabEleId: any, inOpts: any): void;
}
