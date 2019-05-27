import { CHART_TYPES, JSZIP_OUTPUT_TYPE, SLIDE_OBJECT_TYPES } from './enums';
import { inch2Emu, rgbToHex } from './utils';
export declare var jQuery: any;
export declare var fs: any;
export declare var https: any;
export declare var JSZip: any;
export declare var sizeOf: any;
declare type Coord = number | string;
export interface OptsCoords {
    x?: Coord;
    y?: Coord;
    w?: Coord;
    h?: Coord;
}
export interface OptsDataOrPath {
    data?: string;
    path?: string;
}
export interface OptsChartData {
    index?: number;
    name?: string;
    labels?: Array<string>;
    values?: Array<number>;
    sizes?: Array<number>;
}
export interface OptsChartGridLine {
    size?: number;
    color?: string;
    style?: 'solid' | 'dash' | 'dot' | 'none';
}
export interface IBorder {
    color?: string;
    pt?: number;
}
export interface IShadowOpts {
    type: string;
    angle: number;
    opacity: number;
    blur?: number;
    offset?: number;
    color?: string;
}
export interface IChartOpts extends OptsCoords, OptsChartGridLine {
    type: CHART_TYPES;
    layout?: OptsCoords;
    barDir?: string;
    barGrouping?: string;
    barGapWidthPct?: number;
    barGapDepthPct?: number;
    bar3DShape?: string;
    catAxisOrientation?: 'minMax' | 'minMax';
    catGridLine?: OptsChartGridLine;
    valGridLine?: OptsChartGridLine;
    chartColors?: Array<string>;
    chartColorsOpacity?: number;
    showLabel?: boolean;
    lang?: string;
    dataNoEffects?: string;
    dataLabelFormatScatter?: string;
    dataLabelFormatCode?: string;
    dataLabelBkgrdColors?: string;
    dataLabelFontSize?: number;
    dataLabelColor?: string;
    dataLabelFontFace?: string;
    dataLabelPosition?: string;
    displayBlanksAs?: string;
    fill?: string;
    border?: IBorder;
    hasArea?: boolean;
    catAxes?: Array<number>;
    valAxes?: Array<number>;
    lineDataSymbol?: string;
    lineDataSymbolSize?: number;
    lineDataSymbolLineColor?: string;
    lineDataSymbolLineSize?: number;
    showLegend?: boolean;
    legendPos?: string;
    legendFontFace?: string;
    legendFontSize?: number;
    legendColor?: string;
    lineSmooth?: boolean;
    invertedColors?: string;
    serAxisOrientation?: string;
    serAxisHidden?: boolean;
    serGridLine?: OptsChartGridLine;
    showSerAxisTitle?: boolean;
    serLabelFormatCode?: string;
    serAxisLabelPos?: string;
    serAxisLineShow?: boolean;
    serAxisLabelFontSize?: string;
    serAxisLabelColor?: string;
    serAxisLabelFontFace?: string;
    serAxisLabelFrequency?: string;
    serAxisBaseTimeUnit?: string;
    serAxisMajorTimeUnit?: string;
    serAxisMinorTimeUnit?: string;
    serAxisMajorUnit?: number;
    serAxisMinorUnit?: number;
    serAxisTitleColor?: string;
    serAxisTitleFontFace?: string;
    serAxisTitleFontSize?: number;
    serAxisTitleRotate?: number;
    serAxisTitle?: string;
    showDataTable?: boolean;
    showDataTableHorzBorder?: boolean;
    showDataTableVertBorder?: boolean;
    showDataTableOutline?: boolean;
    showDataTableKeys?: boolean;
    title?: string;
    titleFontSize?: number;
    titleColor?: string;
    titleFontFace?: string;
    titleRotate?: number;
    titleAlign?: string;
    titlePos?: string;
    dataLabelFontBold?: boolean;
    valueBarColors?: Array<string>;
    holeSize?: number;
    showTitle?: boolean;
    showValue?: boolean;
    showPercent?: boolean;
    catLabelFormatCode?: string;
    dataBorder?: IBorder;
    lineSize?: number;
    lineDash?: string;
    radarStyle?: string;
    shadow?: IShadowOpts;
    catAxisLabelPos?: string;
    valAxisOrientation?: 'minMax' | 'minMax';
    valAxisMaxVal?: number;
    valAxisMinVal?: number;
    valAxisHidden?: boolean;
    valAxisTitleColor?: string;
    valAxisTitleFontFace?: string;
    valAxisTitleFontSize?: number;
    valAxisTitleRotate?: number;
    valAxisTitle?: string;
    valAxisLabelFormatCode?: string;
    valAxisLineShow?: boolean;
    valAxisLabelRotate?: number;
    valAxisLabelFontSize?: number;
    valAxisLabelFontBold?: boolean;
    valAxisLabelColor?: string;
    valAxisLabelFontFace?: string;
    valAxisMajorUnit?: number;
    showValAxisTitle?: boolean;
    axisPos?: string;
    v3DRotX?: number;
    v3DRotY?: number;
    v3DRAngAx?: boolean;
    v3DPerspective?: string;
}
export interface IMediaOpts extends OptsCoords, OptsDataOrPath {
    onlineVideoLink?: string;
    type?: 'audio' | 'online' | 'video';
}
export interface ITextOpts extends OptsCoords, OptsDataOrPath {
    align?: string;
    autoFit?: boolean;
    color?: string;
    fontSize?: number;
    inset?: number;
    lineSpacing?: number;
    line?: string;
    lineSize?: number;
    placeholder?: object;
    rotate?: number;
    shadow?: IShadowOpts;
    shape?: {
        name: string;
    };
    vert?: 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl';
    valign?: string;
}
export interface ISlideRel {
    Target: string;
    type: SLIDE_OBJECT_TYPES;
    data: string;
    path?: string;
    extn?: string;
    rId: number;
}
export interface ISlideRelChart extends OptsChartData {
    type: CHART_TYPES;
    opts: IChartOpts;
    data: Array<OptsChartData>;
    rId: number;
    Target: string;
    globalId: number;
    fileName: string;
}
export interface ISlideRelMedia {
    type: string;
    opts: IMediaOpts;
    path?: string;
    extn?: string;
    data?: string | ArrayBuffer;
    isSvgPng?: boolean;
    rId: number;
    Target: string;
}
export interface ILayout {
    name: string;
    rels?: object;
    relsChart?: ISlideRelChart;
    relsMedia?: ISlideRelMedia;
    data: Array<object>;
    options: {
        placeholderName: string;
    };
    width: number;
    height: number;
}
export interface ISlideNumber extends OptsCoords {
    fontFace: string;
    fontSize: number;
    color: string;
}
export interface ISlideDataObject {
    type: SLIDE_OBJECT_TYPES;
    text?: string;
    arrTabRows?: Array<Array<{
        text: string | object;
        options?: {
            colspan?: number;
        };
    }>>;
    chartRid?: number;
    image?: string;
    imageRid?: number;
    hyperlink?: {
        rId: number;
        slide?: number;
        tooltip?: string;
        url?: string;
    };
    media?: string;
    mtype?: 'online' | 'other';
    mediaRid?: number;
    options?: {
        x?: number;
        y?: number;
        cx?: number;
        cy?: number;
        w?: number;
        h?: number;
        placeholder?: string;
        shape?: object;
        bodyProp?: {
            lIns?: number;
            rIns?: number;
            bIns?: number;
            tIns?: number;
        };
        isTextBox?: boolean;
        line?: string;
        margin?: number;
        rectRadius?: number;
        fill?: string;
        shadow?: IShadowOpts;
        colW?: number;
        rowH?: number;
        flipH?: boolean;
        flipV?: boolean;
        rotate?: number;
        lineDash?: string;
        lineSize?: number;
        lineHead?: string;
        lineTail?: string;
        sizing?: {
            type?: string;
            x?: number;
            y?: number;
            w?: number;
            h?: number;
        };
        rounding?: string;
    };
}
export interface ISlideLayout {
    name: string;
    slide: ISlide;
    data: Array<object>;
    rels: Array<any>;
    margin: Array<number>;
    slideNumberObj?: ISlideNumber;
    width?: number;
    height?: number;
}
export interface ISlideLayoutChart extends ISlideLayout {
    rels: Array<ISlideRelChart>;
}
export interface ISlideLayoutMedia extends ISlideLayout {
    rels: Array<ISlideRelMedia>;
}
export interface ISlide {
    slide?: {
        back: string;
        bkgdImgRid?: number;
        color: string;
        hidden?: boolean;
    };
    numb?: number;
    name?: string;
    rels?: Array<ISlideRel>;
    relsChart?: Array<ISlideRelChart>;
    relsMedia?: Array<ISlideRelMedia>;
    data?: Array<ISlideDataObject>;
    layoutName?: string;
    layoutObj?: ISlideLayout;
    margin?: object;
    slideNumberObj?: ISlideNumber;
}
export interface IPresentation {
    author: string;
    company: string;
    revision: string;
    subject: string;
    title: string;
    isBrowser: boolean;
    fileName: string;
    fileExtn: string;
    pptLayout: ILayout;
    rtlMode: boolean;
    saveCallback?: null;
    masterSlide?: ISlide;
    chartCounter: number;
    imageCounter: number;
    slides?: ISlide[];
    slideLayouts?: ISlideLayout[];
}
export interface IAddNewSlide {
    getPageNumber: Function;
    slideNumber: Function;
    addChart: Function;
    addImage: Function;
    addMedia: Function;
    addNotes: Function;
    addShape: Function;
    addTable: Function;
    addText: Function;
}
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
export {};
