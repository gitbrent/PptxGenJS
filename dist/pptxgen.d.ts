// Type definitions for pptxgenjs 2.3.0
// Project: https://gitbrent.github.io/PptxGenJS/
// Definitions by: Brent Ely <https://github.com/gitbrent/>
//                 Michael Beaumont <https://github.com/michaelbeaumont>
//                 Nicholas Tietz-Sokolsky <https://github.com/ntietz>
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped
// TypeScript Version: 2.3

export as namespace PptxGenJS;

export = PptxGenJS;

declare class PptxGenJS {
  // Presentation Props
  getLayout(): string;
  setBrowser(isBrowser: boolean): void;
  setLayout(layout: PptxGenJS.LayoutName | PptxGenJS.Layout): void;
  setRTL(isRTL: boolean): void;

  // Presentation Metadata
  setAuthor(author: string): void;
  setCompany(company: string): void;
  setRevision(revision: string): void;
  setSubject(subject: string): void;
  setTitle(title: string): void;

  // Add a new Slide
  addNewSlide(masterLayoutName?: string): PptxGenJS.Slide;
  defineSlideMaster(opts: PptxGenJS.MasterSlideOptions): void;

  // Export
  save(exportFileName: string, callbackFunction?: Function, zipOutputType?: PptxGenJS.JsZipOutputType): void;
}

declare namespace PptxGenJS {
  const version: string;
  export type ChartName = 'AREA' | 'BAR' | 'BUBBLE' | 'DOUGHNUT' | 'LINE' | 'PIE' | 'RADAR' | 'SCATTER'
  export type ChartType = ChartName | { displayName: string; name: string };
  export type JsZipOutputType = 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array';
  export type LayoutName = 'LAYOUT_4x3' | 'LAYOUT_16x9' | 'LAYOUT_16x10' | 'LAYOUT_WIDE';
  export interface Layout {
    name?: string;
    width: number;
    height: number;
  }
  export type HexColor = string; // should match /^[0-9a-fA-F]{6}$/
  export type ThemeColor = 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6';
  export type Color = HexColor | ThemeColor;
  export type Coord = number | string; // string is in form 'n%'

  export type HAlign = 'left' | 'center' | 'right';
  export type VAlign = 'top' | 'middle' | 'bottom';

  export type ChartTimeUnit = 'days' | 'months' | 'years';
  export type ChartAxisOrientation = 'minMax' | 'maxMin';

  export type ChartLabelPos = 'low' | 'high' | 'nextTo';

  export interface PositionOptions {
    x?: Coord;
    y?: Coord;
    w?: Coord;
    h?: Coord;
  }

  export type CommonOptions = PositionOptions; // for backwards compatability

  export type Hyperlink = ({ url: string; slide?: undefined } | { slide: number; url?: undefined }) & { tooltip?: string };

  export type DataOrPath = { data: string; path?: undefined } | { data?: undefined; path: string };

  export interface ImageSizingOptions extends PositionOptions {
    type: 'cover' | 'contain' | 'crop';
  }

  export interface ImageBaseOptions extends PositionOptions {
    hyperlink?: Hyperlink;
    rounding?: boolean;
    sizing?: ImageSizingOptions;
  }

  export type ImageOptions = ImageBaseOptions | DataOrPath;

  export interface MediaBaseOptions extends PositionOptions {
    onlineVideoLink?: string;
    type?: 'audio' | 'online' | 'video';
  }
  export type MediaOptions = MediaBaseOptions | DataOrPath;

  export interface BorderOptions {
    pt: number | string;
    color: HexColor;
  }

  export interface ShadowOptions {
    type?: 'outer' | 'inner';
    angle?: number;
    blur?: number;
    color?: Color;
    offset?: number;
    opacity?: number;
  }

  export interface ChartData {
    name: string;
    labels: string[];
    values: number[];
  }

  export interface GridLineOptions {
    size?: number;
    color?: Color;
    style?: 'solid' | 'dash' | 'dot';
  }

  export interface ChartBaseOptions extends PositionOptions {
    border?: BorderOptions;
    chartColors?: Color[];
    chartColorsOpacity?: number;
    fill?: Color;
    holeSize?: number;
    invertedColors?: Color[];
    legendColor?: Color;
    legendFontFace?: string;
    legendFontSize?: number;
    legendPos?: string;
    layout?: { x: number; y: number; w: number; h: number };  // not `Coord`, must be `number` in 0..1
    radarStyle?: 'standard' | 'marker' | 'filled';
    showDataTable?: true | false;
    showDataTableKeys?: boolean;
    showDataTableHorzBorder?: boolean;
    showDataTableVertBorder?: boolean;
    showDataTableOutline?: boolean;
    showLabel?: boolean;
    showLegend?: boolean;
    showPercent?: boolean;
    showTitle?: boolean;
    showValue?: boolean;
    title?: string;
    titleAlign?: HAlign;
    titleColor?: Color;
    titleFontFace?: string;
    titleFontSize?: number;
    titlePos?: { x: number; y: number };
    titleRotate?: number;
  }

  export interface ChartAxesOptions {
    axisLineColor?: Color;
    catAxisBaseTimeUnit?: ChartTimeUnit;
    catAxisHidden?: boolean;
    catAxisLabelColor?: Color;
    catAxisLabelFontBold?: boolean;
    catAxisLabelFontFace?: string;
    catAxisLabelFontSize?: number;
    catAxisLabelFrequency?: number;
    catAxisLabelPos?: ChartLabelPos;
    catAxisLabelRotate?: number;
    catAxisLineShow?: boolean;
    catAxisMajorTimeUnit?: ChartTimeUnit;
    catAxisMaxVal?: number;
    catAxisMinVal?: number;
    catAxisMinorTimeUnit?: ChartTimeUnit;
    catAxisMajorUnit?: number;
    catAxisMinorUnit?: number;
    catAxisOrientation?: ChartAxisOrientation;
    catAxisTitle?: string;
    catAxisTitleColor?: Color;
    catAxisTitleFontFace?: string;
    catAxisTitleFontSize?: number;
    catAxisTitleRotate?: number;
    catGridLine?: GridLineOptions | 'none';
    showCatAxisTitle?: boolean;
    showValAxisTitle?: boolean;
    valAxisHidden?: boolean;
    valAxisLabelColor?: Color;
    valAxisLabelFontBold?: boolean;
    valAxisLabelFontFace?: string;
    valAxisLabelFontSize?: number;
    valAxisLabelFormatCode?: string;
    valAxisLineShow?: boolean;
    valAxisMajorUnit?: number;
    valAxisMaxVal?: number;
    valAxisMinVal?: number;
    valAxisOrientation?: ChartAxisOrientation;
    valAxisTitle?: string;
    valAxisTitleColor?: Color;
    valAxisTitleFontFace?: string;
    valAxisTitleFontSize?: number;
    valAxisTitleRotate?: number;
    valGridLine?: GridLineOptions | 'none';
  }

  export interface ChartBarDataLineOptions {
    barDir?: 'bar' | 'col';
    barGapWidthPct?: number;
    barGrouping?: 'clustered' | 'stacked' | 'percentStacked';
    catLabelFormatCode?: string;
    dataBorder?: BorderOptions;
    dataLabelColor?: Color;
    dataLabelFormatCode?: string;
    dataLabelFormatScatter?: 'custom' | 'customXY' | 'XY';
    dataLabelFontBold?: boolean;
    dataLabelFontFace?: string;
    dataLabelFontSize?: number;
    dataLabelPosition?: 'bestFit' | 'b' | 'ctr' | 'inBase' | 'inEnd' | 'l' | 'outEnd' | 'r' | 't';
    dataNoEffects?: boolean;
    displayBlanksAs?: 'span' | 'gap';
    gridLineColor?: Color;
    lineDataSymbol?: 'circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'square' | 'triangle';
    lineDataSymbolSize?: number;
    lineDataSymbolLineSize?: number;
    lineDataSymbolLineColor?: Color;
    lineSize?: number;
    lineSmooth?: boolean;
    shadow?: ShadowOptions | 'none';
    valueBarColors?: boolean;
  }

  export interface Chart3DBarOptions {
    bar3DShape?: 'box' | 'cylinder' | 'coneToMax' | 'pyramid' | 'pyramidToMax';
    barGapDepthPct?: number;
    dataLabelBkgrdColors?: boolean;
    serAxisBaseTimeUnit?: ChartTimeUnit;
    serAxisHidden?: boolean;
    serAxisOrientation?: ChartAxisOrientation;
    serAxisLabelColor?: Color;
    serAxisLabelFontBold?: boolean;
    serAxisLabelFontFace?: string;
    serAxisLabelFontSize?: number;
    serAxisLabelFrequency?: number;
    serAxisLabelPos?: ChartLabelPos;
    serAxisLineShow?: boolean;
    serAxisMajorTimeUnit?: ChartTimeUnit;
    serAxisMajorUnit?: number;
    serAxisMinorTimeUnit?: ChartTimeUnit;
    serAxisMinorUnit?: number;
    serAxisTitle?: string;
    serAxisTitleColor?: Color;
    serAxisTitleFontFace?: string;
    serAxisTitleFontSize?: number;
    serAxisTitleRotate?: number;
    serGridLine?: GridLineOptions | 'none';
    v3DRAngAx?: boolean;
    v3DPerspective?: number;
    v3DRotX?: number;
    v3DRotY?: number;
  }

  export type ChartOptions = ChartBaseOptions | ChartAxesOptions | ChartBarDataLineOptions | Chart3DBarOptions;

  export interface ChartMultiTypeOptions {
    catAxes?: ChartAxesOptions[];
    secondaryCatAxis?: boolean;
    secondaryValAxis?: boolean;
    valAxes?: ChartAxesOptions[];
  }

  export interface ChartTypeAndData {
    type: ChartType;
    data: ChartData[];
    options?: ChartOptions;
  }

  export interface Shape {
    displayName: string;
    name: string;
    avLst: { [key: string]: number };
  }

  export interface ShapeOptions extends PositionOptions {
    align?: HAlign;
    fill?: Color | { type: string; color: Color; alpha?: number };
    flipH?: boolean;
    flipV?: boolean;
    lineSize?: number;
    lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot';
    lineHead?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle';
    lineTail?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle';
    line?: Color;
    rectRadius?: number;
    rotate?: number;
  }

  export interface BasicTextFormatting {
    bold?: boolean;
    italic?: boolean;
    strike?: boolean;
    subscript?: boolean;
    superscript?: boolean;
    underline?: boolean;
    color?: Color;
    fontFace?: string;
    fontSize?: number;
  }

  export interface TextOptions extends ShapeOptions, BasicTextFormatting {
    align?: HAlign;
    autoFit?: boolean;
    breakLine?: boolean;
    bullet?: boolean | { type: 'number'; code?: undefined } | { code: string; type?: undefined };
    charSpacing?: number;
    hyperlink?: Hyperlink;
    indentLevel?: number;
    inset?: number;
    isTextBox?: boolean;
    lang?: string;
    lineSpacing?: number;
    margin?: number | [number, number, number, number];
    outline?: { size: number; color: Color };
    paraSpaceAfter?: number;
    paraSpaceBefore?: number;
    rectRadius?: number;
    rtlMode?: boolean;
    shadow?: ShadowOptions;
    shape?: Shape | string;
    shrinkText?: boolean;
    valign?: VAlign;
    vert?: 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl';
  }

  export interface TableFormatting extends TextOptions {
    align?: HAlign;
    border?: 'none' | string | BorderOptions | [number, number, number, number];
    color?: Color;
    colspan?: number;
    fill?: Color;
    margin?: number | [number, number, number, number];
    rowspan?: number;
    valign?: VAlign;
  }

  export type TextDeclaration<T extends TextOptions> = string | number | { text: string | number; options?: T };

  export type TableRow = TextDeclaration<TableFormatting>[];

  export interface SlideNumberOptions extends PositionOptions {
    color?: Color;
    fontFace?: string;
    fontSize?: number;
  }

  export interface MasterSlideOptions {
    title: string;
    bkgd?: string | DataOrPath;
    objects?: object[];
    slideNumber?: SlideNumberOptions;
    margin?: number | number[];
  }

  export class Slide {
    // Slide Number methods
    getPageNumber(): string;
    slideNumber(): SlideNumberOptions;
    slideNumber(options: SlideNumberOptions): void;

    // Core object API Methods
    addChart(type: ChartType, data: ChartData[], options?: ChartOptions): Slide;
    addChart(type: ChartTypeAndData[], options?: ChartOptions | ChartMultiTypeOptions): Slide;
    addImage(options: ImageOptions): Slide;
    addMedia(options: MediaOptions): Slide;
    addNotes(noteText: string): Slide;
    addShape(shape: Shape | string, options: ShapeOptions): Slide;
    addTable(tableData: TableRow[], options: TableFormatting): Slide;
    addText(text: string | number | TextDeclaration<TextOptions>[], options?: TextOptions): Slide;
  }
}
