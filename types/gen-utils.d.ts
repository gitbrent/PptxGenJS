import { IChartOpts, ILayout, ShapeFill } from './core-interfaces';
/**
 * Convert string percentages to number relative to slide size
 * @param {number|string} size - numeric ("5.5") or percentage ("90%")
 * @param {'X' | 'Y'} xyDir - direction
 * @param {ILayout} layout - presentation layout
 * @returns {number} calculated size
 *
 */
export declare function getSmartParseNumber(size: number | string, xyDir: 'X' | 'Y', layout: ILayout): number;
/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
export declare function getUuid(uuidFormat: string): string;
/**
 * TODO: What does this mehtod do again??
 * shallow mix, returns new object
 */
export declare function getMix(o1: any | IChartOpts, o2: any | IChartOpts, etc?: any): {};
/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string} xml - XML string to encode
 * @returns {string} escaped XML
 */
export declare function encodeXmlEntities(xml: string): string;
/**
 * Convert inches into EMU
 * @param {number|string} inches - as string or number
 * @returns {number} EMU value
 */
export declare function inch2Emu(inches: number | string): number;
/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 *
 * @param {number} d - degrees
 * @returns {number} rot - value
 */
export declare function convertRotationDegrees(d: number): number;
/**
 * Converts component value to hex value
 * @param {number} c - component color
 * @returns {string} hex string
 */
export declare function componentToHex(c: number): string;
/**
 * Converts RGB colors from css selectors to Hex for Presentation colors
 * @param {number} r - red value
 * @param {number} g - green value
 * @param {number} b - blue value
 * @returns {string} XML string
 */
export declare function rgbToHex(r: number, g: number, b: number): string;
/**
 * Create either a `a:schemeClr` - (scheme color) or `a:srgbClr` (hexa representation).
 * @param {string} colorStr - hexa representation (eg. "FFFF00") or a scheme color constant (eg. pptx.colors.ACCENT1)
 * @param {string} innerElements - additional elements that adjust the color and are enclosed by the color element
 * @returns {string} XML string
 */
export declare function createColorElement(color: string, innerElements?: string): string;
/**
 * Create color selection
 * @param {ShapeFill} shapeFill - options
 * @param {string} backColor - color string
 * @returns {string} XML string
 */
export declare function genXmlColorSelection(shapeFill: ShapeFill, backColor?: string): string;
export declare const createImageConfig: ({ relId, data, path, extension, Target, fromSvgSize }: {
    relId: any;
    data?: string;
    path?: string;
    extension: any;
    Target: any;
    fromSvgSize: any;
}) => {
    rId: any;
    type: string;
    path: string;
    data: string;
    extn: any;
    Target: any;
    isSvgPng: boolean;
    svgSize: any;
};
export declare const genericParseFloat: (n: string | number) => number;
export declare const translateColor: (color: any) => any;
