/**
 * PptxGenJS: Utility Methods
 */
import { SCHEME_COLORS } from './core-enums';
import { IChartOpts, PresLayout, TextGlowProps, PresSlide, ShapeFillProps, Color, ShapeLineProps } from './core-interfaces';
/**
 * Convert string percentages to number relative to slide size
 * @param {number|string} size - numeric ("5.5") or percentage ("90%")
 * @param {'X' | 'Y'} xyDir - direction
 * @param {PresLayout} layout - presentation layout
 * @returns {number} calculated size
 */
export declare function getSmartParseNumber(size: number | string, xyDir: 'X' | 'Y', layout: PresLayout): number;
/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
export declare function getUuid(uuidFormat: string): string;
/**
 * TODO: What does this method do again??
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
 * Convert `pt` into points (using `ONEPT`)
 *
 * @param {number|string} pt
 * @returns {number} value in points (`ONEPT`)
 */
export declare function valToPts(pt: number | string): number;
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
 * @param {string|SCHEME_COLORS} colorStr - hexa representation (eg. "FFFF00") or a scheme color constant (eg. pptx.SchemeColor.ACCENT1)
 * @param {string} innerElements - additional elements that adjust the color and are enclosed by the color element
 * @returns {string} XML string
 */
export declare function createColorElement(colorStr: string | SCHEME_COLORS, innerElements?: string): string;
/**
 * Creates `a:glow` element
 * @param {TextGlowProps} options glow properties
 * @param {TextGlowProps} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 *	{ size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
export declare function createGlowElement(options: TextGlowProps, defaults: TextGlowProps): string;
/**
 * Create color selection
 * @param shapeFill - options
 * @param backColor - color string
 * @returns XML string
 */
export declare function genXmlColorSelection(shapeFill: Color | ShapeFillProps | ShapeLineProps, backColor?: string): string;
/**
 * Get a new rel ID (rId) for charts, media, etc.
 * @param {PresSlide} target - the slide to use
 * @returns {number} count of all current rels plus 1 for the caller to use as its "rId"
 */
export declare function getNewRelId(target: PresSlide): number;
