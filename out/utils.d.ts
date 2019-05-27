/**
 * PptxGenJS Utils
 */
import { ISlideLayout } from './interfaces';
export declare function getUuid(uuidFormat: string): string;
/**
 * shallow mix, returns new object
 */
export declare function getMix(o1: any, o2: any, etc?: any): {};
/**
 * DESC: Replace special XML characters with HTML-encoded strings
 */
export declare function encodeXmlEntities(inStr: string): string;
/**
 * Convert inches into EMU
 *
 * @param {number|string} `inches`
 * @returns {number} EMU value
 */
export declare function inch2Emu(inches: number | string): number;
export declare function getSmartParseNumber(inVal: number | string, inDir: 'X' | 'Y', pptLayout?: ISlideLayout): number;
/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 *
 * @param {number} `d` - degrees
 * @returns {number} PPT `rot` value
 */
export declare function convertRotationDegrees(d: number): number;
/**
 * Converts component value to hex value
 *
 * @param {number} `c` - component color
 * @returns {string} hex string
 */
export declare function componentToHex(c: number): string;
/**
 * Converts RGB colors from jQuery selectors to Hex for Presentation colors for the `addSlidesForTable()` method
 *
 * @param {number} `r` - red value
 * @param {number} `g` - green value
 * @param {number} `b` - blue value
 */
export declare function rgbToHex(r: number, g: number, b: number): string;
