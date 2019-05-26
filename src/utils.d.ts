/**
 * PptxGenJS Utils
 */
export declare function getUuid(uuidFormat: string): string;
/**
 * shallow mix, returns new object
 */
export declare function getMix(o1: any, o2: any, etc?: any): {};
/**
 * DESC: Replace special XML characters with HTML-encoded strings
 */
export declare function encodeXmlEntities(inStr: string): string;
export declare function inch2Emu(inches: number): number;
export declare function getSmartParseNumber(inVal: number | string, inDir: 'X' | 'Y'): number;
/**
 * DESC: Convert degrees (0..360) to Powerpoint rot value
 */
export declare function convertRotationDegrees(d: number): number;
/**
 * DESC: Convert component value to hex value
 */
export declare function componentToHex(c: number): string;
/**
 * DESC: Used by `addSlidesForTable()` to convert RGB colors from jQuery selectors to Hex for Presentation colors
 */
export declare function rgbToHex(r: number, g: number, b: number): string;
