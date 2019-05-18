/**
 * PptxGenJS Enums
 * NOTE: `enum` wont work for objects, so use `Object.freeze`
 */

// CONST
export const EMU: number = 914400;  // One (1) inch (OfficeXML measures in EMU (English Metric Units))
export const ONEPT: number = 12700; // One (1) point (pt)
export const CRLF: string = '\r\n'; // AKA: Chr(13) & Chr(10)
export const LAYOUT_IDX_SERIES_BASE: number = 2147483649;
export const REGEX_HEX_COLOR: RegExp = /^[0-9a-fA-F]{6}$/;

export const DEF_FONT_TITLE_SIZE: number = 18;
export const DEF_SLIDE_MARGIN_IN: Array<number> = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
export const DEF_FONT_COLOR: string = '000000';
export const DEF_FONT_SIZE: number = 12;
export const DEF_CHART_GRIDLINE = { color: "888888", style: "solid", size: 1 };
export const DEF_SHAPE_SHADOW = { type: 'outer', blur: 3, offset: (23000 / 12700), angle: 90, color: '000000', opacity: 0.35, rotateWithShape: true };
export const DEF_TEXT_SHADOW = { type: 'outer', blur: 8, offset: 4, angle: 270, color: '000000', opacity: 0.75 };

export const AXIS_ID_VALUE_PRIMARY: string = '2094734552';
export const AXIS_ID_VALUE_SECONDARY: string = '2094734553';
export const AXIS_ID_CATEGORY_PRIMARY: string = '2094734554';
export const AXIS_ID_CATEGORY_SECONDARY: string = '2094734555';
export const AXIS_ID_SERIES_PRIMARY: string = '2094734556';

export const LETTERS: Array<string> = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
export const BARCHART_COLORS: Array<string> = [
	'C0504D', '4F81BD', '9BBB59', '8064A2', '4BACC6', 'F79646', '628FC6', 'C86360', 'C0504D', '4F81BD', '9BBB59', '8064A2', '4BACC6', 'F79646', '628FC6', 'C86360'
];
export const PIECHART_COLORS: Array<string> = [
	'5DA5DA', 'FAA43A', '60BD68', 'F17CB0', 'B2912F', 'B276B2', 'DECF3F', 'F15854', 'A7A7A7', '5DA5DA', 'FAA43A', '60BD68', 'F17CB0', 'B2912F', 'B276B2', 'DECF3F', 'F15854', 'A7A7A7'
];

// ENUM
export enum MASTER_OBJECTS {
	"chart" = "chart",
	"image" = "image",
	"line" = "line",
	"rect" = "rect",
	"text" = "text",
	"placeholder" = "placeholder"
}

export enum PLACEHOLDER_TYPES {
	"title" = "title",
	"body" = "body",
	"image" = "pic",
	"chart" = "chart",
	"table" = "tbl",
	"media" = "media"
}

export enum CHART_TYPES {
	'AREA' = 'area',
	'BAR' = 'bar',
	'BAR3D' = 'bar3D',
	'BUBBLE' = 'bubble',
	'DOUGHNUT' = 'doughnut',
	'LINE' = 'line',
	'PIE' = 'pie',
	'RADAR' = 'radar',
	'SCATTER' = 'scatter'
}

/**
* NOTE: 20170304: BULLET_TYPES: Only default is used so far. I'd like to combine the two pieces of code that use these before implementing these as options
* Since we close <p> within the text object bullets, its slightly more difficult than combining into a func and calling to get the paraProp
* and i'm not sure if anyone will even use these... so, skipping for now.
*/
export enum BULLET_TYPES {
	'DEFAULT' = "&#x2022;",
	'CHECK' = "&#x2713;",
	'STAR' = "&#x2605;",
	'TRIANGLE' = "&#x25B6;"
}

export const BASE_SHAPES = Object.freeze({
	'RECTANGLE': { 'displayName': 'Rectangle', 'name': 'rect', 'avLst': {} },
	'LINE': { 'displayName': 'Line', 'name': 'line', 'avLst': {} }
})
