/**
 * PptxGenJS: Utility Methods
 */

import { EMU, REGEX_HEX_COLOR, DEF_FONT_COLOR, ONEPT, SchemeColor, SCHEME_COLORS } from './core-enums'
import { IChartOpts, PresLayout, TextGlowProps, PresSlide, ShapeFillProps, Color, ShapeLineProps, ModifiedThemeColor } from './core-interfaces'

/**
 * Convert string percentages to number relative to slide size
 * @param {number|string} size - numeric ("5.5") or percentage ("90%")
 * @param {'X' | 'Y'} xyDir - direction
 * @param {PresLayout} layout - presentation layout
 * @returns {number} calculated size
 */
export function getSmartParseNumber(size: number | string, xyDir: 'X' | 'Y', layout: PresLayout): number {
	// FIRST: Convert string numeric value if reqd
	if (typeof size === 'string' && !isNaN(Number(size))) size = Number(size)

	// CASE 1: Number in inches
	// Assume any number less than 100 is inches
	if (typeof size === 'number' && size < 100) return inch2Emu(size)

	// CASE 2: Number is already converted to something other than inches
	// Assume any number greater than 100 is not inches! Just return it (its EMU already i guess??)
	if (typeof size === 'number' && size >= 100) return size

	// CASE 3: Percentage (ex: '50%')
	if (typeof size === 'string' && size.indexOf('%') > -1) {
		if (xyDir && xyDir === 'X') return Math.round((parseFloat(size) / 100) * layout.width)
		if (xyDir && xyDir === 'Y') return Math.round((parseFloat(size) / 100) * layout.height)

		// Default: Assume width (x/cx)
		return Math.round((parseFloat(size) / 100) * layout.width)
	}

	// LAST: Default value
	return 0
}

/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
export function getUuid(uuidFormat: string): string {
	return uuidFormat.replace(/[xy]/g, function (c) {
		let r = (Math.random() * 16) | 0,
			v = c === 'x' ? r : (r & 0x3) | 0x8
		return v.toString(16)
	})
}

/**
 * TODO: What does this method do again??
 * shallow mix, returns new object
 */
export function getMix(o1: any | IChartOpts, o2: any | IChartOpts, etc?: any) {
	let objMix = {}
	for (let i = 0; i <= arguments.length; i++) {
		let oN = arguments[i]
		if (oN)
			Object.keys(oN).forEach(key => {
				objMix[key] = oN[key]
			})
	}
	return objMix
}

/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string} xml - XML string to encode
 * @returns {string} escaped XML
 */
export function encodeXmlEntities(xml: string): string {
	// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	if (typeof xml === 'undefined' || xml == null) return ''
	return xml.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

/**
 * Convert inches into EMU
 * @param {number|string} inches - as string or number
 * @returns {number} EMU value
 */
export function inch2Emu(inches: number | string): number {
	// FIRST: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
	// Any value over 100 damn sure isnt inches, must be EMU already, so just return it
	if (typeof inches === 'number' && inches > 100) return inches
	if (typeof inches === 'string') inches = Number(inches.replace(/in*/gi, ''))
	return Math.round(EMU * inches)
}

/**
 * Convert `pt` into points (using `ONEPT`)
 *
 * @param {number|string} pt
 * @returns {number} value in points (`ONEPT`)
 */
export function valToPts(pt: number | string): number {
	let points = Number(pt) || 0
	return isNaN(points) ? 0 : Math.round(points * ONEPT)
}

/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 *
 * @param {number} d - degrees
 * @returns {number} rot - value
 */
export function convertRotationDegrees(d: number): number {
	d = d || 0
	return Math.round((d > 360 ? d - 360 : d) * 60000)
}

/**
 * Converts component value to hex value
 * @param {number} c - component color
 * @returns {string} hex string
 */
export function componentToHex(c: number): string {
	let hex = c.toString(16)
	return hex.length === 1 ? '0' + hex : hex
}

/**
 * Converts RGB colors from css selectors to Hex for Presentation colors
 * @param {number} r - red value
 * @param {number} g - green value
 * @param {number} b - blue value
 * @returns {string} XML string
 */
export function rgbToHex(r: number, g: number, b: number): string {
	return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase()
}

const percentColorModifiers = [
	'alpha',
	'alphaMod',
	'alphaOff',
	'blue',
	'blueMod',
	'blueOff',
	'green',
	'greenMod',
	'greenOff',
	'red',
	'redMod',
	'redOff',
	'hue',
	'hueMod',
	'hueOff',
	'lum',
	'lumMod',
	'lumOff',
	'sat',
	'satMod',
	'satOff',
	'shade',
	'tint',
]

const flagColorModifiers = ['comp', 'gray', 'inv', 'gamma']

function handleModifiedColorProps(color: ModifiedThemeColor): string {
	let output = ''

	for (let modifier of percentColorModifiers) {
		let modifierValue = color[modifier]
		if (modifierValue !== undefined) {
			output += `<a:${modifier} val="${Math.round(modifierValue * 1000)}"/>`
		}
	}

	for (let modifier of flagColorModifiers) {
		if (color[modifier]) {
			output += `<a:${modifier}/>`
		}
	}

	return output
}

/**
 * Create either a `a:schemeClr` - (scheme color) or `a:srgbClr` (hexa representation).
 * @param {string|SCHEME_COLORS} colorStr - hexa representation (eg. "FFFF00") or a scheme color constant (eg. pptx.SchemeColor.ACCENT1)
 * @param {string} innerElements - additional elements that adjust the color and are enclosed by the color element
 * @returns {string} XML string
 */
export function createColorElement(colorInput: string | SCHEME_COLORS | ModifiedThemeColor, innerElements?: string): string {
	let colorStr = typeof colorInput === 'object' ? colorInput.baseColor : colorInput
	let colorVal = (colorStr || '').replace('#', '')
	let isHexaRgb = REGEX_HEX_COLOR.test(colorVal)

	if (typeof colorInput == 'object') {
		innerElements = (innerElements || '') + handleModifiedColorProps(colorInput)
	}

	if (
		!isHexaRgb &&
		colorVal !== SchemeColor.background1 &&
		colorVal !== SchemeColor.background2 &&
		colorVal !== SchemeColor.text1 &&
		colorVal !== SchemeColor.text2 &&
		colorVal !== SchemeColor.accent1 &&
		colorVal !== SchemeColor.accent2 &&
		colorVal !== SchemeColor.accent3 &&
		colorVal !== SchemeColor.accent4 &&
		colorVal !== SchemeColor.accent5 &&
		colorVal !== SchemeColor.accent6
	) {
		console.warn(`"${colorVal}" is not a valid scheme color or hexa RGB! "${DEF_FONT_COLOR}" is used as a fallback. Pass 6-digit RGB or 'pptx.SchemeColor' values`)
		colorVal = DEF_FONT_COLOR
	}

	let tagName = isHexaRgb ? 'srgbClr' : 'schemeClr'
	let colorAttr = 'val="' + (isHexaRgb ? colorVal.toUpperCase() : colorVal) + '"'

	return innerElements ? `<a:${tagName} ${colorAttr}>${innerElements}</a:${tagName}>` : `<a:${tagName} ${colorAttr}/>`
}

/**
 * Creates `a:glow` element
 * @param {TextGlowProps} options glow properties
 * @param {TextGlowProps} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 *	{ size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
export function createGlowElement(options: TextGlowProps, defaults: TextGlowProps): string {
	let strXml = '',
		opts = getMix(defaults, options),
		size = Math.round(opts['size'] * ONEPT),
		color = opts['color'],
		opacity = Math.round(opts['opacity'] * 100000)

	strXml += `<a:glow rad="${size}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += `</a:glow>`

	return strXml
}

/**
 * Create color selection
 * @param {Color | ShapeFillProps | ShapeLineProps} props fill props
 * @returns XML string
 */
export function genXmlColorSelection(props: Color | ShapeFillProps | ShapeLineProps): string {
	let fillType = 'solid'
	let colorVal: string | ModifiedThemeColor = ''
	let internalElements = ''
	let outText = ''

	if (props) {
		if (typeof props === 'string') colorVal = props
		else if ('baseColor' in props) {
			internalElements = handleModifiedColorProps(props)
			colorVal = props.baseColor
		} else {
			if (props.type) fillType = props.type
			if (props.color) colorVal = props.color
			if (props.alpha) internalElements += `<a:alpha val="${Math.round((100 - props.alpha) * 1000)}"/>` // DEPRECATED: @deprecated v3.3.0
			if (props.transparency) internalElements += `<a:alpha val="${Math.round((100 - props.transparency) * 1000)}"/>`
		}

		switch (fillType) {
			case 'solid':
				outText += `<a:solidFill>${createColorElement(colorVal, internalElements)}</a:solidFill>`
				break
			default:
				outText += '' // @note need a statement as having only "break" is removed by rollup, then tiggers "no-default" js-linter
				break
		}
	}

	return outText
}

/**
 * Get a new rel ID (rId) for charts, media, etc.
 * @param {PresSlide} target - the slide to use
 * @returns {number} count of all current rels plus 1 for the caller to use as its "rId"
 */
export function getNewRelId(target: PresSlide): number {
	return target._rels.length + target._relsChart.length + target._relsMedia.length + 1
}
