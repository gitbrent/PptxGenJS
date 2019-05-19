/**
* PptxGenJS Utils
*/

import { EMU } from './enums'

// Basic UUID Generator Adapted from:
// https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
export function getUuid(uuidFormat: string) {
	return uuidFormat.replace(/[xy]/g, function(c) {
		var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
		return v.toString(16);
	});
}

/**
 * shallow mix, returns new object
 */
export function getMix(o1, o2, etc?) {
	var objMix = {};
	for (var i = 0; i <= arguments.length; i++) {
		var oN = arguments[i];
		if (oN) Object.keys(oN).forEach(function(key) { objMix[key] = oN[key]; });
	}
	return objMix;
}

/**
 * DESC: Replace special XML characters with HTML-encoded strings
 */
export function encodeXmlEntities(inStr: string) {
	// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	if (typeof inStr === 'undefined' || inStr == null) return "";
	return inStr.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/\'/g, '&apos;');
}

export function inch2Emu(inches: number) {
	// FIRST: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
	// Any value over 100 damn sure isnt inches, must be EMU already, so just return it
	if (inches > 100) return inches;
	if (typeof inches == 'string') inches = Number(inches.replace(/in*/gi, ''));
	return Math.round(EMU * inches);
}

export function getSmartParseNumber(inVal:number|string, inDir:"X"|"Y") {
	// FIRST: Convert string numeric value if reqd
	if (typeof inVal == 'string' && !isNaN(Number(inVal))) inVal = Number(inVal);

	// CASE 1: Number in inches
	// Figure any number less than 100 is inches
	if (typeof inVal == 'number' && inVal < 100) return inch2Emu(inVal);

	// CASE 2: Number is already converted to something other than inches
	// Figure any number greater than 100 is not inches! :)  Just return it (its EMU already i guess??)
	if (typeof inVal == 'number' && inVal >= 100) return inVal;

	// CASE 3: Percentage (ex: '50%')
	if (typeof inVal == 'string' && inVal.indexOf('%') > -1) {
		if (inDir && inDir == 'X') return Math.round((parseFloat(inVal, 10) / 100) * gObjPptx.pptLayout.width);
		if (inDir && inDir == 'Y') return Math.round((parseFloat(inVal, 10) / 100) * gObjPptx.pptLayout.height);
		// Default: Assume width (x/cx)
		return Math.round((parseFloat(inVal, 10) / 100) * gObjPptx.pptLayout.width);
	}

	// LAST: Default value
	return 0;
}
