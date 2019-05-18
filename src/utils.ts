/**
* PptxGenJS Utils
*/

// Basic UUID Generator Adapted from:
// https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
export function getUuid(uuidFormat:string) {
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
	for (var i=0; i<=arguments.length; i++){
		var oN = arguments[i];
		if ( oN ) Object.keys(oN).forEach(function(key){ objMix[key] = oN[key]; });
	}
	return objMix;
}

/**
 * DESC: Replace special XML characters with HTML-encoded strings
 */
export function encodeXmlEntities(inStr:string) {
	// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	if ( typeof inStr === 'undefined' || inStr == null ) return "";
	return inStr.toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/\'/g,'&apos;');
}
