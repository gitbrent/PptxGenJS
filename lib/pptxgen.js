var PptxGenJS =
/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId])
/******/ 			return installedModules[moduleId].exports;
/******/
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			exports: {},
/******/ 			id: moduleId,
/******/ 			loaded: false
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.loaded = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	var _index = __webpack_require__(1);
	
	var _index2 = _interopRequireDefault(_index);
	
	function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }
	
	module.exports = _index2.default;

/***/ },
/* 1 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();
	
	var _constante = __webpack_require__(2);
	
	var _helpers = __webpack_require__(3);
	
	var _slide = __webpack_require__(4);
	
	var _slide2 = _interopRequireDefault(_slide);
	
	var _exportToPptx = __webpack_require__(5);
	
	var _exportToPptx2 = _interopRequireDefault(_exportToPptx);
	
	var _slideTable = __webpack_require__(28);
	
	var _slideTable2 = _interopRequireDefault(_slideTable);
	
	function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }
	
	function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
	
	var PptxGenJS = function () {
	    function PptxGenJS() {
	        _classCallCheck(this, PptxGenJS);
	
	        //this.slideNum = 0;
	        this.shapes = typeof gObjPptxShapes !== 'undefined' ? gObjPptxShapes : _constante.BASE_SHAPES;
	        this.masters = typeof gObjPptxMasters !== 'undefined' ? gObjPptxMasters : {};
	    }
	
	    /**
	     * Gets the version of this library
	     */
	
	
	    _createClass(PptxGenJS, [{
	        key: 'getVersion',
	        value: function getVersion() {
	            return _constante.APP_VER;
	        }
	    }, {
	        key: 'setTitle',
	
	
	        /**
	         * Sets the Presentation's Title
	         */
	        value: function setTitle(inStrTitle) {
	            _slide2.default.gObjPptx.title = inStrTitle;
	        }
	    }, {
	        key: 'setLayout',
	        value: function setLayout(inLayout) {
	
	            if ($.inArray(inLayout, Object.keys(_constante.LAYOUTS)) > -1) {
	                _slide2.default.gObjPptx.pptLayout = _constante.LAYOUTS[inLayout];
	            } else {
	                try {
	                    console.warn('UNKNOWN LAYOUT! Valid values = ' + Object.keys(_constante.LAYOUTS));
	                } catch (ex) {}
	            }
	            return this;
	        }
	
	        /**
	         * Gets the Presentation's Slide Layout {object}: [screen4x3, screen16x9, widescreen]
	         */
	        /*    getLayout() {
	            return this._layout;
	        }*/
	
	    }, {
	        key: 'addNewSlide',
	        value: function addNewSlide(isGroup) {
	            var slide = new _slide2.default(isGroup);
	            return slide.addNewSlide();
	        }
	    }, {
	        key: 'addSlidesForTable',
	        value: function addSlidesForTable(tabEleId, inOpts) {
	            var slideTable = new _slideTable2.default();
	            slideTable.addSlidesForTable(tabEleId, inOpts);
	        }
	
	        /**
	         * Export the Presentation to an .pptx file
	         * @param {string} [inStrExportName] - Filename to use for the export
	         */
	
	    }, {
	        key: 'save',
	        value: function save(inStrExportName) {
	
	            var exportPptx = new _exportToPptx2.default();
	            exportPptx.save(inStrExportName);
	        }
	    }]);
	
	    return PptxGenJS;
	}();
	
	exports.default = PptxGenJS;

/***/ },
/* 2 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	var LAYOUTS = {
	    'LAYOUT_4x3': {
	        name: 'screen4x3',
	        width: 9144000,
	        height: 6858000
	    },
	    'LAYOUT_16x9': {
	        name: 'screen16x9',
	        width: 9144000,
	        height: 5143500
	    },
	    'LAYOUT_16x10': {
	        name: 'screen16x10',
	        width: 9144000,
	        height: 5715000
	    },
	    'LAYOUT_WIDE': {
	        name: 'custom',
	        width: 12191996,
	        height: 6858000
	    }
	};
	
	var BASE_SHAPES = {
	    RECTANGLE: {
	        'displayName': 'Rectangle',
	        'name': 'rect',
	        'avLst': {}
	    },
	    LINE: {
	        'displayName': 'Line',
	        'name': 'line',
	        'avLst': {}
	    }
	};
	
	var APP_VER = "1.1.0";
	
	exports.LAYOUTS = LAYOUTS;
	exports.BASE_SHAPES = BASE_SHAPES;
	exports.APP_VER = APP_VER;

/***/ },
/* 3 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };
	
	exports.rgbToHex = rgbToHex;
	exports.inch2Emu = inch2Emu;
	exports.getSizeFromImage = getSizeFromImage;
	exports.calcEmuCellHeightForStr = calcEmuCellHeightForStr;
	exports.parseTextToLines = parseTextToLines;
	exports.getShapeInfo = getShapeInfo;
	exports.getSmartParseNumber = getSmartParseNumber;
	exports.decodeXmlEntities = decodeXmlEntities;
	exports.genXmlColorSelection = genXmlColorSelection;
	exports.convertImgToDataURLviaCanvas = convertImgToDataURLviaCanvas;
	exports.genXmlBodyProperties = genXmlBodyProperties;
	exports.genXmlTextCommand = genXmlTextCommand;
	
	var _slide = __webpack_require__(4);
	
	var _slide2 = _interopRequireDefault(_slide);
	
	function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }
	
	var EMU = 914400,
	    CRLF = '\r\n';
	
	function componentToHex(c) {
	    var hex = c.toString(16);
	    return hex.length == 1 ? "0" + hex : hex;
	}
	
	/**
	 * Used by {addSlidesForTable} to convert RGB colors from jQuery selectors to Hex for Presentation colors
	 */
	function rgbToHex(r, g, b) {
	    if (!Number.isInteger(r)) {
	        try {
	            console.warn('Integer expected!');
	        } catch (ex) {}
	    }
	    return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase();
	}
	
	function inch2Emu(inches) {
	    // FIRST: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
	    // Any value over 100 damn sure isnt inches, must be EMU already, so just return it
	    if (inches > 100) return inches;
	    if (typeof inches == 'string') inches = Number(inches.replace(/in*/gi, ''));
	    return Math.round(EMU * inches);
	}
	
	function getSizeFromImage(inImgUrl) {
	    // A: Create
	    var image = new Image();
	
	    // B: Set onload event
	    image.onload = function () {
	        // FIRST: Check for any errors: This is the best method (try/catch wont work, etc.)
	        if (this.width + this.height == 0) {
	            return {
	                width: 0,
	                height: 0
	            };
	        }
	        var obj = {
	            width: this.width,
	            height: this.height
	        };
	        return obj;
	    };
	    image.onerror = function () {
	        try {
	            console.error('[Error] Unable to load image: ' + inImgUrl);
	        } catch (ex) {}
	    };
	
	    // C: Load image
	    image.src = inImgUrl;
	}
	
	function calcEmuCellHeightForStr(cell, inIntWidthInches) {
	    // FORMULA for char-per-inch: (desired chars per line) / (font size [chars-per-inch]) = (reqd print area in inches)
	    var GRATIO = 2.61803398875; // "Golden Ratio"
	    var intCharPerInch = -1,
	        intCalcGratio = 0;
	
	    // STEP 1: Calc chars-per-inch [pitch]
	    // SEE: CPL Formula from http://www.pearsonified.com/2012/01/characters-per-line.php
	    intCharPerInch = 120 / cell.opts.font_size;
	
	    // STEP 2: Calc line count
	    var intLineCnt = Math.floor(cell.text.length / (intCharPerInch * inIntWidthInches));
	    if (intLineCnt < 1) intLineCnt = 1; // Dont allow line count to be 0!
	
	    // STEP 3: Calc cell height
	    var intCellH = intLineCnt * (cell.opts.font_size * 2 / 100);
	    if (intLineCnt > 8) intCellH = intCellH * 0.9;
	
	    // STEP 4: Add cell padding to height
	    if (cell.opts.marginPt && Array.isArray(cell.opts.marginPt)) {
	        intCellH += cell.opts.marginPt[0] / ONEPT * (1 / 72) + cell.opts.marginPt[2] / ONEPT * (1 / 72);
	    } else if (cell.opts.marginPt && Number.isInteger(cell.opts.marginPt)) {
	        intCellH += cell.opts.marginPt / ONEPT * (1 / 72) + cell.opts.marginPt / ONEPT * (1 / 72);
	    }
	
	    // LAST: Return size
	    return inch2Emu(intCellH);
	}
	
	function parseTextToLines(inStr, inFontSize, inWidth) {
	    var U = 2.2; // Character Constant thingy
	    var CPL = inWidth / (inFontSize / U);
	    var arrLines = [];
	    var strCurrLine = '';
	
	    // A: Remove leading/trailing space
	    inStr = $.trim(inStr);
	
	    // B: Build line array
	    $.each(inStr.split('\n'), function (i, line) {
	        $.each(line.split(' '), function (i, word) {
	            if (strCurrLine.length + word.length + 1 < CPL) {
	                strCurrLine += word + " ";
	            } else {
	                if (strCurrLine) arrLines.push(strCurrLine);
	                strCurrLine = word + " ";
	            }
	        });
	        // All words for this line have been exhausted, flush buffer to new line, clear line var
	        if (strCurrLine) arrLines.push($.trim(strCurrLine) + CRLF);
	        strCurrLine = "";
	    });
	
	    // C: Remove trailing linebreak
	    arrLines[arrLines.length - 1] = $.trim(arrLines[arrLines.length - 1]);
	
	    // D: Return lines
	    return arrLines;
	}
	
	function getShapeInfo(shapeName) {
	    if (!shapeName) return gObjPptxShapes.RECTANGLE;
	
	    if ((typeof shapeName === 'undefined' ? 'undefined' : _typeof(shapeName)) == 'object' && shapeName.name && shapeName.displayName && shapeName.avLst) return shapeName;
	
	    if (gObjPptxShapes[shapeName]) return gObjPptxShapes[shapeName];
	
	    var objShape = gObjPptxShapes.filter(function (obj) {
	        return obj.name == shapeName || obj.displayName;
	    })[0];
	    if (typeof objShape !== 'undefined' && objShape != null) return objShape;
	
	    return gObjPptxShapes.RECTANGLE;
	}
	
	function getSmartParseNumber(inVal, inDir, gObjPptx) {
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
	        if (inDir && inDir == 'X') return Math.round(parseInt(inVal, 10) / 100 * gObjPptx.pptLayout.width);
	        if (inDir && inDir == 'Y') return Math.round(parseInt(inVal, 10) / 100 * gObjPptx.pptLayout.height);
	        // Default: Assume width (x/cx)
	        return Math.round(parseInt(inVal, 10) / 100 * gObjPptx.pptLayout.width);
	    }
	
	    // LAST: Default value
	    return 0;
	}
	
	function decodeXmlEntities(inStr) {
	    // NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	    if (typeof inStr === 'undefined' || inStr == null) return "";
	    return inStr.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/\'/g, '&apos;');
	}
	
	function genXmlColorSelection(color_info, back_info) {
	    var outText = '';
	    var colorVal;
	    var fillType = 'solid';
	    var internalElements = '';
	
	    if (back_info) {
	        outText += '<p:bg><p:bgPr>';
	        outText += genXmlColorSelection(back_info, false);
	        outText += '<a:effectLst/>';
	        outText += '</p:bgPr></p:bg>';
	    }
	
	    if (color_info) {
	        if (typeof color_info == 'string') colorVal = color_info;else {
	            if (color_info.type) fillType = color_info.type;
	            if (color_info.color) colorVal = color_info.color;
	            if (color_info.alpha) internalElements += '<a:alpha val="' + (100 - color_info.alpha) + '000"/>';
	        }
	
	        switch (fillType) {
	            case 'solid':
	                outText += '<a:solidFill><a:srgbClr val="' + colorVal + '">' + internalElements + '</a:srgbClr></a:solidFill>';
	                break;
	        }
	    }
	
	    return outText;
	}
	
	function convertImgToDataURLviaCanvas(slideRel) {
	    // A: Create
	    var image = new Image();
	    // B: Set onload event
	    image.onload = function () {
	        // First: Check for any errors: This is the best method (try/catch wont work, etc.)
	        if (this.width + this.height == 0) {
	            this.onerror();
	            return;
	        }
	        var canvas = document.createElement('CANVAS');
	        var ctx = canvas.getContext('2d');
	        canvas.height = this.height;
	        canvas.width = this.width;
	        ctx.drawImage(this, 0, 0);
	        // Users running on local machine will get the following error:
	        // "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
	        // when the canvas.toDataURL call executes below.
	        try {
	            callbackImgToDataURLDone(canvas.toDataURL(slideRel.type), slideRel);
	        } catch (ex) {
	            this.onerror();
	            console.log("NOTE: Browsers wont let you load/convert local images! (search for --allow-file-access-from-files)");
	            return;
	        }
	        canvas = null;
	    };
	    image.onerror = function () {
	        try {
	            console.error('[Error] Unable to load image: ' + slideRel.path);
	        } catch (ex) {}
	        // Return a predefined "Broken image" graphic so the user will see something on the slide
	        callbackImgToDataURLDone('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGQAAAB3CAYAAAD1oOVhAAAGAUlEQVR4Xu2dT0xcRRzHf7tAYSsc0EBSIq2xEg8mtTGebVzEqOVIolz0siRE4gGTStqKwdpWsXoyGhMuyAVJOHBgqyvLNgonDkabeCBYW/8kTUr0wsJC+Wfm0bfuvn37Znbem9mR9303mJnf/Pb7ed95M7PDI5JIJPYJV5EC7e3t1N/fT62trdqViQCIu+bVgpIHEo/Hqbe3V/sdYVKHyWSSZmZm8ilVA0oeyNjYmEnaVC2Xvr6+qg5fAOJAz4DU1dURGzFSqZRVqtMpAFIGyMjICC0vL9PExIRWKADiAYTNshYWFrRCARAOEFZcCKWtrY0GBgaUTYkBRACIE4rKZwqACALR5RQAqQCIDqcASIVAVDsFQCSAqHQKgEgCUeUUAPEBRIVTAMQnEBvK5OQkbW9vk991CoAEAMQJxc86BUACAhKUUwAkQCBBOAVAAgbi1ykAogCIH6cAiCIgsk4BEIVAZJwCIIqBVLqiBxANQFgXS0tLND4+zl08AogmIG5OSSQS1gGKwgtANAIRcQqAaAbCe6YASBWA2E6xDyeyDUl7+AKQMkDYYevm5mZHabA/Li4uUiaTsYLau8QA4gLE/hU7wajyYtv1hReDAiAOxQcHBymbzark4BkbQKom/X8dp9Npmpqasn4BIAYAYSnYp+4BBEAMUcCwNOCQsAKZnp62NtQOw8WmwT09PUo+ijaHsOMx7GppaaH6+nolH0Z10K2tLVpdXbW6UfV3mNqBdHd3U1NTk2rtlMRfW1uj2dlZAFGirkRQAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAGHqrm8caPzQ0WC1logbeiC7X3xJm0PvUmRzh45cuki1588FAmVn9BO6P3yF9utrqGH0MtW82S8UN9RA9v/4k7InjhcJFTs/TLVXLwmJV67S7vD7tHF5pKi46fYdosdOcOOGG8j1OcqefbFEJD9Q3GCwDhqT31HklS4A8VRgfYM2Op6k3bt/BQJl58J7lPvwg5JYNccepaMry0LPqFA7hCm39+NNyp2J0172b19QysGINj5CsRtpij57musOViH0QPJQXn6J9u7dlYJSFkbrMYolrwvDAJAC+WWdEpQz7FTgECeUCpzi6YxvvqXoM6eEhqnCSgDikEzUKUE7Aw7xuHctKB5OYU3dZlNR9syQdAaAcAYTC0pXF+39c09o2Ik+3EqxVKqiB7hbYAxZkk4pbBaEM+AQofv+wTrFwylBOQNABIGwavdfe4O2pg5elO+86l99nY58/VUF0byrYsjiSFluNlXYrOHcBar7+EogUADEQ0YRGHbzoKAASBkg2+9cpM1rV0tK2QOcXW7bLEFAARAXIF4w2DrDWoeUWaf4hQIgDiA8GPZ2iNfi0Q8UACkAIgrDbrJ385eDxaPLLrEsFAB5oG6lMPJQPLZZZKAACBGVhcG2Q+bmuLu2nk55e4jqPv1IeEoceiBeX7s2zCa5MAqdstl91vfXwaEGsv/rb5TtOFk6tWXOuJGh6KmnhO9sayrMninPx103JBtXblHkice58cINZP4Hyr5wpkgkdiChEmc4FWazLzenNKa/p0jncwDiqcD6BuWePk07t1asatZGoYQzSqA4nFJ7soNiP/+EUyfc25GI2GG53dHPrKo1g/1Cw4pIXLrzO+1c+/wg7tBbFDle/EbQcjFCPWQJCau5EoBoFpzXHYDwFNJcDiCaBed1ByA8hTSXA4hmwXndAQhPIc3lAKJZcF53AMJTSHM5gGgWnNcdgPAU0lwOIJoF53UHIDyFNJcfSiCdnZ0Ui8U0SxlMd7lcjubn561gh+Y1scFIU/0o/3sgeLO12E2k7UXKYumgFoAYdg8ACIAYpoBh6cAhAGKYAoalA4cAiGEKGJYOHAIghilgWDpwCIAYpoBh6cAhAGKYAoalA4cAiGEKGJYOHAIghilgWDpwCIAYpoBh6ZQ4JB6PKzviYthnNy4d9h+1M5mMlVckkUjsG5dhiBMCEMPg/wuOfrZZ/RSywQAAAABJRU5ErkJggg==', slideRel);
	    };
	    // C: Load image
	    image.src = slideRel.path;
	}
	
	function genXmlBodyProperties(objOptions) {
	    var bodyProperties = '<a:bodyPr';
	
	    if (objOptions && objOptions.bodyProp) {
	        // A: Enable or disable textwrapping none or square:
	        objOptions.bodyProp.wrap ? bodyProperties += ' wrap="' + objOptions.bodyProp.wrap + '" rtlCol="0"' : bodyProperties += ' wrap="square" rtlCol="0"';
	
	        // B: Set anchorPoints bottom, center or top:
	        if (objOptions.bodyProp.anchor) bodyProperties += ' anchor="' + objOptions.bodyProp.anchor + '"';
	        if (objOptions.bodyProp.anchorCtr) bodyProperties += ' anchorCtr="' + objOptions.bodyProp.anchorCtr + '"';
	
	        // C: Textbox margins [padding]:
	        if (objOptions.bodyProp.bIns || objOptions.bodyProp.bIns == 0) bodyProperties += ' bIns="' + objOptions.bodyProp.bIns + '"';
	        if (objOptions.bodyProp.lIns || objOptions.bodyProp.lIns == 0) bodyProperties += ' lIns="' + objOptions.bodyProp.lIns + '"';
	        if (objOptions.bodyProp.rIns || objOptions.bodyProp.rIns == 0) bodyProperties += ' rIns="' + objOptions.bodyProp.rIns + '"';
	        if (objOptions.bodyProp.tIns || objOptions.bodyProp.tIns == 0) bodyProperties += ' tIns="' + objOptions.bodyProp.tIns + '"';
	
	        // D: Close <a:bodyPr element
	        bodyProperties += '>';
	
	        // E: NEW: Add auto-fit type tags
	        if (objOptions.shrinkText) bodyProperties += '<a:normAutofit fontScale="85000" lnSpcReduction="20000" />'; // MS-PPT > Format Shape > Text Options: "Shrink text on overflow"
	        else if (objOptions.bodyProp.autoFit !== false) bodyProperties += '<a:spAutoFit/>'; // MS-PPT > Format Shape > Text Options: "Resize shape to fit text"
	
	        // LAST: Close bodyProp
	        bodyProperties += '</a:bodyPr>';
	    } else {
	        // DEFAULT:
	        bodyProperties += ' wrap="square" rtlCol="0"></a:bodyPr>';
	    }
	
	    return bodyProperties;
	}
	
	function genXmlTextCommand(text_info, text_string, slide_obj, slide_num) {
	
	    var area_opt_data = genXmlTextData(text_info, slide_obj);
	    var parsedText;
	    //var startInfo = '<a:rPr lang="en-US"' + area_opt_data.font_size + area_opt_data.bold + area_opt_data.italic + area_opt_data.underline + area_opt_data.char_spacing + ' dirty="0" smtClean="0"' + (area_opt_data.rpr_info != '' ? ('>' + area_opt_data.rpr_info) : '/>') + '<a:t>';
	    var startInfo = '<a:rPr lang="en-US"' + area_opt_data.font_size + area_opt_data.bold + area_opt_data.underline + area_opt_data.char_spacing + ' dirty="0" smtClean="0"' + (area_opt_data.rpr_info != '' ? '>' + area_opt_data.rpr_info : '/>') + '<a:t>';
	    var endTag = '</a:r>';
	    var outData = '<a:r>' + startInfo;
	
	    if (text_string.field) {
	        endTag = '</a:fld>';
	        var outTextField = pptxFields[text_string.field];
	        if (outTextField === null) {
	            for (var fieldIntName in pptxFields) {
	                if (pptxFields[fieldIntName] === text_string.field) {
	                    outTextField = text_string.field;
	                    break;
	                }
	            }
	
	            if (outTextField === null) outTextField = 'datetime';
	        }
	
	        outData = '<a:fld id="{' + gen_private.plugs.type.msoffice.makeUniqueID('5C7A2A3D') + '}" type="' + outTextField + '">' + startInfo;
	        outData += CreateFieldText(outTextField, slide_num);
	    } else {
	        // Automatic support for newline - split it into multi-p:
	        parsedText = text_string.split("\n");
	        if (parsedText.length > 1) {
	            var outTextData = '';
	            for (var i = 0, total_size_i = parsedText.length; i < total_size_i; i++) {
	                outTextData += outData + decodeXmlEntities(parsedText[i]);
	
	                if (i + 1 < total_size_i) {
	                    outTextData += '</a:t></a:r></a:p><a:p>';
	                }
	            }
	
	            outData = outTextData;
	        } else {
	            outData += text_string.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
	        }
	    }
	
	    var outBreakP = '';
	    if (text_info.breakLine) outBreakP += '</a:p><a:p>';
	
	    return outData + '</a:t>' + endTag + outBreakP;
	}
	
	function genXmlTextData(text_info, slide_obj) {
	    var out_obj = {};
	
	    out_obj.font_size = '';
	    out_obj.bold = '';
	    out_obj.underline = '';
	    out_obj.rpr_info = '';
	    out_obj.char_spacing = '';
	
	    if ((typeof text_info === 'undefined' ? 'undefined' : _typeof(text_info)) == 'object') {
	        if (text_info.bold) {
	            out_obj.bold = ' b="1"';
	        }
	
	        if (text_info.underline) {
	            out_obj.underline = ' u="sng"';
	        }
	
	        if (text_info.font_size) {
	            out_obj.font_size = ' sz="' + text_info.font_size + '00"';
	        }
	
	        if (text_info.char_spacing) {
	            out_obj.char_spacing = ' spc="' + text_info.char_spacing * 100 + '"';
	            // must also disable kerning; otherwise text won't actually expand
	            out_obj.char_spacing += ' kern="0"';
	        }
	
	        if (text_info.color) {
	            out_obj.rpr_info += genXmlColorSelection(text_info.color);
	        } else if (slide_obj && slide_obj.color) {
	            out_obj.rpr_info += genXmlColorSelection(slide_obj.color);
	        }
	
	        if (text_info.font_face) {
	            out_obj.rpr_info += '<a:latin typeface="' + text_info.font_face + '" pitchFamily="34" charset="0"/><a:cs typeface="' + text_info.font_face + '" pitchFamily="34" charset="0"/>';
	        }
	    } else {
	        if (slide_obj && slide_obj.color) out_obj.rpr_info += genXmlColorSelection(slide_obj.color);
	    }
	
	    if (out_obj.rpr_info != '') out_obj.rpr_info += '</a:rPr>';
	
	    return out_obj;
	}
	
	function callbackImgToDataURLDone(inStr, slideRel) {
	    var intEmpty = 0;
	
	    // STEP 1: Store base64 data for this image
	    slideRel.data = inStr;
	
	    // STEP 2: Call export function once all async processes have completed
	    $.each(_slide2.default.gObjPptx.slides, function (i, slide) {
	        $.each(slide.rels, function (i, rel) {
	            if (rel.path == slideRel.path) rel.data = inStr;
	            if (!rel.data) intEmpty++;
	        });
	    });
	
	    // STEP 3: Continue export process
	    if (intEmpty == 0) this; //.save();
	}

/***/ },
/* 4 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };
	
	var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();
	
	var _helpers = __webpack_require__(3);
	
	var _constante = __webpack_require__(2);
	
	function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
	
	var EMU = 914400,
	    SLDNUMFLDID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}';
	
	var Slide = function () {
	    function Slide() {
	        var isGroup = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : false;
	
	        _classCallCheck(this, Slide);
	
	        this.group = isGroup;
	        this.slideNum = 0;
	        this.slideObjNum = 0;
	    }
	
	    _createClass(Slide, [{
	        key: 'hasSlideNumber',
	        value: function hasSlideNumber(inBool) {
	            if (inBool) Slide.gObjPptx.slides[this.slideNum].hasSlideNumber = inBool;else return Slide.gObjPptx.slides[this.slideNum].hasSlideNumber;
	        }
	    }, {
	        key: 'getPageNumber',
	        value: function getPageNumber() {
	            return this.slideNum;
	        }
	    }, {
	        key: 'addNewSlide',
	        value: function addNewSlide(inMaster) {
	            var _this = this;
	
	            this.slideNum = Slide.gObjPptx.slides.length;
	            var pageNum = this.slideNum + 1;
	
	            // A: Add this SLIDE to PRESENTATION, Add default values as well
	            Slide.gObjPptx.slides[this.slideNum] = {};
	            Slide.gObjPptx.slides[this.slideNum].slide = new Slide(this.group);
	            Slide.gObjPptx.slides[this.slideNum].name = 'Slide ' + pageNum;
	            Slide.gObjPptx.slides[this.slideNum].numb = pageNum;
	            Slide.gObjPptx.slides[this.slideNum].data = [];
	            Slide.gObjPptx.slides[this.slideNum].rels = [];
	            Slide.gObjPptx.slides[this.slideNum].hasSlideNumber = false;
	
	            // C: Add 'Master Slide' attr to Slide if a valid master was provided
	            if (inMaster && this.masters) {
	                // A: Add images (do this before adding slide bkgd)
	                if (inMaster.images && inMaster.images.length > 0) {
	                    $.each(inMaster.images, function (i, image) {
	                        _this.addImage(image.src, (0, _helpers.inch2Emu)(image.x), (0, _helpers.inch2Emu)(image.y), (0, _helpers.inch2Emu)(image.cx), (0, _helpers.inch2Emu)(image.cy), image.data || '');
	                    });
	                }
	
	                // B: Add any Slide Background: Image or Fill
	                if (inMaster.bkgd && inMaster.bkgd.src) {
	                    var slideObjRels = Slide.gObjPptx.slides[this.slideNum].rels;
	                    var strImgExtn = inMaster.bkgd.src.substring(inMaster.bkgd.src.indexOf('.') + 1).toLowerCase();
	                    if (strImgExtn == 'jpg') strImgExtn = 'jpeg';
	                    if (strImgExtn == 'gif') strImgExtn = 'png'; // MS-PPT: canvas.toDataURL for gif comes out image/png, and PPT will show "needs repair" unless we do this
	                    // TODO 1.5: The next few lines are copies from .addImage above. A bad idea thats already bit my once! So of course it's makred as future :)
	                    var intRels = 1;
	                    for (var idx = 0; idx < Slide.gObjPptx.slides.length; idx++) {
	                        intRels += Slide.gObjPptx.slides[idx].rels.length;
	                    }
	                    slideObjRels.push({
	                        path: inMaster.bkgd.src,
	                        type: 'image/' + strImgExtn,
	                        extn: strImgExtn,
	                        data: inMaster.bkgd.data || '',
	                        rId: intRels + 1,
	                        Target: '../media/image' + intRels + '.' + strImgExtn
	                    });
	                    slide.bkgdImgRid = slideObjRels[slideObjRels.length - 1].rId;
	                } else if (inMaster.bkgd) {
	                    slide.back = inMaster.bkgd;
	                }
	
	                // C: Add shapes
	                if (inMaster.shapes && inMaster.shapes.length > 0) {
	                    $.each(inMaster.shapes, function (i, shape) {
	                        // 1: Grab all options (x, y, color, etc.)
	                        var objOpts = {};
	                        $.each(Object.keys(shape), function (i, key) {
	                            if (shape[key] != 'type') objOpts[key] = shape[key];
	                        });
	                        // 2: Create object using 'type'
	                        if (shape.type == 'text') slide.addText(shape.text, objOpts);else if (shape.type == 'line') slide.addShape(_this.shapes.LINE, objOpts);
	                    });
	                }
	
	                // D: Slide Number
	                if (typeof inMaster.isNumbered !== 'undefined') this.slide.hasSlideNumber(inMaster.isNumbered);
	            }
	            return this;
	        }
	    }, {
	        key: 'addTable',
	        value: function addTable(arrTabRows, inOpt, tabOpt) {
	
	            var opt = (typeof inOpt === 'undefined' ? 'undefined' : _typeof(inOpt)) === 'object' ? inOpt : {};
	            if (opt.w) opt.cx = opt.w;
	            if (opt.h) opt.cy = opt.h;
	
	            // STEP 1: REALITY-CHECK
	            if (arrTabRows == null || arrTabRows.length == 0 || !Array.isArray(arrTabRows)) {
	                try {
	                    console.warn('[warn] addTable: Array expected!');
	                } catch (ex) {}
	                return null;
	            }
	
	            // STEP 2: Grab Slide object count
	            this.slideObjNum = Slide.gObjPptx.slides[this.slideNum].data.length;
	
	            // STEP 3: Set default options if needed
	            if (typeof opt.x === 'undefined') opt.x = EMU / 2;
	            if (typeof opt.y === 'undefined') opt.y = EMU;
	            if (typeof opt.cx === 'undefined') opt.cx = Slide.gObjPptx.pptLayout.width - EMU / 2;
	            // Dont do this for cy - leaving it null triggers auto-rowH in makeXMLSlide function
	
	            // STEP 4: We use different logic in makeSlide (smartCalc is not used), so convert to EMU now
	            if (opt.x < 20) opt.x = (0, _helpers.inch2Emu)(opt.x);
	            if (opt.y < 20) opt.y = (0, _helpers.inch2Emu)(opt.y);
	            if (opt.w < 20) opt.w = (0, _helpers.inch2Emu)(opt.w);
	            if (opt.h < 20) opt.h = (0, _helpers.inch2Emu)(opt.h);
	            if (opt.cx < 20) opt.cx = (0, _helpers.inch2Emu)(opt.cx);
	            if (opt.cy && opt.cy < 20) opt.cy = (0, _helpers.inch2Emu)(opt.cy);
	            //
	            if (tabOpt && Array.isArray(tabOpt.colW)) {
	                $.each(tabOpt.colW, function (i, colW) {
	                    if (colW < 20) tabOpt.colW[i] = (0, _helpers.inch2Emu)(colW);
	                });
	            }
	
	            // Handle case where user passed in a simple array
	            var arrTemp = $.extend(true, [], arrTabRows);
	            if (!Array.isArray(arrTemp[0])) arrTemp = [$.extend(true, [], arrTabRows)];
	
	            // STEP 5: Add data
	            // NOTE: Use extend to avoid mutation
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum] = {
	                type: 'table',
	                arrTabRows: arrTemp,
	                options: $.extend(true, {}, opt),
	                objTabOpts: $.extend(true, {}, tabOpt) || {}
	            };
	
	            // LAST: Return this Slide object
	            return this;
	        }
	    }, {
	        key: 'addText',
	        value: function addText(text, opt) {
	            // STEP 1: Grab Slide object count
	            this.slideObjNum = Slide.gObjPptx.slides[this.slideNum].data.length;
	
	            // ROBUST: Convert attr values that will likely be passed by users to valid OOXML values
	            if (opt.valign) opt.valign = opt.valign.toLowerCase().replace(/^c.*/i, 'ctr').replace(/^m.*/i, 'ctr').replace(/^t.*/i, 't').replace(/^b.*/i, 'b');
	            if (opt.align) opt.align = opt.align.toLowerCase().replace(/^c.*/i, 'center').replace(/^m.*/i, 'center').replace(/^l.*/i, 'left').replace(/^r.*/i, 'right');
	
	            // STEP 2: Set props
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum] = {};
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].type = 'text';
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].text = text;
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options = (typeof opt === 'undefined' ? 'undefined' : _typeof(opt)) === 'object' ? opt : {};
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp = jQuery.extend({}, opt.bodyProp);
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.autoFit = opt.autoFit || false; // If true, shape will collapse to text size (Fit To Shape)
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.anchor = opt.valign || 'ctr'; // VALS: [t,ctr,b]
	            if (opt.inset && !isNaN(Number(opt.inset)) || opt.inset == 0) {
	                Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.lIns = (0, _helpers.inch2Emu)(opt.inset);
	                Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.rIns = (0, _helpers.inch2Emu)(opt.inset);
	                Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.tIns = (0, _helpers.inch2Emu)(opt.inset);
	                Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.bIns = (0, _helpers.inch2Emu)(opt.inset);
	            }
	
	            // LAST: Return
	            return this;
	        }
	    }, {
	        key: 'addShape',
	        value: function addShape(shape, opt) {
	            // STEP 1: Grab Slide object count
	            this.slideObjNum = Slide.gObjPptx.slides[this.slideNum].data.length;
	
	            // STEP 2: Set props
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum] = {};
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].type = 'text';
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options = (typeof opt === 'undefined' ? 'undefined' : _typeof(opt)) == 'object' ? opt : {};
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.shape = shape;
	
	            return this;
	        }
	    }, {
	        key: 'addImage',
	        value: function addImage(strImagePath, intPosX, intPosY, intSizeX, intSizeY, strImageData, strImgData) {
	            var intRels = 1;
	
	            // FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
	            // TODO: FUTURE: DEPRECATED: Only allow object param in 1.5 or 2.0
	            if ((typeof strImagePath === 'undefined' ? 'undefined' : _typeof(strImagePath)) === 'object') {
	                intPosX = strImagePath.x || 0;
	                intPosY = strImagePath.y || 0;
	                intSizeX = strImagePath.cx || strImagePath.w || 0;
	                intSizeY = strImagePath.cy || strImagePath.h || 0;
	                strImageData = strImagePath.data || '';
	                strImagePath = strImagePath.path || ''; // This line must be last as were about to ovewrite ourself!
	            }
	            // REALITY-CHECK:
	            if (!strImagePath && !strImgData) {
	                try {
	                    console.error("ERROR: Image can't be empty");
	                } catch (ex) {}
	                return null;
	            }
	
	            // STEP 1: Set vars for this Slide
	            this.slideObjNum = Slide.gObjPptx.slides[this.slideNum].data.length;
	            var slideObjRels = Slide.gObjPptx.slides[this.slideNum].rels;
	            var strImgExtn = 'png'; // Every image is encoded via canvas>base64, so they all come out as png (use of another extn will cause "needs repair" dialog on open in PPT)
	
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum] = {};
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].type = 'image';
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].image = strImagePath;
	
	            // STEP 2: Set image properties & options
	            // TODO 1.1: Measure actual image when no intSizeX/intSizeY params passed
	            // ....: This is an async process: we need to make getSizeFromImage use callback, then set H/W...
	            // if ( !intSizeX || !intSizeY ) { var imgObj = getSizeFromImage(strImagePath);
	            var imgObj = {
	                width: 1,
	                height: 1
	            };
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options = {};
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.x = intPosX || 0;
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.y = intPosY || 0;
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.cx = intSizeX || imgObj.width;
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.cy = intSizeY || imgObj.height;
	
	            // STEP 3: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	            // NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
	            $.each(Slide.gObjPptx.slides, function (i, slide) {
	                intRels += slide.rels.length;
	            });
	            slideObjRels.push({
	                path: strImagePath,
	                type: 'image/' + strImgExtn,
	                extn: strImgExtn,
	                data: strImgData || '',
	                rId: intRels + 1,
	                Target: '../media/image' + intRels + '.' + strImgExtn
	            });
	            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].imageRid = slideObjRels[slideObjRels.length - 1].rId;
	
	            // LAST: Return this Slide
	            return this;
	        }
	    }], [{
	        key: 'header',
	        value: function header(inSlide) {
	
	            var strSlideXml = void 0,
	                aStr = [],
	                propertBg = [],
	                propertySpTree = [],
	                propertySp = [];
	
	            var head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="' + inSlide.name + '">';
	            aStr.push(head);
	
	            // STEP 2: Add background color or background image (if any)
	            // A: Background color
	            /*if ( inSlide.slide.back ) strSlideXml += genXmlColorSelection(false, inSlide.slide.back);
	            // B: Add background image (using Strech) (if any)
	            if ( inSlide.slide.bkgdImgRid ) {
	                // TODO 1.0: We should be doing this in the slideLayout...
	                strSlideXml += `<p:bg>
	                                <p:bgPr><a:blipFill dpi="0" rotWithShape="1">
	                                    <a:blip r:embed="rId${inSlide.slide.bkgdImgRid}"><a:lum/></a:blip>
	                                    <a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
	                                    <a:effectLst/></p:bgPr>
	                             </p:bg>`;
	            } */
	
	            if (inSlide.slide.back) aStrSlideXml.push((0, _helpers.genXmlColorSelection)(false, inSlide.slide.back));
	            // B: Add background image (using Strech) (if any)
	            if (inSlide.slide.bkgdImgRid) {
	                // TODO 1.0: We should be doing this in the slideLayout...
	                propertBg = ['<p:bg>', '<p:bgPr><a:blipFill dpi="0" rotWithShape="1">', '<a:blip r:embed="rId' + inSlide.slide.bkgdImgRid + '"><a:lum/></a:blip>', '<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>', '<a:effectLst/></p:bgPr>', '</p:bg>'];
	                aStr.push(propertBg.join(''));
	            }
	            propertySpTree = ['<p:spTree>', '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>', '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>', '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'];
	
	            aStr.push(propertySpTree.join(''));
	
	            // STEP 4: Add slide numbers if selected
	            // TODO 1.0: Fixed location sucks! Place near bottom corner using slide.size !!!
	            if (inSlide.hasSlideNumber) {
	                propertySp = ['<p:sp>', '<p:nvSpPr>', '<p:cNvPr id="25" name="Shape 25"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr></p:nvSpPr>', '<p:spPr>', '<a:xfrm><a:off x="' + EMU * 0.3 + '" y="' + EMU * 5.2 + '"/><a:ext cx="400000" cy="300000"/></a:xfrm>', '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', '<a:extLst>', '<a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">', '<ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext>', '</a:extLst>', '</p:spPr>', '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr/><a:fld id="' + SLDNUMFLDID + '" type="slidenum"/></a:p></p:txBody>', '</p:sp>'];
	                aStr.push(propertySp.join(''));
	            }
	            strSlideXml = aStr.join('');
	            return strSlideXml;
	        }
	    }, {
	        key: 'footer',
	        value: function footer() {
	            var footer = ['</p:spTree>', '</p:cSld>', '<p:clrMapOvr>', '<a:masterClrMapping/>', '</p:clrMapOvr>', '</p:sld>'];
	            return footer.join('');
	        }
	    }]);
	
	    return Slide;
	}();
	
	Slide.gObjPptx = {
	    title: 'PresePptxGenJS Presentation',
	    fileName: 'Presentation',
	    fileExtn: '.pptx',
	    pptLayout: _constante.LAYOUTS['LAYOUT_WIDE'],
	    slides: []
	};
	exports.default = Slide;

/***/ },
/* 5 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();
	
	var _contentType = __webpack_require__(6);
	
	var _contentType2 = _interopRequireDefault(_contentType);
	
	var _rels = __webpack_require__(7);
	
	var _rels2 = _interopRequireDefault(_rels);
	
	var _appXml = __webpack_require__(8);
	
	var _appXml2 = _interopRequireDefault(_appXml);
	
	var _coreXml = __webpack_require__(9);
	
	var _coreXml2 = _interopRequireDefault(_coreXml);
	
	var _presentationXmlRels = __webpack_require__(10);
	
	var _presentationXmlRels2 = _interopRequireDefault(_presentationXmlRels);
	
	var _slideLayoutXml = __webpack_require__(11);
	
	var _slideLayoutXml2 = _interopRequireDefault(_slideLayoutXml);
	
	var _slideLayoutRelXml = __webpack_require__(12);
	
	var _slideLayoutRelXml2 = _interopRequireDefault(_slideLayoutRelXml);
	
	var _slideMasterXml = __webpack_require__(13);
	
	var _slideMasterXml2 = _interopRequireDefault(_slideMasterXml);
	
	var _slideMasterRelXml = __webpack_require__(14);
	
	var _slideMasterRelXml2 = _interopRequireDefault(_slideMasterRelXml);
	
	var _themeXml = __webpack_require__(15);
	
	var _themeXml2 = _interopRequireDefault(_themeXml);
	
	var _presentationXml = __webpack_require__(16);
	
	var _presentationXml2 = _interopRequireDefault(_presentationXml);
	
	var _presPropsXml = __webpack_require__(17);
	
	var _presPropsXml2 = _interopRequireDefault(_presPropsXml);
	
	var _tableStyleXml = __webpack_require__(18);
	
	var _tableStyleXml2 = _interopRequireDefault(_tableStyleXml);
	
	var _viewPropsXml = __webpack_require__(19);
	
	var _viewPropsXml2 = _interopRequireDefault(_viewPropsXml);
	
	var _slideXml = __webpack_require__(20);
	
	var _slideXml2 = _interopRequireDefault(_slideXml);
	
	var _slideXmlRel = __webpack_require__(27);
	
	var _slideXmlRel2 = _interopRequireDefault(_slideXmlRel);
	
	var _slide = __webpack_require__(4);
	
	var _slide2 = _interopRequireDefault(_slide);
	
	function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }
	
	function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
	
	var ExportPptx = function () {
	    function ExportPptx() {
	        _classCallCheck(this, ExportPptx);
	
	        this.gObjPptx = _slide2.default.gObjPptx;
	        this.zip = new JSZip();
	        //this.zip = new com.sap.powerdesigner.web.galilei.common.util.Zip.create();
	        this.intSlideNum = 0;
	        this.intRels = 0;
	    }
	
	    _createClass(ExportPptx, [{
	        key: 'save',
	        value: function save(inStrExportName) {
	            var _this = this;
	
	            var intRels = 0,
	                arrImages = [];
	
	            // STEP 1: Set export title (if any)
	            if (inStrExportName) this.gObjPptx.fileName = inStrExportName;
	
	            // STEP 2: Total all images (rels) across the Presentation
	            // PERF: Only send unique image paths for encoding (encoding func will find and fill ALL matching img paths and fill)
	            var slides = this.gObjPptx.slides;
	            $.each(slides, function (i, slide) {
	                $.each(slide.rels, function (i, rel) {
	                    if (!rel.data && $.inArray(rel.path, arrImages) == -1) {
	                        intRels++;
	                        _this.convertImgToDataURLviaCanvas(rel, _this.callbackImgToDataURLDone);
	                        arrImages.push(rel.path);
	                    }
	                });
	            });
	
	            // STEP 3: Export now if there's no images to encode (otherwise, last async imgConvert call above will call exportFile)
	            if (intRels == 0) {
	                this.doExportPresentation();
	            };
	        }
	    }, {
	        key: 'doExportPresentation',
	        value: function doExportPresentation() {
	            this.zip.folder("_rels");
	            this.zip.folder("docProps");
	            this.zip.folder("ppt").folder("_rels");
	            this.zip.folder("ppt/media");
	            this.zip.folder("ppt/slideLayouts").folder("_rels");
	            this.zip.folder("ppt/slideMasters").folder("_rels");
	            this.zip.folder("ppt/slides").folder("_rels");
	            this.zip.folder("ppt/theme");
	
	            this.zip.file("[Content_Types].xml", (0, _contentType2.default)(this.gObjPptx));
	            this.zip.file("_rels/.rels", (0, _rels2.default)());
	            this.zip.file("docProps/app.xml", (0, _appXml2.default)(this.gObjPptx));
	            this.zip.file("docProps/core.xml", (0, _coreXml2.default)(this.gObjPptx));
	            this.zip.file("ppt/_rels/presentation.xml.rels", (0, _presentationXmlRels2.default)(this.gObjPptx));
	
	            // Create a Layout/Master/Rel/Slide file for each SLIDE
	            for (var idx = 0; idx < this.gObjPptx.slides.length; idx++) {
	                this.intSlideNum++;
	                this.zip.file("ppt/slideLayouts/slideLayout" + this.intSlideNum + ".xml", (0, _slideLayoutXml2.default)());
	                this.zip.file("ppt/slideLayouts/_rels/slideLayout" + this.intSlideNum + ".xml.rels", (0, _slideLayoutRelXml2.default)());
	                this.zip.file("ppt/slides/slide" + this.intSlideNum + ".xml", (0, _slideXml2.default)(this.gObjPptx.slides[idx], this.gObjPptx));
	                this.zip.file("ppt/slides/_rels/slide" + this.intSlideNum + ".xml.rels", (0, _slideXmlRel2.default)(this.intSlideNum, this.gObjPptx));
	            }
	            this.zip.file("ppt/slideMasters/slideMaster1.xml", (0, _slideMasterXml2.default)(this.gObjPptx));
	            this.zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", (0, _slideMasterRelXml2.default)(this.gObjPptx));
	
	            // Add all images
	            this.addAllImages();
	
	            this.zip.file("ppt/theme/theme1.xml", (0, _themeXml2.default)());
	            this.zip.file("ppt/presentation.xml", (0, _presentationXml2.default)(this.gObjPptx));
	            this.zip.file("ppt/presProps.xml", (0, _presPropsXml2.default)());
	            this.zip.file("ppt/tableStyles.xml", (0, _tableStyleXml2.default)());
	            this.zip.file("ppt/viewProps.xml", (0, _viewPropsXml2.default)());
	
	            // =======
	            // STEP 3: Push the PPTX file to browser
	            // =======
	            var strExportName = this.gObjPptx.fileName.toLowerCase().indexOf('.ppt') > -1 ? this.gObjPptx.fileName : this.gObjPptx.fileName + this.gObjPptx.fileExtn;
	            this.zip.generateAsync({
	                type: "blob"
	            }).then(function (content) {
	                //sap.galilei.ui.common.FileManager.saveAs( content, strExportName );
	                saveAs(content, strExportName);
	                _slide2.default.gObjPptx.slides = [];
	            });
	        }
	    }, {
	        key: 'addAllImages',
	        value: function addAllImages() {
	            for (var idx = 0; idx < this.gObjPptx.slides.length; idx++) {
	                for (var idy = 0; idy < this.gObjPptx.slides[idx].rels.length; idy++) {
	                    var id = this.gObjPptx.slides[idx].rels[idy].rId - 1;
	                    var data = this.gObjPptx.slides[idx].rels[idy].data;
	                    // data:image/png;base64
	                    var header = data.substring(0, data.indexOf(","));
	                    // NOTE: Trim the leading 'data:image/png;base64,' text as it is not needed (and image wont render correctly with it)
	                    var content = data.substring(data.indexOf(",") + 1);
	                    var extn = /data:image\/(\w+)/.exec(header)[1];
	                    var isBase64 = /base64/.test(header);
	                    this.zip.file("ppt/media/image" + id + "." + extn, content, {
	                        base64: isBase64
	                    });
	                }
	            }
	        }
	    }, {
	        key: 'convertImgToDataURLviaCanvas',
	        value: function convertImgToDataURLviaCanvas(slideRel) {
	            // A: Create
	            var self = this;
	            var image = new Image();
	            // B: Set onload event
	            image.onload = function () {
	                // First: Check for any errors: This is the best method (try/catch wont work, etc.)
	                if (this.width + this.height == 0) {
	                    this.onerror();
	                    return;
	                }
	                var canvas = document.createElement('CANVAS');
	                var ctx = canvas.getContext('2d');
	                canvas.height = this.height;
	                canvas.width = this.width;
	                ctx.drawImage(this, 0, 0);
	                // Users running on local machine will get the following error:
	                // "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
	                // when the canvas.toDataURL call executes below.
	                try {
	                    self.callbackImgToDataURLDone(canvas.toDataURL(slideRel.type), slideRel);
	                } catch (ex) {
	                    this.onerror();
	                    console.log("NOTE: Browsers wont let you load/convert local images! (search for --allow-file-access-from-files)");
	                    return;
	                }
	                canvas = null;
	            };
	            image.onerror = function () {
	                try {
	                    console.error('[Error] Unable to load image: ' + slideRel.path);
	                } catch (ex) {}
	                // Return a predefined "Broken image" graphic so the user will see something on the slide
	                self.callbackImgToDataURLDone('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGQAAAB3CAYAAAD1oOVhAAAGAUlEQVR4Xu2dT0xcRRzHf7tAYSsc0EBSIq2xEg8mtTGebVzEqOVIolz0siRE4gGTStqKwdpWsXoyGhMuyAVJOHBgqyvLNgonDkabeCBYW/8kTUr0wsJC+Wfm0bfuvn37Znbem9mR9303mJnf/Pb7ed95M7PDI5JIJPYJV5EC7e3t1N/fT62trdqViQCIu+bVgpIHEo/Hqbe3V/sdYVKHyWSSZmZm8ilVA0oeyNjYmEnaVC2Xvr6+qg5fAOJAz4DU1dURGzFSqZRVqtMpAFIGyMjICC0vL9PExIRWKADiAYTNshYWFrRCARAOEFZcCKWtrY0GBgaUTYkBRACIE4rKZwqACALR5RQAqQCIDqcASIVAVDsFQCSAqHQKgEgCUeUUAPEBRIVTAMQnEBvK5OQkbW9vk991CoAEAMQJxc86BUACAhKUUwAkQCBBOAVAAgbi1ykAogCIH6cAiCIgsk4BEIVAZJwCIIqBVLqiBxANQFgXS0tLND4+zl08AogmIG5OSSQS1gGKwgtANAIRcQqAaAbCe6YASBWA2E6xDyeyDUl7+AKQMkDYYevm5mZHabA/Li4uUiaTsYLau8QA4gLE/hU7wajyYtv1hReDAiAOxQcHBymbzark4BkbQKom/X8dp9Npmpqasn4BIAYAYSnYp+4BBEAMUcCwNOCQsAKZnp62NtQOw8WmwT09PUo+ijaHsOMx7GppaaH6+nolH0Z10K2tLVpdXbW6UfV3mNqBdHd3U1NTk2rtlMRfW1uj2dlZAFGirkRQAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAGHqrm8caPzQ0WC1logbeiC7X3xJm0PvUmRzh45cuki1588FAmVn9BO6P3yF9utrqGH0MtW82S8UN9RA9v/4k7InjhcJFTs/TLVXLwmJV67S7vD7tHF5pKi46fYdosdOcOOGG8j1OcqefbFEJD9Q3GCwDhqT31HklS4A8VRgfYM2Op6k3bt/BQJl58J7lPvwg5JYNccepaMry0LPqFA7hCm39+NNyp2J0172b19QysGINj5CsRtpij57musOViH0QPJQXn6J9u7dlYJSFkbrMYolrwvDAJAC+WWdEpQz7FTgECeUCpzi6YxvvqXoM6eEhqnCSgDikEzUKUE7Aw7xuHctKB5OYU3dZlNR9syQdAaAcAYTC0pXF+39c09o2Ik+3EqxVKqiB7hbYAxZkk4pbBaEM+AQofv+wTrFwylBOQNABIGwavdfe4O2pg5elO+86l99nY58/VUF0byrYsjiSFluNlXYrOHcBar7+EogUADEQ0YRGHbzoKAASBkg2+9cpM1rV0tK2QOcXW7bLEFAARAXIF4w2DrDWoeUWaf4hQIgDiA8GPZ2iNfi0Q8UACkAIgrDbrJ385eDxaPLLrEsFAB5oG6lMPJQPLZZZKAACBGVhcG2Q+bmuLu2nk55e4jqPv1IeEoceiBeX7s2zCa5MAqdstl91vfXwaEGsv/rb5TtOFk6tWXOuJGh6KmnhO9sayrMninPx103JBtXblHkice58cINZP4Hyr5wpkgkdiChEmc4FWazLzenNKa/p0jncwDiqcD6BuWePk07t1asatZGoYQzSqA4nFJ7soNiP/+EUyfc25GI2GG53dHPrKo1g/1Cw4pIXLrzO+1c+/wg7tBbFDle/EbQcjFCPWQJCau5EoBoFpzXHYDwFNJcDiCaBed1ByA8hTSXA4hmwXndAQhPIc3lAKJZcF53AMJTSHM5gGgWnNcdgPAU0lwOIJoF53UHIDyFNJcfSiCdnZ0Ui8U0SxlMd7lcjubn561gh+Y1scFIU/0o/3sgeLO12E2k7UXKYumgFoAYdg8ACIAYpoBh6cAhAGKYAoalA4cAiGEKGJYOHAIghilgWDpwCIAYpoBh6cAhAGKYAoalA4cAiGEKGJYOHAIghilgWDpwCIAYpoBh6ZQ4JB6PKzviYthnNy4d9h+1M5mMlVckkUjsG5dhiBMCEMPg/wuOfrZZ/RSywQAAAABJRU5ErkJggg==', slideRel);
	            };
	            // C: Load image
	            image.src = slideRel.path;
	        }
	    }, {
	        key: 'callbackImgToDataURLDone',
	        value: function callbackImgToDataURLDone(inStr, slideRel) {
	            var intEmpty = 0;
	
	            // STEP 1: Store base64 data for this image
	            slideRel.data = inStr;
	
	            // STEP 2: Call export function once all async processes have completed
	            $.each(this.gObjPptx.slides, function (i, slide) {
	                $.each(slide.rels, function (i, rel) {
	                    if (rel.path == slideRel.path) rel.data = inStr;
	                    if (!rel.data) intEmpty++;
	                });
	            });
	
	            // STEP 3: Continue export process
	            if (intEmpty == 0) this.doExportPresentation();
	        }
	    }]);
	
	    return ExportPptx;
	}();
	
	exports.default = ExportPptx;

/***/ },
/* 6 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	exports.default = makeXmlContTypes;
	function makeXmlContTypes(gObjPptx) {
	    var CRLF = '\r\n';
	
	    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + ' <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + ' <Default Extension="xml" ContentType="application/xml"/>' + ' <Default Extension="jpeg" ContentType="image/jpeg"/>' + ' <Default Extension="png" ContentType="image/png"/>' + ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' + ' <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' + ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' + ' <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' + ' <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' + ' <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>' + ' <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>';
	    $.each(gObjPptx.slides, function (idx, slide) {
	        strXml += '<Override PartName="/ppt/slideMasters/slideMaster' + (idx + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>';
	        strXml += '<Override PartName="/ppt/slideLayouts/slideLayout' + (idx + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
	        strXml += '<Override PartName="/ppt/slides/slide' + (idx + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
	    });
	    strXml += '</Types>';
	    return strXml;
	}

/***/ },
/* 7 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	exports.default = makeXmlRootRels;
	function makeXmlRootRels() {
	    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' + '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' + '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' + '</Relationships>';
	    return strXml;
	}

/***/ },
/* 8 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
					value: true
	});
	exports.default = makeXmlApp;
	function makeXmlApp(gObjPptx) {
					var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n\
						<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\
							<TotalTime>0</TotalTime>\
							<Words>0</Words>\
							<Application>Microsoft Office PowerPoint</Application>\
							<PresentationFormat>On-screen Show (4:3)</PresentationFormat>\
							<Paragraphs>0</Paragraphs>\
							<Slides>' + gObjPptx.slides.length + '</Slides>\
							<Notes>0</Notes>\
							<HiddenSlides>0</HiddenSlides>\
							<MMClips>0</MMClips>\
							<ScaleCrop>false</ScaleCrop>\
							<HeadingPairs>\
							  <vt:vector size="4" baseType="variant">\
							    <vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>\
							    <vt:variant><vt:i4>1</vt:i4></vt:variant>\
							    <vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>\
							    <vt:variant><vt:i4>' + gObjPptx.slides.length + '</vt:i4></vt:variant>\
							  </vt:vector>\
							</HeadingPairs>\
							<TitlesOfParts>';
					strXml += '<vt:vector size="' + (gObjPptx.slides.length + 1) + '" baseType="lpstr">';
					strXml += '<vt:lpstr>Office Theme</vt:lpstr>';
					$.each(gObjPptx.slides, function (idx, slideObj) {
									strXml += '<vt:lpstr>Slide ' + (idx + 1) + '</vt:lpstr>';
					});
					strXml += ' </vt:vector>\r\n\n          </TitlesOfParts>\r\n\n          <Company>PptxGenJS</Company>\r\n\n          <LinksUpToDate>false</LinksUpToDate>\r\n\n          <SharedDoc>false</SharedDoc>\r\n\n          <HyperlinksChanged>false</HyperlinksChanged>\r\n\n          <AppVersion>15.0000</AppVersion>\r\n\n        </Properties>';
					return strXml;
	}

/***/ },
/* 9 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
		value: true
	});
	exports.default = makeXmlCore;
	function makeXmlCore(gObjPptx) {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n\
							<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"\
								 xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"\
								 xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\
								<dc:title>' + gObjPptx.title + '</dc:title>\
								<dc:creator>PptxGenJS</dc:creator>\
								<cp:lastModifiedBy>PptxGenJS</cp:lastModifiedBy>\
								<cp:revision>1</cp:revision>\
								<dcterms:created xsi:type="dcterms:W3CDTF">' + new Date().toISOString() + '</dcterms:created>\
								<dcterms:modified xsi:type="dcterms:W3CDTF">' + new Date().toISOString() + '</dcterms:modified>\
							</cp:coreProperties>';
		return strXml;
	}

/***/ },
/* 10 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	  value: true
	});
	exports.default = makeXmlPresentationRels;
	function makeXmlPresentationRels(gObjPptx) {
	  var intRelNum = 0;
	  var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
	
	  strXml += '  <Relationship Id="rId1" Target="slideMasters/slideMaster1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
	  intRelNum++;
	
	  for (var idx = 1; idx <= gObjPptx.slides.length; idx++) {
	    intRelNum++;
	    strXml += '  <Relationship Id="rId' + intRelNum + '" Target="slides/slide' + idx + '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>';
	  }
	  intRelNum++;
	  strXml += '  <Relationship Id="rId' + intRelNum + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>' + '  <Relationship Id="rId' + (intRelNum + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>' + '  <Relationship Id="rId' + (intRelNum + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' + '  <Relationship Id="rId' + (intRelNum + 3) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>' + '</Relationships>';
	  return strXml;
	}

/***/ },
/* 11 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	  value: true
	});
	exports.default = makeXmlSlideLayout;
	function makeXmlSlideLayout() {
	  var strXml = void 0;
	  var SLDNUMFLDID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}';
	
	  strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
	  strXml += '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1">\r\n' + '<p:cSld name="Title Slide">' + '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="2130425"/><a:ext cx="7772400" cy="1470025"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>' + '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1371600" y="3886200"/><a:ext cx="6400800" cy="1752600"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle>' + '  <a:lvl1pPr marL="0"       indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr>' + '  <a:lvl2pPr marL="457200"  indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl2pPr>' + '  <a:lvl3pPr marL="914400"  indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl3pPr>' + '  <a:lvl4pPr marL="1371600" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl4pPr>' + '  <a:lvl5pPr marL="1828800" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl5pPr>' + '  <a:lvl6pPr marL="2286000" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl6pPr>' + '  <a:lvl7pPr marL="2743200" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl7pPr>' + '  <a:lvl8pPr marL="3200400" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl8pPr>' + '  <a:lvl9pPr marL="3657600" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl9pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master subtitle style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>' + '<p:cNvPr id="4" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{F8166F1F-CE9B-4651-A6AA-CD717754106B}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>01/01/2016</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>' + '<p:cNvPr id="5" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>' + '<p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="' + SLDNUMFLDID + '" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t></a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld>' + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
	  //
	  return strXml;
	}

/***/ },
/* 12 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	  value: true
	});
	exports.default = makeXmlSlideLayoutRel;
	//export function makeXmlSlideLayoutRel(inSlideNum) {
	function makeXmlSlideLayoutRel() {
	  var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
	  strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n';
	  //?strXml += '  <Relationship Id="rId'+ inSlideNum +'" Target="../slideMasters/slideMaster'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
	  //strXml += '  <Relationship Id="rId1" Target="../slideMasters/slideMaster'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
	  strXml += '  <Relationship Id="rId1" Target="../slideMasters/slideMaster1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>\r\n';
	  strXml += '</Relationships>';
	  //
	  return strXml;
	}

/***/ },
/* 13 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	      value: true
	});
	exports.default = makeXmlSlideMaster;
	function makeXmlSlideMaster(gObjPptx) {
	      var intSlideLayoutId = 2147483649;
	      var SLDNUMFLDID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}';
	      var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">\r\n' + '  <p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr>\r\n' + '<p:cNvPr id="2" name="Title Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>\r\n' + '<p:cNvPr id="3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>\r\n' + '<p:cNvPr id="4" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="6356350"/><a:ext cx="2133600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{F8166F1F-CE9B-4651-A6AA-CD717754106B}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>12/25/2015</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>\r\n' + '<p:cNvPr id="5" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3124200" y="6356350"/><a:ext cx="2895600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>\r\n' + '<p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="6553200" y="6356350"/><a:ext cx="2133600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="' + SLDNUMFLDID + '" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t></a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>\r\n' + '<p:sldLayoutIdLst>\r\n';
	      // Create a sldLayout for each SLIDE
	      for (var idx = 1; idx <= gObjPptx.slides.length; idx++) {
	            strXml += ' <p:sldLayoutId id="' + intSlideLayoutId + '" r:id="rId' + idx + '"/>\r\n';
	            intSlideLayoutId++;
	      }
	      strXml += '</p:sldLayoutIdLst>\r\n' + '<p:txStyles>\r\n' + ' <p:titleStyle>\r\n' + '  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>\r\n' + ' </p:titleStyle>' + ' <p:bodyStyle>\r\n' + '  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' + '  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' + '  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' + '  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' + '  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' + '  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' + '  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' + '  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' + '  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' + ' </p:bodyStyle>\r\n' + ' <p:otherStyle>\r\n' + '  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' + '  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' + '  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' + '  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' + '  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' + '  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' + '  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' + '  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' + '  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' + '  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' + ' </p:otherStyle>\r\n' + '</p:txStyles>\r\n' + '</p:sldMaster>';
	      //
	      return strXml;
	}

/***/ },
/* 14 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	exports.default = makeXmlSlideMasterRel;
	function makeXmlSlideMasterRel(gObjPptx) {
	    // TODO 1.1: create a slideLayout for each SLDIE
	    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n';
	    for (var idx = 1; idx <= gObjPptx.slides.length; idx++) {
	        strXml += '  <Relationship Id="rId' + idx + '" Target="../slideLayouts/slideLayout' + idx + '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"/>\r\n';
	    }
	    strXml += '  <Relationship Id="rId' + (gObjPptx.slides.length + 1) + '" Target="../theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"/>\r\n';
	    strXml += '</Relationships>';
	    //
	    return strXml;
	}

/***/ },
/* 15 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
					value: true
	});
	exports.default = makeXmlTheme;
	function makeXmlTheme() {
					var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n\
							<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">\
							<a:themeElements>\
							  <a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>\
							  <a:dk2><a:srgbClr val="1F497D"/></a:dk2>\
							  <a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3>\
							  <a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5>\
							  <a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/>\
							  </a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/>\
							</a:theme>';
					return strXml;
	}

/***/ },
/* 16 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	  value: true
	});
	exports.default = makeXmlPresentation;
	function makeXmlPresentation(gObjPptx) {
	  var intCurPos = 0;
	  var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">\r\n';
	
	  // STEP 1: Build SLIDE master list
	  strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>\r\n';
	  strXml += '<p:sldIdLst>\r\n';
	  for (var idx = 0; idx < gObjPptx.slides.length; idx++) {
	    strXml += '<p:sldId id="' + (idx + 256) + '" r:id="rId' + (idx + 2) + '"/>\r\n';
	  }
	  strXml += '</p:sldIdLst>\r\n';
	
	  // STEP 2: Build SLIDE text styles
	  strXml += '<p:sldSz cx="' + gObjPptx.pptLayout.width + '" cy="' + gObjPptx.pptLayout.height + '" type="' + gObjPptx.pptLayout.name + '"/>\r\n' + '<p:notesSz cx="' + gObjPptx.pptLayout.height + '" cy="' + gObjPptx.pptLayout.width + '"/>' + '<p:defaultTextStyle>';
	  +'  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>';
	  for (var idx = 1; idx < 10; idx++) {
	    strXml += '  <a:lvl' + idx + 'pPr marL="' + intCurPos + '" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">' + '    <a:defRPr sz="1800" kern="1200">' + '      <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' + '      <a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>' + '    </a:defRPr>' + '  </a:lvl' + idx + 'pPr>';
	    intCurPos += 457200;
	  }
	  strXml += '</p:defaultTextStyle>\r\n';
	
	  strXml += '<p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext></p:extLst>\r\n' + '</p:presentation>';
	  //
	  return strXml;
	}

/***/ },
/* 17 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	exports.default = makeXmlPresProps;
	function makeXmlPresProps() {
	    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">\r\n' + '  <p:extLst>\r\n' + '    <p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="0"/></p:ext>\r\n' + '    <p:ext uri="{D31A062A-798A-4329-ABDD-BBA856620510}"><p14:defaultImageDpi xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="220"/></p:ext>\r\n' + '    <p:ext uri="{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}"><p15:chartTrackingRefBased xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" val="1"/></p:ext>\r\n' + '  </p:extLst>\r\n' + '</p:presentationPr>';
	    return strXml;
	}

/***/ },
/* 18 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	exports.default = makeXmlTableStyles;
	function makeXmlTableStyles() {
	    // SEE: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
	    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>';
	    return strXml;
	}

/***/ },
/* 19 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	exports.default = makeXmlViewProps;
	function makeXmlViewProps() {
	    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' + '<p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>' + '<p:slideViewPr>' + '  <p:cSldViewPr>' + '    <p:cViewPr varScale="1"><p:scale><a:sx n="64" d="100"/><a:sy n="64" d="100"/></p:scale><p:origin x="-1392" y="-96"/></p:cViewPr>' + '    <p:guideLst><p:guide orient="horz" pos="2160"/><p:guide pos="2880"/></p:guideLst>' + '  </p:cSldViewPr>' + '</p:slideViewPr>' + '<p:notesTextViewPr>' + '  <p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr>' + '</p:notesTextViewPr>' + '<p:gridSpacing cx="78028800" cy="78028800"/>' + '</p:viewPr>';
	    return strXml;
	}

/***/ },
/* 20 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	exports.default = makeXmlSlide;
	
	var _exportTable = __webpack_require__(21);
	
	var _exportTable2 = _interopRequireDefault(_exportTable);
	
	var _exportImage = __webpack_require__(22);
	
	var _exportImage2 = _interopRequireDefault(_exportImage);
	
	var _optionAdapter = __webpack_require__(23);
	
	var _optionAdapter2 = _interopRequireDefault(_optionAdapter);
	
	var _slideGroup = __webpack_require__(24);
	
	var _slideGroup2 = _interopRequireDefault(_slideGroup);
	
	var _slide = __webpack_require__(4);
	
	var _slide2 = _interopRequireDefault(_slide);
	
	var _shape = __webpack_require__(26);
	
	var _shape2 = _interopRequireDefault(_shape);
	
	function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }
	
	function makeXmlSlide(inSlide, gObjPptx) {
	
	    var EMU = 914400,
	        ONEPT = 12700;
	    var intTableNum = 1,
	        objSlideData = inSlide.data,
	        strSlideXml = void 0;
	
	    if (inSlide.slide.group) {
	        var startGroup,
	            endGroup,
	            group = new _slideGroup2.default(gObjPptx).generateGroup();
	
	        startGroup = group.groupStart;
	        endGroup = group.groupEnd;
	    }
	    // STEP 1: Start slide XML
	    strSlideXml = _slide2.default.header(inSlide);
	    inSlide.slide.group ? strSlideXml += startGroup : strSlideXml;
	
	    // STEP 5: Loop over all Slide objects and add them to this slide:
	    $.each(objSlideData, function (idx, slideObj) {
	
	        // A: Set option vars
	        if (slideObj.options) {
	            var _OptionAdapter = (0, _optionAdapter2.default)(slideObj, gObjPptx),
	                x = _OptionAdapter.x,
	                y = _OptionAdapter.y,
	                cx = _OptionAdapter.cx,
	                cy = _OptionAdapter.cy,
	                shapeType = _OptionAdapter.shapeType,
	                locationAttr = _OptionAdapter.locationAttr;
	        } else {
	            var x = 0,
	                y = 0,
	                cx = EMU * 10,
	                cy = 0,
	                locationAttr = '',
	                shapeType = null;
	        }
	
	        // B: Create this particular object on Slide
	        switch (slideObj.type) {
	            case 'table':
	                strSlideXml += (0, _exportTable2.default)(inSlide, slideObj, intTableNum, x, y, cx, cy);
	                break;
	
	            case 'text':
	                strSlideXml += new _shape2.default(slideObj).generateShape(idx, inSlide);
	                break;
	
	            case 'image':
	                strSlideXml += (0, _exportImage2.default)(idx, slideObj, locationAttr, x, y, cx, cy);
	                break;
	        }
	    });
	    inSlide.slide.group ? strSlideXml += endGroup : strSlideXml;
	    // STEP 6: Close spTree and finalize slide XML
	    strSlideXml += _slide2.default.footer();
	
	    // LAST: Return
	    return strSlideXml;
	}

/***/ },
/* 21 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };
	
	exports.default = ExportTable;
	
	var _helpers = __webpack_require__(3);
	
	function ExportTable(inSlide, slideObj, intTableNum, x, y, cx, cy) {
	
	    var ONEPT = 12700,
	        EMU = 914400;
	    var arrRowspanCells = [],
	        arrTabRows = slideObj.arrTabRows,
	        objTabOpts = slideObj.objTabOpts,
	        intColCnt = 0,
	        intColW = 0;
	
	    // NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
	    // ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
	    for (var tmp = 0; tmp < arrTabRows[0].length; tmp++) {
	        intColCnt += arrTabRows[0][tmp].opts && arrTabRows[0][tmp].opts.colspan ? Number(arrTabRows[0][tmp].opts.colspan) : 1;
	    }
	
	    // STEP 1: Start Table XML
	    // NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
	    var strXml = '<p:graphicFrame>' + '  <p:nvGraphicFramePr>' + '    <p:cNvPr id="' + (intTableNum * inSlide.numb + 1) + '" name="Table ' + intTableNum * inSlide.numb + '"/>' + '    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' + '    <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>' + '  </p:nvGraphicFramePr>' + '  <p:xfrm>' + '    <a:off  x="' + (x || EMU) + '"  y="' + (y || EMU) + '"/>' + '    <a:ext cx="' + (cx || EMU) + '" cy="' + (cy || EMU) + '"/>' + '  </p:xfrm>' + '  <a:graphic>' + '    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' + '      <a:tbl>' + '        <a:tblPr/>';
	    // + '        <a:tblPr bandRow="1"/>';
	    // TODO 1.5: Support banded rows, first/last row, etc.
	    // NOTE: Banding, etc. only shows when using a table style! (or set alt row color if banding)
	    // <a:tblPr firstCol="0" firstRow="0" lastCol="0" lastRow="0" bandCol="0" bandRow="1">
	
	    // STEP 2: Set column widths
	    // Evenly distribute cols/rows across size provided when applicable (calc them if only overall dimensions were provided)
	    // A: Col widths provided?
	    if (Array.isArray(objTabOpts.colW)) {
	        strXml += '<a:tblGrid>';
	        for (var col = 0; col < intColCnt; col++) {
	            strXml += '  <a:gridCol w="' + (objTabOpts.colW[col] || slideObj.options.cx / intColCnt) + '"/>';
	        }
	        strXml += '</a:tblGrid>';
	    }
	    // B: Table Width provided without colW? Then distribute cols
	    else {
	            intColW = objTabOpts.colW ? objTabOpts.colW : EMU;
	            if (slideObj.options.cx && !objTabOpts.colW) intColW = slideObj.options.cx / intColCnt;
	            strXml += '<a:tblGrid>';
	            for (var col = 0; col < intColCnt; col++) {
	                strXml += '<a:gridCol w="' + intColW + '"/>';
	            }
	            strXml += '</a:tblGrid>';
	        }
	    // C: Table Height provided without rowH? Then distribute rows
	    var intRowH = objTabOpts.rowH ? (0, _helpers.inch2Emu)(objTabOpts.rowH) : 0;
	    if (slideObj.options.cy && !objTabOpts.rowH) intRowH = slideObj.options.cy / arrTabRows.length;
	
	    // STEP 3: Build an array of rowspan cells now so we can add stubs in as we loop below
	    $.each(arrTabRows, function (rIdx, row) {
	        $(row).each(function (cIdx, cell) {
	            var colIdx = cIdx;
	            if (cell.opts && cell.opts.rowspan && Number.isInteger(cell.opts.rowspan)) {
	                for (var idy = 1; idy < cell.opts.rowspan; idy++) {
	                    arrRowspanCells.push({ row: rIdx + idy, col: colIdx });
	                    colIdx++; // For cases where we already have a rowspan in this row - we need to Increment to account for this extra cell!
	                }
	            }
	        });
	    });
	
	    // STEP 4: Build table rows/cells
	    $.each(arrTabRows, function (rIdx, row) {
	        if (Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx]) intRowH = (0, _helpers.inch2Emu)(objTabOpts.rowH[rIdx]);
	
	        // A: Start row
	        strXml += '<a:tr h="' + intRowH + '">';
	
	        // B: Loop over each CELL
	        $(row).each(function (cIdx, cell) {
	            // 1: OPTIONS: Build/set cell options (blocked for code folding)
	            {
	                // 1: Load/Create options
	                var cellOpts = cell.opts || {};
	
	                // 2: Do Important/Override Opts
	                // Feature: TabOpts Default Values (tabOpts being used when cellOpts dont exist):
	                // SEE: http://officeopenxml.com/drwTableCellProperties-alignment.php
	                $.each(['align', 'bold', 'border', 'color', 'fill', 'font_face', 'font_size', 'underline', 'valign'], function (i, name) {
	                    if (objTabOpts[name] && !cellOpts[name]) cellOpts[name] = objTabOpts[name];
	                });
	
	                var cellB = cellOpts.bold ? ' b="1"' : ''; // [0,1] or [false,true]
	                var cellU = cellOpts.underline ? ' u="sng"' : ''; // [none,sng (single), et al.]
	                var cellFont = cellOpts.font_face ? ' <a:latin typeface="' + cellOpts.font_face + '"/>' : '';
	                var cellFontPt = cellOpts.font_size ? ' sz="' + cellOpts.font_size + '00"' : '';
	                var cellAlign = cellOpts.align ? ' algn="' + cellOpts.align.replace(/^c$/i, 'ctr').replace('center', 'ctr').replace('left', 'l').replace('right', 'r') + '"' : '';
	                var cellValign = cellOpts.valign ? ' anchor="' + cellOpts.valign.replace(/^c$/i, 'ctr').replace(/^m$/i, 'ctr').replace('center', 'ctr').replace('middle', 'ctr').replace('top', 't').replace('btm', 'b').replace('bottom', 'b') + '"' : '';
	                var cellColspan = cellOpts.colspan ? ' gridSpan="' + cellOpts.colspan + '"' : '';
	                var cellRowspan = cellOpts.rowspan ? ' rowSpan="' + cellOpts.rowspan + '"' : '';
	                var cellFontClr = cell.optImp && cell.optImp.color || cellOpts.color ? ' <a:solidFill><a:srgbClr val="' + (cell.optImp && cell.optImp.color || cellOpts.color) + '"/></a:solidFill>' : '';
	                var cellFill = cell.optImp && cell.optImp.fill || cellOpts.fill ? ' <a:solidFill><a:srgbClr val="' + (cell.optImp && cell.optImp.fill || cellOpts.fill) + '"/></a:solidFill>' : '';
	                var intMarginPt = cellOpts.marginPt || cellOpts.marginPt == 0 ? cellOpts.marginPt * ONEPT : 0;
	                // Margin/Padding:
	                var cellMargin = '';
	                if (cellOpts.marginPt && Array.isArray(cellOpts.marginPt)) {
	                    cellMargin = ' marL="' + cellOpts.marginPt[3] + '" marR="' + cellOpts.marginPt[1] + '" marT="' + cellOpts.marginPt[0] + '" marB="' + cellOpts.marginPt[2] + '"';
	                } else if (cellOpts.marginPt && Number.isInteger(cellOpts.marginPt)) {
	                    cellMargin = ' marL="' + intMarginPt + '" marR="' + intMarginPt + '" marT="' + intMarginPt + '" marB="' + intMarginPt + '"';
	                }
	            }
	
	            // 2: Cell Content: Either the text element or the cell itself (for when users just pass a string - no object or options)
	            var strCellText = (typeof cell === 'undefined' ? 'undefined' : _typeof(cell)) === 'object' ? cell.text : cell;
	
	            // TODO 1.5: Cell NOWRAP property (text wrap: add to a:tcPr (horzOverflow="overflow" or whatev opts exist)
	
	            // 3: ROWSPAN: Add dummy cells for any active rowspan
	            // TODO 1.5: ROWSPAN & COLSPAN in same cell is not yet handled!
	            if (arrRowspanCells.filter(function (obj) {
	                return obj.row == rIdx && obj.col == cIdx;
	            }).length > 0) {
	                strXml += '<a:tc vMerge="1"><a:tcPr/></a:tc>';
	            }
	
	            // 4: Start Table Cell, add Align, add Text content
	            strXml += ' <a:tc' + cellColspan + cellRowspan + '>' + '  <a:txBody>' + '    <a:bodyPr/>' + '    <a:lstStyle/>' + '    <a:p>' + '      <a:pPr' + cellAlign + '/>' + '      <a:r>' + '        <a:rPr lang="en-US" dirty="0" smtClean="0"' + cellFontPt + cellB + cellU + '>' + cellFontClr + cellFont + '</a:rPr>' + '        <a:t>' + (0, _helpers.decodeXmlEntities)(strCellText) + '</a:t>' + '      </a:r>' + '      <a:endParaRPr lang="en-US" dirty="0"/>' + '    </a:p>' + '  </a:txBody>' + '  <a:tcPr' + cellMargin + cellValign + '>';
	
	            // 5: Borders: Add any borders
	            if (cellOpts.border && typeof cellOpts.border === 'string') {
	                strXml += '  <a:lnL w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnL>';
	                strXml += '  <a:lnR w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnR>';
	                strXml += '  <a:lnT w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnT>';
	                strXml += '  <a:lnB w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnB>';
	            } else if (cellOpts.border && Array.isArray(cellOpts.border)) {
	                $.each([{ idx: 3, name: 'lnL' }, { idx: 1, name: 'lnR' }, { idx: 0, name: 'lnT' }, { idx: 2, name: 'lnB' }], function (i, obj) {
	                    if (cellOpts.border[obj.idx]) {
	                        var strC = '<a:solidFill><a:srgbClr val="' + (cellOpts.border[obj.idx].color ? cellOpts.border[obj.idx].color : '666666') + '"/></a:solidFill>';
	                        var intW = cellOpts.border[obj.idx] && (cellOpts.border[obj.idx].pt || cellOpts.border[obj.idx].pt == 0) ? ONEPT * Number(cellOpts.border[obj.idx].pt) : ONEPT;
	                        strXml += '<a:' + obj.name + ' w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strC + '</a:' + obj.name + '>';
	                    } else strXml += '<a:' + obj.name + ' w="0"><a:miter lim="400000" /></a:' + obj.name + '>';
	                });
	            } else if (cellOpts.border && _typeof(cellOpts.border) === 'object') {
	                var intW = cellOpts.border && (cellOpts.border.pt || cellOpts.border.pt == 0) ? ONEPT * Number(cellOpts.border.pt) : ONEPT;
	                var strClr = '<a:solidFill><a:srgbClr val="' + (cellOpts.border.color ? cellOpts.border.color : '666666') + '"/></a:solidFill>';
	                var strAttr = '<a:prstDash val="';
	                strAttr += cellOpts.border.type && cellOpts.border.type.toLowerCase().indexOf('dash') > -1 ? "sysDash" : "solid";
	                strAttr += '"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>';
	                // *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
	                strXml += '<a:lnL w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnL>';
	                strXml += '<a:lnR w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnR>';
	                strXml += '<a:lnT w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnT>';
	                strXml += '<a:lnB w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnB>';
	                // *** IMPORTANT! *** LRTB order matters!
	            }
	
	            // 6: Close cell Properties & Cell
	            strXml += cellFill + '  </a:tcPr>' + ' </a:tc>';
	
	            // LAST: COLSPAN: Add a 'merged' col for each column being merged (SEE: http://officeopenxml.com/drwTableGrid.php)
	            if (cellOpts.colspan) {
	                for (var tmp = 1; tmp < Number(cellOpts.colspan); tmp++) {
	                    strXml += '<a:tc hMerge="1"><a:tcPr/></a:tc>';
	                }
	            }
	        });
	
	        // B-2: Handle Rowspan as last col case
	        // We add dummy cells inside cell loop, but cases where last col is rowspaned
	        // by prev row wont be created b/c cell loop above exhausted before the col
	        // index of the final col was reached... ANYHOO, add it here when necc.
	        if (arrRowspanCells.filter(function (obj) {
	            return obj.row == rIdx && obj.col + 1 >= $(row).length;
	        }).length > 0) {
	            strXml += '<a:tc vMerge="1"><a:tcPr/></a:tc>';
	        }
	
	        // C: Complete row
	        strXml += '</a:tr>';
	    });
	
	    // STEP 5: Complete table
	    strXml += '      </a:tbl>' + '    </a:graphicData>' + '  </a:graphic>' + '</p:graphicFrame>';
	
	    // STEP 6: Set table XML
	    var strSlideXml = strXml;
	
	    // LAST: Increment counter
	    intTableNum++;
	
	    return strSlideXml;
	}

/***/ },
/* 22 */
/***/ function(module, exports) {

	"use strict";
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	exports.default = function (idx, slideObj, locationAttr, x, y, cx, cy) {
	
	    var strSlideXml = "<p:pic>\n                        <p:nvPicPr>\n                          <p:cNvPr id=\"" + (idx + 2) + "\" name=\"Object " + (idx + 1) + "\" descr=\"" + slideObj.image + "\"/>\n                                <p:cNvPicPr>\n                                    <a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/>\n                                </p:nvPicPr>\n                                <p:blipFill>\n                                    <a:blip r:embed=\"rId" + slideObj.imageRid + "\" cstate=\"print\"/><a:stretch><a:fillRect/></a:stretch>\n                                </p:blipFill>\n                            <p:spPr>\n                                <a:xfrm" + locationAttr + ">\n                                    <a:off  x=\"" + x + "\"  y=\"" + y + "\"/>\n                                    <a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/>\n                                </a:xfrm>\n                                <a:prstGeom prst=\"rect\">\n                                    <a:avLst/>\n                                </a:prstGeom>\n                            </p:spPr>\n                    </p:pic>";
	
	    return strSlideXml;
	};

/***/ },
/* 23 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	exports.default = function (slideObj, gObjPptx) {
	    var EMU = 914400;
	    var x = 0,
	        y = 0,
	        cx = EMU * 10,
	        cy = 0,
	        locationAttr = '',
	        shapeType = null;
	
	    if (slideObj.options.w || slideObj.options.w == 0) slideObj.options.cx = slideObj.options.w;
	    if (slideObj.options.h || slideObj.options.h == 0) slideObj.options.cy = slideObj.options.h;
	    //
	    if (slideObj.options.x || slideObj.options.x == 0) x = (0, _helpers.getSmartParseNumber)(slideObj.options.x, 'X', gObjPptx);
	    if (slideObj.options.y || slideObj.options.y == 0) y = (0, _helpers.getSmartParseNumber)(slideObj.options.y, 'Y', gObjPptx);
	    if (slideObj.options.cx || slideObj.options.cx == 0) cx = (0, _helpers.getSmartParseNumber)(slideObj.options.cx, 'X', gObjPptx);
	    if (slideObj.options.cy || slideObj.options.cy == 0) cy = (0, _helpers.getSmartParseNumber)(slideObj.options.cy, 'Y', gObjPptx);
	    if (slideObj.options.flipH) locationAttr += ' flipH="1"';
	    if (slideObj.options.flipV) locationAttr += ' flipV="1"';
	    if (slideObj.options.shape) shapeType = (0, _helpers.getShapeInfo)(slideObj.options.shape);
	    if (slideObj.options.rotate) {
	        var rotateVal = slideObj.options.rotate > 360 ? slideObj.options.rotate - 360 : slideObj.options.rotate;
	        rotateVal *= 60000;
	        locationAttr += ' rot="' + rotateVal + '"';
	    }
	    return { x: x, y: y, cx: cx, cy: cy, shapeType: shapeType, locationAttr: locationAttr };
	};
	
	var _helpers = __webpack_require__(3);

/***/ },
/* 24 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();
	
	var _baseGroup = __webpack_require__(25);
	
	var _baseGroup2 = _interopRequireDefault(_baseGroup);
	
	var _helpers = __webpack_require__(3);
	
	function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }
	
	function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
	
	function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }
	
	function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }
	
	var SlideGroup = function (_BaseGroup) {
	    _inherits(SlideGroup, _BaseGroup);
	
	    function SlideGroup(gObjPptx) {
	        _classCallCheck(this, SlideGroup);
	
	        var _this = _possibleConstructorReturn(this, (SlideGroup.__proto__ || Object.getPrototypeOf(SlideGroup)).call(this));
	
	        _this.allShapes = gObjPptx.slides;
	        return _this;
	    }
	
	    _createClass(SlideGroup, [{
	        key: 'generateGroup',
	        value: function generateGroup() {
	            var _this2 = this;
	
	            var nLength = this.allShapes.length,
	                nIndex = void 0;
	
	            var _loop = function _loop() {
	                var oShapes = _this2.allShapes[nIndex].data,
	                    groupIndex = 0;
	
	                var _loop2 = function _loop2(j) {
	                    var oShape = oShapes[j].options;
	                    Object.keys(_this2.wrapperGroupCoordinate).map(function (i) {
	                        var nValue = void 0;
	                        if (i === 'name') {
	                            _this2.wrapperGroupCoordinate[i] = _this2.allShapes[nIndex].name;
	                        }
	                        // get the x,y of the first shape
	                        else if (i === 'x' || i === 'y') {
	                                if (groupIndex < 2) {
	                                    _this2.wrapperGroupCoordinate[i] = oShape[i];
	                                    groupIndex++;
	                                } else {
	                                    // if an unordred shapes
	                                    _this2.wrapperGroupCoordinate[i] = Math.min(_this2.wrapperGroupCoordinate[i], oShape[i]);
	                                }
	                            } else if (i === 'cx' && (oShape[i] || oShape['w'])) {
	                                nValue = oShape[i] ? oShape[i] : oShape['w'];
	
	                                if (groupIndex === 2) {
	                                    _this2.wrapperGroupCoordinate[i] = nValue;
	                                    groupIndex++;
	                                }
	
	                                if (oShape['x'] > _this2.wrapperGroupCoordinate[i] + _this2.wrapperGroupCoordinate['x']) {
	                                    _this2.wrapperGroupCoordinate[i] = oShape['x'] + nValue;
	                                }
	                            } else if (i === 'cy' && (oShape[i] || oShape['h'])) {
	                                nValue = oShape[i] ? oShape[i] : oShape['h'];
	
	                                if (groupIndex === 3) {
	                                    _this2.wrapperGroupCoordinate[i] = nValue;
	                                    groupIndex++;
	                                }
	                                if (oShape['y'] > _this2.wrapperGroupCoordinate[i] + _this2.wrapperGroupCoordinate['y']) {
	                                    _this2.wrapperGroupCoordinate[i] = oShape['y'] + nValue;
	                                }
	                            }
	                    });
	                };
	
	                for (var j = 0; j < oShapes.length; j++) {
	                    _loop2(j);
	                }
	            };
	
	            for (nIndex = 0; nIndex < nLength; nIndex++) {
	                _loop();
	            }
	            var sStart = ['<p:grpSp>', ' <p:nvGrpSpPr>', ' <p:cNvPr id="' + this.wrapperGroupCoordinate.id + '" name="' + this.wrapperGroupCoordinate.name + '"/>', '<p:cNvGrpSpPr/>', '<p:nvPr/>', '</p:nvGrpSpPr>', '<p:grpSpPr>', '<a:xfrm>', '<a:off x="' + (0, _helpers.inch2Emu)(this.wrapperGroupCoordinate.x) + '"  y="' + (0, _helpers.inch2Emu)(this.wrapperGroupCoordinate.y) + '"/>', '<a:ext cx="' + (0, _helpers.inch2Emu)(this.wrapperGroupCoordinate.cx) + '" cy="' + (0, _helpers.inch2Emu)(this.wrapperGroupCoordinate.cy) + '"/>', '<a:chOff x="' + (0, _helpers.inch2Emu)(this.wrapperGroupCoordinate.x) + '"  y="' + (0, _helpers.inch2Emu)(this.wrapperGroupCoordinate.y) + '" />', '<a:chExt cx="' + (0, _helpers.inch2Emu)(this.wrapperGroupCoordinate.cx) + '" cy="' + (0, _helpers.inch2Emu)(this.wrapperGroupCoordinate.cy) + '"/>', '</a:xfrm>', '</p:grpSpPr>'];
	            this.groupStart = sStart.join('');
	            this.groupEnd = '</p:grpSp>';
	            this.id++;
	            return this;
	        }
	    }, {
	        key: 'groupStart',
	        set: function set(sGroupStart) {
	            this._groupStart = sGroupStart;
	        },
	        get: function get() {
	            return this._groupStart;
	        }
	    }, {
	        key: 'groupEnd',
	        set: function set(sGroupEnd) {
	            this._groupEnd = sGroupEnd;
	        },
	        get: function get() {
	            return this._groupEnd;
	        }
	    }]);
	
	    return SlideGroup;
	}(_baseGroup2.default);
	
	exports.default = SlideGroup;

/***/ },
/* 25 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
	
	var BaseGroup = function BaseGroup(name) {
	    _classCallCheck(this, BaseGroup);
	
	    this.id = 2;
	    this.wrapperGroupCoordinate = { id: this.id, name: name, x: 0, y: 0, cx: 0, cy: 0 };
	    this.groupStart = '';
	    this.groupEnd = '';
	};
	
	exports.default = BaseGroup;

/***/ },
/* 26 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };
	
	var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();
	
	var _helpers = __webpack_require__(3);
	
	var _optionAdapter = __webpack_require__(23);
	
	var _optionAdapter2 = _interopRequireDefault(_optionAdapter);
	
	var _slide = __webpack_require__(4);
	
	var _slide2 = _interopRequireDefault(_slide);
	
	function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }
	
	function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
	
	var ONEPT = 12700,
	    EMU = 914400;
	
	var Shape = function () {
	    function Shape(shapeObject) {
	        _classCallCheck(this, Shape);
	
	        this.shapeObject = shapeObject;
	        this.coordinate = (0, _optionAdapter2.default)(this.shapeObject, _slide2.default.gObjPptx);
	    }
	
	    _createClass(Shape, [{
	        key: 'fixMargin',
	        value: function fixMargin() {
	            // Lines can have zero cy, but text should not
	            var cy = this.coordinate.cy;
	
	            if (!this.shapeObject.options.line && cy == 0) cy = EMU * 0.3;
	
	            // Margin/Padding/Inset for textboxes
	            if (this.shapeObject.options.margin && Array.isArray(this.shapeObject.options.margin)) {
	
	                this.shapeObject.options.bodyProp.lIns = this.shapeObject.options.margin[0] * ONEPT || 0;
	                this.shapeObject.options.bodyProp.rIns = this.shapeObject.options.margin[1] * ONEPT || 0;
	                this.shapeObject.options.bodyProp.bIns = this.shapeObject.options.margin[2] * ONEPT || 0;
	                this.shapeObject.options.bodyProp.tIns = this.shapeObject.options.margin[3] * ONEPT || 0;
	            } else if ((this.shapeObject.options.margin || this.shapeObject.options.margin == 0) && Number.isInteger(this.shapeObject.options.margin)) {
	                this.shapeObject.options.bodyProp.lIns = this.shapeObject.options.margin * ONEPT;
	                this.shapeObject.options.bodyProp.rIns = this.shapeObject.options.margin * ONEPT;
	                this.shapeObject.options.bodyProp.bIns = this.shapeObject.options.margin * ONEPT;
	                this.shapeObject.options.bodyProp.tIns = this.shapeObject.options.margin * ONEPT;
	            }
	            return this; //.shapeObject;
	        }
	    }, {
	        key: 'setShapeProperty',
	        value: function setShapeProperty(idx) {
	            var isTextBox = this.shapeObject.options && this.shapeObject.options.isTextBox,
	                _coordinate = this.coordinate,
	                x = _coordinate.x,
	                y = _coordinate.y,
	                cx = _coordinate.cx,
	                cy = _coordinate.cy,
	                locationAttr = _coordinate.locationAttr;
	
	            //B: The addition of the "txBox" attribute is the sole determiner of if an object is a Shape or Textbox
	            var aStr = void 0;
	            aStr = ['<p:sp>', '<p:nvSpPr><p:cNvPr id="' + (idx + 2) + '" name="Object ' + (idx + 1) + '"/>', '<p:cNvSpPr ' + (isTextBox ? ' txBox="1"/><p:nvPr/>' : '/><p:nvPr/>'), '</p:nvSpPr>', '<p:spPr><a:xfrm' + locationAttr + '>', '<a:off x="' + x + '" y="' + y + '"/>', '<a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>'];
	
	            this._xmlShape = aStr.join('');
	            return this;
	        }
	    }, {
	        key: 'setPrstGeom',
	        value: function setPrstGeom() {
	            var shapeType = this.coordinate.shapeType;
	
	
	            if (shapeType == null) shapeType = (0, _helpers.getShapeInfo)(null);
	            if (this.shapeObject.options && this.shapeObject.options.customPreset) {
	                this._xmlShape += '<a:prstGeom prst="' + shapeType.name + '">' + this.shapeObject.options.customPreset + '</a:prstGeom>';
	            } else {
	                this._xmlShape += '<a:prstGeom prst="' + shapeType.name + '"><a:avLst/></a:prstGeom>';
	            }
	
	            return this;
	        }
	    }, {
	        key: 'isFillOption',
	        value: function isFillOption() {
	            if (this.shapeObject.options) {
	
	                this.shapeObject.options.fill ? this._xmlShape += (0, _helpers.genXmlColorSelection)(this.shapeObject.options.fill) : this._xmlShape += '<a:noFill/>';
	
	                if (this.shapeObject.options.line) {
	                    var lineAttr = '';
	
	                    if (this.shapeObject.options.line_size) lineAttr += ' w="' + this.shapeObject.options.line_size * ONEPT + '"';
	                    this._xmlShape += '<a:ln' + lineAttr + '>';
	                    this._xmlShape += (0, _helpers.genXmlColorSelection)(this.shapeObject.options.line);
	                    if (this.shapeObject.options.line_head) this._xmlShape += '<a:headEnd type="' + this.shapeObject.options.line_head + '"/>';
	                    if (this.shapeObject.options.line_tail) this._xmlShape += '<a:tailEnd type="' + this.shapeObject.options.line_tail + '"/>';
	                    this._xmlShape += '</a:ln>';
	                }
	            } else {
	                this._xmlShape += '<a:noFill/>';
	            }
	            return this;
	        }
	    }, {
	        key: 'isEffect',
	        value: function isEffect() {
	            if (this.shapeObject.options.effects) {
	
	                for (var ii = 0, total_size_ii = this.shapeObject.options.effects.length; ii < total_size_ii; ii++) {
	                    switch (this.shapeObject.options.effects[ii].type) {
	
	                        case 'outerShadow':
	                            effectsList += cbGenerateEffects(this.shapeObject.options.effects[ii], 'outerShdw');
	                            break;
	                        case 'innerShadow':
	                            effectsList += cbGenerateEffects(this.shapeObject.options.effects[ii], 'innerShdw');
	                            break;
	                    }
	                }
	            }
	            return this;
	        }
	    }, {
	        key: 'closeShapeProperty',
	        value: function closeShapeProperty() {
	            this._xmlShape += '</p:spPr>';
	            return this;
	        }
	    }, {
	        key: 'styleProperty',
	        value: function styleProperty() {
	
	            var moreStyles = '',
	                moreStylesAttr = '',
	                outStyles = '',
	                styleData = '';
	
	            if (this.shapeObject.options) {
	
	                if (this.shapeObject.options.align) {
	                    switch (this.shapeObject.options.align) {
	                        case 'right':
	                            moreStylesAttr += ' algn="r"';
	                            break;
	                        case 'center':
	                            moreStylesAttr += ' algn="ctr"';
	                            break;
	                        case 'justify':
	                            moreStylesAttr += ' algn="just"';
	                            break;
	                    }
	                }
	
	                if (this.shapeObject.options.indentLevel > 0) {
	                    moreStylesAttr += ' lvl="' + this.shapeObject.options.indentLevel + '"';
	                }
	            }
	
	            if (moreStyles != '') this._xmlShape += '<a:pPr' + moreStylesAttr + '>' + moreStyles + '</a:pPr>';else if (moreStylesAttr != '') this._xmlShape += '<a:pPr' + moreStylesAttr + '/>';
	
	            if (styleData != '') this._xmlShape += '<p:style>' + styleData + '</p:style>';
	
	            return outStyles;
	        }
	    }, {
	        key: 'txBody',
	        value: function txBody(inSlide) {
	
	            if (typeof this.shapeObject.text == 'string') {
	
	                this._xmlShape += '<p:txBody>' + (0, _helpers.genXmlBodyProperties)(this.shapeObject.options) + '<a:lstStyle/><a:p>' + this.styleProperty();
	                this._xmlShape += (0, _helpers.genXmlTextCommand)(this.shapeObject.options, this.shapeObject.text, inSlide.slide, inSlide.slide.getPageNumber());
	            } else if (typeof this.shapeObject.text == 'number') {
	
	                this._xmlShape += '<p:txBody>' + (0, _helpers.genXmlBodyProperties)(this.shapeObject.options) + '<a:lstStyle/><a:p>' + this.styleProperty();
	                this._xmlShape += (0, _helpers.genXmlTextCommand)(this.shapeObject.options, this.shapeObject.text + '', inSlide.slide, inSlide.slide.getPageNumber());
	            } else if (this.shapeObject.text && this.shapeObject.text.length) {
	                var outBodyOpt = (0, _helpers.genXmlBodyProperties)(this.shapeObject.options);
	                this._xmlShape += '<p:txBody>' + outBodyOpt + '<a:lstStyle/><a:p>' + this.styleProperty();
	
	                for (var j = 0, total_size_j = this.shapeObject.text.length; j < total_size_j; j++) {
	                    if (_typeof(this.shapeObject.text[j]) == 'object' && this.shapeObject.text[j].text) {
	                        this._xmlShape += (0, _helpers.genXmlTextCommand)(this.shapeObject.text[j].options || this.shapeObject.options, this.shapeObject.text[j].text, inSlide.slide, outBodyOpt, this.styleProperty(), inSlide.slide.getPageNumber());
	                    } else if (typeof this.shapeObject.text[j] == 'string') {
	                        this._xmlShape += (0, _helpers.genXmlTextCommand)(this.shapeObject.options, this.shapeObject.text[j], inSlide.slide, outBodyOpt, this.styleProperty(), inSlide.slide.getPageNumber());
	                    } else if (typeof this.shapeObject.text[j] == 'number') {
	                        this._xmlShape += (0, _helpers.genXmlTextCommand)(this.shapeObject.options, this.shapeObject.text[j] + '', inSlide.slide, outBodyOpt, this.styleProperty(), inSlide.slide.getPageNumber());
	                    } else if (_typeof(this.shapeObject.text[j]) == 'object' && this.shapeObject.text[j].field) {
	                        this._xmlShape += (0, _helpers.genXmlTextCommand)(this.shapeObject.options, this.shapeObject.text[j], inSlide.slide, outBodyOpt, this.styleProperty(), inSlide.slide.getPageNumber());
	                    }
	                }
	            } else if (_typeof(this.shapeObject.text) == 'object' && this.shapeObject.text.field) {
	                this._xmlShape += '<p:txBody>' + (0, _helpers.genXmlBodyProperties)(this.shapeObject.options) + '<a:lstStyle/><a:p>' + this.styleProperty();
	                this._xmlShape += (0, _helpers.genXmlTextCommand)(this.shapeObject.options, this.shapeObject.text, inSlide.slide, inSlide.slide.getPageNumber());
	            }
	
	            // We must add that at the end of every paragraph with text:
	            if (typeof this.shapeObject.text !== 'undefined') {
	                var font_size = '';
	                if (this.shapeObject.options && this.shapeObject.options.font_size) font_size = ' sz="' + this.shapeObject.options.font_size + '00"';
	                this._xmlShape += '<a:endParaRPr lang="en-US"' + font_size + ' dirty="0"/></a:p></p:txBody>';
	            }
	            return this;
	        }
	    }, {
	        key: 'closeShape',
	        value: function closeShape() {
	            var sEndShape = void 0;
	            this.shapeObject.type == 'cxn' ? sEndShape = '</p:cxnSp>' : sEndShape = '</p:sp>';
	            this._xmlShape += sEndShape;
	            return this;
	        }
	    }, {
	        key: 'generateShape',
	        value: function generateShape(idx, inSlide) {
	            this.fixMargin().setShapeProperty(idx).setPrstGeom().isFillOption().isEffect().closeShapeProperty().txBody(inSlide).closeShape();
	            return this._xmlShape;
	        }
	    }]);
	
	    return Shape;
	}();
	
	exports.default = Shape;

/***/ },
/* 27 */
/***/ function(module, exports) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	exports.default = makeXmlSlideRel;
	function makeXmlSlideRel(inSlideNum, gObjPptx) {
	    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n' + '  <Relationship Id="rId1" Target="../slideLayouts/slideLayout' + inSlideNum + '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"/>\r\n';
	
	    // Add any IMAGEs for this Slide
	    for (var idx = 0; idx < gObjPptx.slides[inSlideNum - 1].rels.length; idx++) {
	        strXml += '  <Relationship Id="rId' + gObjPptx.slides[inSlideNum - 1].rels[idx].rId + '" Target="' + gObjPptx.slides[inSlideNum - 1].rels[idx].Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>\r\n';
	    }
	
	    strXml += '</Relationships>';
	    //
	    return strXml;
	}

/***/ },
/* 28 */
/***/ function(module, exports, __webpack_require__) {

	'use strict';
	
	Object.defineProperty(exports, "__esModule", {
	    value: true
	});
	
	var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();
	
	var _helpers = __webpack_require__(3);
	
	var _slide = __webpack_require__(4);
	
	var _slide2 = _interopRequireDefault(_slide);
	
	function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }
	
	function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
	
	var ONEPT = 12700,
	    EMU = 914400;
	
	var SlideTable = function () {
	    function SlideTable() {
	        _classCallCheck(this, SlideTable);
	    }
	
	    _createClass(SlideTable, [{
	        key: 'addSlidesForTable',
	        value: function addSlidesForTable(tabEleId, inOpts) {
	
	            var opts = inOpts || {},
	                arrObjTabHeadRows = [],
	                arrObjTabBodyRows = [],
	                arrObjTabFootRows = [],
	                arrObjSlides = [],
	                arrRows = [],
	                arrColW = [],
	                arrTabColW = [],
	                intTabW = 0,
	                emuTabCurrH = 0;
	
	            // NOTE: Look for opts.margin first as user can override Slide Master settings if they want
	            var arrInchMargins = this.lookMargin(opts);
	
	            var emuSlideTabW = _slide2.default.gObjPptx.pptLayout.width - (0, _helpers.inch2Emu)(arrInchMargins[1] + arrInchMargins[3]);
	
	            var emuSlideTabH = _slide2.default.gObjPptx.pptLayout.height - (0, _helpers.inch2Emu)(arrInchMargins[0] + arrInchMargins[2]);
	
	            // STEP 1: Grab overall table style/col widths
	            this.tableStyle(tabEleId, arrTabColW);
	            $.each(arrTabColW, function (i, colW) {
	                intTabW += colW;
	            });
	
	            // STEP 2: Calc/Set column widths by using same column width percent from HTML table
	            this.calcColWidth(arrTabColW, tabEleId, arrColW, emuSlideTabW, intTabW);
	
	            // STEP 3: Iterate over each table element and create data arrays (text and opts)
	            // NOTE: We create 3 arrays instead of one so we can loop over body then show header/footer rows on first and last page
	            this.createDataArray(tabEleId, arrObjTabHeadRows, arrObjTabBodyRows, arrObjTabFootRows);
	
	            // STEP 4: Paginate data: Iterate over all table rows, divide into slides/pages based upon the row height>overall height
	            this.paginateData(arrObjTabHeadRows, arrObjTabBodyRows, arrObjTabFootRows, arrColW, emuTabCurrH, emuSlideTabH, arrRows, arrObjSlides, opts); // tab loop
	            // Flush final row buffer to slide
	            arrObjSlides.push($.extend(true, [], arrRows));
	
	            // STEP 5: Create a SLIDE for each of our 1-N table pieces
	            this.createSlides(arrObjSlides, opts, arrInchMargins, emuSlideTabW, arrColW);
	        }
	    }, {
	        key: 'tableStyle',
	        value: function tableStyle(tabEleId, arrTabColW) {
	            $.each(['thead', 'tbody', 'tfoot'], function (i, val) {
	                if ($('#' + tabEleId + ' ' + val + ' tr').length > 0) {
	                    $('#' + tabEleId + ' ' + val + ' tr:first-child').find('th, td').each(function (i, cell) {
	                        // TODO 1.5: This is a hack - guessing at col widths when colspan
	                        if ($(this).attr('colspan')) {
	                            for (var idx = 0; idx < $(this).attr('colspan'); idx++) {
	                                arrTabColW.push(Math.round($(this).outerWidth() / $(this).attr('colspan')));
	                            }
	                        } else {
	                            arrTabColW.push($(this).outerWidth());
	                        }
	                    });
	                    return false; // break out of .each loop
	                }
	            });
	        }
	    }, {
	        key: 'calcColWidth',
	        value: function calcColWidth(arrTabColW, tabEleId, arrColW, emuSlideTabW, intTabW) {
	            $.each(arrTabColW, function (i, colW) {
	                $('#' + tabEleId + ' thead tr:first-child th:nth-child(' + (i + 1) + ')').data('pptx-min-width') ? arrColW.push((0, _helpers.inch2Emu)($('#' + tabEleId + ' thead tr:first-child th:nth-child(' + (i + 1) + ')').data('pptx-min-width'))) : arrColW.push(Math.round(emuSlideTabW * (colW / intTabW * 100) / 100));
	            });
	        }
	    }, {
	        key: 'createDataArray',
	        value: function createDataArray(tabEleId, arrObjTabHeadRows, arrObjTabBodyRows, arrObjTabFootRows) {
	            $.each(['thead', 'tbody', 'tfoot'], function (i, val) {
	                $('#' + tabEleId + ' ' + val + ' tr').each(function (i, row) {
	                    var arrObjTabCells = [];
	                    $(row).find('th, td').each(function (i, cell) {
	                        // A: Covert colors to Hex from RGB
	                        var arrRGB1 = [];
	                        var arrRGB2 = [];
	                        arrRGB1 = $(cell).css('color').replace(/\s+/gi, '').replace('rgb(', '').replace(')', '').split(',');
	                        arrRGB2 = $(cell).css('background-color').replace(/\s+/gi, '').replace('rgb(', '').replace(')', '').split(',');
	
	                        // B: Create option object
	                        var objOpts = {
	                            font_size: $(cell).css('font-size').replace(/\D/gi, ''),
	                            bold: $(cell).css('font-weight') == "bold" || Number($(cell).css('font-weight')) >= 500 ? true : false,
	                            color: (0, _helpers.rgbToHex)(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
	                            fill: (0, _helpers.rgbToHex)(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2]))
	                        };
	                        if ($.inArray($(cell).css('text-align'), ['left', 'center', 'right', 'start', 'end']) > -1) objOpts.align = $(cell).css('text-align').replace('start', 'left').replace('end', 'right');
	                        if ($.inArray($(cell).css('vertical-align'), ['top', 'middle', 'bottom']) > -1) objOpts.valign = $(cell).css('vertical-align');
	
	                        // C: Add padding [margin] (if any)
	                        // NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
	                        if ($(cell).css('padding-left')) {
	                            objOpts.marginPt = [];
	                            $.each(['padding-top', 'padding-right', 'padding-bottom', 'padding-left'], function (i, val) {
	                                objOpts.marginPt.push(Math.round($(cell).css(val).replace(/\D/gi, '') * ONEPT));
	                            });
	                        }
	
	                        // D: Add colspan (if any)
	                        if ($(cell).attr('colspan')) objOpts.colspan = $(cell).attr('colspan');
	
	                        // E: Add border (if any)
	                        if ($(cell).css('border-top-width') || $(cell).css('border-right-width') || $(cell).css('border-bottom-width') || $(cell).css('border-left-width')) {
	                            objOpts.border = [];
	                            $.each(['top', 'right', 'bottom', 'left'], function (i, val) {
	                                var intBorderW = Math.round(Number($(cell).css('border-' + val + '-width').replace('px', '')));
	                                var arrRGB = [];
	                                arrRGB = $(cell).css('border-' + val + '-color').replace(/\s+/gi, '').replace('rgba(', '').replace('rgb(', '').replace(')', '').split(',');
	                                var strBorderC = (0, _helpers.rgbToHex)(Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2]));
	                                objOpts.border.push({
	                                    pt: intBorderW,
	                                    color: strBorderC
	                                });
	                            });
	                        }
	
	                        // F: Massage cell text so we honor linebreak tag as a line break during line parsing
	                        var $cell = $(cell).clone();
	                        $cell.html($(cell).html().replace(/<br[^>]*>/gi, '\n'));
	
	                        // LAST: Add cell
	                        arrObjTabCells.push({
	                            text: $cell.text(),
	                            opts: objOpts
	                        });
	                    });
	                    switch (val) {
	                        case 'thead':
	                            arrObjTabHeadRows.push(arrObjTabCells);
	                            break;
	                        case 'tbody':
	                            arrObjTabBodyRows.push(arrObjTabCells);
	                            break;
	                        case 'tfoot':
	                            arrObjTabFootRows.push(arrObjTabCells);
	                            break;
	                        default:
	                    }
	                });
	            });
	        }
	    }, {
	        key: 'paginateData',
	        value: function paginateData(arrObjTabHeadRows, arrObjTabBodyRows, arrObjTabFootRows, arrColW, emuTabCurrH, emuSlideTabH, arrRows, arrObjSlides, opts) {
	            $.each([arrObjTabHeadRows, arrObjTabBodyRows, arrObjTabFootRows], function (iTab, tab) {
	                var currRow = [];
	                $.each(tab, function (iRow, row) {
	                    // A: Reset ROW variables
	                    var arrCellsLines = [],
	                        arrCellsLineHeights = [],
	                        emuRowH = 0,
	                        intMaxLineCnt = 0,
	                        intMaxColIdx = 0;
	
	                    // B: Parse and store each cell's text into line array (*MAGIC HAPPENS HERE*)
	                    $(row).each(function (iCell, cell) {
	                        // 1: Create a cell object for each table column
	                        currRow.push({
	                            text: '',
	                            opts: cell.opts
	                        });
	
	                        // 2: Parse cell contents into lines (**MAGIC HAPENSS HERE**)
	                        var lines = (0, _helpers.parseTextToLines)(cell.text, cell.opts.font_size, arrColW[iCell] / ONEPT);
	                        arrCellsLines.push(lines);
	
	                        // 3: Keep track of max line count within all row cells
	                        if (lines.length > intMaxLineCnt) {
	                            intMaxLineCnt = lines.length;
	                            intMaxColIdx = iCell;
	                        }
	                    });
	
	                    // C: Calculate Line-Height
	                    // FYI: Line-Height =~ font-size [px~=pt] * 1.65 / 100 = inches high
	                    // FYI: 1px = 14288 EMU (0.156 inches) @96 PPI - I ended up going with 20000 EMU as margin spacing needed a bit more than 1:1
	                    $(row).each(function (iCell, cell) {
	                        var lineHeight = (0, _helpers.inch2Emu)(cell.opts.font_size * 1.65 / 100);
	                        if (Array.isArray(cell.opts.marginPt) && cell.opts.marginPt[0]) lineHeight += cell.opts.marginPt[0] / intMaxLineCnt;
	                        if (Array.isArray(cell.opts.marginPt) && cell.opts.marginPt[2]) lineHeight += cell.opts.marginPt[2] / intMaxLineCnt;
	                        arrCellsLineHeights.push(Math.round(lineHeight));
	                    });
	
	                    // D: AUTO-PAGING: Add text one-line-a-time to this row's cells until: lines are exhausted OR table H limit is hit
	                    for (var idx = 0; idx < intMaxLineCnt; idx++) {
	                        // 1: Add the current line to cell
	                        for (var col = 0; col < arrCellsLines.length; col++) {
	                            // A: Commit this slide to Presenation if table Height limit is hit
	                            if (emuTabCurrH + arrCellsLineHeights[intMaxColIdx] > emuSlideTabH) {
	                                // 1: Add the current row to table
	                                // NOTE: Edge cases can occur where we create a new slide only to have no more lines
	                                // ....: and then a blank row sits at the bottom of a table!
	                                // ....: Hence, we very all cells have text before adding this final row.
	                                $.each(currRow, function (i, cell) {
	                                    if (cell.text.length > 0) {
	                                        // IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
	                                        arrRows.push($.extend(true, [], currRow));
	                                        return false; // break out of .each loop
	                                    }
	                                });
	                                // 2: Add new Slide with current array of table rows
	                                arrObjSlides.push($.extend(true, [], arrRows));
	                                // 3: Empty rows for new Slide
	                                arrRows.length = 0;
	                                // 4: Reset curr table height for new Slide
	                                emuTabCurrH = 0; // This row's emuRowH w/b added below
	                                // 5: Empty current row's text (continue adding lines where we left off below)
	                                $.each(currRow, function (i, cell) {
	                                    cell.text = '';
	                                });
	                                // 6: Auto-Paging Options: addHeaderToEach
	                                if (opts.addHeaderToEach) {
	                                    var headRow = [];
	                                    $.each(arrObjTabHeadRows[0], function (iCell, cell) {
	                                        headRow.push({
	                                            text: cell.text,
	                                            opts: cell.opts
	                                        });
	                                        var lines = (0, _helpers.parseTextToLines)(cell.text, cell.opts.font_size, arrColW[iCell] / ONEPT);
	                                        if (lines.length > intMaxLineCnt) {
	                                            intMaxLineCnt = lines.length;
	                                            intMaxColIdx = iCell;
	                                        }
	                                    });
	                                    arrRows.push($.extend(true, [], headRow));
	                                }
	                            }
	
	                            // B: Add next line of text to this cell
	                            if (arrCellsLines[col][idx]) currRow[col].text += arrCellsLines[col][idx];
	                        }
	
	                        // 2: Add this new rows H to overall (The cell with the longest line array is the one we use as the determiner for overall row Height)
	                        emuTabCurrH += arrCellsLineHeights[intMaxColIdx];
	                    }
	
	                    // E: Flush row buffer - Add the current row to table, then truncate row cell array
	                    // IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
	                    arrRows.push($.extend(true, [], currRow));
	                    currRow.length = 0;
	                }); // row loop
	            });
	        }
	    }, {
	        key: 'createSlides',
	        value: function createSlides(arrObjSlides, opts, arrInchMargins, emuSlideTabW, arrColW) {
	            $.each(arrObjSlides, function (i, slide) {
	                // A: Create table row array
	                var arrTabRows = [];
	
	                // B: Create new Slide
	                var newSlide = opts && opts.master && gObjPptxMasters ? new _slide2.default().addNewSlide(opts.master) : new _slide2.default().addNewSlide();
	
	                // C: Create array of Rows
	                $.each(slide, function (i, row) {
	                    var arrTabRowCells = [];
	                    $.each(row, function (i, cell) {
	                        arrTabRowCells.push(cell);
	                    });
	                    arrTabRows.push(arrTabRowCells);
	                });
	
	                // D: Add table to Slide
	                newSlide.addTable(arrTabRows, {
	                    x: arrInchMargins[3],
	                    y: arrInchMargins[0],
	                    cx: emuSlideTabW / EMU
	                }, {
	                    colW: arrColW
	                });
	
	                // E: Add any additional objects
	                if (opts.addImage) newSlide.addImage(opts.addImage.url, opts.addImage.x, opts.addImage.y, opts.addImage.w, opts.addImage.h);
	                if (opts.addText) newSlide.addText(opts.addText.text, opts.addText.opts || {});
	                if (opts.addShape) newSlide.addShape(opts.addShape.shape, opts.addShape.opts || {});
	                if (opts.addTable) newSlide.addTable(opts.addTable.rows, opts.addTable.opts || {}, opts.addTable.tabOpts || {});
	            });
	        }
	    }, {
	        key: 'lookMargin',
	        value: function lookMargin(opts) {
	            var arrInchMargins = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
	            if (opts && opts.margin) {
	                if (Array.isArray(opts.margin)) arrInchMargins = opts.margin;else if (!isNaN(opts.margin)) arrInchMargins = [opts.margin, opts.margin, opts.margin, opts.margin];
	            } else if (opts && opts.master && opts.master.margin && gObjPptxMasters) {
	                if (Array.isArray(opts.master.margin)) arrInchMargins = opts.master.margin;else if (!isNaN(opts.master.margin)) arrInchMargins = [opts.master.margin, opts.master.margin, opts.master.margin, opts.master.margin];
	            }
	            return arrInchMargins;
	        }
	    }]);
	
	    return SlideTable;
	}();
	
	exports.default = SlideTable;

/***/ }
/******/ ]);
//# sourceMappingURL=pptxgen.js.map