/*\
|*|  :: pptxgen.js ::
|*|
|*|  JavaScript framework that creates PowerPoint (pptx) presentations
|*|  https://github.com/gitbrent/PptxGenJS
|*|
|*|  This framework is released under the MIT Public License (MIT)
|*|
|*|  PptxGenJS (C) 2015-2019 Brent Ely -- https://github.com/gitbrent
|*|
|*|  Some code derived from the OfficeGen project:
|*|  github.com/Ziv-Barber/officegen/ (Copyright 2013 Ziv Barber)
|*|
|*|  Permission is hereby granted, free of charge, to any person obtaining a copy
|*|  of this software and associated documentation files (the "Software"), to deal
|*|  in the Software without restriction, including without limitation the rights
|*|  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
|*|  copies of the Software, and to permit persons to whom the Software is
|*|  furnished to do so, subject to the following conditions:
|*|
|*|  The above copyright notice and this permission notice shall be included in all
|*|  copies or substantial portions of the Software.
|*|
|*|  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
|*|  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
|*|  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
|*|  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
|*|  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
|*|  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
|*|  SOFTWARE.
\*/

/*
	PPTX Units are "DXA" (except for font sizing)
	....: There are 1440 DXA per inch. 1 inch is 72 points. 1 DXA is 1/20th's of a point (20 DXA is 1 point).
	....: There is also something called EMU's (914400 EMUs is 1 inch, 12700 EMUs is 1pt).
	SEE: https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
	|
	OBJECT LAYOUTS: 16x9 (10" x 5.625"), 16x10 (10" x 6.25"), 4x3 (10" x 7.5"), Wide (13.33" x 7.5") and Custom (any size)
	|
	REFS:
	* "Structure of a PresentationML document (Open XML SDK)"
	* @see: https://msdn.microsoft.com/en-us/library/office/gg278335.aspx
	* TableStyleId enumeration
	* @see: https://msdn.microsoft.com/en-us/library/office/hh273476(v=office.14).aspx
*/

import {
	EMU, ONEPT, CRLF, DEF_SLIDE_MARGIN_IN, LETTERS, BARCHART_COLORS,
	DEF_FONT_COLOR, PIECHART_COLORS, CHART_TYPES, MASTER_OBJECTS, BASE_SHAPES
} from './enums';
import { getMix, encodeXmlEntities } from './utils'
//import { TODO-What funs need this } from './gen-xml';

// Detect Node.js (NODEJS is ultimately used to determine how to save: either `fs` or web-based, so using fs-detection is perfect)
var NODEJS: boolean = false;
var APPJS: boolean = false;
{
	// NOTE: `NODEJS` determines which network library to use, so using fs-detection is apropos.
	if (typeof module !== 'undefined' && module.exports && typeof require === 'function' && typeof window === 'undefined') {
		try {
			require.resolve('fs');
			NODEJS = true;
		}
		catch (ex) {
			NODEJS = false;
		}
	}
	else if (typeof module !== 'undefined' && module.exports && typeof require === 'function' && typeof window !== 'undefined') {
		APPJS = true;
	}
}

// Require [include] colors/shapes for Node/Angular/React, etc.
if (NODEJS || APPJS) {
	var gObjPptxColors = require('../dist/pptxgen.colors.js');
	var gObjPptxShapes = require('../dist/pptxgen.shapes.js');
}

// Polyfill for IE11 (https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Number/isInteger)
Number.isInteger = Number.isInteger || function(value) {
	return typeof value === "number" && isFinite(value) && Math.floor(value) === value;
};

interface ILayout {
	name: string
	width: number
	height: number
}
interface ISlide {
	layout: ILayout
}
interface ISlideLayout {
	name: string
	slide: ISlide
	data: Array<object>
	rels: Array<object>
	margin: Array<number>
	slideNumberObj?: null
}
interface IPresentation {
	author: string
	company: string
	revision: string
	subject: string
	title: string
	isBrowser: boolean
	fileName: string
	fileExtn: string
	pptLayout: ILayout
	rtlMode: false
	saveCallback: null
	masterSlide: object
	chartCounter: number
	imageCounter: number
	slides: ISlide[]
	slideLayouts: ISlideLayout[]
}

let LAYOUTS = {
	"LAYOUT_4x3": { "name": "screen4x3", "width": 9144000, "height": 6858000 } as ILayout,
	"LAYOUT_16x9": { "name": "screen16x9", "width": 9144000, "height": 5143500 } as ILayout,
	"LAYOUT_16x10": { "name": "screen16x10", "width": 9144000, "height": 5715000 } as ILayout,
	"LAYOUT_WIDE": { "name": "custom", "width": 12192000, "height": 6858000 } as ILayout,
	"LAYOUT_USER": { "name": "custom", "width": 12192000, "height": 6858000 } as ILayout
}

export var gObjPptx: IPresentation = {
	// Core
	author: 'PptxGenJS',
	company: 'PptxGenJS',
	revision: '1',
	subject: 'PptxGenJS Presentation',
	title: 'PptxGenJS Presentation',

	// PptxGenJS props
	isBrowser: false,
	fileName: 'Presentation',
	fileExtn: '.pptx',
	pptLayout: LAYOUTS['LAYOUT_16x9'],
	rtlMode: false,
	saveCallback: null,

	// PptxGenJS data
	/** @type {object} master slide layout object */
	masterSlide: {
		slide: {},
		data: [],
		rels: [],
		slideNumberObj: null
	},

	/** @type {Number} global counter for included charts (used for index in their filenames) */
	chartCounter: 0,

	/** @type {Number} global counter for included images (used for index in their filenames) */
	imageCounter: 0,

	/** @type {object[]} this Presentation's Slide objects */
	slides: [],

	/** @type {object[]} slide layout definition objects, used for generating slide layout files */
	slideLayouts: [{
		name: 'BLANK',
		slide: null,
		data: [],
		rels: [],
		margin: DEF_SLIDE_MARGIN_IN,
		slideNumberObj: null
	}]
};
export var gObjPptxGenerators = {}; // TODO: TODO-3: stubbed here -need to attach it to generators!

var PptxGenJS = function() {
	// APP
	var APP_VER = "3.0.0-beta";
	var APP_BLD = "20190517";

	// CONSTANTS
	// TODO-3:

	// IMAGES (base64)
	{
		var IMG_BROKEN = 'data:image/png;base64,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';
		var IMG_PLAYBTN = 'data:image/png;base64,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';
	}
	//
	var JSZIP_OUTPUT_TYPES = ['arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array']; /** @see https://stuk.github.io/jszip/documentation/api_jszip/generate_async.html */
	var SLDNUMFLDID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}';
	//
	var DEF_CELL_BORDER = { color: "666666" };
	var DEF_CELL_MARGIN_PT = [3, 3, 3, 3]; // TRBL-style
	var DEF_SLIDE_BKGD = 'FFFFFF';

	// A: Create internal pptx object
	// B: Set Presentation property defaults
	// TODO-3: use `state` instead of global object

	// C: Expose shape library to clients
	this.charts = CHART_TYPES;
	this.colors = (typeof gObjPptxColors !== 'undefined' ? gObjPptxColors : {});
	this.shapes = (typeof gObjPptxShapes !== 'undefined' ? gObjPptxShapes : BASE_SHAPES);
	// Declare only after `this.colors` is initialized
	var SCHEME_COLOR_NAMES = Object.keys(this.colors).map(function(clrKey) { return this.colors[clrKey] }.bind(this));

	// D: Fall back to base shapes if shapes file was not linked
	gObjPptxShapes = (gObjPptxShapes || this.shapes);

	/* ===============================================================================================
	|
	 #####
	#     # ###### #    # ###### #####    ##   #####  ####  #####   ####
	#       #      ##   # #      #    #  #  #    #   #    # #    # #
	#  #### #####  # #  # #####  #    # #    #   #   #    # #    #  ####
	#     # #      #  # # #      #####  ######   #   #    # #####       #
	#     # #      #   ## #      #   #  #    #   #   #    # #   #  #    #
	 #####  ###### #    # ###### #    # #    #   #    ####  #    #  ####
	|
	==================================================================================================
	*/

	var gObjPptxGenerators = {
		/**
		 * Adds a background image or color to a slide definition.
		 * @param {String|Object} bkg color string or an object with image definition
		 * @param {Object} target slide object that the background is set to
		 */
		addBackgroundDefinition: function addBackgroundDefinition(bkg, target) {
			if (typeof bkg === 'object' && (bkg.src || bkg.path || bkg.data)) {
				// Allow the use of only the data key (`path` isnt reqd)
				bkg.src = bkg.src || bkg.path || null;
				if (!bkg.src) bkg.src = 'preencoded.png';
				var targetRels = target.rels;
				var strImgExtn = (bkg.src.split('.').pop() || 'png').split('?')[0]; // Handle "blah.jpg?width=540" etc.
				if (strImgExtn == 'jpg') strImgExtn = 'jpeg'; // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup

				var intRels = targetRels.length + 1;
				targetRels.push({
					path: bkg.src,
					type: 'image/' + strImgExtn,
					extn: strImgExtn,
					data: (bkg.data || ''),
					rId: intRels,
					Target: '../media/image' + (++gObjPptx.imageCounter) + '.' + strImgExtn
				});
				target.slide.bkgdImgRid = intRels;
			}
			else if (bkg && typeof bkg === 'string') {
				target.slide.back = bkg;
			}
		},

		/**
		 * Adds a text object to a slide definition.
		 * @param {String} text
		 * @param {Object} opt
		 * @param {Object} target slide object that the text should be added to
		 * @since: 1.0.0
		 */
		addTextDefinition: function addTextDefinition(text, opt, target, isPlaceholder) {
			var opt = (opt && typeof opt === 'object' ? opt : {});
			var text = (text || '');
			if (Array.isArray(text) && text.length == 0) text = '';
			var resultObject = {
				type: null,
				text: null,
				options: null
			};

			// STEP 2: Set some options
			// Placeholders should inherit their colors or override them, so don't default them
			if (!opt.placeholder) {
				opt.color = (opt.color || target.slide.color || DEF_FONT_COLOR); // Set color (options > inherit from Slide > default to black)
			}

			// ROBUST: Convert attr values that will likely be passed by users to valid OOXML values
			if (opt.valign) opt.valign = opt.valign.toLowerCase().replace(/^c.*/i, 'ctr').replace(/^m.*/i, 'ctr').replace(/^t.*/i, 't').replace(/^b.*/i, 'b');
			if (opt.align) opt.align = opt.align.toLowerCase().replace(/^c.*/i, 'center').replace(/^m.*/i, 'center').replace(/^l.*/i, 'left').replace(/^r.*/i, 'right');

			// ROBUST: Set rational values for some shadow props if needed
			correctShadowOptions(opt.shadow);

			// STEP 3: Set props
			resultObject.type = isPlaceholder ? 'placeholder' : 'text';
			resultObject.text = text;

			// STEP 4: Set options
			resultObject.options = opt;
			if (opt.shape && opt.shape.name == 'line') {
				opt.line = (opt.line || '333333');
				opt.lineSize = (opt.lineSize || 1);
			}
			resultObject.options.bodyProp = {};
			resultObject.options.bodyProp.autoFit = (opt.autoFit || false); // If true, shape will collapse to text size (Fit To Shape)
			resultObject.options.bodyProp.anchor = (opt.valign || (!opt.placeholder ? 'ctr' : null)); // VALS: [t,ctr,b]
			resultObject.options.bodyProp.rot = (opt.rotate || null); // VALS: degree * 60,000
			resultObject.options.bodyProp.vert = (opt.vert || null); // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
			resultObject.options.lineSpacing = (opt.lineSpacing && !isNaN(opt.lineSpacing) ? opt.lineSpacing : null);

			if ((opt.inset && !isNaN(Number(opt.inset))) || opt.inset == 0) {
				resultObject.options.bodyProp.lIns = inch2Emu(opt.inset);
				resultObject.options.bodyProp.rIns = inch2Emu(opt.inset);
				resultObject.options.bodyProp.tIns = inch2Emu(opt.inset);
				resultObject.options.bodyProp.bIns = inch2Emu(opt.inset);
			}

			target.data.push(resultObject);
			createHyperlinkRels(text || '', target.rels);

			return resultObject;
		},

		/**
		 * Adds Notes to a slide.
		 * @param {String} notes
		 * @param {Object} opt (*unused*)
		 * @param {Object} target slide object
		 * @since 2.3.0
		 */
		addNotesDefinition: function addNotesDefinition(notes, opt, target) {
			var opt = (opt && typeof opt === 'object' ? opt : {});
			var resultObject = {
				type: null,
				text: null
			};

			resultObject.type = 'notes';
			resultObject.text = notes;

			target.data.push(resultObject);

			return resultObject;
		},

		/**
		 * Adds a placeholder object to a slide definition.
		 * @param {String} text
		 * @param {Object} opt
		 * @param {Object} target slide object that the placeholder should be added to
		 */
		addPlaceholderDefinition: function addPlaceholderDefinition(text, opt, target) {
			return gObjPptxGenerators.addTextDefinition(text, opt, target, true);
		},

		/**
		 * Adds a shape object to a slide definition.
		 * @param {Object} shape shape const object (pptx.shapes)
		 * @param {Object} opt
		 * @param {Object} target slide object that the shape should be added to
		 * @return {Object} shape object
		 */
		addShapeDefinition: function addShapeDefinition(shape, opt, target) {
			var options = (typeof opt === 'object' ? opt : {});
			var resultObject = {
				type: null,
				text: null,
				options: {}
			};

			if (!shape || typeof shape !== 'object') {
				console.error("Missing/Invalid shape parameter! Example: `addShape(pptx.shapes.LINE, {x:1, y:1, w:1, h:1});` ");
				return;
			}

			resultObject.type = 'text';
			resultObject.options = options;
			options.shape = shape;
			options.x = (options.x || (options.x == 0 ? 0 : 1));
			options.y = (options.y || (options.y == 0 ? 0 : 1));
			options.w = (options.w || (options.w == 0 ? 0 : 1));
			options.h = (options.h || (options.h == 0 ? 0 : 1));
			options.line = (options.line || (shape.name == 'line' ? '333333' : null));
			options.lineSize = (options.lineSize || (shape.name == 'line' ? 1 : null));
			if (['dash', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot', 'solid', 'sysDash', 'sysDot'].indexOf(options.lineDash || '') < 0) options.lineDash = 'solid';

			target.data.push(resultObject);
			return resultObject;
		},

		/**
		 * Adds an image object to a slide definition.
		 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
		 * @param {Object} objImage - object containing `path`/`data`, `x`, `y`, etc.
		 * @param {Object} target - slide that the image should be added to (if not specified as the 2nd arg)
		 * @return {Object} image object
		 */
		addImageDefinition: function addImageDefinition(objImage, target) {
			var resultObject = {
				type: null,
				text: null,
				options: null,
				image: null,
				imageRid: null,
				hyperlink: null
			};
			// FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
			var intPosX = (objImage.x || 0);
			var intPosY = (objImage.y || 0);
			var intWidth = (objImage.w || 0);
			var intHeight = (objImage.h || 0);
			var sizing = objImage.sizing || null;
			var objHyperlink = (objImage.hyperlink || '');
			var strImageData = (objImage.data || '');
			var strImagePath = (objImage.path || '');
			var imageRelId = target.rels.length + 1;

			// REALITY-CHECK:
			if (!strImagePath && !strImageData) {
				console.error("ERROR: `addImage()` requires either 'data' or 'path' parameter!");
				return null;
			}
			else if (strImageData && strImageData.toLowerCase().indexOf('base64,') == -1) {
				console.error("ERROR: Image `data` value lacks a base64 header! Ex: 'image/png;base64,NMP[...]')");
				return null;
			}

			// STEP 1: Set extension
			// NOTE: Split to address URLs with params (eg: `path/brent.jpg?someParam=true`)
			var strImgExtn = strImagePath.split('.').pop().split("?")[0].split("#")[0] || 'png';
			// However, pre-encoded images can be whatever mime-type they want (and good for them!)
			if (strImageData && /image\/(\w+)\;/.exec(strImageData) && /image\/(\w+)\;/.exec(strImageData).length > 0) {
				strImgExtn = /image\/(\w+)\;/.exec(strImageData)[1];
			}
			else if (strImageData && strImageData.toLowerCase().indexOf('image/svg+xml') > -1) {
				strImgExtn = 'svg';
			}
			// STEP 2: Set type/path
			resultObject.type = 'image';
			resultObject.image = (strImagePath || 'preencoded.png');

			// STEP 3: Set image properties & options
			// FIXME: Measure actual image when no intWidth/intHeight params passed
			// ....: This is an async process: we need to make getSizeFromImage use callback, then set H/W...
			// if ( !intWidth || !intHeight ) { var imgObj = getSizeFromImage(strImagePath);
			var imgObj = { width: 1, height: 1 };
			resultObject.options = {
				x: (intPosX || 0),
				y: (intPosY || 0),
				cx: (intWidth || imgObj.width),
				cy: (intHeight || imgObj.height),
				rounding: (objImage.rounding || false),
				sizing: sizing,
				placeholder: objImage.placeholder
			};

			// STEP 4: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
			if (strImgExtn == 'svg') {
				// SVG files consume *TWO* rId's: (a png version and the svg image)
				// <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
				// <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/>

				target.rels.push({
					path: (strImagePath || strImageData + 'png'),
					type: 'image/png',
					extn: 'png',
					data: (strImageData || ''),
					rId: imageRelId,
					Target: '../media/image' + (++gObjPptx.imageCounter) + '.png',
					isSvgPng: true,
					svgSize: { w: resultObject.options.cx, h: resultObject.options.cy }
				});
				resultObject.imageRid = imageRelId;
				target.rels.push({
					path: (strImagePath || strImageData),
					type: 'image/' + strImgExtn,
					extn: strImgExtn,
					data: (strImageData || ''),
					rId: (imageRelId + 1),
					Target: '../media/image' + (++gObjPptx.imageCounter) + '.' + strImgExtn
				});
				resultObject.imageRid = (imageRelId + 1);
			}
			else {
				target.rels.push({
					path: (strImagePath || 'preencoded.' + strImgExtn),
					type: 'image/' + strImgExtn,
					extn: strImgExtn,
					data: (strImageData || ''),
					rId: imageRelId,
					Target: '../media/image' + (++gObjPptx.imageCounter) + '.' + strImgExtn
				});
				resultObject.imageRid = imageRelId;
			}

			// STEP 5: (Issue#77) Hyperlink support
			if (typeof objHyperlink === 'object') {
				if (!objHyperlink.url && !objHyperlink.slide) console.log("ERROR: 'hyperlink requires either: `url` or `slide`'");
				else {
					var intRelId = imageRelId + 1;

					target.rels.push({
						type: 'hyperlink',
						data: (objHyperlink.slide ? 'slide' : 'dummy'),
						rId: intRelId,
						Target: objHyperlink.url || objHyperlink.slide
					});

					objHyperlink.rId = intRelId;
					resultObject.hyperlink = objHyperlink;
				}
			}

			target.data.push(resultObject);
			return resultObject;
		},

		/**
		 * Generate the chart based on input data.
		 * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
		 *
		 * @param {object} type should belong to: 'column', 'pie'
		 * @param {object} data a JSON object with follow the following format
		 * @param {object} opt
		 * @param {object} target slide object that the chart should be added to
		 * @return {Object} chart object
		 * {
		 *   title: 'eSurvey chart',
		 *   data: [
		 *		{
		 *			name: 'Income',
		 *			labels: ['2005', '2006', '2007', '2008', '2009'],
		 *			values: [23.5, 26.2, 30.1, 29.5, 24.6]
		 *		},
		 *		{
		 *			name: 'Expense',
		 *			labels: ['2005', '2006', '2007', '2008', '2009'],
		 *			values: [18.1, 22.8, 23.9, 25.1, 25]
		 *		}
		 *	 ]
		 *	}
		 */
		addChartDefinition: function addChartDefinition(type, data, opt, target) {
			var targetRels = target.rels;
			var chartId = (++gObjPptx.chartCounter);
			var chartRelId = target.rels.length + 1;
			var resultObject = {
				type: null,
				text: null,
				options: null,
				chartRid: null
			};
			// DESIGN: `type` can an object (ex: `pptx.charts.DOUGHNUT`) or an array of chart objects
			// EX: addChartDefinition([ { type:pptx.charts.BAR, data:{name:'', labels:[], values[]} }, {<etc>} ])
			// Multi-Type Charts
			var tmpOpt;
			var tmpData = [], options;
			if (Array.isArray(type)) {
				// For multi-type charts there needs to be data for each type,
				// as well as a single data source for non-series operations.
				// The data is indexed below to keep the data in order when segmented
				// into types.
				type.forEach(function(obj) {
					tmpData = tmpData.concat(obj.data)
				});
				tmpOpt = data || opt;
			}
			else {
				tmpData = data;
				tmpOpt = opt;
			}
			tmpData.forEach(function(item, i) { item.index = i; });
			options = (tmpOpt && typeof tmpOpt === 'object' ? tmpOpt : {});

			// STEP 1: TODO: check for reqd fields, correct type, etc
			// `type` exists in CHART_TYPES
			// Array.isArray(data)
			/*
			if ( Array.isArray(rel.data) && rel.data.length > 0 && typeof rel.data[0] === 'object'
				&& rel.data[0].labels && Array.isArray(rel.data[0].labels)
				&& rel.data[0].values && Array.isArray(rel.data[0].values) ) {
				obj = rel.data[0];
			}
			else {
				console.warn("USAGE: addChart( 'pie', [ {name:'Sales', labels:['Jan','Feb'], values:[10,20]} ], {x:1, y:1} )");
				return;
			}
			*/

			// STEP 2: Set default options/decode user options
			// A: Core
			options.type = type;
			options.x = (typeof options.x !== 'undefined' && options.x != null && !isNaN(options.x) ? options.x : 1);
			options.y = (typeof options.y !== 'undefined' && options.y != null && !isNaN(options.y) ? options.y : 1);
			options.w = (options.w || '50%');
			options.h = (options.h || '50%');

			// B: Options: misc
			if (['bar', 'col'].indexOf(options.barDir || '') < 0) options.barDir = 'col';
			// IMPORTANT: 'bestFit' will cause issues with PPT-Online in some cases, so defualt to 'ctr'!
			if (['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't'].indexOf(options.dataLabelPosition || '') < 0) options.dataLabelPosition = (options.type.name == 'pie' || options.type.name == 'doughnut' ? 'bestFit' : 'ctr');
			options.dataLabelBkgrdColors = (options.dataLabelBkgrdColors == true || options.dataLabelBkgrdColors == false ? options.dataLabelBkgrdColors : false);
			if (['b', 'l', 'r', 't', 'tr'].indexOf(options.legendPos || '') < 0) options.legendPos = 'r';
			// barGrouping: "21.2.3.17 ST_Grouping (Grouping)"
			if (['clustered', 'standard', 'stacked', 'percentStacked'].indexOf(options.barGrouping || '') < 0) options.barGrouping = 'standard';
			if (options.barGrouping.indexOf('tacked') > -1) {
				options.dataLabelPosition = 'ctr'; // IMPORTANT: PPT-Online will not open Presentation when 'outEnd' etc is used on stacked!
				if (!options.barGapWidthPct) options.barGapWidthPct = 50;
			}
			// 3D bar: ST_Shape
			if (['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax'].indexOf(options.bar3DShape || '') < 0) options.bar3DShape = 'box';
			// lineDataSymbol: http://www.datypic.com/sc/ooxml/a-val-32.html
			// Spec has [plus,star,x] however neither PPT2013 nor PPT-Online support them
			if (['circle', 'dash', 'diamond', 'dot', 'none', 'square', 'triangle'].indexOf(options.lineDataSymbol || '') < 0) options.lineDataSymbol = 'circle';
			if (['gap', 'span'].indexOf(options.displayBlanksAs || '') < 0) options.displayBlanksAs = 'span';
			if (['standard', 'marker', 'filled'].indexOf(options.radarStyle || '') < 0) options.radarStyle = 'standard';
			options.lineDataSymbolSize = (options.lineDataSymbolSize && !isNaN(options.lineDataSymbolSize) ? options.lineDataSymbolSize : 6);
			options.lineDataSymbolLineSize = (options.lineDataSymbolLineSize && !isNaN(options.lineDataSymbolLineSize) ? options.lineDataSymbolLineSize * ONEPT : 0.75 * ONEPT);
			// `layout` allows the override of PPT defaults to maximize space
			if (options.layout) {
				['x', 'y', 'w', 'h'].forEach(function(key) {
					var val = options.layout[key];
					if (isNaN(Number(val)) || val < 0 || val > 1) {
						console.warn('Warning: chart.layout.' + key + ' can only be 0-1');
						delete options.layout[key]; // remove invalid value so that default will be used
					}
				});
			}

			// Set gridline defaults
			options.catGridLine = options.catGridLine || (type.name == 'scatter' ? { color: 'D9D9D9', pt: 1 } : 'none');
			options.valGridLine = options.valGridLine || (type.name == 'scatter' ? { color: 'D9D9D9', pt: 1 } : {});
			options.serGridLine = options.serGridLine || (type.name == 'scatter' ? { color: 'D9D9D9', pt: 1 } : 'none');
			correctGridLineOptions(options.catGridLine);
			correctGridLineOptions(options.valGridLine);
			correctGridLineOptions(options.serGridLine);
			correctShadowOptions(options.shadow);

			// C: Options: plotArea
			options.showDataTable = (options.showDataTable == true || options.showDataTable == false ? options.showDataTable : false);
			options.showDataTableHorzBorder = (options.showDataTableHorzBorder == true || options.showDataTableHorzBorder == false ? options.showDataTableHorzBorder : true);
			options.showDataTableVertBorder = (options.showDataTableVertBorder == true || options.showDataTableVertBorder == false ? options.showDataTableVertBorder : true);
			options.showDataTableOutline = (options.showDataTableOutline == true || options.showDataTableOutline == false ? options.showDataTableOutline : true);
			options.showDataTableKeys = (options.showDataTableKeys == true || options.showDataTableKeys == false ? options.showDataTableKeys : true);
			options.showLabel = (options.showLabel == true || options.showLabel == false ? options.showLabel : false);
			options.showLegend = (options.showLegend == true || options.showLegend == false ? options.showLegend : false);
			options.showPercent = (options.showPercent == true || options.showPercent == false ? options.showPercent : true);
			options.showTitle = (options.showTitle == true || options.showTitle == false ? options.showTitle : false);
			options.showValue = (options.showValue == true || options.showValue == false ? options.showValue : false);
			options.catAxisLineShow = (typeof options.catAxisLineShow !== 'undefined' ? options.catAxisLineShow : true);
			options.valAxisLineShow = (typeof options.valAxisLineShow !== 'undefined' ? options.valAxisLineShow : true);
			options.serAxisLineShow = (typeof options.serAxisLineShow !== 'undefined' ? options.serAxisLineShow : true);

			options.v3DRotX = (!isNaN(options.v3DRotX) && options.v3DRotX >= -90 && options.v3DRotX <= 90 ? options.v3DRotX : 30);
			options.v3DRotY = (!isNaN(options.v3DRotY) && options.v3DRotY >= 0 && options.v3DRotY <= 360 ? options.v3DRotY : 30);
			options.v3DRAngAx = (options.v3DRAngAx == true || options.v3DRAngAx == false ? options.v3DRAngAx : true);
			options.v3DPerspective = (!isNaN(options.v3DPerspective) && options.v3DPerspective >= 0 && options.v3DPerspective <= 240 ? options.v3DPerspective : 30);

			// D: Options: chart
			options.barGapWidthPct = (!isNaN(options.barGapWidthPct) && options.barGapWidthPct >= 0 && options.barGapWidthPct <= 1000 ? options.barGapWidthPct : 150);
			options.barGapDepthPct = (!isNaN(options.barGapDepthPct) && options.barGapDepthPct >= 0 && options.barGapDepthPct <= 1000 ? options.barGapDepthPct : 150);

			options.chartColors = (Array.isArray(options.chartColors) ? options.chartColors : (options.type.name == 'pie' || options.type.name == 'doughnut' ? PIECHART_COLORS : BARCHART_COLORS));
			options.chartColorsOpacity = (options.chartColorsOpacity && !isNaN(options.chartColorsOpacity) ? options.chartColorsOpacity : null);
			//
			options.border = (options.border && typeof options.border === 'object' ? options.border : null);
			if (options.border && (!options.border.pt || isNaN(options.border.pt))) options.border.pt = 1;
			if (options.border && (!options.border.color || typeof options.border.color !== 'string' || options.border.color.length != 6)) options.border.color = '363636';
			//
			options.dataBorder = (options.dataBorder && typeof options.dataBorder === 'object' ? options.dataBorder : null);
			if (options.dataBorder && (!options.dataBorder.pt || isNaN(options.dataBorder.pt))) options.dataBorder.pt = 0.75;
			if (options.dataBorder && (!options.dataBorder.color || typeof options.dataBorder.color !== 'string' || options.dataBorder.color.length != 6)) options.dataBorder.color = 'F9F9F9';
			//
			if (!options.dataLabelFormatCode && options.type.name === 'scatter') options.dataLabelFormatCode = "General";
			options.dataLabelFormatCode = options.dataLabelFormatCode && typeof options.dataLabelFormatCode === 'string' ? options.dataLabelFormatCode : (options.type.name == 'pie' || options.type.name == 'doughnut') ? '0%' : '#,##0';
			//
			// Set default format for Scatter chart labels to custom string if not defined
			if (!options.dataLabelFormatScatter && options.type.name === 'scatter') options.dataLabelFormatScatter = 'custom';
			//
			options.lineSize = (typeof options.lineSize === 'number' ? options.lineSize : 2);
			options.valAxisMajorUnit = (typeof options.valAxisMajorUnit === 'number' ? options.valAxisMajorUnit : null);
			options.valAxisCrossesAt = (options.valAxisCrossesAt || 'autoZero');

			// STEP 4: Set props
			resultObject.type = 'chart';
			resultObject.options = options;

			// STEP 5: Add this chart to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
			targetRels.push({
				rId: chartRelId,
				data: tmpData,
				opts: options,
				type: 'chart',
				globalId: chartId,
				fileName: 'chart' + chartId + '.xml',
				Target: '/ppt/charts/chart' + chartId + '.xml'
			});
			resultObject.chartRid = chartRelId;

			target.data.push(resultObject);
			return resultObject;
		},

		/* ===== */

		/**
		 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
		 * The following object is expected as a slide definition:
		 * {
		 *   bkgd: 'FF00FF',
		 *   objects: [{
		 *     text: {
		 *       text: 'Hello World',
		 *       x: 1,
		 *       y: 1
		 *     }
		 *   }]
		 * }
		 * @param {Object} slideDef slide definition
		 * @param {Object} target empty slide object that should be updated by the passed definition
		 */
		createSlideObject: function createSlideObject(slideDef, target) {
			// STEP 1: Add background
			if (slideDef.bkgd) {
				gObjPptxGenerators.addBackgroundDefinition(slideDef.bkgd, target);
			}

			// STEP 2: Add all Slide Master objects in the order they were given (Issue#53)
			if (slideDef.objects && Array.isArray(slideDef.objects) && slideDef.objects.length > 0) {
				slideDef.objects.forEach(function(object, idx) {
					var key = Object.keys(object)[0];
					if (MASTER_OBJECTS[key] && key == 'chart') gObjPptxGenerators.addChartDefinition(CHART_TYPES[(object.chart.type || '').toUpperCase()], object.chart.data, object.chart.opts, target);
					else if (MASTER_OBJECTS[key] && key == 'image') gObjPptxGenerators.addImageDefinition(object[key], target);
					else if (MASTER_OBJECTS[key] && key == 'line') gObjPptxGenerators.addShapeDefinition(gObjPptxShapes.LINE, object[key], target);
					else if (MASTER_OBJECTS[key] && key == 'rect') gObjPptxGenerators.addShapeDefinition(gObjPptxShapes.RECTANGLE, object[key], target);
					else if (MASTER_OBJECTS[key] && key == 'text') gObjPptxGenerators.addTextDefinition(object[key].text, object[key].options, target, false);
					else if (MASTER_OBJECTS[key] && key == 'placeholder') {
						// TODO: 20180820: Check for existing `name`?
						object[key].options.placeholderName = object[key].options.name; delete object[key].options.name; // remap name for earier handling internally
						object[key].options.placeholderType = object[key].options.type; delete object[key].options.type; // remap name for earier handling internally
						object[key].options.placeholderIdx = (100 + idx);
						gObjPptxGenerators.addPlaceholderDefinition(object[key].text, object[key].options, target);
					}
				});
			}

			// STEP 3: Add Slide Numbers (NOTE: Do this last so numbers are not covered by objects!)
			if (slideDef.slideNumber && typeof slideDef.slideNumber === 'object') {
				target.slideNumberObj = slideDef.slideNumber;
			};
		},

		/**
		 * Transforms a slide object to resulting XML string.
		 * @param {Object} slideObject slide object created within gObjPptxGenerators.createSlideObject
		 * @return {String} XML string with <p:cSld> as the root
		 */
		slideObjectToXml: function slideObjectToXml(slideObject) {
			var strSlideXml = slideObject.name ? '<p:cSld name="' + slideObject.name + '">' : '<p:cSld>';
			var intTableNum = 1;

			// STEP 1: Add background
			if (slideObject.slide.back) {
				strSlideXml += genXmlColorSelection(false, slideObject.slide.back);
			}

			// STEP 2: Add background image (using Strech) (if any)
			if (slideObject.slide.bkgdImgRid) {
				// FIXME: We should be doing this in the slideLayout...
				strSlideXml += '<p:bg>'
					+ '<p:bgPr><a:blipFill dpi="0" rotWithShape="1">'
					+ '<a:blip r:embed="rId' + slideObject.slide.bkgdImgRid + '"><a:lum/></a:blip>'
					+ '<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
					+ '<a:effectLst/></p:bgPr>'
					+ '</p:bg>';
			}

			// STEP 3: Continue slide by starting spTree node
			strSlideXml += '<p:spTree>';
			strSlideXml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>';
			strSlideXml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>';
			strSlideXml += '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

			// STEP 4: Loop over all Slide.data objects and add them to this slide ===============================
			slideObject.data.forEach(function(slideItemObj, idx) {
				var x = 0, y = 0, cx = getSmartParseNumber('75%', 'X'), cy = 0, placeholderObj;
				var locationAttr = "", shapeType = null;

				if (slideObject.layoutObj && slideObject.layoutObj.data && slideItemObj.options && slideItemObj.options.placeholder) {
					placeholderObj = slideObject.layoutObj.data.filter(function(layoutObj) { return layoutObj.options.placeholderName == slideItemObj.options.placeholder })[0];
				}

				// A: Set option vars
				slideItemObj.options = slideItemObj.options || {};

				if (slideItemObj.options.w || slideItemObj.options.w == 0) slideItemObj.options.cx = slideItemObj.options.w;
				if (slideItemObj.options.h || slideItemObj.options.h == 0) slideItemObj.options.cy = slideItemObj.options.h;
				//
				if (slideItemObj.options.x || slideItemObj.options.x == 0) x = getSmartParseNumber(slideItemObj.options.x, 'X');
				if (slideItemObj.options.y || slideItemObj.options.y == 0) y = getSmartParseNumber(slideItemObj.options.y, 'Y');
				if (slideItemObj.options.cx || slideItemObj.options.cx == 0) cx = getSmartParseNumber(slideItemObj.options.cx, 'X');
				if (slideItemObj.options.cy || slideItemObj.options.cy == 0) cy = getSmartParseNumber(slideItemObj.options.cy, 'Y');

				// If using a placeholder then inherit it's position
				if (placeholderObj) {
					if (placeholderObj.options.x || placeholderObj.options.x == 0) x = getSmartParseNumber(placeholderObj.options.x, 'X');
					if (placeholderObj.options.y || placeholderObj.options.y == 0) y = getSmartParseNumber(placeholderObj.options.y, 'Y');
					if (placeholderObj.options.cx || placeholderObj.options.cx == 0) cx = getSmartParseNumber(placeholderObj.options.cx, 'X');
					if (placeholderObj.options.cy || placeholderObj.options.cy == 0) cy = getSmartParseNumber(placeholderObj.options.cy, 'Y');
				}
				//
				if (slideItemObj.options.shape) shapeType = getShapeInfo(slideItemObj.options.shape);
				//
				if (slideItemObj.options.flipH) locationAttr += ' flipH="1"';
				if (slideItemObj.options.flipV) locationAttr += ' flipV="1"';
				if (slideItemObj.options.rotate) locationAttr += ' rot="' + convertRotationDegrees(slideItemObj.options.rotate) + '"';

				// B: Add OBJECT to current Slide ----------------------------
				switch (slideItemObj.type) {
					case 'table':
						// FIRST: Ensure we have rows - otherwise, bail!
						if (!slideItemObj.arrTabRows || (Array.isArray(slideItemObj.arrTabRows) && slideItemObj.arrTabRows.length == 0)) break;

						// Set table vars
						var objTableGrid = {};
						var arrTabRows = slideItemObj.arrTabRows;
						var objTabOpts = slideItemObj.options;
						var intColCnt = 0, intColW = 0;

						// Calc number of columns
						// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
						// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
						arrTabRows[0].forEach(function(cell, idx) {
							var cellOpts = cell.options || cell.opts || null;
							intColCnt += (cellOpts && cellOpts.colspan ? Number(cellOpts.colspan) : 1);
						});

						// STEP 1: Start Table XML =============================
						// NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
						var strXml = '<p:graphicFrame>'
							+ '  <p:nvGraphicFramePr>'
							+ '    <p:cNvPr id="' + (intTableNum * slideObject.numb + 1) + '" name="Table ' + (intTableNum * slideObject.numb) + '"/>'
							+ '    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
							+ '    <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>'
							+ '  </p:nvGraphicFramePr>'
							+ '  <p:xfrm>'
							+ '    <a:off  x="' + (x || (x == 0 ? 0 : EMU)) + '"  y="' + (y || (y == 0 ? 0 : EMU)) + '"/>'
							+ '    <a:ext cx="' + (cx || (cx == 0 ? 0 : EMU)) + '" cy="' + (cy || (cy == 0 ? 0 : EMU)) + '"/>'
							+ '  </p:xfrm>'
							+ '  <a:graphic>'
							+ '    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
							+ '      <a:tbl>'
							+ '        <a:tblPr/>';
						// + '        <a:tblPr bandRow="1"/>';

						// FIXME: Support banded rows, first/last row, etc.
						// NOTE: Banding, etc. only shows when using a table style! (or set alt row color if banding)
						// <a:tblPr firstCol="0" firstRow="0" lastCol="0" lastRow="0" bandCol="0" bandRow="1">

						// STEP 2: Set column widths
						// Evenly distribute cols/rows across size provided when applicable (calc them if only overall dimensions were provided)
						// A: Col widths provided?
						if (Array.isArray(objTabOpts.colW)) {
							strXml += '<a:tblGrid>';
							for (var col = 0; col < intColCnt; col++) {
								strXml += '  <a:gridCol w="' + Math.round(inch2Emu(objTabOpts.colW[col]) || (slideItemObj.options.cx / intColCnt)) + '"/>';
							}
							strXml += '</a:tblGrid>';
						}
						// B: Table Width provided without colW? Then distribute cols
						else {
							intColW = (objTabOpts.colW ? objTabOpts.colW : EMU);
							if (slideItemObj.options.cx && !objTabOpts.colW) intColW = Math.round(slideItemObj.options.cx / intColCnt); // FIX: Issue#12
							strXml += '<a:tblGrid>';
							for (var col = 0; col < intColCnt; col++) { strXml += '<a:gridCol w="' + intColW + '"/>'; }
							strXml += '</a:tblGrid>';
						}

						// STEP 3: Build our row arrays into an actual grid to match the XML we will be building next (ISSUE #36)
						// Note row arrays can arrive "lopsided" as in row1:[1,2,3] row2:[3] when first two cols rowspan!,
						// so a simple loop below in XML building wont suffice to build table correctly.
						// We have to build an actual grid now
						/*
							EX: (A0:rowspan=3, B1:rowspan=2, C1:colspan=2)

							/------|------|------|------\
							|  A0  |  B0  |  C0  |  D0  |
							|      |  B1  |  C1  |      |
							|      |      |  C2  |  D2  |
							\------|------|------|------/
						*/
						jQuery.each(arrTabRows, function(rIdx, row) {
							// A: Create row if needed (recall one may be created in loop below for rowspans, so dont assume we need to create one each iteration)
							if (!objTableGrid[rIdx]) objTableGrid[rIdx] = {};

							// B: Loop over all cells
							jQuery(row).each(function(cIdx, cell) {
								// DESIGN: NOTE: Row cell arrays can be "uneven" (diff cell count in each) due to rowspan/colspan
								// Therefore, for each cell we run 0->colCount to determien the correct slot for it to reside
								// as the uneven/mixed nature of the data means we cannot use the cIdx value alone.
								// E.g.: the 2nd element in the row array may actually go into the 5th table grid row cell b/c of colspans!
								for (var idx = 0; (cIdx + idx) < intColCnt; idx++) {
									var currColIdx = (cIdx + idx);

									if (!objTableGrid[rIdx][currColIdx]) {
										// A: Set this cell
										objTableGrid[rIdx][currColIdx] = cell;

										// B: Handle `colspan` or `rowspan` (a {cell} cant have both! FIXME: FUTURE: ROWSPAN & COLSPAN in same cell)
										if (cell && cell.opts && cell.opts.colspan && !isNaN(Number(cell.opts.colspan))) {
											for (var idy = 1; idy < Number(cell.opts.colspan); idy++) {
												objTableGrid[rIdx][currColIdx + idy] = { "hmerge": true, text: "hmerge" };
											}
										}
										else if (cell && cell.opts && cell.opts.rowspan && !isNaN(Number(cell.opts.rowspan))) {
											for (var idz = 1; idz < Number(cell.opts.rowspan); idz++) {
												if (!objTableGrid[rIdx + idz]) objTableGrid[rIdx + idz] = {};
												objTableGrid[rIdx + idz][currColIdx] = { "vmerge": true, text: "vmerge" };
											}
										}

										// C: Break out of colCnt loop now that slot has been filled
										break;
									}
								}
							});
						});

						/* Only useful for rowspan/colspan testing
						if ( objTabOpts.debug ) {
							console.table(objTableGrid);
							var arrText = [];
							jQuery.each(objTableGrid, function(i,row){ var arrRow = []; jQuery.each(row,function(i,cell){ arrRow.push(cell.text); }); arrText.push(arrRow); });
							console.table( arrText );
						}
						*/

						// STEP 4: Build table rows/cells ============================
						jQuery.each(objTableGrid, function(rIdx, rowObj) {
							// A: Table Height provided without rowH? Then distribute rows
							var intRowH = 0; // IMPORTANT: Default must be zero for auto-sizing to work
							if (Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx]) intRowH = inch2Emu(Number(objTabOpts.rowH[rIdx]));
							else if (objTabOpts.rowH && !isNaN(Number(objTabOpts.rowH))) intRowH = inch2Emu(Number(objTabOpts.rowH));
							else if (slideItemObj.options.cy || slideItemObj.options.h) intRowH = (slideItemObj.options.h ? inch2Emu(slideItemObj.options.h) : slideItemObj.options.cy) / arrTabRows.length;

							// B: Start row
							strXml += '<a:tr h="' + intRowH + '">';

							// C: Loop over each CELL
							jQuery.each(rowObj, function(cIdx, cell) {
								// 1: "hmerge" cells are just place-holders in the table grid - skip those and go to next cell
								if (cell.hmerge) return;

								// 2: OPTIONS: Build/set cell options ===========================
								{
									var cellOpts = cell.options || cell.opts || {};
									if (typeof cell === 'number' || typeof cell === 'string') cell = { text: cell.toString() };
									cellOpts.isTableCell = true; // Used to create textBody XML
									cell.options = cellOpts;

									// B: Apply default values (tabOpts being used when cellOpts dont exist):
									// SEE: http://officeopenxml.com/drwTableCellProperties-alignment.php
									['align', 'bold', 'border', 'color', 'fill', 'fontFace', 'fontSize', 'margin', 'underline', 'valign']
										.forEach(function(name, idx) {
											if (objTabOpts[name] && !cellOpts[name] && cellOpts[name] != 0) cellOpts[name] = objTabOpts[name];
										});

									var cellValign = (cellOpts.valign) ? ' anchor="' + cellOpts.valign.replace(/^c$/i, 'ctr').replace(/^m$/i, 'ctr').replace('center', 'ctr').replace('middle', 'ctr').replace('top', 't').replace('btm', 'b').replace('bottom', 'b') + '"' : '';
									var cellColspan = (cellOpts.colspan) ? ' gridSpan="' + cellOpts.colspan + '"' : '';
									var cellRowspan = (cellOpts.rowspan) ? ' rowSpan="' + cellOpts.rowspan + '"' : '';
									var cellFill = ((cell.optImp && cell.optImp.fill) || cellOpts.fill) ? ' <a:solidFill><a:srgbClr val="' + ((cell.optImp && cell.optImp.fill) || cellOpts.fill.replace('#', '')) + '"/></a:solidFill>' : '';
									var cellMargin = (cellOpts.margin == 0 || cellOpts.margin ? cellOpts.margin : DEF_CELL_MARGIN_PT);
									if (!Array.isArray(cellMargin) && typeof cellMargin === 'number') cellMargin = [cellMargin, cellMargin, cellMargin, cellMargin];
									cellMargin = ' marL="' + cellMargin[3] * ONEPT + '" marR="' + cellMargin[1] * ONEPT + '" marT="' + cellMargin[0] * ONEPT + '" marB="' + cellMargin[2] * ONEPT + '"';
								}

								// FIXME: Cell NOWRAP property (text wrap: add to a:tcPr (horzOverflow="overflow" or whatev opts exist)

								// 3: ROWSPAN: Add dummy cells for any active rowspan
								if (cell.vmerge) {
									strXml += '<a:tc vMerge="1"><a:tcPr/></a:tc>';
									return;
								}

								// 4: Set CELL content and properties ==================================
								strXml += '<a:tc' + cellColspan + cellRowspan + '>' + genXmlTextBody(cell) + '<a:tcPr' + cellMargin + cellValign + '>';

								// 5: Borders: Add any borders
								if (cellOpts.border && typeof cellOpts.border === 'string' && cellOpts.border.toLowerCase() == 'none') {
									strXml += '  <a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL>';
									strXml += '  <a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR>';
									strXml += '  <a:lnT w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnT>';
									strXml += '  <a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB>';
								}
								else if (cellOpts.border && typeof cellOpts.border === 'string') {
									strXml += '  <a:lnL w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnL>';
									strXml += '  <a:lnR w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnR>';
									strXml += '  <a:lnT w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnT>';
									strXml += '  <a:lnB w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnB>';
								}
								else if (cellOpts.border && Array.isArray(cellOpts.border)) {
									jQuery.each([{ idx: 3, name: 'lnL' }, { idx: 1, name: 'lnR' }, { idx: 0, name: 'lnT' }, { idx: 2, name: 'lnB' }], function(i, obj) {
										if (cellOpts.border[obj.idx]) {
											var strC = '<a:solidFill><a:srgbClr val="' + ((cellOpts.border[obj.idx].color) ? cellOpts.border[obj.idx].color : DEF_CELL_BORDER.color) + '"/></a:solidFill>';
											var intW = (cellOpts.border[obj.idx] && (cellOpts.border[obj.idx].pt || cellOpts.border[obj.idx].pt == 0)) ? (ONEPT * Number(cellOpts.border[obj.idx].pt)) : ONEPT;
											strXml += '<a:' + obj.name + ' w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strC + '</a:' + obj.name + '>';
										}
										else strXml += '<a:' + obj.name + ' w="0"><a:miter lim="400000" /></a:' + obj.name + '>';
									});
								}
								else if (cellOpts.border && typeof cellOpts.border === 'object') {
									var intW = (cellOpts.border && (cellOpts.border.pt || cellOpts.border.pt == 0)) ? (ONEPT * Number(cellOpts.border.pt)) : ONEPT;
									var strClr = '<a:solidFill><a:srgbClr val="' + ((cellOpts.border.color) ? cellOpts.border.color.replace('#', '') : DEF_CELL_BORDER.color) + '"/></a:solidFill>';
									var strAttr = '<a:prstDash val="';
									strAttr += ((cellOpts.border.type && cellOpts.border.type.toLowerCase().indexOf('dash') > -1) ? "sysDash" : "solid");
									strAttr += '"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>';
									// *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
									strXml += '<a:lnL w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnL>';
									strXml += '<a:lnR w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnR>';
									strXml += '<a:lnT w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnT>';
									strXml += '<a:lnB w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnB>';
									// *** IMPORTANT! *** LRTB order matters!
								}

								// 6: Close cell Properties & Cell
								strXml += cellFill;
								strXml += '  </a:tcPr>';
								strXml += ' </a:tc>';

								// LAST: COLSPAN: Add a 'merged' col for each column being merged (SEE: http://officeopenxml.com/drwTableGrid.php)
								if (cellOpts.colspan) {
									for (var tmp = 1; tmp < Number(cellOpts.colspan); tmp++) { strXml += '<a:tc hMerge="1"><a:tcPr/></a:tc>'; }
								}
							});

							// D: Complete row
							strXml += '</a:tr>';
						});

						// STEP 5: Complete table
						strXml += '      </a:tbl>';
						strXml += '    </a:graphicData>';
						strXml += '  </a:graphic>';
						strXml += '</p:graphicFrame>';

						// STEP 6: Set table XML
						strSlideXml += strXml;

						// LAST: Increment counter
						intTableNum++;
						break;

					case 'text':
					case 'placeholder':
						// Lines can have zero cy, but text should not
						if (!slideItemObj.options.line && cy == 0) cy = (EMU * 0.3);

						// Margin/Padding/Inset for textboxes
						if (slideItemObj.options.margin && Array.isArray(slideItemObj.options.margin)) {
							slideItemObj.options.bodyProp.lIns = (slideItemObj.options.margin[0] * ONEPT || 0);
							slideItemObj.options.bodyProp.rIns = (slideItemObj.options.margin[1] * ONEPT || 0);
							slideItemObj.options.bodyProp.bIns = (slideItemObj.options.margin[2] * ONEPT || 0);
							slideItemObj.options.bodyProp.tIns = (slideItemObj.options.margin[3] * ONEPT || 0);
						}
						else if ((slideItemObj.options.margin || slideItemObj.options.margin == 0) && Number.isInteger(slideItemObj.options.margin)) {
							slideItemObj.options.bodyProp.lIns = (slideItemObj.options.margin * ONEPT);
							slideItemObj.options.bodyProp.rIns = (slideItemObj.options.margin * ONEPT);
							slideItemObj.options.bodyProp.bIns = (slideItemObj.options.margin * ONEPT);
							slideItemObj.options.bodyProp.tIns = (slideItemObj.options.margin * ONEPT);
						}

						var effectsList = '';
						if (shapeType == null) shapeType = getShapeInfo(null);

						// A: Start SHAPE =======================================================
						strSlideXml += '<p:sp>';

						// B: The addition of the "txBox" attribute is the sole determiner of if an object is a Shape or Textbox
						strSlideXml += '<p:nvSpPr><p:cNvPr id="' + (idx + 2) + '" name="Object ' + (idx + 1) + '"/>';
						strSlideXml += '<p:cNvSpPr' + ((slideItemObj.options && slideItemObj.options.isTextBox) ? ' txBox="1"/>' : '/>');
						strSlideXml += '<p:nvPr>';
						strSlideXml += slideItemObj.type === 'placeholder' ? genXmlPlaceholder(slideItemObj) : genXmlPlaceholder(placeholderObj);
						strSlideXml += '</p:nvPr>';
						strSlideXml += '</p:nvSpPr><p:spPr>';
						strSlideXml += '<a:xfrm' + locationAttr + '>';
						strSlideXml += '<a:off x="' + x + '" y="' + y + '"/>';
						strSlideXml += '<a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>';
						strSlideXml += '<a:prstGeom prst="' + shapeType.name + '"><a:avLst>'
							+ (slideItemObj.options.rectRadius ? '<a:gd name="adj" fmla="val ' + Math.round(slideItemObj.options.rectRadius * EMU * 100000 / Math.min(cx, cy)) + '" />' : '')
							+ '</a:avLst></a:prstGeom>';

						// Option: FILL
						strSlideXml += (slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : '<a:noFill/>');

						// Shape Type: LINE: line color
						if (slideItemObj.options.line) {
							strSlideXml += '<a:ln' + (slideItemObj.options.lineSize ? ' w="' + (slideItemObj.options.lineSize * ONEPT) + '"' : '') + '>';
							strSlideXml += genXmlColorSelection(slideItemObj.options.line);
							if (slideItemObj.options.lineDash) strSlideXml += '<a:prstDash val="' + slideItemObj.options.lineDash + '"/>';
							if (slideItemObj.options.lineHead) strSlideXml += '<a:headEnd type="' + slideItemObj.options.lineHead + '"/>';
							if (slideItemObj.options.lineTail) strSlideXml += '<a:tailEnd type="' + slideItemObj.options.lineTail + '"/>';
							strSlideXml += '</a:ln>';
						}

						// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
						if (slideItemObj.options.shadow) {
							slideItemObj.options.shadow.type = (slideItemObj.options.shadow.type || 'outer');
							slideItemObj.options.shadow.blur = (slideItemObj.options.shadow.blur || 8) * ONEPT;
							slideItemObj.options.shadow.offset = (slideItemObj.options.shadow.offset || 4) * ONEPT;
							slideItemObj.options.shadow.angle = (slideItemObj.options.shadow.angle || 270) * 60000;
							slideItemObj.options.shadow.color = (slideItemObj.options.shadow.color || '000000');
							slideItemObj.options.shadow.opacity = (slideItemObj.options.shadow.opacity || 0.75) * 100000;

							strSlideXml += '<a:effectLst>';
							strSlideXml += '<a:' + slideItemObj.options.shadow.type + 'Shdw sx="100000" sy="100000" kx="0" ky="0" ';
							strSlideXml += ' algn="bl" rotWithShape="0" blurRad="' + slideItemObj.options.shadow.blur + '" ';
							strSlideXml += ' dist="' + slideItemObj.options.shadow.offset + '" dir="' + slideItemObj.options.shadow.angle + '">';
							strSlideXml += '<a:srgbClr val="' + slideItemObj.options.shadow.color + '">';
							strSlideXml += '<a:alpha val="' + slideItemObj.options.shadow.opacity + '"/></a:srgbClr>'
							strSlideXml += '</a:outerShdw>';
							strSlideXml += '</a:effectLst>';
						}

						/* FIXME: FUTURE: Text wrapping (copied from MS-PPTX export)
						// Commented out b/c i'm not even sure this works - current code produces text that wraps in shapes and textboxes, so...
						if ( slideItemObj.options.textWrap ) {
							strSlideXml += '<a:extLst>'
										+ '<a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">'
										+ '<ma14:wrappingTextBoxFlag xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main" val="1" />'
										+ '</a:ext>'
										+ '</a:extLst>';
						}
						*/

						// B: Close Shape Properties
						strSlideXml += '</p:spPr>';

						// Add formatted text
						strSlideXml += genXmlTextBody(slideItemObj);

						// LAST: Close SHAPE =======================================================
						strSlideXml += '</p:sp>';
						break;

					case 'image':
						var sizing = slideItemObj.options.sizing,
							rounding = slideItemObj.options.rounding,
							width = cx,
							height = cy;

						strSlideXml += '<p:pic>';
						strSlideXml += '  <p:nvPicPr>'
						strSlideXml += '    <p:cNvPr id="' + (idx + 2) + '" name="Object ' + (idx + 1) + '" descr="' + encodeXmlEntities(slideItemObj.image) + '">';
						if (slideItemObj.hyperlink && slideItemObj.hyperlink.url) strSlideXml += '<a:hlinkClick r:id="rId' + slideItemObj.hyperlink.rId + '" tooltip="' + (slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : '') + '" />';
						if (slideItemObj.hyperlink && slideItemObj.hyperlink.slide) strSlideXml += '<a:hlinkClick r:id="rId' + slideItemObj.hyperlink.rId + '" tooltip="' + (slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : '') + '" action="ppaction://hlinksldjump" />';
						strSlideXml += '    </p:cNvPr>';
						strSlideXml += '    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>';
						strSlideXml += '    <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>';
						strSlideXml += '  </p:nvPicPr>';
						strSlideXml += '<p:blipFill>';
						// NOTE: This works for both cases: either `path` or `data` contains the SVG
						if (slideObject.rels.filter(function(rel) { return rel.rId == slideItemObj.imageRid })[0].extn == 'svg') {
							strSlideXml += '<a:blip r:embed="rId' + (slideItemObj.imageRid - 1) + '"/>';
							strSlideXml += '<a:extLst>';
							strSlideXml += '  <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">';
							strSlideXml += '    <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId' + slideItemObj.imageRid + '"/>';
							strSlideXml += '  </a:ext>';
							strSlideXml += '</a:extLst>';
						}
						else {
							strSlideXml += '<a:blip r:embed="rId' + slideItemObj.imageRid + '"/>';
						}
						if (sizing && sizing.type) {
							var boxW = sizing.w ? getSmartParseNumber(sizing.w, 'X') : cx,
								boxH = sizing.h ? getSmartParseNumber(sizing.h, 'Y') : cy,
								boxX = getSmartParseNumber(sizing.x || 0, 'X'),
								boxY = getSmartParseNumber(sizing.y || 0, 'Y');

							strSlideXml += gObjPptxGenerators.imageSizingXml[sizing.type]({ w: width, h: height }, { w: boxW, h: boxH, x: boxX, y: boxY });
							width = boxW;
							height = boxH;
						}
						else {
							strSlideXml += '  <a:stretch><a:fillRect/></a:stretch>';
						}
						strSlideXml += '</p:blipFill>';
						strSlideXml += '<p:spPr>'
						strSlideXml += ' <a:xfrm' + locationAttr + '>'
						strSlideXml += '  <a:off  x="' + x + '"  y="' + y + '"/>'
						strSlideXml += '  <a:ext cx="' + width + '" cy="' + height + '"/>'
						strSlideXml += ' </a:xfrm>'
						strSlideXml += ' <a:prstGeom prst="' + (rounding ? 'ellipse' : 'rect') + '"><a:avLst/></a:prstGeom>'
						strSlideXml += '</p:spPr>';
						strSlideXml += '</p:pic>';
						break;

					case 'media':
						if (slideItemObj.mtype == 'online') {
							strSlideXml += '<p:pic>';
							strSlideXml += ' <p:nvPicPr>';
							// IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preview image rId, PowerPoint throws error!
							strSlideXml += ' <p:cNvPr id="' + (slideItemObj.mediaRid + 2) + '" name="Picture' + (idx + 1) + '"/>';
							strSlideXml += ' <p:cNvPicPr/>';
							strSlideXml += ' <p:nvPr>';
							strSlideXml += '  <a:videoFile r:link="rId' + slideItemObj.mediaRid + '"/>';
							strSlideXml += ' </p:nvPr>';
							strSlideXml += ' </p:nvPicPr>';
							// NOTE: `blip` is diferent than videos; also there's no preview "p:extLst" above but exists in videos
							strSlideXml += ' <p:blipFill><a:blip r:embed="rId' + (slideItemObj.mediaRid + 1) + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'; // NOTE: Preview image is required!
							strSlideXml += ' <p:spPr>';
							strSlideXml += '  <a:xfrm' + locationAttr + '>';
							strSlideXml += '   <a:off x="' + x + '" y="' + y + '"/>';
							strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>';
							strSlideXml += '  </a:xfrm>';
							strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
							strSlideXml += ' </p:spPr>';
							strSlideXml += '</p:pic>';
						}
						else {
							strSlideXml += '<p:pic>';
							strSlideXml += ' <p:nvPicPr>';
							// IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preiew image rId, PowerPoint throws error!
							strSlideXml += ' <p:cNvPr id="' + (slideItemObj.mediaRid + 2) + '" name="' + slideItemObj.media.split('/').pop().split('.').shift() + '"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>';
							strSlideXml += ' <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>';
							strSlideXml += ' <p:nvPr>';
							strSlideXml += '  <a:videoFile r:link="rId' + slideItemObj.mediaRid + '"/>';
							strSlideXml += '  <p:extLst>';
							strSlideXml += '   <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">';
							strSlideXml += '    <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId' + (slideItemObj.mediaRid + 1) + '"/>';
							strSlideXml += '   </p:ext>';
							strSlideXml += '  </p:extLst>';
							strSlideXml += ' </p:nvPr>';
							strSlideXml += ' </p:nvPicPr>';
							strSlideXml += ' <p:blipFill><a:blip r:embed="rId' + (slideItemObj.mediaRid + 2) + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'; // NOTE: Preview image is required!
							strSlideXml += ' <p:spPr>';
							strSlideXml += '  <a:xfrm' + locationAttr + '>';
							strSlideXml += '   <a:off x="' + x + '" y="' + y + '"/>';
							strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>';
							strSlideXml += '  </a:xfrm>';
							strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
							strSlideXml += ' </p:spPr>';
							strSlideXml += '</p:pic>';
						}
						break;

					case 'chart':
						strSlideXml += '<p:graphicFrame>';
						strSlideXml += ' <p:nvGraphicFramePr>';
						strSlideXml += '   <p:cNvPr id="' + (idx + 2) + '" name="Chart ' + (idx + 1) + '"/>';
						strSlideXml += '   <p:cNvGraphicFramePr/>';
						strSlideXml += '   <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>';
						strSlideXml += ' </p:nvGraphicFramePr>';
						strSlideXml += ' <p:xfrm>'
						strSlideXml += '  <a:off  x="' + x + '"  y="' + y + '"/>'
						strSlideXml += '  <a:ext cx="' + cx + '" cy="' + cy + '"/>'
						strSlideXml += ' </p:xfrm>'
						strSlideXml += ' <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
						strSlideXml += '  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">';
						strSlideXml += '   <c:chart r:id="rId' + (slideItemObj.chartRid) + '" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>';
						strSlideXml += '  </a:graphicData>';
						strSlideXml += ' </a:graphic>';
						strSlideXml += '</p:graphicFrame>';
						break;
				}
			});

			// STEP 5: Add slide numbers last (if any)
			if (slideObject.slideNumberObj) {
				if (!slideObject.slideNumberObj) slideObject.slideNumberObj = { x: 0.3, y: '90%' }

				strSlideXml += '<p:sp>'
					+ '  <p:nvSpPr>'
					+ '    <p:cNvPr id="25" name="Slide Number Placeholder 24"/>'
					+ '    <p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr>'
					+ '    <p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr>'
					+ '  </p:nvSpPr>'
					+ '  <p:spPr>'
					+ '    <a:xfrm>'
					+ '      <a:off x="' + getSmartParseNumber(slideObject.slideNumberObj.x, 'X') + '" y="' + getSmartParseNumber(slideObject.slideNumberObj.y, 'Y') + '"/>'
					+ '      <a:ext cx="' + (slideObject.slideNumberObj.w ? getSmartParseNumber(slideObject.slideNumberObj.w, 'X') : 800000) + '" cy="' + (slideObject.slideNumberObj.h ? getSmartParseNumber(slideObject.slideNumberObj.h, 'Y') : 300000) + '"/>'
					+ '    </a:xfrm>'
					+ '    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
					+ '    <a:extLst><a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}"><ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext></a:extLst>'
					+ '  </p:spPr>';
				// ISSUE #68: "Page number styling"
				strSlideXml += '<p:txBody>';
				strSlideXml += '  <a:bodyPr/>';
				strSlideXml += '  <a:lstStyle><a:lvl1pPr>';
				if (slideObject.slideNumberObj.fontFace || slideObject.slideNumberObj.fontSize || slideObject.slideNumberObj.color) {
					strSlideXml += '<a:defRPr sz="' + (slideObject.slideNumberObj.fontSize ? Math.round(slideObject.slideNumberObj.fontSize) : '12') + '00">';
					if (slideObject.slideNumberObj.color) strSlideXml += genXmlColorSelection(slideObject.slideNumberObj.color);
					if (slideObject.slideNumberObj.fontFace) strSlideXml += '<a:latin typeface="' + slideObject.slideNumberObj.fontFace + '"/><a:ea typeface="' + slideObject.slideNumberObj.fontFace + '"/><a:cs typeface="' + slideObject.slideNumberObj.fontFace + '"/>';
					strSlideXml += '</a:defRPr>';
				}
				strSlideXml += '</a:lvl1pPr></a:lstStyle>';
				strSlideXml += '<a:p><a:fld id="' + SLDNUMFLDID + '" type="slidenum">'
					+ '<a:rPr lang="en-US" smtClean="0"/><a:t></a:t></a:fld>'
					+ '<a:endParaRPr lang="en-US"/></a:p>';
				strSlideXml += '</p:txBody></p:sp>';
			}

			// STEP 6: Close spTree and finalize slide XML
			strSlideXml += '</p:spTree>';
			strSlideXml += '</p:cSld>';

			// LAST: Return
			return strSlideXml;
		},

		/**
		 * Transforms slide relations to XML string.
		 * Extra relations that are not dynamic can be passed using the 2nd arg (e.g. theme relation in master file).
		 * These relations use rId series that starts with 1-increased maximum of rIds used for dynamic relations.
		 * @param {Object} slideObject slide object whose relations are being transformed
		 * @param {Object[]} defaultRels array of default relations (such objects expected: { target: <filepath>, type: <schemepath> })
		 * @return {String} complete XML string ready to be saved as a file
		 */
		slideObjectRelationsToXml: function slideObjectRelationsToXml(slideObject, defaultRels) {
			var lastRid = 0; // stores maximum rId used for dynamic relations
			var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
			strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
			// Add any rels for this Slide (image/audio/video/youtube/chart)
			slideObject.rels.forEach(function(rel, idx) {
				lastRid = Math.max(lastRid, rel.rId);
				if (rel.type.toLowerCase().indexOf('image') > -1) {
					strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>';
				}
				else if (rel.type.toLowerCase().indexOf('chart') > -1) {
					strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>';
				}
				else if (rel.type.toLowerCase().indexOf('audio') > -1) {
					// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
					if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
						strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>';
					else
						strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"/>';
				}
				else if (rel.type.toLowerCase().indexOf('video') > -1) {
					// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
					if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
						strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>';
					else
						strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>';
				}
				else if (rel.type.toLowerCase().indexOf('online') > -1) {
					// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
					if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
						strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.microsoft.com/office/2007/relationships/image"/>';
					else
						strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>';
				}
				else if (rel.type.toLowerCase().indexOf('hyperlink') > -1) {
					if (rel.data == 'slide') {
						strXml += '<Relationship Id="rId' + rel.rId + '" Target="slide' + rel.Target + '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>';
					}
					else {
						strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"/>';
					}
				}
				else if (rel.type.toLowerCase().indexOf('notesSlide') > -1) {
					strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"/>';
				}
			});

			defaultRels.forEach(function(rel, idx) {
				strXml += '<Relationship Id="rId' + (lastRid + idx + 1) + '" Target="' + rel.target + '" Type="' + rel.type + '"/>';
			});

			strXml += '</Relationships>';
			return strXml;
		},

		imageSizingXml: {
			cover: function(imgSize, boxDim) {
				var imgRatio = imgSize.h / imgSize.w,
					boxRatio = boxDim.h / boxDim.w,
					isBoxBased = boxRatio > imgRatio,
					width = isBoxBased ? (boxDim.h / imgRatio) : boxDim.w,
					height = isBoxBased ? boxDim.h : (boxDim.w * imgRatio),
					hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)),
					vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height));
				return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '" /><a:stretch/>';
			},
			contain: function(imgSize, boxDim) {
				var imgRatio = imgSize.h / imgSize.w,
					boxRatio = boxDim.h / boxDim.w,
					widthBased = boxRatio > imgRatio,
					width = widthBased ? boxDim.w : (boxDim.h / imgRatio),
					height = widthBased ? (boxDim.w * imgRatio) : boxDim.h,
					hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)),
					vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height));
				return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '" /><a:stretch/>';

			},
			crop: function(imageSize, boxDim) {
				var l = boxDim.x,
					r = imageSize.w - (boxDim.x + boxDim.w),
					t = boxDim.y,
					b = imageSize.h - (boxDim.y + boxDim.h),
					lPerc = Math.round(1e5 * (l / imageSize.w)),
					rPerc = Math.round(1e5 * (r / imageSize.w)),
					tPerc = Math.round(1e5 * (t / imageSize.h)),
					bPerc = Math.round(1e5 * (b / imageSize.h));
				return '<a:srcRect l="' + lPerc + '" r="' + rPerc + '" t="' + tPerc + '" b="' + bPerc + '" /><a:stretch/>';
			}
		},

		/**
		 * Based on passed data, creates Excel Worksheet that is used as a data source for a chart.
		 * @param {Object} chartObject chart object
		 * @param {ZipObject} zip zip file that the resulting XLSX should be added to
		 * @return {Promise} promise of generating the XLSX file
		 */
		createExcelWorksheet: function createExcelWorksheet(chartObject, zip) {
			var data = chartObject.data;

			return new Promise(function(resolve, reject) {
				var zipExcel = new JSZip();
				var intBubbleCols = (((data.length - 1) * 2) + 1) // 1 for "X-Values", then 2 for every Y-Axis

				// A: Add folders
				zipExcel.folder("_rels");
				zipExcel.folder("docProps");
				zipExcel.folder("xl/_rels");
				zipExcel.folder("xl/tables");
				zipExcel.folder("xl/theme");
				zipExcel.folder("xl/worksheets");
				zipExcel.folder("xl/worksheets/_rels");

				// B: Add core contents
				{
					zipExcel.file("[Content_Types].xml",
						'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
						+ '  <Default Extension="xml" ContentType="application/xml"/>'
						+ '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
						//+ '  <Default Extension="jpeg" ContentType="image/jpg"/><Default Extension="png" ContentType="image/png"/>'
						//+ '  <Default Extension="bmp" ContentType="image/bmp"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="tif" ContentType="image/tif"/><Default Extension="pdf" ContentType="application/pdf"/><Default Extension="mov" ContentType="application/movie"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
						//+ '  <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
						+ '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
						+ '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
						+ '  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
						+ '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
						+ '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
						+ '  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
						+ '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
						+ '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
						+ '</Types>\n'
					);
					zipExcel.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
						+ '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
						+ '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
						+ '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
						+ '</Relationships>\n');
					zipExcel.file("docProps/app.xml",
						'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
						+ '<Application>Microsoft Excel</Application>'
						+ '<DocSecurity>0</DocSecurity>'
						+ '<ScaleCrop>false</ScaleCrop>'
						+ '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>'
						+ '</Properties>\n'
					);
					zipExcel.file("docProps/core.xml",
						'<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
						+ '<dc:creator>PptxGenJS</dc:creator>'
						+ '<cp:lastModifiedBy>Ely, Brent</cp:lastModifiedBy>'
						+ '<dcterms:created xsi:type="dcterms:W3CDTF">' + new Date().toISOString() + '</dcterms:created>'
						+ '<dcterms:modified xsi:type="dcterms:W3CDTF">' + new Date().toISOString() + '</dcterms:modified>'
						+ '</cp:coreProperties>\n');
					zipExcel.file("xl/_rels/workbook.xml.rels",
						'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
						+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
						+ '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
						+ '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
						+ '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
						+ '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
						+ '</Relationships>\n'
					);
					zipExcel.file("xl/styles.xml",
						'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="0" formatCode="General"/></numFmts><fonts count="4"><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="10"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="18"/><color indexed="8"/>'
						+ '<name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><dxfs count="0"/><tableStyles count="0"/><colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ffff0000"/><rgbColor rgb="ff00ff00"/><rgbColor rgb="ff0000ff"/>'
						+ '<rgbColor rgb="ffffff00"/><rgbColor rgb="ffff00ff"/><rgbColor rgb="ff00ffff"/><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ff878787"/><rgbColor rgb="fff9f9f9"/></indexedColors></colors></styleSheet>\n'
					);
					zipExcel.file("xl/theme/theme1.xml",
						'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>'
					);
					zipExcel.file("xl/workbook.xml",
						'<?xml version="1.0" encoding="UTF-8"?>'
						+ '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
						+ '<fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/>'
						+ '<workbookPr />'
						+ '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="15960" windowHeight="18080"/></bookViews>'
						+ '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" /></sheets>'
						+ '<calcPr calcId="171026" concurrentCalc="0"/>'
						+ '</workbook>\n'
					);
					zipExcel.file("xl/worksheets/_rels/sheet1.xml.rels",
						'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
						+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
						+ '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'
						+ '</Relationships>\n'
					);
				}

				// sharedStrings.xml
				{
					// A: Start XML
					var strSharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
					if (chartObject.opts.type.name === 'bubble') {
						strSharedStrings += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (intBubbleCols + 1) + '" uniqueCount="' + (intBubbleCols + 1) + '">';
					}
					else if (chartObject.opts.type.name === 'scatter') {
						strSharedStrings += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (data.length + 1) + '" uniqueCount="' + (data.length + 1) + '">';
					}
					else {
						strSharedStrings += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (data[0].labels.length + data.length + 1) + '" uniqueCount="' + (data[0].labels.length + data.length + 1) + '">';
						// B: Add 'blank' for A1
						strSharedStrings += '<si><t xml:space="preserve"></t></si>';
					}

					// C: Add `name`/Series
					if (chartObject.opts.type.name === 'bubble') {
						data.forEach(function(objData, idx) {
							if (idx == 0) strSharedStrings += '<si><t>' + 'X-Axis' + '</t></si>';
							else {
								strSharedStrings += '<si><t>' + encodeXmlEntities(objData.name || ' ') + '</t></si>';
								strSharedStrings += '<si><t>' + encodeXmlEntities('Size ' + idx) + '</t></si>';
							}
						});
					}
					else {
						data.forEach(function(objData, idx) { strSharedStrings += '<si><t>' + encodeXmlEntities((objData.name || ' ').replace('X-Axis', 'X-Values')) + '</t></si>'; });
					}

					// D: Add `labels`/Categories
					if (chartObject.opts.type.name != 'bubble' && chartObject.opts.type.name != 'scatter') {
						data[0].labels.forEach(function(label, idx) { strSharedStrings += '<si><t>' + encodeXmlEntities(label) + '</t></si>'; });
					}

					strSharedStrings += '</sst>\n';
					zipExcel.file("xl/sharedStrings.xml", strSharedStrings);
				}

				// tables/table1.xml
				{
					var strTableXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
					if (chartObject.opts.type.name == 'bubble') {
						/*
						strTableXml += '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:'+ LETTERS[data.length-1] + (data[0].values.length+1) +'" totalsRowShown="0">';
						strTableXml += '<tableColumns count="' + (data.length) +'">';
						data.forEach(function(obj,idx){ strTableXml += '<tableColumn id="'+ (idx+1) +'" name="'+ (idx==0 ? 'X-Values' : 'Y-Value '+idx) +'" />' });
						*/
					}
					else if (chartObject.opts.type.name == 'scatter') {
						strTableXml += '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:' + LETTERS[data.length - 1] + (data[0].values.length + 1) + '" totalsRowShown="0">';
						strTableXml += '<tableColumns count="' + (data.length) + '">';
						data.forEach(function(obj, idx) { strTableXml += '<tableColumn id="' + (idx + 1) + '" name="' + (idx == 0 ? 'X-Values' : 'Y-Value ' + idx) + '" />' });
					}
					else {
						strTableXml += '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:' + LETTERS[data.length] + (data[0].labels.length + 1) + '" totalsRowShown="0">';
						strTableXml += '<tableColumns count="' + (data.length + 1) + '">';
						strTableXml += '<tableColumn id="1" name=" " />';
						data.forEach(function(obj, idx) { strTableXml += '<tableColumn id="' + (idx + 2) + '" name="' + encodeXmlEntities(obj.name) + '" />' });
					}
					strTableXml += '</tableColumns>';
					strTableXml += '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0" />';
					strTableXml += '</table>';
					zipExcel.file("xl/tables/table1.xml", strTableXml);
				}

				// worksheets/sheet1.xml
				{
					var strSheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
					strSheetXml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
					if (chartObject.opts.type.name === 'bubble') {
						strSheetXml += '<dimension ref="A1:' + LETTERS[(intBubbleCols - 1)] + (data[0].values.length + 1) + '" />';
					}
					else if (chartObject.opts.type.name === 'scatter') {
						strSheetXml += '<dimension ref="A1:' + LETTERS[(data.length - 1)] + (data[0].values.length + 1) + '" />';
					}
					else {
						strSheetXml += '<dimension ref="A1:' + LETTERS[data.length] + (data[0].labels.length + 1) + '" />';
					}

					strSheetXml += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1" /></sheetView></sheetViews>';
					strSheetXml += '<sheetFormatPr baseColWidth="10" defaultColWidth="11.5" defaultRowHeight="12" />';
					if (chartObject.opts.type.name == 'bubble') {
						strSheetXml += '<cols>';
						strSheetXml += '<col min="1" max="' + data.length + '" width="11" customWidth="1" />';
						strSheetXml += '</cols>';
						/* EX: INPUT: `data`
						[
							{ name:'X-Axis'  , values:[10,11,12,13,14,15,16,17,18,19,20] },
							{ name:'Y-Axis 1', values:[ 1, 6, 7, 8, 9], sizes:[ 4, 5, 6, 7, 8] },
							{ name:'Y-Axis 2', values:[33,32,42,53,63], sizes:[11,12,13,14,15] }
						];
						*/
						/* EX: OUTPUT: bubbleChart Worksheet:
							-|----A-----|------B-----|------C-----|------D-----|------E-----|
							1| X-Values | Y-Values 1 | Y-Sizes 1  | Y-Values 2 | Y-Sizes 2  |
							2|    11    |     22     |      4     |     33     |      8     |
							-|----------|------------|------------|------------|------------|
						*/
						strSheetXml += '<sheetData>';

						// A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
						strSheetXml += '<row r="1" spans="1:' + intBubbleCols + '">';
						strSheetXml += '<c r="A1" t="s"><v>0</v></c>';
						for (var idx = 1; idx < intBubbleCols; idx++) {
							strSheetXml += '<c r="' + (idx < 26 ? LETTERS[idx] : 'A' + LETTERS[idx % LETTERS.length]) + '1" t="s">'; // NOTE: use `t="s"` for label cols!
							strSheetXml += '<v>' + idx + '</v>';
							strSheetXml += '</c>';
						}
						strSheetXml += '</row>';

						// B: Add row for each X-Axis value (Y-Axis* value is optional)
						data[0].values.forEach(function(val, idx) {
							// Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
							strSheetXml += '<row r="' + (idx + 2) + '" spans="1:' + intBubbleCols + '">';
							strSheetXml += '<c r="A' + (idx + 2) + '"><v>' + val + '</v></c>';
							// Add Y-Axis 1->N (idy=0 = Xaxis)
							var idxColLtr = 1;
							for (var idy = 1; idy < data.length; idy++) {
								// y-value
								strSheetXml += '<c r="' + (idxColLtr < 26 ? LETTERS[idxColLtr] : 'A' + LETTERS[idxColLtr % LETTERS.length]) + '' + (idx + 2) + '">';
								strSheetXml += '<v>' + (data[idy].values[idx] || '') + '</v>';
								strSheetXml += '</c>';
								idxColLtr++;
								// y-size
								strSheetXml += '<c r="' + (idxColLtr < 26 ? LETTERS[idxColLtr] : 'A' + LETTERS[idxColLtr % LETTERS.length]) + '' + (idx + 2) + '">';
								strSheetXml += '<v>' + (data[idy].sizes[idx] || '') + '</v>';
								strSheetXml += '</c>';
								idxColLtr++;
							};
							strSheetXml += '</row>';
						});
					}
					else if (chartObject.opts.type.name == 'scatter') {
						strSheetXml += '<cols>';
						strSheetXml += '<col min="1" max="' + data.length + '" width="11" customWidth="1" />';
						//data.forEach(function(obj,idx){ strSheetXml += '<col min="'+(idx+1)+'" max="'+(idx+1)+'" width="11" customWidth="1" />' });
						strSheetXml += '</cols>';
						/* EX: INPUT: `data`
						[
							{ name:'X-Axis'  , values:[10,11,12,13,14,15,16,17,18,19,20] },
							{ name:'Y-Axis 1', values:[ 1, 6, 7, 8, 9] },
							{ name:'Y-Axis 2', values:[33,32,42,53,63] }
						];
						*/
						/* EX: OUTPUT: scatterChart Worksheet:
							-|----A-----|------B-----|
							1| X-Values | Y-Values 1 |
							2|    11    |     22     |
							-|----------|------------|
						*/
						strSheetXml += '<sheetData>';

						// A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
						strSheetXml += '<row r="1" spans="1:' + data.length + '">';
						strSheetXml += '<c r="A1" t="s"><v>0</v></c>';
						for (var idx = 1; idx < data.length; idx++) {
							strSheetXml += '<c r="' + (idx < 26 ? LETTERS[idx] : 'A' + LETTERS[idx % LETTERS.length]) + '1" t="s">'; // NOTE: use `t="s"` for label cols!
							strSheetXml += '<v>' + idx + '</v>';
							strSheetXml += '</c>';
						}
						strSheetXml += '</row>';

						// B: Add row for each X-Axis value (Y-Axis* value is optional)
						data[0].values.forEach(function(val, idx) {
							// Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
							strSheetXml += '<row r="' + (idx + 2) + '" spans="1:' + data.length + '">';
							strSheetXml += '<c r="A' + (idx + 2) + '"><v>' + val + '</v></c>';
							// Add Y-Axis 1->N
							for (var idy = 1; idy < data.length; idy++) {
								strSheetXml += '<c r="' + (idy < 26 ? LETTERS[idy] : 'A' + LETTERS[idy % LETTERS.length]) + '' + (idx + 2) + '">';
								strSheetXml += '<v>' + (data[idy].values[idx] || data[idy].values[idx] == 0 ? data[idy].values[idx] : '') + '</v>';
								strSheetXml += '</c>';
							};
							strSheetXml += '</row>';
						});
					}
					else {
						strSheetXml += '<cols>';
						strSheetXml += '<col min="1" max="1" width="11" customWidth="1" />';
						//data.forEach(function(){ strSheetXml += '<col min="10" max="100" width="10" customWidth="1" />' });
						strSheetXml += '</cols>';
						strSheetXml += '<sheetData>';

						/* EX: INPUT: `data`
						[
							{ name:'Red', labels:['Jan..May-17'], values:[11,13,14,15,16] },
							{ name:'Amb', labels:['Jan..May-17'], values:[22, 6, 7, 8, 9] },
							{ name:'Grn', labels:['Jan..May-17'], values:[33,32,42,53,63] }
						];
						*/
						/* EX: OUTPUT: lineChart Worksheet:
							-|---A---|--B--|--C--|--D--|
							1|       | Red | Amb | Grn |
							2|Jan-17 |   11|   22|   33|
							3|Feb-17 |   55|   43|   70|
							4|Mar-17 |   56|  143|   99|
							5|Apr-17 |   65|    3|  120|
							6|May-17 |   75|   93|  170|
							-|-------|-----|-----|-----|
						*/

						// A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
						strSheetXml += '<row r="1" spans="1:' + (data.length + 1) + '">';
						strSheetXml += '<c r="A1" t="s"><v>0</v></c>';
						for (var idx = 1; idx <= data.length; idx++) {
							// FIXME: Max cols is 52
							strSheetXml += '<c r="' + (idx < 26 ? LETTERS[idx] : 'A' + LETTERS[idx % LETTERS.length]) + '1" t="s">'; // NOTE: use `t="s"` for label cols!
							strSheetXml += '<v>' + idx + '</v>';
							strSheetXml += '</c>';
						}
						strSheetXml += '</row>';

						// B: Add data row(s) for each category
						data[0].labels.forEach(function(cat, idx) {
							// Leading col is reserved for the label, so hard-code it, then loop over col values
							strSheetXml += '<row r="' + (idx + 2) + '" spans="1:' + (data.length + 1) + '">';
							strSheetXml += '<c r="A' + (idx + 2) + '" t="s">';
							strSheetXml += '<v>' + (data.length + idx + 1) + '</v>';
							strSheetXml += '</c>';
							for (var idy = 0; idy < data.length; idy++) {
								strSheetXml += '<c r="' + ((idy + 1) < 26 ? LETTERS[(idy + 1)] : 'A' + LETTERS[(idy + 1) % LETTERS.length]) + '' + (idx + 2) + '">';
								strSheetXml += '<v>' + (data[idy].values[idx] || '') + '</v>';
								strSheetXml += '</c>';
							}
							strSheetXml += '</row>';
						});
					}
					strSheetXml += '</sheetData>';
					strSheetXml += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />';
					// Link the `table1.xml` file to define an actual Table in Excel
					// NOTE: This onyl works with scatter charts - all others give a "cannot find linked file" error
					// ....: Since we dont need the table anyway (chart data can be edited/range selected, etc.), just dont use this
					// ....: Leaving this so nobody foolishly attempts to add this in the future
					// strSheetXml += '<tableParts count="1"><tablePart r:id="rId1" /></tableParts>';
					strSheetXml += '</worksheet>\n';
					zipExcel.file("xl/worksheets/sheet1.xml", strSheetXml);
				}

				// C: Add XLSX to PPTX export
				zipExcel.generateAsync({ type: 'base64' })
					.then(function(content) {
						// 1: Create the embedded Excel worksheet with labels and data
						zip.file("ppt/embeddings/Microsoft_Excel_Worksheet" + chartObject.globalId + ".xlsx", content, { base64: true });

						// 2: Create the chart.xml and rels files
						zip.file("ppt/charts/_rels/" + chartObject.fileName + ".rels",
							'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
							+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
							+ '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet' + chartObject.globalId + '.xlsx"/>'
							+ '</Relationships>'
						);
						zip.file("ppt/charts/" + chartObject.fileName, makeXmlCharts(chartObject));

						// 3: Done
						resolve();
					})
					.catch(function(strErr) {
						reject(strErr);
					});
			});
		}
	};

	/* ===============================================================================================
	|
	#     #
	#     #  ######  #       #####   ######  #####    ####
	#     #  #       #       #    #  #       #    #  #
	#######  #####   #       #    #  #####   #    #   ####
	#     #  #       #       #####   #       #####        #
	#     #  #       #       #       #       #   #   #    #
	#     #  ######  ######  #       ######  #    #   ####
	|
	==================================================================================================
	*/

	/**
	 * DESC: Export the .pptx file
	 */
	function doExportPresentation(outputType) {
		var arrChartPromises = [];
		var intSlideNum = 0, intRels = 0, intNotesRels = 0;

		// STEP 1: Create new JSZip file
		var zip = new JSZip();

		// STEP 2: Add all required folders and files
		zip.folder("_rels");
		zip.folder("docProps");
		zip.folder("ppt").folder("_rels");
		zip.folder("ppt/charts").folder("_rels");
		zip.folder("ppt/embeddings");
		zip.folder("ppt/media");
		zip.folder("ppt/slideLayouts").folder("_rels");
		zip.folder("ppt/slideMasters").folder("_rels");
		zip.folder("ppt/slides").folder("_rels");
		zip.folder("ppt/theme");
		zip.folder("ppt/notesMasters").folder("_rels");
		zip.folder("ppt/notesSlides").folder("_rels");
		//
		zip.file("[Content_Types].xml", makeXmlContTypes());
		zip.file("_rels/.rels", makeXmlRootRels());
		zip.file("docProps/app.xml", makeXmlApp());
		zip.file("docProps/core.xml", makeXmlCore());
		zip.file("ppt/_rels/presentation.xml.rels", makeXmlPresentationRels());
		//
		zip.file("ppt/theme/theme1.xml", makeXmlTheme());
		zip.file("ppt/presentation.xml", makeXmlPresentation());
		zip.file("ppt/presProps.xml", makeXmlPresProps());
		zip.file("ppt/tableStyles.xml", makeXmlTableStyles());
		zip.file("ppt/viewProps.xml", makeXmlViewProps());

		// Create a Layout/Master/Rel/Slide file for each SLIDE
		for (var idx = 1; idx <= gObjPptx.slideLayouts.length; idx++) {
			zip.file("ppt/slideLayouts/slideLayout" + idx + ".xml", makeXmlLayout(gObjPptx.slideLayouts[idx - 1]));
			zip.file("ppt/slideLayouts/_rels/slideLayout" + idx + ".xml.rels", makeXmlSlideLayoutRel(idx));
		}

		for (var idx = 0; idx < gObjPptx.slides.length; idx++) {
			intSlideNum++;
			zip.file('ppt/slides/slide' + intSlideNum + '.xml', makeXmlSlide(gObjPptx.slides[idx]));
			zip.file('ppt/slides/_rels/slide' + intSlideNum + '.xml.rels', makeXmlSlideRel(intSlideNum));

			// Here we will create all slide notes related items. Notes of empty strings
			// are created for slides which do not have notes specified, to keep track of _rels.
			zip.file('ppt/notesSlides/notesSlide' + intSlideNum + '.xml', makeXmlNotesSlide(gObjPptx.slides[idx]));
			zip.file('ppt/notesSlides/_rels/notesSlide' + intSlideNum + '.xml.rels', makeXmlNotesSlideRel(intSlideNum));
		}

		zip.file("ppt/slideMasters/slideMaster1.xml", makeXmlMaster(gObjPptx.masterSlide));
		zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", makeXmlMasterRel(gObjPptx.masterSlide));
		zip.file('ppt/notesMasters/notesMaster1.xml', makeXmlNotesMaster());
		zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', makeXmlNotesMasterRel());

		// Create all Rels (images, media, chart data)
		gObjPptx.slideLayouts.forEach(function(layout) { createMediaFiles(layout, zip, arrChartPromises); });
		gObjPptx.slides.forEach(function(slide) { createMediaFiles(slide, zip, arrChartPromises); });
		createMediaFiles(gObjPptx.masterSlide, zip, arrChartPromises);

		// STEP 3: Wait for Promises (if any) then generate the PPTX file
		Promise.all(arrChartPromises)
			.then(function(arrResults) {
				var strExportName = ((gObjPptx.fileName.toLowerCase().indexOf('.ppt') > -1) ? gObjPptx.fileName : gObjPptx.fileName + gObjPptx.fileExtn);
				if (outputType && JSZIP_OUTPUT_TYPES.indexOf(outputType) >= 0) {
					zip.generateAsync({ type: outputType }).then(gObjPptx.saveCallback);
				}
				else if (NODEJS && !gObjPptx.isBrowser) {
					if (gObjPptx.saveCallback) {
						if (strExportName.indexOf('http') == 0) {
							zip.generateAsync({ type: 'nodebuffer' }).then(function(content) { gObjPptx.saveCallback(content); });
						}
						else {
							zip.generateAsync({ type: 'nodebuffer' }).then(function(content) { fs.writeFile(strExportName, content, function() { gObjPptx.saveCallback(strExportName); }); });
						}
					}
					else {
						// Starting in late 2017 (Node ~8.9.1), `fs` requires a callback so use a dummy func
						zip.generateAsync({ type: 'nodebuffer' }).then(function(content) { fs.writeFile(strExportName, content, function() { }); });
					}
				}
				else {
					zip.generateAsync({ type: 'blob' }).then(function(content) { writeFileToBrowser(strExportName, content); });
				}
			})
			.catch(function(strErr) {
				console.error(strErr);
			});
	}

	function writeFileToBrowser(strExportName, content) {
		// STEP 1: Create element
		var a = document.createElement("a");
		document.body.appendChild(a);
		a.style = "display: none";

		// STEP 2: Download file to browser
		// DESIGN: Use `createObjectURL()` (or MS-specific func for IE11) to D/L files in client browsers (FYI: synchronously executed)
		if (window.navigator.msSaveOrOpenBlob) {
			// REF: https://docs.microsoft.com/en-us/microsoft-edge/dev-guide/html5/file-api/blob
			let blobObject = new Blob([content]);
			jQuery(a).click(function() {
				window.navigator.msSaveOrOpenBlob(blobObject, strExportName);
			});
			a.click();

			// Clean-up
			document.body.removeChild(a);

			// LAST: Callback (if any)
			if (gObjPptx.saveCallback) gObjPptx.saveCallback(strExportName);
		}
		else if (window.URL.createObjectURL) {
			var blob = new Blob([content], { type: "octet/stream" });
			var url = window.URL.createObjectURL(blob);
			a.href = url;
			a.download = strExportName;
			a.click();

			// Clean-up (NOTE: Add a slight delay before removing to avoid 'blob:null' error in Firefox Issue#81)
			setTimeout(function() {
				window.URL.revokeObjectURL(url);
				document.body.removeChild(a);
			}, 100);

			// LAST: Callback (if any)
			if (gObjPptx.saveCallback) gObjPptx.saveCallback(strExportName);
		}

		// STEP 3: Clear callback func post-save
		gObjPptx.saveCallback = null;
	}

	function createMediaFiles(layout, zip, chartPromises) {
		layout.rels.forEach(function(rel) {
			if (rel.type == 'chart') {
				chartPromises.push(gObjPptxGenerators.createExcelWorksheet(rel, zip));
			}
			else if (rel.type != 'online' && rel.type != 'hyperlink') {
				// A: Loop vars
				var data = rel.data;

				// B: Users will undoubtedly pass various string formats, so correct prefixes as needed
				if (data.indexOf(',') == -1 && data.indexOf(';') == -1) data = 'image/png;base64,' + data;
				else if (data.indexOf(',') == -1) data = 'image/png;base64,' + data;
				else if (data.indexOf(';') == -1) data = 'image/png;' + data;

				// C: Add media
				zip.file(rel.Target.replace('..', 'ppt'), data.split(',').pop(), { base64: true });
			}
		});
	}

	/**
	 * DESC: Convert component value to hex value
	 */
	function componentToHex(c) {
		var hex = c.toString(16);
		return hex.length == 1 ? "0" + hex : hex;
	}

	/**
	 * DESC: Used by `addSlidesForTable()` to convert RGB colors from jQuery selectors to Hex for Presentation colors
	 */
	function rgbToHex(r, g, b) {
		if (!Number.isInteger(r)) { try { console.warn('Integer expected!'); } catch (ex) { } }
		return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase();
	}

	function inch2Emu(inches) {
		// FIRST: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
		// Any value over 100 damn sure isnt inches, must be EMU already, so just return it
		if (inches > 100) return inches;
		if (typeof inches == 'string') inches = Number(inches.replace(/in*/gi, ''));
		return Math.round(EMU * inches);
	}

	function addPlaceholdersToSlides(slide) {
		// Add all placeholders on this Slide that dont already exist
		slide.layoutObj.data.forEach(function(slideLayoutObj) {
			if (slideLayoutObj.type === MASTER_OBJECTS.placeholder.name) {
				// A: Search for this placeholder on Slide before we add
				// NOTE: Check to ensure a placeholder does not already exist on the Slide
				// They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
				if (slide.data.filter(function(slideObj) { return slideObj.options && slideObj.options.placeholder == slideLayoutObj.options.placeholderName }).length == 0) {
					gObjPptxGenerators.addTextDefinition('', { placeholder: slideLayoutObj.options.placeholderName }, slide, false);
				}
			}
		});
	}

	// IMAGE METHODS:

	function getSizeFromImage(inImgUrl) {
		if (NODEJS) {
			try {
				var dimensions = sizeOf(inImgUrl);
				return { width: dimensions.width, height: dimensions.height };
			}
			catch (ex) {
				console.error('ERROR: Unable to read image: ' + inImgUrl);
				return { width: 0, height: 0 };
			}
		}

		// A: Create
		var image = new Image();

		// B: Set onload event
		image.onload = () => {
			// FIRST: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (this.width + this.height == 0) { return { width: 0, height: 0 }; }
			var obj = { width: this.width, height: this.height };
			return obj;
		};
		image.onerror = () => {
			try { console.error('[Error] Unable to load image: ' + inImgUrl); } catch (ex) { }
		};

		// C: Load image
		image.src = inImgUrl;
	}

	/* Encode Image/Audio/Video into base64 */
	function encodeSlideMediaRels(layout, arrRelsDone) {
		var intRels = 0;

		layout.rels.forEach(function(rel) {
			// Read and Encode each media lacking `data` into base64 (for use in export)
			if (rel.type != 'online' && rel.type != 'chart' && !rel.data && arrRelsDone.indexOf(rel.path) == -1) {
				// Node local-file encoding is syncronous, so we can load all images here, then call export with a callback (if any)
				if (NODEJS && rel.path.indexOf('http') != 0) {
					try {
						var bitmap = fs.readFileSync(rel.path);
						rel.data = Buffer.from(bitmap).toString('base64');
					}
					catch (ex) {
						console.error('ERROR....: Unable to read media: "' + rel.path + '"');
						console.error('DETAILS..: ' + ex);
						rel.data = IMG_BROKEN;
					}
				}
				else if (NODEJS && rel.path.indexOf('http') == 0) {
					intRels++;
					convertRemoteMediaToDataURL(rel);
					arrRelsDone.push(rel.path);
				}
				else {
					intRels++;
					convertImgToDataURL(rel);
					arrRelsDone.push(rel.path);
				}
			}
			else if (rel.isSvgPng && rel.data && rel.data.toLowerCase().indexOf('image/svg') > -1) {
				// The SVG base64 must be converted to PNG SVG before export
				intRels++;
				callbackImgToDataURLDone(rel.data, rel);
				arrRelsDone.push(rel.path);
			}
		});

		return intRels;
	}

	/* `FileReader()` + `readAsDataURL` = Ablity to read any file into base64! */
	function convertImgToDataURL(slideRel) {
		var xhr = new XMLHttpRequest();
		xhr.onload = function() {
			var reader = new FileReader();
			reader.onloadend = function() { callbackImgToDataURLDone(reader.result, slideRel); }
			reader.readAsDataURL(xhr.response);
		};
		xhr.onerror = function(ex) {
			// TODO: xhr.error/catch whatever! then return
			console.error('Unable to load image: "' + slideRel.path);
			console.error(ex || '');
			// Return a predefined "Broken image" graphic so the user will see something on the slide
			callbackImgToDataURLDone(IMG_BROKEN, slideRel);
		};
		xhr.open('GET', slideRel.path);
		xhr.responseType = 'blob';
		xhr.send();
	}

	/* Node equivalent of `convertImgToDataURL()`: Use https to fetch, then use Buffer to encode to base64 */
	function convertRemoteMediaToDataURL(slideRel) {
		https.get(slideRel.path, function(res) {
			var rawData = "";
			res.setEncoding('binary'); // IMPORTANT: Only binary encoding works
			res.on("data", function(chunk) { rawData += chunk; });
			res.on("end", function() {
				var data = Buffer.from(rawData, 'binary').toString('base64');
				callbackImgToDataURLDone(data, slideRel);
			});
			res.on("error", function(e) {
				reject(e);
			});
		});
	}

	/* browser: Convert SVG-base64 data to PNG-base64 */
	function convertSvgToPngViaCanvas(slideRel) {
		// A: Create
		var image = new Image();

		// B: Set onload event
		image.onload = function() {
			// First: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (this.width + this.height == 0) { this.onerror('h/w=0'); return; }
			var canvas = document.createElement('CANVAS');
			var ctx = canvas.getContext('2d');
			canvas.width = this.width;
			canvas.height = this.height;
			ctx.drawImage(this, 0, 0);
			// Users running on local machine will get the following error:
			// "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
			// when the canvas.toDataURL call executes below.
			try { callbackImgToDataURLDone(canvas.toDataURL(slideRel.type), slideRel); }
			catch (ex) {
				this.onerror(ex);
				return;
			}
			canvas = null;
		};
		image.onerror = function(ex) {
			console.error(ex || '');
			// Return a predefined "Broken image" graphic so the user will see something on the slide
			callbackImgToDataURLDone(IMG_BROKEN, slideRel);
		};

		// C: Load image
		image.src = slideRel.data; // use pre-encoded SVG base64 data
	}

	function callbackImgToDataURLDone(base64Data, slideRel) {
		// SVG images were retrieved via `convertImgToDataURL()`, but have to be encoded to PNG now
		if (slideRel.isSvgPng && base64Data.indexOf('image/svg') > -1) {
			// Pass the SVG XML as base64 for conversion to PNG
			slideRel.data = base64Data;
			if (NODEJS) console.log('SVG is not supported in Node');
			else convertSvgToPngViaCanvas(slideRel);
			return;
		}

		var intEmpty = 0;
		var funcCallback = function(rel) {
			if (rel.path == slideRel.path) rel.data = base64Data;
			if (!rel.data) intEmpty++;
		}

		// STEP 1: Set data for this rel, count outstanding
		gObjPptx.slides.forEach(function(slide) { slide.rels.forEach(funcCallback); });
		gObjPptx.slideLayouts.forEach(function(layout) { layout.rels.forEach(funcCallback); });
		gObjPptx.masterSlide.rels.forEach(funcCallback);

		// STEP 2: Continue export process if all rels have base64 `data` now
		if (intEmpty == 0) doExportPresentation();
	}

	function getShapeInfo(shapeName) {
		if (!shapeName) return gObjPptxShapes.RECTANGLE;

		if (typeof shapeName == 'object' && shapeName.name && shapeName.displayName && shapeName.avLst) return shapeName;

		if (gObjPptxShapes[shapeName]) return gObjPptxShapes[shapeName];

		var objShape = gObjPptxShapes.filter(function(obj) { return obj.name == shapeName || obj.displayName; })[0];
		if (typeof objShape !== 'undefined' && objShape != null) return objShape;

		return gObjPptxShapes.RECTANGLE;
	}

	function getSmartParseNumber(inVal, inDir) {
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

	function createHyperlinkRels(inText, slideRels) {
		var arrTextObjects = [];

		// Only text objects can have hyperlinks, so return if this is plain text/number
		if (typeof inText === 'string' || typeof inText === 'number') return;
		// IMPORTANT: Check for isArray before typeof=object, or we'll exhaust recursion!
		else if (Array.isArray(inText)) arrTextObjects = inText;
		else if (typeof inText === 'object') arrTextObjects = [inText];

		arrTextObjects.forEach(function(text, idx) {
			// `text` can be an array of other `text` objects (table cell word-level formatting), so use recursion
			if (Array.isArray(text)) createHyperlinkRels(text, slideRels);
			else if (text && typeof text === 'object' && text.options && text.options.hyperlink && !text.options.hyperlink.rId) {
				if (typeof text.options.hyperlink !== 'object') console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:'https://github.com'}` ");
				else if (!text.options.hyperlink.url && !text.options.hyperlink.slide) console.log("ERROR: 'hyperlink requires either: `url` or `slide`'");
				else {
					var intRels = 0;
					gObjPptx.slides.forEach(function(slide, idx) { intRels += slide.rels.length; });
					var intRelId = intRels + 1;

					slideRels.push({
						type: 'hyperlink',
						data: (text.options.hyperlink.slide ? 'slide' : 'dummy'),
						rId: intRelId,
						Target: text.options.hyperlink.url || text.options.hyperlink.slide
					});

					text.options.hyperlink.rId = intRelId;
				}
			}
		});
	}

	/**
	* Magic happens here
	*/
	function parseTextToLines(cell, inWidth) {
		var CHAR = 2.2 + (cell.opts && cell.opts.lineWeight ? cell.opts.lineWeight : 0); // Character Constant (An approximation of the Golden Ratio)
		var CPL = (inWidth * EMU / ((cell.opts.fontSize || DEF_FONT_SIZE) / CHAR)); // Chars-Per-Line
		var arrLines = [];
		var strCurrLine = '';

		// Allow a single space/whitespace as cell text
		if (cell.text && cell.text.trim() == '') return [' '];

		// A: Remove leading/trailing space
		var inStr = (cell.text || '').toString().trim();

		// B: Build line array
		jQuery.each(inStr.split('\n'), function(i, line) {
			jQuery.each(line.split(' '), function(i, word) {
				if (strCurrLine.length + word.length + 1 < CPL) {
					strCurrLine += (word + " ");
				}
				else {
					if (strCurrLine) arrLines.push(strCurrLine);
					strCurrLine = (word + " ");
				}
			});
			// All words for this line have been exhausted, flush buffer to new line, clear line var
			if (strCurrLine) arrLines.push(jQuery.trim(strCurrLine) + CRLF);
			strCurrLine = '';
		});

		// C: Remove trailing linebreak
		arrLines[(arrLines.length - 1)] = jQuery.trim(arrLines[(arrLines.length - 1)]);

		// D: Return lines
		return arrLines;
	}

	/**
	* Magic happens here
	*/
	function getSlidesForTableRows(inArrRows, opts) {
		var LINEH_MODIFIER = 1.9;
		var opts = opts || {};
		var arrInchMargins = DEF_SLIDE_MARGIN_IN; // (0.5" on all sides)
		var arrObjTabHeadRows = [], arrObjTabBodyRows = [], arrObjTabFootRows = [];
		var arrObjSlides = [], arrRows = [], currRow = [];
		var intTabW = 0, emuTabCurrH = 0;
		var emuSlideTabW = EMU * 1, emuSlideTabH = EMU * 1;
		var arrObjTabHeadRows = opts.arrObjTabHeadRows || '';
		var numCols = 0;

		if (opts.debug) console.log('------------------------------------');
		if (opts.debug) console.log('opts.w ............. = ' + (opts.w || '').toString());
		if (opts.debug) console.log('opts.colW .......... = ' + (opts.colW || '').toString());
		if (opts.debug) console.log('opts.slideMargin ... = ' + (opts.slideMargin || '').toString());

		// NOTE: Use default size as zero cell margin is causing our tables to be too large and touch bottom of slide!
		if (!opts.slideMargin && opts.slideMargin != 0) opts.slideMargin = DEF_SLIDE_MARGIN_IN[0];

		// STEP 1: Calc margins/usable space
		if (opts.slideMargin || opts.slideMargin == 0) {
			if (Array.isArray(opts.slideMargin)) arrInchMargins = opts.slideMargin;
			else if (!isNaN(opts.slideMargin)) arrInchMargins = [opts.slideMargin, opts.slideMargin, opts.slideMargin, opts.slideMargin];
		}
		else if (opts && opts.master && opts.master.margin) {
			if (Array.isArray(opts.master.margin)) arrInchMargins = opts.master.margin;
			else if (!isNaN(opts.master.margin)) arrInchMargins = [opts.master.margin, opts.master.margin, opts.master.margin, opts.master.margin];
		}

		// STEP 2: Calc number of columns
		// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
		// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
		inArrRows[0].forEach(function(cell, idx) {
			if (!cell) cell = {};
			var cellOpts = cell.options || cell.opts || null;
			numCols += (cellOpts && cellOpts.colspan ? cellOpts.colspan : 1);
		});

		if (opts.debug) console.log('arrInchMargins ..... = ' + arrInchMargins.toString());
		if (opts.debug) console.log('numCols ............ = ' + numCols);

		// Calc opts.w if we can
		if (!opts.w && opts.colW) {
			if (Array.isArray(opts.colW)) opts.colW.forEach(function(val, idx) { opts.w += val });
			else { opts.w = opts.colW * numCols }
		}

		// STEP 2: Calc usable space/table size now that we have usable space calc'd
		emuSlideTabW = (opts.w ? inch2Emu(opts.w) : (gObjPptx.pptLayout.width - inch2Emu((opts.x || arrInchMargins[1]) + arrInchMargins[3])));
		if (opts.debug) console.log('emuSlideTabW (in) ........ = ' + (emuSlideTabW / EMU).toFixed(1));
		if (opts.debug) console.log('gObjPptx.pptLayout.h ..... = ' + (gObjPptx.pptLayout.height / EMU));

		// STEP 3: Calc column widths if needed so we can subsequently calc lines (we need `emuSlideTabW`!)
		if (!opts.colW || !Array.isArray(opts.colW)) {
			if (opts.colW && !isNaN(Number(opts.colW))) {
				var arrColW = [];
				inArrRows[0].forEach(function(cell, idx) { arrColW.push(opts.colW) });
				opts.colW = [];
				arrColW.forEach(function(val, idx) { opts.colW.push(val) });
			}
			// No column widths provided? Then distribute cols.
			else {
				opts.colW = [];
				for (var iCol = 0; iCol < numCols; iCol++) { opts.colW.push((emuSlideTabW / EMU / numCols)); }
			}
		}

		// STEP 4: Iterate over each line and perform magic =========================
		// NOTE: inArrRows will be an array of {text:'', opts{}} whether from `addSlidesForTable()` or `.addTable()`
		inArrRows.forEach(function(row, iRow) {
			// A: Reset ROW variables
			var arrCellsLines = [], arrCellsLineHeights = [], emuRowH = 0, intMaxLineCnt = 0, intMaxColIdx = 0;

			// B: Calc usable vertical space/table height
			// NOTE: Use margins after the first Slide (dont re-use opt.y - it could've been halfway down the page!) (ISSUE#43,ISSUE#47,ISSUE#48)
			if (arrObjSlides.length > 0) {
				emuSlideTabH = (gObjPptx.pptLayout.height - inch2Emu((opts.y / EMU < arrInchMargins[0] ? opts.y / EMU : arrInchMargins[0]) + arrInchMargins[2]));
				// Use whichever is greater: area between margins or the table H provided (dont shrink usable area - the whole point of over-riding X on paging is to *increarse* usable space)
				if (emuSlideTabH < opts.h) emuSlideTabH = opts.h;
			}
			else emuSlideTabH = (opts.h ? opts.h : (gObjPptx.pptLayout.height - inch2Emu((opts.y / EMU || arrInchMargins[0]) + arrInchMargins[2])));
			if (opts.debug) console.log('* Slide ' + arrObjSlides.length + ': emuSlideTabH (in) ........ = ' + (emuSlideTabH / EMU).toFixed(1));

			// C: Parse and store each cell's text into line array (**MAGIC HAPPENS HERE**)
			row.forEach(function(cell, iCell) {
				// FIRST: REALITY-CHECK:
				if (!cell) cell = {};

				// DESIGN: Cells are henceforth {objects} with `text` and `opts`
				var lines = [];

				// 1: Cleanse data
				if (!isNaN(cell) || typeof cell === 'string') {
					// Grab table formatting `opts` to use here so text style/format inherits as it should
					cell = { text: cell.toString(), opts: opts };
				}
				else if (typeof cell === 'object') {
					// ARG0: `text`
					if (typeof cell.text === 'number') cell.text = cell.text.toString();
					else if (typeof cell.text === 'undefined' || cell.text == null) cell.text = "";

					// ARG1: `options`
					var opt = cell.options || cell.opts || {};
					cell.opts = opt;
				}
				// Capture some table options for use in other functions
				cell.opts.lineWeight = opts.lineWeight;

				// 2: Create a cell object for each table column
				currRow.push({ text: '', opts: cell.opts });

				// 3: Parse cell contents into lines (**MAGIC HAPPENSS HERE**)
				var lines = parseTextToLines(cell, (opts.colW[iCell] / ONEPT));
				arrCellsLines.push(lines);
				//if (opts.debug) console.log('Cell:'+iCell+' - lines:'+lines.length);

				// 4: Keep track of max line count within all row cells
				if (lines.length > intMaxLineCnt) { intMaxLineCnt = lines.length; intMaxColIdx = iCell; }
				var lineHeight = inch2Emu((cell.opts.fontSize || opts.fontSize || DEF_FONT_SIZE) * LINEH_MODIFIER / 100);
				// NOTE: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
				if (cell.opts && cell.opts.rowspan) lineHeight = 0;

				// 5: Add cell margins to lineHeight (if any)
				if (cell.opts.margin) {
					if (cell.opts.margin[0]) lineHeight += (cell.opts.margin[0] * ONEPT) / intMaxLineCnt;
					if (cell.opts.margin[2]) lineHeight += (cell.opts.margin[2] * ONEPT) / intMaxLineCnt;
				}

				// Add to array
				arrCellsLineHeights.push(Math.round(lineHeight));
			});

			// D: AUTO-PAGING: Add text one-line-a-time to this row's cells until: lines are exhausted OR table H limit is hit
			for (var idx = 0; idx < intMaxLineCnt; idx++) {
				// 1: Add the current line to cell
				for (var col = 0; col < arrCellsLines.length; col++) {
					// A: Commit this slide to Presenation if table Height limit is hit
					if (emuTabCurrH + arrCellsLineHeights[intMaxColIdx] > emuSlideTabH) {
						if (opts.debug) console.log('--------------- New Slide Created ---------------');
						if (opts.debug) console.log(' (calc) ' + (emuTabCurrH / EMU).toFixed(1) + '+' + (arrCellsLineHeights[intMaxColIdx] / EMU).toFixed(1) + ' > ' + emuSlideTabH / EMU.toFixed(1));
						if (opts.debug) console.log('--------------- New Slide Created ---------------');
						// 1: Add the current row to table
						// NOTE: Edge cases can occur where we create a new slide only to have no more lines
						// ....: and then a blank row sits at the bottom of a table!
						// ....: Hence, we verify all cells have text before adding this final row.
						jQuery.each(currRow, function(i, cell) {
							if (cell.text.length > 0) {
								// IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
								arrRows.push(jQuery.extend(true, [], currRow));
								return false; // break out of .each loop
							}
						});
						// 2: Add new Slide with current array of table rows
						arrObjSlides.push(jQuery.extend(true, [], arrRows));
						// 3: Empty rows for new Slide
						arrRows.length = 0;
						// 4: Reset current table height for new Slide
						emuTabCurrH = 0; // This row's emuRowH w/b added below
						// 5: Empty current row's text (continue adding lines where we left off below)
						jQuery.each(currRow, function(i, cell) { cell.text = ''; });
						// 6: Auto-Paging Options: addHeaderToEach
						if (opts.addHeaderToEach && arrObjTabHeadRows) arrRows = arrRows.concat(arrObjTabHeadRows);
					}

					// B: Add next line of text to this cell
					if (arrCellsLines[col][idx]) currRow[col].text += arrCellsLines[col][idx];
				}

				// 2: Add this new rows H to overall (use cell with the most lines as the determiner for overall row Height)
				emuTabCurrH += arrCellsLineHeights[intMaxColIdx];
			}

			if (opts.debug) console.log('-> ' + iRow + ' row done!');
			if (opts.debug) console.log('-> emuTabCurrH (in) . = ' + (emuTabCurrH / EMU).toFixed(1));

			// E: Flush row buffer - Add the current row to table, then truncate row cell array
			// IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
			if (currRow.length) arrRows.push(jQuery.extend(true, [], currRow));
			currRow.length = 0;
		});

		// STEP 4-2: Flush final row buffer to slide
		arrObjSlides.push(jQuery.extend(true, [], arrRows));

		// LAST:
		if (opts.debug) { console.log('arrObjSlides count = ' + arrObjSlides.length); console.log(arrObjSlides); }
		return arrObjSlides;
	}

	/**
	 * NOTE: Used by both: text and lineChart
	 * Creates `a:innerShdw` or `a:outerShdw` depending on pass options `opts`.
	 * @param {Object} opts optional shadow properties
	 * @param {Object} defaults defaults for unspecified properties in `opts`
	 * @see http://officeopenxml.com/drwSp-effects.php
	 *	{ type: 'outer', blur: 3, offset: (23000 / 12700), angle: 90, color: '000000', opacity: 0.35, rotateWithShape: true };
	 */
	function createShadowElement(options, defaults, isShape) {
		if (options === 'none') {
			return '<a:effectLst/>';
		}
		var
			strXml = '<a:effectLst>',
			opts = getMix(defaults, options),
			type = opts.type || 'outer',
			blur = opts.blur * ONEPT,
			offset = opts.offset * ONEPT,
			angle = opts.angle * 60000,
			color = opts.color,
			opacity = opts.opacity * 100000,
			rotateWithShape = opts.rotateWithShape ? 1 : 0;

		strXml += '<a:' + type + 'Shdw sx="100000" sy="100000" kx="0" ky="0"  algn="bl" blurRad="' + blur + '" ';
		strXml += 'rotWithShape="' + (+rotateWithShape) + '"';
		strXml += ' dist="' + offset + '" dir="' + angle + '">';
		strXml += '<a:srgbClr val="' + color + '">'; // TODO: should accept scheme colors implemented in Issue #135
		strXml += '<a:alpha val="' + opacity + '"/></a:srgbClr>';
		strXml += '</a:' + type + 'Shdw>';
		strXml += '</a:effectLst>';

		return strXml;
	}

	/**
	 * Checks shadow options passed by user and performs corrections if needed.
	 * @param {Object} shadowOpts
	 */
	function correctShadowOptions(shadowOpts) {
		if (!shadowOpts || shadowOpts === 'none') return;

		// OPT: `type`
		if (shadowOpts.type != 'outer' && shadowOpts.type != 'inner') {
			console.warn('Warning: shadow.type options are `outer` or `inner`.');
			shadowOpts.type = 'outer';
		}

		// OPT: `angle`
		if (shadowOpts.angle) {
			// A: REALITY-CHECK
			if (isNaN(Number(shadowOpts.angle)) || shadowOpts.angle < 0 || shadowOpts.angle > 359) {
				console.warn('Warning: shadow.angle can only be 0-359');
				shadowOpts.angle = 270;
			}

			// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
			shadowOpts.angle = Math.round(Number(shadowOpts.angle));
		}

		// OPT: `opacity`
		if (shadowOpts.opacity) {
			// A: REALITY-CHECK
			if (isNaN(Number(shadowOpts.opacity)) || shadowOpts.opacity < 0 || shadowOpts.opacity > 1) {
				console.warn('Warning: shadow.opacity can only be 0-1');
				shadowOpts.opacity = 0.75;
			}

			// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
			shadowOpts.opacity = Number(shadowOpts.opacity)
		}
	}

	function correctGridLineOptions(glOpts) {
		if (!glOpts || glOpts === 'none') return;
		if (glOpts.size !== undefined && (isNaN(Number(glOpts.size)) || glOpts.size <= 0)) {
			console.warn('Warning: chart.gridLine.size must be greater than 0.');
			delete glOpts.size; // delete prop to used defaults
		}
		if (glOpts.style && ['solid', 'dash', 'dot'].indexOf(glOpts.style) < 0) {
			console.warn('Warning: chart.gridLine.style options: `solid`, `dash`, `dot`.');
			delete glOpts.style;
		}
	}



	/* ===============================================================================================
	|
	######                                             #     ######   ###
	#     #  #    #  #####   #       #   ####         # #    #     #   #
	#     #  #    #  #    #  #       #  #    #       #   #   #     #   #
	######   #    #  #####   #       #  #           #     #  ######    #
	#        #    #  #    #  #       #  #           #######  #         #
	#        #    #  #    #  #       #  #    #      #     #  #         #
	#         ####   #####   ######  #   ####       #     #  #        ###
	|
	==================================================================================================
	*/

	/**
	 * Library version
	 */
	this.version = APP_VER + '.' + APP_BLD;

	/**
	 * Expose a couple private helper functions from above
	 */
	this.inch2Emu = inch2Emu;
	this.rgbToHex = rgbToHex;

	/**
	 * Gets the Presentation's Slide Layout {object} from `LAYOUTS`
	 */
	this.getLayout = function getLayout() {
		return gObjPptx.pptLayout;
	};

	/**
	 * Set Right-to-Left (RTL) mode for users whose language requires this setting
	 */
	this.setRTL = function setRTL(inBool) {
		if (typeof inBool !== 'boolean') return;
		else {
			gObjPptx.rtlMode = inBool;
		}
	}

	/**
	 * Sets the Presentation's Slide Layout {object}: [screen4x3, screen16x9, widescreen]
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 * @param {string} inLayout - a const name from LAYOUTS variable
	 * @param {object} inLayout - an object with user-defined w/h
	 */
	this.setLayout = function setLayout(inLayout) {
		// Allow custom slide size (inches) [ISSUE #29]
		if (typeof inLayout === 'object' && inLayout.width && inLayout.height) {
			LAYOUTS['LAYOUT_USER'].width = Math.round(Number(inLayout.width) * EMU);
			LAYOUTS['LAYOUT_USER'].height = Math.round(Number(inLayout.height) * EMU);

			gObjPptx.pptLayout = LAYOUTS['LAYOUT_USER'];
		}
		else if (Object.keys(LAYOUTS).indexOf(inLayout) > -1) {
			gObjPptx.pptLayout = LAYOUTS[inLayout];
		}
		else {
			try { console.warn('UNKNOWN LAYOUT! Valid values = ' + Object.keys(LAYOUTS)); } catch (ex) { }
		}
	}

	/**
	 * Sets the Presentation's Title
	 */
	this.setTitle = function setTitle(inStrTitle) {
		gObjPptx.title = inStrTitle || 'PptxGenJS Presentation';
	};

	/**
	 * Sets the Presentation Option: `isBrowser`
	 * Target: Angular/React/Webpack, etc.
	 * This setting affects how files are saved: using `fs` for Node.js or browser libs
	 */
	this.setBrowser = function setBrowser(inBool) {
		gObjPptx.isBrowser = inBool || false;
	};

	/**
	 * Sets the Presentation's Author
	 */
	this.setAuthor = function setAuthor(inStrAuthor) {
		gObjPptx.author = inStrAuthor || 'PptxGenJS';
	};

	/**
	 * DESC: Sets the Presentation's Revision
	 * NOTE: PowerPoint requires `revision` be: number only (without "." or ",") otherwise, PPT will throw errors upon opening Presentation.
	 */
	this.setRevision = function setRevision(inStrRevision) {
		gObjPptx.revision = inStrRevision || '1';
		gObjPptx.revision = gObjPptx.revision.replace(/[\.\,\-]+/gi, '');
		if (isNaN(gObjPptx.revision)) gObjPptx.revision = '1';
	};

	/**
	 * Sets the Presentation's Subject
	 */
	this.setSubject = function setSubject(inStrSubject) {
		gObjPptx.subject = inStrSubject || 'PptxGenJS Presentation';
	};

	/**
	 * Sets the Presentation's Company
	 */
	this.setCompany = function setCompany(inStrCompany) {
		gObjPptx.company = inStrCompany || 'PptxGenJS';
	};

	/**
	 * Export the Presentation to an .pptx file
	 * @param {string} [inStrExportName] - Filename to use for the export
	 */
	this.save = function save(inStrExportName: string, funcCallback?: Function, outputType?: string) {
		var intRels = 0, arrRelsDone = [];

		// STEP 1: Add empty placeholder objects to slides that don't already have them
		gObjPptx.slides.forEach(function(slide, idx) {
			if (slide.layoutObj) addPlaceholdersToSlides(slide);
		});

		// STEP 2: Set export properties
		if (funcCallback) gObjPptx.saveCallback = funcCallback;
		if (inStrExportName) gObjPptx.fileName = inStrExportName;

		// STEP 3: Read/Encode Images
		// PERF: Only send unique paths for encoding (encoding func will find and fill *ALL* matching paths across the Presentation)

		// A: Slide rels
		gObjPptx.slides.forEach(function(slide, idx) { intRels += encodeSlideMediaRels(slide, arrRelsDone); });

		// B: Layout rels
		gObjPptx.slideLayouts.forEach(function(layout, idx) { intRels += encodeSlideMediaRels(layout, arrRelsDone); });

		// C: Master Slide rels
		intRels += encodeSlideMediaRels(gObjPptx.masterSlide, arrRelsDone);

		// STEP 4: Export now if there's no images to encode (otherwise, last async imgConvert call above will call exportFile)
		if (intRels == 0) doExportPresentation(outputType);
	};

	/**
	 * Add a new Slide to the Presentation
	 * @returns {Object[]} slideObj - The new Slide object
	 */
	this.addNewSlide = function addNewSlide(inMasterName): object[] {
		var slideObj = {};
		var slideNum = gObjPptx.slides.length;
		var slideObjNum = 0;
		var pageNum = (slideNum + 1);
		var objLayout = gObjPptx.slideLayouts.filter(function(layout) { return layout.name == inMasterName })[0];

		// A: Add this SLIDE to PRESENTATION, Add default values as well
		gObjPptx.slides[slideNum] = {
			slide: slideObj,
			name: 'Slide ' + pageNum,
			numb: pageNum,
			data: [],
			rels: [],
			slideNumberObj: null,
			layout: inMasterName || '[ default ]',
			layoutObj: objLayout
		};

		// ==========================================================================
		// PUBLIC METHODS:
		// ==========================================================================

		slideObj.getPageNumber = function() {
			return pageNum;
		};

		slideObj.slideNumber = function(inObj) {
			if (inObj && typeof inObj === 'object') {
				// A:
				gObjPptx.slides[slideNum].slideNumberObj = inObj;

				// B: Add slideNumber to slideMaster1.xml
				if (!gObjPptx.masterSlide.slideNumberObj) gObjPptx.masterSlide.slideNumberObj = inObj;

				// C: Add slideNumber to `BLANK` (default) layout
				if (!gObjPptx.slideLayouts[0].slideNumberObj) gObjPptx.slideLayouts[0].slideNumberObj = inObj;
			}
			else {
				return gObjPptx.slides[slideNum].slideNumberObj;
			}
		};

		/**
		 * Generate the chart based on input data.
		 *
		 * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
		 *
		 * @param {object} renderType should belong to: 'column', 'pie'
		 * @param {object} data a JSON object with follow the following format
		 * {
		 *   title: 'eSurvey chart',
		 *   data: [
		 *		{
		 *			name: 'Income',
		 *			labels: ['2005', '2006', '2007', '2008', '2009'],
		 *			values: [23.5, 26.2, 30.1, 29.5, 24.6]
		 *		},
		 *		{
		 *			name: 'Expense',
		 *			labels: ['2005', '2006', '2007', '2008', '2009'],
		 *			values: [18.1, 22.8, 23.9, 25.1, 25]
		 *		}
		 *	 ]
		 * }
		 */
		slideObj.addChart = function(type, data, opt) {
			gObjPptxGenerators.addChartDefinition(type, data, opt, gObjPptx.slides[slideNum]);
			return this;
		}

		/**
		 * NOTE: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
		 */
		slideObj.addImage = function(objImage) {
			gObjPptxGenerators.addImageDefinition(objImage, gObjPptx.slides[slideNum]);
			return this;
		};

		slideObj.addMedia = function(opt) {
			var intRels = 1;
			var intImages = ++gObjPptx.imageCounter;
			var intPosX = (opt.x || 0);
			var intPosY = (opt.y || 0);
			var intSizeX = (opt.w || 2);
			var intSizeY = (opt.h || 2);
			var strData = (opt.data || '');
			var strLink = (opt.link || '');
			var strPath = (opt.path || '');
			var strType = (opt.type || "audio");
			var strExtn = "mp3";

			// STEP 1: REALITY-CHECK
			if (!strPath && !strData && strType != 'online') {
				console.error("ERROR: `addMedia()` requires either 'data' or 'path' values!");
				return null;
			}
			else if (strData && strData.toLowerCase().indexOf('base64,') == -1) {
				console.error("ERROR: Media `data` value lacks a base64 header! Ex: 'video/mpeg;base64,NMP[...]')");
				return null;
			}
			// Online Video: requires `link`
			if (strType == 'online' && !strLink) {
				console.error('addMedia() error: online videos require `link` value')
				return null;
			}

			// STEP 2: Set vars for this Slide
			var slideObjNum = gObjPptx.slides[slideNum].data.length;
			var slideObjRels = gObjPptx.slides[slideNum].rels;

			strType = (strData ? strData.split(';')[0].split('/')[0] : strType);
			strExtn = (strData ? strData.split(';')[0].split('/')[1] : strPath.split('.').pop());

			gObjPptx.slides[slideNum].data[slideObjNum] = {};
			gObjPptx.slides[slideNum].data[slideObjNum].type = 'media';
			gObjPptx.slides[slideNum].data[slideObjNum].mtype = strType;
			gObjPptx.slides[slideNum].data[slideObjNum].media = (strPath || 'preencoded.mov');

			// STEP 3: Set media properties & options
			gObjPptx.slides[slideNum].data[slideObjNum].options = {};
			gObjPptx.slides[slideNum].data[slideObjNum].options.x = intPosX;
			gObjPptx.slides[slideNum].data[slideObjNum].options.y = intPosY;
			gObjPptx.slides[slideNum].data[slideObjNum].options.cx = intSizeX;
			gObjPptx.slides[slideNum].data[slideObjNum].options.cy = intSizeY;

			// STEP 4: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
			// NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
			gObjPptx.slides.forEach(function(slide) { intRels += slide.rels.length; });

			if (strType == 'online') {
				// Add video
				slideObjRels.push({
					path: (strPath || 'preencoded' + strExtn),
					data: 'dummy',
					type: 'online',
					extn: strExtn,
					rId: (intRels + 1),
					Target: strLink
				});
				gObjPptx.slides[slideNum].data[slideObjNum].mediaRid = slideObjRels[slideObjRels.length - 1].rId;

				// Add preview/overlay image
				slideObjRels.push({
					path: 'preencoded.png',
					data: IMG_PLAYBTN,
					type: 'image/png',
					extn: 'png',
					rId: (intRels + 2),
					Target: '../media/image' + intRels + '.png'
				});
			}
			else {
				// Audio/Video files consume *TWO* rId's:
				// <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
				// <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
				slideObjRels.push({
					path: (strPath || 'preencoded' + strExtn),
					type: strType + '/' + strExtn,
					extn: strExtn,
					data: (strData || ''),
					rId: (intRels + 0),
					Target: '../media/media' + intImages + '.' + strExtn
				});
				gObjPptx.slides[slideNum].data[slideObjNum].mediaRid = slideObjRels[slideObjRels.length - 1].rId;
				slideObjRels.push({
					path: (strPath || 'preencoded' + strExtn),
					type: strType + '/' + strExtn,
					extn: strExtn,
					data: (strData || ''),
					rId: (intRels + 1),
					Target: '../media/media' + intImages + '.' + strExtn
				});
				// Add preview/overlay image
				slideObjRels.push({
					data: IMG_PLAYBTN,
					path: 'preencoded.png',
					type: 'image/png',
					extn: 'png',
					rId: (intRels + 2),
					Target: '../media/image' + intImages + '.png'
				});
			}

			// LAST: Return this Slide
			return this;
		}

		slideObj.addNotes = function(notes, options) {
			gObjPptxGenerators.addNotesDefinition(notes, options, gObjPptx.slides[slideNum]);
			return this;
		};

		slideObj.addShape = function(shape, opt) {
			gObjPptxGenerators.addShapeDefinition(shape, opt, gObjPptx.slides[slideNum]);
			return this;
		};

		// RECURSIVE: (sometimes)
		// FUTURE: slideObj.addTable = function(arrTabRows, inOpt){
		// FIXME: Move to gObjPptxGenerators (as every other object uses a generator #consistency)
		// TODO: dont forget to update the "this.color" refs below to "target.slide.color"!!!
		slideObj.addTable = function(arrTabRows, inOpt) {
			var opt = (inOpt && typeof inOpt === 'object' ? inOpt : {});

			// STEP 1: REALITY-CHECK
			if (arrTabRows == null || arrTabRows.length == 0 || !Array.isArray(arrTabRows)) {
				try { console.warn('[warn] addTable: Array expected! USAGE: slide.addTable( [rows], {options} );'); } catch (ex) { }
				return null;
			}

			// STEP 2: Row setup: Handle case where user passed in a simple 1-row array. EX: `["cell 1", "cell 2"]`
			//var arrRows = jQuery.extend(true,[],arrTabRows);
			//if ( !Array.isArray(arrRows[0]) ) arrRows = [ jQuery.extend(true,[],arrTabRows) ];
			var arrRows = arrTabRows;
			if (!Array.isArray(arrRows[0])) arrRows = [arrTabRows];

			// STEP 3: Set options
			opt.x = getSmartParseNumber(opt.x || (opt.x == 0 ? 0 : EMU / 2), 'X');
			opt.y = getSmartParseNumber(opt.y || (opt.y == 0 ? 0 : EMU), 'Y');
			opt.cy = opt.h || opt.cy; // NOTE: Dont set default `cy` - leaving it null triggers auto-rowH in `makeXMLSlide()`
			if (opt.cy) opt.cy = getSmartParseNumber(opt.cy, 'Y');
			opt.h = opt.cy;
			opt.autoPage = (opt.autoPage == false ? false : true);
			opt.fontSize = opt.fontSize || DEF_FONT_SIZE;
			opt.lineWeight = (typeof opt.lineWeight !== 'undefined' && !isNaN(Number(opt.lineWeight)) ? Number(opt.lineWeight) : 0);
			opt.margin = (opt.margin == 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_PT);
			if (!isNaN(opt.margin)) opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)]
			if (opt.lineWeight > 1) opt.lineWeight = 1;
			else if (opt.lineWeight < -1) opt.lineWeight = -1;
			// Set default color if needed (table option > inherit from Slide > default to black)
			if (!opt.color) opt.color = opt.color || this.color || DEF_FONT_COLOR;

			// Set/Calc table width
			// Get slide margins - start with default values, then adjust if master or slide margins exist
			var arrTableMargin = DEF_SLIDE_MARGIN_IN;
			// Case 1: Master margins
			if (objLayout && typeof objLayout.margin !== 'undefined') {
				if (Array.isArray(objLayout.margin)) arrTableMargin = objLayout.margin;
				else if (!isNaN(Number(objLayout.margin))) arrTableMargin = [Number(objLayout.margin), Number(objLayout.margin), Number(objLayout.margin), Number(objLayout.margin)];
			}
			// Case 2: Table margins
			/* FIXME: add `margin` option to slide options
				else if ( slideObj.margin ) {
					if ( Array.isArray(slideObj.margin) ) arrTableMargin = slideObj.margin;
					else if ( !isNaN(Number(slideObj.margin)) ) arrTableMargin = [Number(slideObj.margin), Number(slideObj.margin), Number(slideObj.margin), Number(slideObj.margin)];
				}
			*/

			// Calc table width depending upon what data we have - several scenarios exist (including bad data, eg: colW doesnt match col count)
			if (opt.w || opt.cx) {
				opt.cx = getSmartParseNumber((opt.w || opt.cx), 'X');
				opt.w = opt.cx;
			}
			else if (opt.colW) {
				if (typeof opt.colW === 'string' || typeof opt.colW === 'number') {
					opt.cx = Math.floor(Number(opt.colW) * arrRows[0].length);
					opt.w = opt.cx;
				}
				else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length != arrRows[0].length) {
					console.warn('addTable: colW.length != data.length! Defaulting to evenly distributed col widths.');

					var numColWidth = Math.floor(((gObjPptx.pptLayout.width / EMU) - arrTableMargin[1] - arrTableMargin[3]) / arrRows[0].length);
					opt.colW = [];
					for (var idx = 0; idx < arrRows[0].length; idx++) { opt.colW.push(numColWidth); }
					opt.cx = Math.floor(numColWidth * arrRows[0].length);
					opt.w = opt.cx;
				}
			}
			else {
				var numTabWidth = ((gObjPptx.pptLayout.width / EMU) - arrTableMargin[1] - arrTableMargin[3]);
				opt.cx = Math.floor(numTabWidth);
				opt.w = opt.cx;
			}

			// STEP 4: Convert units to EMU now (we use different logic in makeSlide->table - smartCalc is not used)
			if (opt.x < 20) opt.x = inch2Emu(opt.x);
			if (opt.y < 20) opt.y = inch2Emu(opt.y);
			if (opt.cx < 20) opt.cx = inch2Emu(opt.cx);
			if (opt.cy && opt.cy < 20) opt.cy = inch2Emu(opt.cy);

			// STEP 5: Check for fine-grained formatting, disable auto-page when found
			// Since genXmlTextBody already checks for text array ( text:[{},..{}] ) we're done!
			// Text in individual cells will be formatted as they are added by calls to genXmlTextBody within table builder
			arrRows.forEach(function(row, rIdx) {
				row.forEach(function(cell, cIdx) {
					if (cell && cell.text && Array.isArray(cell.text)) opt.autoPage = false;
				});
			});

			// STEP 6: Create hyperlink rels
			createHyperlinkRels(arrRows, gObjPptx.slides[slideNum].rels);

			// STEP 7: Auto-Paging: (via {options} and used internally)
			// (used internally by `addSlidesForTable()` to not engage recursion - we've already paged the table data, just add this one)
			if (opt && opt.autoPage == false) {
				// Add data (NOTE: Use `extend` to avoid mutation)
				gObjPptx.slides[slideNum].data[gObjPptx.slides[slideNum].data.length] = {
					type: 'table',
					arrTabRows: arrRows,
					options: jQuery.extend(true, {}, opt)
				};
			}
			else {
				// Loop over rows and create 1-N tables as needed (ISSUE#21)
				getSlidesForTableRows(arrRows, opt).forEach(function(arrRows, idx) {
					// A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
					var currSlide = (!gObjPptx.slides[slideNum + idx] ? addNewSlide(inMasterName) : gObjPptx.slides[slideNum + idx].slide);

					// B: Reset opt.y to `option`/`margin` after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
					if (idx > 0) opt.y = inch2Emu(opt.newPageStartY || arrTableMargin[0]);

					// C: Add this table to new Slide
					opt.autoPage = false;
					currSlide.addTable(arrRows, jQuery.extend(true, {}, opt));
				});
			}

			// LAST: Return this Slide
			return this;
		};

		slideObj.addText = function(text, options) {
			gObjPptxGenerators.addTextDefinition(text, options, gObjPptx.slides[slideNum], false);
			return this;
		};

		// ==========================================================================
		// POST-METHODS:
		// ==========================================================================

		// NOTE: Slide Numbers: In order for Slide Numbers to work normally, they need to be in all 3 files: master/layout/slide
		// `defineSlideMaster` and `slideObj.slideNumber` will add {slideNumber} to `gObjPptx.masterSlide` and `gObjPptx.slideLayouts`
		// so, lastly, add to the Slide now.
		if (objLayout && objLayout.slideNumberObj && !slideObj.slideNumber()) gObjPptx.slides[slideNum].slideNumberObj = objLayout.slideNumberObj;

		// LAST: Return this Slide
		return slideObj;
	};

	/**
	 * Adds a new slide master [layout] to the presentation.
	 * @param {Object} inObjMasterDef - layout definition
	 * @return {Object} this
	 */
	this.defineSlideMaster = function defineSlideMaster(inObjMasterDef) {
		if (!inObjMasterDef.title) { throw Error("defineSlideMaster() object argument requires a `title` value."); }

		var objLayout = {
			name: inObjMasterDef.title,
			slide: {},
			data: [],
			rels: [],
			margin: inObjMasterDef.margin || DEF_SLIDE_MARGIN_IN,
			slideNumberObj: inObjMasterDef.slideNumber || null
		};

		// STEP 1: Create the Slide Master/Layout
		gObjPptxGenerators.createSlideObject(inObjMasterDef, objLayout);

		// STEP 2: Add it to layout defs
		gObjPptx.slideLayouts.push(objLayout);

		// STEP 3: Add slideNumber to master slide (if any)
		if (objLayout.slideNumberObj && !gObjPptx.masterSlide.slideNumberObj) gObjPptx.masterSlide.slideNumberObj = objLayout.slideNumberObj;

		// LAST:
		return this;
	};

	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * "Auto-Paging is the future!" --Elon Musk
	 *
	 * @param {string} tabEleId - The HTML Element ID of the table
	 * @param {array} inOpts - An array of options (e.g.: tabsize)
	 */
	this.addSlidesForTable = function addSlidesForTable(tabEleId, inOpts) {
		var api = this;
		var opts = inOpts || {};
		var arrObjTabHeadRows = [], arrObjTabBodyRows = [], arrObjTabFootRows = [];
		var arrObjSlides = [], arrRows = [], arrColW = [], arrTabColW = [];
		var intTabW = 0, emuTabCurrH = 0;

		// REALITY-CHECK:
		if (jQuery('#' + tabEleId).length == 0) { console.error('Table "' + tabEleId + '" does not exist!'); return; }

		var arrInchMargins = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
		opts.margin = (opts.margin || opts.margin == 0 ? opts.margin : 0.5);

		if (opts.master && typeof opts.master === 'string') {
			var objLayout = gObjPptx.slideLayouts.filter(function(layout) { return layout.name == opts.master })[0];
			if (objLayout && objLayout.margin) {
				if (Array.isArray(objLayout.margin)) arrInchMargins = objLayout.margin;
				else if (!isNaN(objLayout.margin)) arrInchMargins = [objLayout.margin, objLayout.margin, objLayout.margin, objLayout.margin];
				opts.margin = arrInchMargins;
			}
		}
		else if (opts && opts.margin) {
			if (Array.isArray(opts.margin)) arrInchMargins = opts.margin;
			else if (!isNaN(opts.margin)) arrInchMargins = [opts.margin, opts.margin, opts.margin, opts.margin];
		}

		var emuSlideTabW = (opts.w ? inch2Emu(opts.w) : (gObjPptx.pptLayout.width - inch2Emu(arrInchMargins[1] + arrInchMargins[3])));
		var emuSlideTabH = (opts.h ? inch2Emu(opts.h) : (gObjPptx.pptLayout.height - inch2Emu(arrInchMargins[0] + arrInchMargins[2])));

		// STEP 1: Grab table col widths
		jQuery.each(['thead', 'tbody', 'tfoot'], function(i, val) {
			if (jQuery('#' + tabEleId + ' > ' + val + ' > tr').length > 0) {
				jQuery('#' + tabEleId + ' > ' + val + ' > tr:first-child').find('> th, > td').each(function(i, cell) {
					// FIXME: This is a hack - guessing at col widths when colspan
					if (jQuery(this).attr('colspan')) {
						for (var idx = 0; idx < jQuery(this).attr('colspan'); idx++) {
							arrTabColW.push(Math.round(jQuery(this).outerWidth() / jQuery(this).attr('colspan')));
						}
					}
					else {
						arrTabColW.push(jQuery(this).outerWidth());
					}
				});
				return false; // break out of .each loop
			}
		});
		jQuery.each(arrTabColW, function(i, colW) { intTabW += colW; });

		// STEP 2: Calc/Set column widths by using same column width percent from HTML table
		jQuery.each(arrTabColW, function(i, colW) {
			var intCalcWidth = Number(((emuSlideTabW * (colW / intTabW * 100)) / 100 / EMU).toFixed(2));
			var intMinWidth = jQuery('#' + tabEleId + ' thead tr:first-child th:nth-child(' + (i + 1) + ')').data('pptx-min-width');
			var intSetWidth = jQuery('#' + tabEleId + ' thead tr:first-child th:nth-child(' + (i + 1) + ')').data('pptx-width');
			arrColW.push((intSetWidth ? intSetWidth : (intMinWidth > intCalcWidth ? intMinWidth : intCalcWidth)));
		});

		// STEP 3: Iterate over each table element and create data arrays (text and opts)
		// NOTE: We create 3 arrays instead of one so we can loop over body then show header/footer rows on first and last page
		jQuery.each(['thead', 'tbody', 'tfoot'], function(i, val) {
			jQuery('#' + tabEleId + ' > ' + val + ' > tr').each(function(i, row) {
				var arrObjTabCells = [];
				jQuery(row).find('> th, > td').each(function(i, cell) {
					// A: Get RGB text/bkgd colors
					var arrRGB1 = [];
					var arrRGB2 = [];
					arrRGB1 = jQuery(cell).css('color').replace(/\s+/gi, '').replace('rgba(', '').replace('rgb(', '').replace(')', '').split(',');
					arrRGB2 = jQuery(cell).css('background-color').replace(/\s+/gi, '').replace('rgba(', '').replace('rgb(', '').replace(')', '').split(',');
					// ISSUE#57: jQuery default is this rgba value of below giving unstyled tables a black bkgd, so use white instead (FYI: if cell has `background:#000000` jQuery returns 'rgb(0, 0, 0)', so this soln is pretty solid)
					if (jQuery(cell).css('background-color') == 'rgba(0, 0, 0, 0)' || jQuery(cell).css('background-color') == 'transparent') arrRGB2 = [255, 255, 255];

					// B: Create option object
					var objOpts = {
						fontSize: jQuery(cell).css('font-size').replace(/[a-z]/gi, ''),
						bold: ((jQuery(cell).css('font-weight') == "bold" || Number(jQuery(cell).css('font-weight')) >= 500) ? true : false),
						color: rgbToHex(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
						fill: rgbToHex(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2])),
						border: null,
						margin: null,
						colspan: null,
						rowspan: null
					};
					if (['left', 'center', 'right', 'start', 'end'].indexOf(jQuery(cell).css('text-align')) > -1) objOpts.align = jQuery(cell).css('text-align').replace('start', 'left').replace('end', 'right');
					if (['top', 'middle', 'bottom'].indexOf(jQuery(cell).css('vertical-align')) > -1) objOpts.valign = jQuery(cell).css('vertical-align');

					// C: Add padding [margin] (if any)
					// NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
					if (jQuery(cell).css('padding-left')) {
						objOpts.margin = [];
						jQuery.each(['padding-top', 'padding-right', 'padding-bottom', 'padding-left'], function(i, val) {
							objOpts.margin.push(Math.round(jQuery(cell).css(val).replace(/\D/gi, '')));
						});
					}

					// D: Add colspan/rowspan (if any)
					if (jQuery(cell).attr('colspan')) objOpts.colspan = jQuery(cell).attr('colspan');
					if (jQuery(cell).attr('rowspan')) objOpts.rowspan = jQuery(cell).attr('rowspan');

					// E: Add border (if any)
					if (jQuery(cell).css('border-top-width') || jQuery(cell).css('border-right-width') || jQuery(cell).css('border-bottom-width') || jQuery(cell).css('border-left-width')) {
						objOpts.border = [];
						jQuery.each(['top', 'right', 'bottom', 'left'], function(i, val) {
							var intBorderW = Math.round(Number(jQuery(cell).css('border-' + val + '-width').replace('px', '')));
							var arrRGB = [];
							arrRGB = jQuery(cell).css('border-' + val + '-color').replace(/\s+/gi, '').replace('rgba(', '').replace('rgb(', '').replace(')', '').split(',');
							var strBorderC = rgbToHex(Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2]));
							objOpts.border.push({ pt: intBorderW, color: strBorderC });
						});
					}

					// F: Massage cell text so we honor linebreak tag as a line break during line parsing
					var $cell2 = jQuery(cell).clone();
					$cell2.html(jQuery(cell).html().replace(/<br[^>]*>/gi, '\n'));

					// LAST: Add cell
					arrObjTabCells.push({
						text: jQuery.trim($cell2.text()),
						opts: objOpts
					});
				});
				switch (val) {
					case 'thead': arrObjTabHeadRows.push(arrObjTabCells); break;
					case 'tbody': arrObjTabBodyRows.push(arrObjTabCells); break;
					case 'tfoot': arrObjTabFootRows.push(arrObjTabCells); break;
					default:
				}
			});
		});

		// STEP 4: NOTE: `margin` is "cell margin (pt)" everywhere else tables are used, so explicitly convert to "slide margin" here
		if (opts.margin) {
			opts.slideMargin = opts.margin;
			delete (opts.margin);
		}

		// STEP 5: Break table into Slides as needed
		// Pass head-rows as there is an option to add to each table and the parse func needs this daa to fulfill that option
		opts.arrObjTabHeadRows = arrObjTabHeadRows || '';
		opts.colW = arrColW;

		getSlidesForTableRows(arrObjTabHeadRows.concat(arrObjTabBodyRows).concat(arrObjTabFootRows), opts)
			.forEach(function(arrTabRows, idx) {
				// A: Create new Slide
				var newSlide = (opts.master ? api.addNewSlide(opts.master) : api.addNewSlide());

				// B: DESIGN: Reset `y` to `newPageStartY` or margin after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
				if (idx == 0) opts.y = opts.y || arrInchMargins[0];
				if (idx > 0) opts.y = opts.newPageStartY || arrInchMargins[0];
				if (opts.debug) console.log('opts.newPageStartY:' + opts.newPageStartY + ' / arrInchMargins[0]:' + arrInchMargins[0] + ' => opts.y = ' + opts.y);

				// C: Add table to Slide
				newSlide.addTable(arrTabRows, { x: (opts.x || arrInchMargins[3]), y: opts.y, w: (emuSlideTabW / EMU), colW: arrColW, autoPage: false });

				// D: Add any additional objects
				if (opts.addImage) newSlide.addImage({ path: opts.addImage.url, x: opts.addImage.x, y: opts.addImage.y, w: opts.addImage.w, h: opts.addImage.h });
				if (opts.addShape) newSlide.addShape(opts.addShape.shape, (opts.addShape.opts || opts.addShape.options || {}));
				if (opts.addTable) newSlide.addTable(opts.addTable.rows, (opts.addTable.opts || opts.addTable.options || {}));
				if (opts.addText) newSlide.addText(opts.addText.text, (opts.addText.opts || opts.addText.options || {}));
			});
	}
};

// NodeJS support
if (NODEJS) {
	var jQuery = null;
	var fs = null;
	var https = null;
	var JSZip = null;
	var sizeOf = null;

	// A: jQuery dependency
	try {
		var jsdom = require("jsdom");
		var dom = new jsdom.JSDOM("<!DOCTYPE html>");
		jQuery = require("jquery")(dom.window);
	} catch (ex) { console.error("Unable to load `jquery`!\n" + ex); throw 'LIB-MISSING-JQUERY'; }

	// B: Other dependencies
	try { fs = require("fs"); } catch (ex) { console.error("Unable to load `fs`"); throw 'LIB-MISSING-FS'; }
	try { https = require("https"); } catch (ex) { console.error("Unable to load `https`"); throw 'LIB-MISSING-HTTPS'; }
	try { JSZip = require("jszip"); } catch (ex) { console.error("Unable to load `jszip`"); throw 'LIB-MISSING-JSZIP'; }
	try { sizeOf = require("image-size"); } catch (ex) { console.error("Unable to load `image-size`"); throw 'LIB-MISSING-IMGSIZE'; }

	// LAST: Export module
	module.exports = PptxGenJS;
}
// Angular/React/etc support
else if (APPJS) {
	// A: jQuery dependency
	try { jQuery = require("jquery"); } catch (ex) { console.error("Unable to load `jquery`!\n" + ex); throw 'LIB-MISSING-JQUERY'; }

	// B: Other dependencies
	try { JSZip = require("jszip"); } catch (ex) { console.error("Unable to load `jszip`"); throw 'LIB-MISSING-JSZIP'; }

	// LAST: Export module
	module.exports = PptxGenJS;
}
