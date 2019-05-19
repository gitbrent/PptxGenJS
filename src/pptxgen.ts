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
	EMU, ONEPT, CRLF, DEF_SLIDE_MARGIN_IN, LETTERS, BARCHART_COLORS, SCHEME_COLOR_NAMES,
	DEF_FONT_COLOR, PIECHART_COLORS, CHART_TYPES, MASTER_OBJECTS, BASE_SHAPES
} from './enums';
import { getSmartParseNumber, inch2Emu } from './utils'
//import { gObjPptxGenerators } from './gen-xml';
import * as genXml from './gen-xml';
//import {gObjPptxShapes} from './shapes'

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
	//var gObjPptxColors = require('../dist/pptxgen.colors.js');
	//var gObjPptxShapes = require('../dist/pptxgen.shapes.js');
}

// Polyfill for IE11 (https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Number/isInteger)
Number.isInteger = Number.isInteger || function(value) {
	return typeof value === "number" && isFinite(value) && Math.floor(value) === value;
};

// Common
export type Coord = number | string; // string is in form 'n%'
export interface OptsCoords {
  x?: Coord
  y?: Coord
  w?: Coord
  h?: Coord
}
export interface OptsDataOrPath {
  // Exactly one must be set
  data?: string
  path?: string
}
// Opts
export interface IBorder {
    color?: string // '#696969'
	pt?: number
}
export interface IChartOpts {
	barDir?: string, barGrouping?: string, barGapWidthPct:number, barGapDepthPct:number, bar3DShape:string,
	chartColors?: Array<string>, chartColorsOpacity?: number,
	showLabel:boolean,
	lang:string,dataNoEffects:string,
	dataLabelFormatScatter: string,
	dataLabelFormatCode:string,
	dataLabelBkgrdColors:string,
	dataLabelFontSize:number,
	dataLabelColor:string,
	dataLabelFontFace:string,
	dataLabelPosition:string,
	lineDataSymbol:string,
	lineDataSymbolSize:number,
	lineDataSymbolLineColor:string,
	lineDataSymbolLineSize:number,
	lineSmooth:boolean,
	invertedColors:string,
	dataLabelFontBold:boolean,
	valueBarColors:Array<string>,
	type:CHART_TYPES,
	holeSize:number,
	showValue:boolean,
	showPercent:boolean,
	catLabelFormatCode?: string, dataBorder?: IBorder, lineSize?: number, lineDash?: string, radarStyle?: string, shadow:IShadowOpts
}
export interface IMediaOpts extends OptsCoords, OptsDataOrPath {
  onlineVideoLink?: string;
  type?: "audio" | "online" | "video";
}
export interface IShadowOpts {
	type: string, angle: number, opacity: number
}
export interface ITextOpts extends OptsCoords, OptsDataOrPath {
  align?: string // "left" | "center" | "right"
  autoFit?: boolean
  color?: string
  fontSize?: number
  inset?:number
  lineSpacing?:number
  line?: string // color
  lineSize?: number
  placeholder?: object
  rotate?:number // VALS: degree * 60,000
  shadow?: IShadowOpts
  shape?: {name:string}
  vert?: 'eaVert'|'horz'|'mongolianVert'|'vert'|'vert270'|'wordArtVert'|'wordArtVertRtl'
  valign?: string //"top" | "middle" | "bottom"
}
// Core: `slide` and `presentation`
export interface ILayout {
	name: string
	width: number
	height: number
}
export interface ISlideNumber extends OptsCoords {
	fontFace:string
	fontSize:number
	color: string
}
export interface ISlideRel {
	path: string
	type: string
	extn: string
	data: string
	rId: number
	Target: string
}
export interface ISlideLayout {
	name: string
	slide: ISlide
	data: Array<object>
	rels: Array<ISlideRel>
	margin: Array<number>
	slideNumberObj?: ISlideNumber
}
export interface ISlide {
	slide: {
		back:string
		bkgdImgRid?: number
		color:string
	}
	rels: Array<ISlideRel>
	data: Array<object>
	layoutName: string
	layoutObj?: ILayout
	slideNumberObj?: ISlideNumber
}
export interface IPresentation {
	author: string
	company: string
	revision: string
	subject: string
	title: string
	isBrowser: boolean
	fileName: string
	fileExtn: string
	pptLayout: ILayout
	rtlMode: boolean
	saveCallback?: null
	masterSlide?: ISlide
	chartCounter: number
	imageCounter: number
	slides?: ISlide[]
	slideLayouts?: ISlideLayout[]
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
		slide: null,
		layoutName: null,
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

///////////
/*
export default class PptxGenJS {
    author: string;

  constructor(name, sound){
    this.author = name;
   }
 addText(){
  console.log(this.author + `${this.author}`);
}
*/
///////////

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

	// A: Create internal pptx object
	// B: Set Presentation property defaults
	// TODO-3: use `state` instead of global object

	// C: Expose shape library to clients
	this.charts = CHART_TYPES;
	///this.colors = (typeof gObjPptxColors !== 'undefined' ? gObjPptxColors : {});
	///this.shapes = (typeof gObjPptxShapes !== 'undefined' ? gObjPptxShapes : BASE_SHAPES);
	// Declare only after `this.colors` is initialized
	//var SCHEME_COLOR_NAMES = Object.keys(this.colors).map(function(clrKey) { return this.colors[clrKey] }.bind(this));

	// D: Fall back to base shapes if shapes file was not linked
	///gObjPptxShapes = (gObjPptxShapes || this.shapes);

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
		zip.file("[Content_Types].xml", genXml.makeXmlContTypes());
		zip.file("_rels/.rels", genXml.makeXmlRootRels());
		zip.file("docProps/app.xml", genXml.makeXmlApp());
		zip.file("docProps/core.xml", genXml.makeXmlCore());
		zip.file("ppt/_rels/presentation.xml.rels", genXml.makeXmlPresentationRels());
		//
		zip.file("ppt/theme/theme1.xml", genXml.makeXmlTheme());
		zip.file("ppt/presentation.xml", genXml.makeXmlPresentation());
		zip.file("ppt/presProps.xml", genXml.makeXmlPresProps());
		zip.file("ppt/tableStyles.xml", genXml.makeXmlTableStyles());
		zip.file("ppt/viewProps.xml", genXml.makeXmlViewProps());

		// Create a Layout/Master/Rel/Slide file for each SLIDE
		for (var idx = 1; idx <= gObjPptx.slideLayouts.length; idx++) {
			zip.file("ppt/slideLayouts/slideLayout" + idx + ".xml", genXml.makeXmlLayout(gObjPptx.slideLayouts[idx - 1]));
			zip.file("ppt/slideLayouts/_rels/slideLayout" + idx + ".xml.rels", genXml.makeXmlSlideLayoutRel(idx));
		}

		for (var idx = 0; idx < gObjPptx.slides.length; idx++) {
			intSlideNum++;
			zip.file('ppt/slides/slide' + intSlideNum + '.xml', genXml.makeXmlSlide(gObjPptx.slides[idx]));
			zip.file('ppt/slides/_rels/slide' + intSlideNum + '.xml.rels', genXml.makeXmlSlideRel(intSlideNum));

			// Here we will create all slide notes related items. Notes of empty strings
			// are created for slides which do not have notes specified, to keep track of _rels.
			zip.file('ppt/notesSlides/notesSlide' + intSlideNum + '.xml', genXml.makeXmlNotesSlide(gObjPptx.slides[idx]));
			zip.file('ppt/notesSlides/_rels/notesSlide' + intSlideNum + '.xml.rels', genXml.makeXmlNotesSlideRel(intSlideNum));
		}

		zip.file("ppt/slideMasters/slideMaster1.xml", genXml.makeXmlMaster(gObjPptx.masterSlide));
		zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", genXml.makeXmlMasterRel(gObjPptx.masterSlide));
		zip.file('ppt/notesMasters/notesMaster1.xml', genXml.makeXmlNotesMaster());
		zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', genXml.makeXmlNotesMasterRel());

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
	this.setLayout = function setLayout(inLayout?) {
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
	this.setRevision = function setRevision(inStrRevision:string) {
		gObjPptx.revision = inStrRevision || '1';
		gObjPptx.revision = gObjPptx.revision.replace(/[\.\,\-]+/gi, '');
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
		gObjPptx.slides.forEach(function(slide) {
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
	this.addNewSlide = function addNewSlide(inMasterName:string): object[] {
		var slideObj = {};
		var slideNum = gObjPptx.slides.length;
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
			layoutName: inMasterName || '[ default ]',
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
			genXml.createHyperlinkRels(arrRows, gObjPptx.slides[slideNum].rels);

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
			genXml.gObjPptxGenerators.addTextDefinition(text, options, gObjPptx.slides[slideNum], false);
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

		var objLayout:ISlideLayout = {
			name: inObjMasterDef.title,
			slide: null,
			data: [],
			rels: [],
			margin: inObjMasterDef.margin || DEF_SLIDE_MARGIN_IN,
			slideNumberObj: inObjMasterDef.slideNumber || null
		};

		// STEP 1: Create the Slide Master/Layout
		genXml.gObjPptxGenerators.createSlideObject(inObjMasterDef, objLayout);

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
