/*\
|*|  :: pptxgen.js ::
|*|
|*|  JavaScript framework that creates PowerPoint (pptx) presentations
|*|  https://github.com/gitbrent/PptxGenJS
|*|
|*|  This framework is released under the MIT Public License (MIT)
|*|
|*|  PptxGenJS (C) 2015-2017 Brent Ely -- https://github.com/gitbrent
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

// Polyfill for IE11 (https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Number/isInteger)
Number.isInteger = Number.isInteger || function(value) {
	return typeof value === "number" && isFinite(value) && Math.floor(value) === value;
};

// Detect Node.js
var NODEJS = ( typeof module !== 'undefined' && module.exports );

// [Node.js] <script> includes
if ( NODEJS ) {
	var gObjPptxMasters = require('../dist/pptxgen.masters.js');
	var gObjPptxShapes  = require('../dist/pptxgen.shapes.js');
}

var PptxGenJS = function(){
	// CONSTANTS
	var APP_VER = "1.2.1";
	var APP_REL = "20170220";
	var LAYOUTS = {
		'LAYOUT_4x3'  : { name: 'screen4x3',   width:  9144000, height: 6858000 },
		'LAYOUT_16x9' : { name: 'screen16x9',  width:  9144000, height: 5143500 },
		'LAYOUT_16x10': { name: 'screen16x10', width:  9144000, height: 5715000 },
		'LAYOUT_WIDE' : { name: 'custom',      width: 12191996, height: 6858000 },
		'LAYOUT_USER' : { name: 'custom',      width: 12191996, height: 6858000 }
	};
	var BASE_SHAPES = {
		RECTANGLE: { 'displayName': 'Rectangle', 'name': 'rect', 'avLst': {} },
		LINE:      { 'displayName': 'Line', 'name': 'line', 'avLst': {} }
	};
	var SLDNUMFLDID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}';
	var IMG_BROKEN = 'data:image/png;base64,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';
	var IMG_PLAYBTN = 'data:image/png;base64,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';
	var EMU = 914400;	// One (1) Inch - OfficeXML measures in EMU (English Metric Units)
	var ONEPT = 12700;	// One (1) point (pt)
	var CRLF = '\r\n';

	// A: Create internal pptx object
	var gObjPptx = {};

	// B: Set Presentation Property Defaults
	gObjPptx.title = 'PptxGenJS Presentation';
	gObjPptx.fileName = 'Presentation';
	gObjPptx.fileExtn = '.pptx';
	gObjPptx.pptLayout = LAYOUTS['LAYOUT_16x9'];
	gObjPptx.slides = [];

	// C: Expose shape library to clients
	this.shapes  = (typeof gObjPptxShapes  !== 'undefined') ? gObjPptxShapes  : BASE_SHAPES;
	this.masters = (typeof gObjPptxMasters !== 'undefined') ? gObjPptxMasters : {};

	// D: Fall back to base shapes if shapes file was not linked
	if ( typeof gObjPptxShapes === 'undefined' ) gObjPptxShapes = BASE_SHAPES;

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
	function doExportPresentation(callback) {
		var intSlideNum = 0, intRels = 0;

		// STEP 1: Create new JSZip file
		var zip = new JSZip();

		// STEP 2: Add all required folders and files
		zip.folder("_rels");
		zip.folder("docProps");
		zip.folder("ppt").folder("_rels");
		zip.folder("ppt/media");
		zip.folder("ppt/slideLayouts").folder("_rels");
		zip.folder("ppt/slideMasters").folder("_rels");
		zip.folder("ppt/slides").folder("_rels");
		zip.folder("ppt/theme");

		zip.file("[Content_Types].xml", makeXmlContTypes());
		zip.file("_rels/.rels", makeXmlRootRels());
		zip.file("docProps/app.xml", makeXmlApp());
		zip.file("docProps/core.xml", makeXmlCore());
		zip.file("ppt/_rels/presentation.xml.rels", makeXmlPresentationRels());

		// Create a Layout/Master/Rel/Slide file for each SLIDE
		for ( var idx=0; idx<gObjPptx.slides.length; idx++ ) {
			intSlideNum++;
			zip.file("ppt/slideLayouts/slideLayout"+ intSlideNum +".xml", makeXmlSlideLayout( intSlideNum ));
			zip.file("ppt/slideLayouts/_rels/slideLayout"+ intSlideNum +".xml.rels", makeXmlSlideLayoutRel( intSlideNum ));
			zip.file("ppt/slides/slide"+ intSlideNum +".xml", makeXmlSlide(gObjPptx.slides[idx]));
			zip.file("ppt/slides/_rels/slide"+ intSlideNum +".xml.rels", makeXmlSlideRel( intSlideNum ));
		}
		zip.file("ppt/slideMasters/slideMaster1.xml", makeXmlSlideMaster());
		zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", makeXmlSlideMasterRel());

		// Loop over all Rels (images/media) and add them
		gObjPptx.slides.forEach(function(slide,idx){
			slide.rels.forEach(function(rel,idy){
				var data = rel.data;

				// A: Users will undoubtedly pass in string in various formats, so modify as needed
				if      ( data.indexOf(',') == -1 && data.indexOf(';') == -1 ) data = 'image/png;base64,' + data;
				else if ( data.indexOf(',') == -1                            ) data = 'image/png;base64,' + data;
				else if ( data.indexOf(';') == -1                            ) data = 'image/png;' + data;

				// B: Add media
				if ( rel.type != 'online' ) zip.file( rel.Target.replace('..','ppt'), data.split(',').pop(), {base64:true} );
			});
		});

		zip.file("ppt/theme/theme1.xml", makeXmlTheme());
		zip.file("ppt/presentation.xml", makeXmlPresentation());
		zip.file("ppt/presProps.xml",    makeXmlPresProps());
		zip.file("ppt/tableStyles.xml",  makeXmlTableStyles());
		zip.file("ppt/viewProps.xml",    makeXmlViewProps());

		// STEP 3: Push the PPTX file to browser
		var strExportName = ((gObjPptx.fileName.toLowerCase().indexOf('.ppt') > -1) ? gObjPptx.fileName : gObjPptx.fileName+gObjPptx.fileExtn);
		if ( NODEJS ) {
			zip.generateAsync({type:'nodebuffer'}).then(function(content){ fs.writeFile(strExportName, content, callback(strExportName)); });
		}
		else {
			zip.generateAsync({type:'blob'}).then(function(content){ saveAs(content, strExportName); });
		}
	}

	function componentToHex(c) {
		var hex = c.toString(16);
		return hex.length == 1 ? "0" + hex : hex;
	}

	/**
	 * DESC: Used by `addSlidesForTable()` to convert RGB colors from jQuery selectors to Hex for Presentation colors
	 */
	function rgbToHex(r, g, b) {
		if (! Number.isInteger(r)) { try { console.warn('Integer expected!'); } catch(ex){} }
		return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase();
	}

	function inch2Emu(inches) {
		// FIRST: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
		// Any value over 100 damn sure isnt inches, must be EMU already, so just return it
		if (inches > 100) return inches;
		if ( typeof inches == 'string' ) inches = Number( inches.replace(/in*/gi,'') );
		return Math.round(EMU * inches);
	}

	function getSizeFromImage(inImgUrl) {
		if ( NODEJS ) {
			try {
				var dimensions = sizeOf(inImgUrl);
				return { width:dimensions.width, height:dimensions.height };
			}
			catch(ex) {
				console.error('ERROR: Unable to read image: '+inImgUrl);
				return { width:0, height:0 };
			}
		}

		// A: Create
		var image = new Image();

		// B: Set onload event
		image.onload = function(){
			// FIRST: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (this.width + this.height == 0) { return { width:0, height:0 }; }
			var obj = { width:this.width, height:this.height };
			return obj;
		};
		image.onerror = function(){
			try { console.error( '[Error] Unable to load image: ' + inImgUrl ); } catch(ex){}
		};

		// C: Load image
		image.src = inImgUrl;
	}

	function convertImgToDataURLviaCanvas(slideRel){
		// A: Create
		var image = new Image();

		// B: Set onload event
		image.onload = function(){
			// First: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (this.width + this.height == 0) { this.onerror(); return; }
			var canvas = document.createElement('CANVAS');
			var ctx = canvas.getContext('2d');
			canvas.height = this.height;
			canvas.width  = this.width;
			ctx.drawImage(this, 0, 0);
			// Users running on local machine will get the following error:
			// "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
			// when the canvas.toDataURL call executes below.
			try { callbackImgToDataURLDone( canvas.toDataURL(slideRel.type), slideRel ); }
			catch(ex) {
				this.onerror();
				if ( window.location.href.indexOf('file:') == 0 ) {
					console.warn("WARNING: You are running this in a local web browser, which means you cant read local files! (use '--allow-file-access-from-files' flag with Chrome, etc.)");
				}
				return;
			}
			canvas = null;
		};
		image.onerror = function(){
			try {
				if ( typeof window !== 'undefined' && window.location.href.indexOf('file:') == 0 ) {
					console.warn("WARNING: You are running this in a local web browser, which means you cant read local files! (use '--allow-file-access-from-files' flag with Chrome, etc.)");
				}
				console.error('Unable to load image: "'+ slideRel.path +'"\nPlease check the image URL:\n'+ ( slideRel.path.indexOf('/') == 0 ? slideRel.path : window.location.href.substring(0,window.location.href.lastIndexOf('/')+1) + slideRel.path ) );
			} catch(ex){}
			// Return a predefined "Broken image" graphic so the user will see something on the slide
			callbackImgToDataURLDone(IMG_BROKEN, slideRel);
		};

		// C: Load image
		image.src = slideRel.path;
	}

	function callbackImgToDataURLDone(inStr, slideRel){
		var intEmpty = 0;

		// STEP 1: Set data for this rel, count outstanding
		$.each(gObjPptx.slides, function(i,slide){
			$.each(slide.rels, function(i,rel){
				if ( rel.path == slideRel.path ) rel.data = inStr;
				if ( !rel.data ) intEmpty++;
			});
		});

		// STEP 2: Continue export process if all rels have base64 `data` now
		if ( intEmpty == 0 ) doExportPresentation();
	}

	function calcEmuCellHeightForStr(cell, inIntWidthInches) {
		// FORMULA for char-per-inch: (desired chars per line) / (font size [chars-per-inch]) = (reqd print area in inches)
		var GRATIO = 2.61803398875; // "Golden Ratio"
		var intCharPerInch = -1, intCalcGratio = 0;

		// STEP 1: Calc chars-per-inch [pitch]
		// SEE: CPL Formula from http://www.pearsonified.com/2012/01/characters-per-line.php
		intCharPerInch = (120 / cell.opts.font_size);

		// STEP 2: Calc line count
		var intLineCnt = Math.floor( cell.text.length / (intCharPerInch * inIntWidthInches) );
		if (intLineCnt < 1) intLineCnt = 1; // Dont allow line count to be 0!

		// STEP 3: Calc cell height
		var intCellH = ( intLineCnt * ((cell.opts.font_size * 2) / 100) );
		if ( intLineCnt > 8 ) intCellH = (intCellH * 0.9);

		// STEP 4: Add cell padding to height
		if ( cell.opts.margin && Array.isArray(cell.opts.margin) ) {
			intCellH += (cell.opts.margin[0]/ONEPT*(1/72)) + (cell.opts.margin[2]/ONEPT*(1/72));
		}
		else if ( cell.opts.margin && Number.isInteger(cell.opts.margin) ) {
			intCellH += (cell.opts.margin/ONEPT*(1/72)) + (cell.opts.margin/ONEPT*(1/72));
		}

		// LAST: Return size
		return inch2Emu( intCellH );
	}

	function getShapeInfo( shapeName ) {
		if ( !shapeName ) return gObjPptxShapes.RECTANGLE;

		if ( typeof shapeName == 'object' && shapeName.name && shapeName.displayName && shapeName.avLst ) return shapeName;

		if ( gObjPptxShapes[shapeName] ) return gObjPptxShapes[shapeName];

		var objShape = gObjPptxShapes.filter(function(obj){ return obj.name == shapeName || obj.displayName; })[0];
		if ( typeof objShape !== 'undefined' && objShape != null ) return objShape;

		return gObjPptxShapes.RECTANGLE;
	}

	function getSmartParseNumber( inVal, inDir ) {
		// FIRST: Convert string numeric value if reqd
		if ( typeof inVal == 'string' && !isNaN(Number(inVal)) ) inVal = Number(inVal);

		// CASE 1: Number in inches
		// Figure any number less than 100 is inches
		if ( typeof inVal == 'number' && inVal < 100 ) return inch2Emu(inVal);

		// CASE 2: Number is already converted to something other than inches
		// Figure any number greater than 100 is not inches! :)  Just return it (its EMU already i guess??)
		if ( typeof inVal == 'number' && inVal >= 100 ) return inVal;

		// CASE 3: Percentage (ex: '50%')
		if ( typeof inVal == 'string' && inVal.indexOf('%') > -1 ) {
			if ( inDir && inDir == 'X') return Math.round( (parseInt(inVal,10) / 100) * gObjPptx.pptLayout.width  );
			if ( inDir && inDir == 'Y') return Math.round( (parseInt(inVal,10) / 100) * gObjPptx.pptLayout.height );
			// Default: Assume width (x/cx)
			return Math.round( (parseInt(inVal,10) / 100) * gObjPptx.pptLayout.width );
		}

		// LAST: Default value
		return 0;
	}

	function decodeXmlEntities( inStr ) {
		// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
		if ( typeof inStr === 'undefined' || inStr == null ) return "";
		return inStr.toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/\'/g,'&apos;');
	}

	/**
	* Magic happens here
	*/
	function parseTextToLines(inStr, inFontSize, inWidth) {
		var U = 2.2; // Character Constant thingy
		var CPL = (inWidth*EMU / ( inFontSize/U ));
		var arrLines = [];
		var strCurrLine = '';

		// A: Remove leading/trailing space
		inStr = $.trim(inStr);

		// B: Build line array
		$.each(inStr.split('\n'), function(i,line){
			$.each(line.split(' '), function(i,word){
				if ( strCurrLine.length + word.length + 1 < CPL ) {
					strCurrLine += (word + " ");
				}
				else {
					if ( strCurrLine ) arrLines.push( strCurrLine );
					strCurrLine = (word + " ");
				}
			});
			// All words for this line have been exhausted, flush buffer to new line, clear line var
			if ( strCurrLine ) arrLines.push( $.trim(strCurrLine) + CRLF );
			strCurrLine = "";
		});

		// C: Remove trailing linebreak
		arrLines[(arrLines.length-1)] = $.trim(arrLines[(arrLines.length-1)]);

		// D: Return lines
		return arrLines;
	}

	/**
	* Magic happens here
	*/
	function getSlidesForTableRows( inArrRows, opts ) {
		var LINEH_MODIFIER = 1.8;
		var DEF_FONT_SIZE = 12;
		var opts = opts || {};
		var arrInchMargins = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
		var arrObjTabHeadRows = [], arrObjTabBodyRows = [], arrObjTabFootRows = [];
		var arrObjSlides = [], arrRows = [], currRow = [];
		var intTabW = 0, emuTabCurrH = 0;
		var emuSlideTabW = EMU*1, emuSlideTabH = EMU*1;
		var arrObjTabHeadRows = opts.arrObjTabHeadRows || '';
		var numCols = 0;

		if (opts.debug) console.log('opts.w ............. = '+ (opts.w||'').toString());
		if (opts.debug) console.log('opts.colW .......... = '+ (opts.colW||'').toString());
		if (opts.debug) console.log('opts.slideMargin ... = '+ (opts.slideMargin||'').toString());

		// NOTE: Use default size of 1 as zero cell margin is causing our tables to be too large and touch bottom of slide!
		if ( !opts.slideMargin && opts.slideMargin != 0 ) opts.slideMargin = 1;

		// STEP 1: Calc margins/usable space
		if ( opts.slideMargin || opts.slideMargin == 0 ) {
			if ( Array.isArray(opts.slideMargin) ) arrInchMargins = opts.slideMargin;
			else if ( !isNaN(opts.slideMargin) ) arrInchMargins = [opts.slideMargin, opts.slideMargin, opts.slideMargin, opts.slideMargin];
		}
		else if ( opts && opts.master && opts.master.margin && gObjPptxMasters) {
			if ( Array.isArray(opts.master.margin) ) arrInchMargins = opts.master.margin;
			else if ( !isNaN(opts.master.margin) ) arrInchMargins = [opts.master.margin, opts.master.margin, opts.master.margin, opts.master.margin];
		}

		// STEP 2: Calc number of columns
		// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
		// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
		inArrRows[0].forEach(function(cell,idx){
			var cellOpts = cell.options || cell.opts || null; // DEPRECATED (`opts`)
			numCols += ( cellOpts && cellOpts.colspan ? cellOpts.colspan : 1 );
		});

		if (opts.debug) console.log('arrInchMargins ..... = '+ arrInchMargins.toString());
		if (opts.debug) console.log('numCols ............ = '+ numCols );

		// Calc opts.w if we can
		if ( !opts.w && opts.colW ) {
			if ( Array.isArray(opts.colW) ) opts.colW.forEach(function(val,idx){ opts.w += val });
			else { opts.w = opts.colW * numCols }
		}

		// STEP 2: Set table size now that we have usable space calc'd
		emuSlideTabW = ( opts.w ? inch2Emu(opts.w) : (gObjPptx.pptLayout.width  - inch2Emu((opts.x || arrInchMargins[1]) + arrInchMargins[3])) );
		emuSlideTabH = ( opts.h ? inch2Emu(opts.h) : (gObjPptx.pptLayout.height - inch2Emu((opts.y || arrInchMargins[0]) + arrInchMargins[2])) );

		if (opts.debug) console.log('gObjPptx.pptLayout.height (in) = '+ (gObjPptx.pptLayout.height/EMU) );
		if (opts.debug) console.log('emuSlideTabW (in) ............ = '+ (emuSlideTabW/EMU).toFixed(1) );
		if (opts.debug) console.log('emuSlideTabH (in) ............ = '+ (emuSlideTabH/EMU).toFixed(1) );

		// STEP 3: Calc column widths if needed so we can subsequently calc lines (we need `emuSlideTabW`!)
		if ( !opts.colW || !Array.isArray(opts.colW) ) {
			if ( opts.colW && !isNaN(Number(opts.colW)) ) {
				var arrColW = [];
				inArrRows[0].forEach(function(cell,idx){ arrColW.push( opts.colW ) });
				opts.colW = [];
				arrColW.forEach(function(val,idx){ opts.colW.push(val) });
			}
			// No column widths provided? Then distribute cols.
			else {
				opts.colW = [];
				for (var iCol=0; iCol<numCols; iCol++) { opts.colW.push( (emuSlideTabW/EMU/numCols) ); }
			}
		}

		if (opts.debug) console.log('opts.colW..... = '+ opts.colW.toString());

		// STEP 4: Iterate over each line and perform magic
		// ROBUST: inArrRows will be an array of {text:'', opts{}} whether from `addSlidesForTable()` or `.addTable()`
		inArrRows.forEach(function(row,iRow){
			// A: Reset ROW variables
			var arrCellsLines = [], arrCellsLineHeights = [], emuRowH = 0, intMaxLineCnt = 0, intMaxColIdx = 0;

			// B: Parse and store each cell's text into line array (*MAGIC HAPPENS HERE*)
			row.forEach(function(cell,iCell){
				// DESIGN: Cells are henceforth {objects} with `text` and `opts`

				// 1: Cleanse data
				if ( !isNaN(cell) || typeof cell === 'string' ) {
					// Grab table formatting `opts` to use here so text style/format inherits as it should
					cell = { text:cell.toString(), opts:opts };
				}
				else if ( typeof cell === 'object' ) {
					// ARG0: `text`
					if ( !cell.text ) cell.text = "?";

					// ARG1: `options`
					var opt = cell.options || cell.opts || {}; // Legacy support for `opts` (<= v1.2.0)
					cell.opts = opt; // This odd soln is needed until `opts` can be safely discarded (DEPRECATED)
				}

				// 2: Create a cell object for each table column
				currRow.push({ text:'', opts:cell.opts });

				// 3: Parse cell contents into lines (**MAGIC HAPENSS HERE**)
				var lines = parseTextToLines(cell.text.trim(), (cell.opts.font_size || opts.font_size || DEF_FONT_SIZE), (opts.colW[iCell]/ONEPT));
				arrCellsLines.push( lines );

				// 4: Keep track of max line count within all row cells
				if ( lines.length > intMaxLineCnt ) { intMaxLineCnt = lines.length; intMaxColIdx = iCell; }
				var lineHeight = inch2Emu((cell.opts.font_size || opts.font_size || DEF_FONT_SIZE) * LINEH_MODIFIER / 100);
				// NOTE: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
				if (cell.opts && cell.opts.rowspan) lineHeight = 0;
				if ( Array.isArray(cell.opts.margin) && cell.opts.margin[0] ) lineHeight += cell.opts.margin[0] / intMaxLineCnt;
				if ( Array.isArray(cell.opts.margin) && cell.opts.margin[2] ) lineHeight += cell.opts.margin[2] / intMaxLineCnt;
				arrCellsLineHeights.push( Math.round(lineHeight) );
				if (opts.debug) console.log('lineHeight (in)....: '+ (lineHeight/EMU) + ' ... (intMaxLineCnt = '+intMaxLineCnt+')');
			});

			// C: AUTO-PAGING: Add text one-line-a-time to this row's cells until: lines are exhausted OR table H limit is hit
			for (var idx=0; idx<intMaxLineCnt; idx++) {
				// 1: Add the current line to cell
				for (var col=0; col<arrCellsLines.length; col++) {
					// A: Commit this slide to Presenation if table Height limit is hit
					if ( emuTabCurrH + arrCellsLineHeights[intMaxColIdx] > emuSlideTabH ) {
						// 1: Add the current row to table
						// NOTE: Edge cases can occur where we create a new slide only to have no more lines
						// ....: and then a blank row sits at the bottom of a table!
						// ....: Hence, we verify all cells have text before adding this final row.
						$.each(currRow, function(i,cell){
							if (cell.text.length > 0 ) {
								// IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
								arrRows.push( $.extend(true, [], currRow) );
								return false; // break out of .each loop
							}
						});
						// 2: Add new Slide with current array of table rows
						arrObjSlides.push( $.extend(true, [], arrRows) );
						// 3: Empty rows for new Slide
						arrRows.length = 0;
						// 4: Reset current table height for new Slide
						emuTabCurrH = 0; // This row's emuRowH w/b added below
						// 5: Empty current row's text (continue adding lines where we left off below)
						$.each(currRow,function(i,cell){ cell.text = ''; });
						// 6: Auto-Paging Options: addHeaderToEach
						if ( opts.addHeaderToEach && arrObjTabHeadRows ) {
							var headRow = [];
							$.each(arrObjTabHeadRows[0], function(iCell,cell){
								headRow.push({ text:cell.text, opts:cell.opts });
								var lines = parseTextToLines(cell.text, cell.opts.font_size, (opts.colW[iCell]/ONEPT));
								if ( lines.length > intMaxLineCnt ) { intMaxLineCnt = lines.length; intMaxColIdx = iCell; }
							});
							arrRows.push( $.extend(true, [], headRow) );
						}
					}

					// B: Add next line of text to this cell
					if ( arrCellsLines[col][idx] ) currRow[col].text += arrCellsLines[col][idx];
				}

				// 2: Add this new rows H to overall (The cell with the longest line array is the one we use as the determiner for overall row Height)
				emuTabCurrH += arrCellsLineHeights[intMaxColIdx];
			}

			// D: Flush row buffer - Add the current row to table, then truncate row cell array
			// IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
			if (currRow.length) arrRows.push( $.extend(true,[],currRow) );
			currRow.length = 0;
		});

		// STEP 4-2: Flush final row buffer to slide
		arrObjSlides.push( $.extend(true,[],arrRows) );

		// LAST:
		if (opts.debug) { console.log('arrObjSlides count = '+arrObjSlides.length); console.log(arrObjSlides); }
		return arrObjSlides;
	}

	/* =======================================================================================================
	|
	#     #  #     #  #             #####
	 #   #   ##   ##  #            #     #  ######  #    #  ######  #####     ##    #####  #   ####   #    #
	  # #    # # # #  #            #        #       ##   #  #       #    #   #  #     #    #  #    #  ##   #
	   #     #  #  #  #            #  ####  #####   # #  #  #####   #    #  #    #    #    #  #    #  # #  #
	  # #    #     #  #            #     #  #       #  # #  #       #####   ######    #    #  #    #  #  # #
	 #   #   #     #  #            #     #  #       #   ##  #       #   #   #    #    #    #  #    #  #   ##
	#     #  #     #  #######       #####   ######  #    #  ######  #    #  #    #    #    #   ####   #    #
	|
	=========================================================================================================
	*/

	function genXmlBodyProperties( objOptions ) {
		var bodyProperties = '<a:bodyPr';

		if ( objOptions && objOptions.bodyProp ) {
			// A: Enable or disable textwrapping none or square:
			( objOptions.bodyProp.wrap ) ? bodyProperties += ' wrap="' + objOptions.bodyProp.wrap + '" rtlCol="0"' : bodyProperties += ' wrap="square" rtlCol="0"';

			// B: Set anchorPoints bottom, center or top:
			if ( objOptions.bodyProp.anchor    ) bodyProperties += ' anchor="' + objOptions.bodyProp.anchor + '"';
			if ( objOptions.bodyProp.anchorCtr ) bodyProperties += ' anchorCtr="' + objOptions.bodyProp.anchorCtr + '"';

			// C: Textbox margins [padding]:
			if ( objOptions.bodyProp.bIns || objOptions.bodyProp.bIns == 0 ) bodyProperties += ' bIns="' + objOptions.bodyProp.bIns + '"';
			if ( objOptions.bodyProp.lIns || objOptions.bodyProp.lIns == 0 ) bodyProperties += ' lIns="' + objOptions.bodyProp.lIns + '"';
			if ( objOptions.bodyProp.rIns || objOptions.bodyProp.rIns == 0 ) bodyProperties += ' rIns="' + objOptions.bodyProp.rIns + '"';
			if ( objOptions.bodyProp.tIns || objOptions.bodyProp.tIns == 0 ) bodyProperties += ' tIns="' + objOptions.bodyProp.tIns + '"';

			// D: Close <a:bodyPr element
			bodyProperties += '>';

			// E: NEW: Add auto-fit type tags
			if ( objOptions.shrinkText ) bodyProperties += '<a:normAutofit fontScale="85000" lnSpcReduction="20000" />'; // MS-PPT > Format Shape > Text Options: "Shrink text on overflow"
			else if ( objOptions.bodyProp.autoFit !== false ) bodyProperties += '<a:spAutoFit/>'; // MS-PPT > Format Shape > Text Options: "Resize shape to fit text"

			// LAST: Close bodyProp
			bodyProperties += '</a:bodyPr>';
		}
		else {
			// DEFAULT:
			bodyProperties += ' wrap="square" rtlCol="0"></a:bodyPr>';
		}

		return bodyProperties;
	}

	function genXmlTextCommand( text_info, text_string, slide_obj, slide_num ) {
		var area_opt_data = genXmlTextData( text_info, slide_obj );
		var parsedText;
		//
		var startInfo = '<a:rPr lang="en-US"' + area_opt_data.font_size + area_opt_data.bold
				+ area_opt_data.italic + area_opt_data.underline + area_opt_data.char_spacing
				+ ' dirty="0" smtClean="0"' + (area_opt_data.rpr_info != '' ? ('>' + area_opt_data.rpr_info) : '/>')
				+ '<a:t>';
		var endTag = '</a:r>';
		var outData = '<a:r>' + startInfo;

		if ( text_string.field ) {
			endTag = '</a:fld>';
			var outTextField = pptxFields[text_string.field];
			if ( outTextField === null ) {
				for ( var fieldIntName in pptxFields ) {
					if ( pptxFields[fieldIntName] === text_string.field ) {
						outTextField = text_string.field;
						break;
					}
				}

				if ( outTextField === null ) outTextField = 'datetime';
			}

			outData = '<a:fld id="{' + gen_private.plugs.type.msoffice.makeUniqueID ( '5C7A2A3D' ) + '}" type="' + outTextField + '">' + startInfo;
			outData += CreateFieldText( outTextField, slide_num );

		}
		else {
			// LINE-BREAKS/MULTI-LINE: Split text into multi-p:
			parsedText = text_string.split("\n");
			if ( parsedText.length > 1 ) {
				var outTextData = '';
				for ( var i = 0, total_size_i = parsedText.length; i < total_size_i; i++ ) {
					// ISSUE #34: Text with linebreaks doesnt allow for align formatting
					// IMPORTANT: Closing <p> is fine and all, but one option `align` isnt set at the <r>un-level - its a `pPr` so do that here:
					var outData = ( (text_info.align && (i > 0 || text_info.lineIdx > 0)) ? '<a:pPr algn="'+ text_info.align.toLowerCase().replace(/^c.*/i,'ctr').replace(/^l.*/i,'l').replace(/^r.*/i,'r') +'"/>' : '') + '<a:r>' + startInfo;
					outTextData += outData + decodeXmlEntities(parsedText[i]);
					// Stop/Start <p>aragraph as long as there is more lines ahead (otherwise its closed at the end of this function)
					if ( (i + 1) < total_size_i ) {
						outTextData += '</a:t></a:r>'+ ( text_info.align ? '<a:endParaRPr lang="en-US"/>' : '') +'</a:p><a:p>';
					}
				}
				outData = outTextData;
			}
			else {
				// Handle cases where addText `text` was an array of objects - if a text object doesnt contain a '\n' it still need alignment!
				// The first pPr-align is done in makeXml - use line countr to ensure we only add subsequently as needed
				var outData = ( (text_info.align && (i > 0 || text_info.lineIdx > 0)) ? '<a:pPr algn="'+ text_info.align.toLowerCase().replace(/^c.*/i,'ctr').replace(/^l.*/i,'l').replace(/^r.*/i,'r') +'"/>' : '') + '<a:r>' + startInfo;
				outData += decodeXmlEntities(text_string);
			}
		}

		// Return paragraph with text run
		return outData + '</a:t>' + endTag + (text_info.breakLine ? '</a:p><a:p>' : '');
	}

	function genXmlTextData( text_info, slide_obj ) {
		var out_obj = {};

		out_obj.font_size = '';
		out_obj.bold = '';
		out_obj.italic = '';
		out_obj.underline = '';
		out_obj.rpr_info = '';
        out_obj.char_spacing = '';

		if ( typeof text_info == 'object' ) {
			if ( text_info.bold ) {
				out_obj.bold = ' b="1"';
			}

			if ( text_info.italic ) {
				out_obj.italic = ' i="1"';
			}

			if ( text_info.underline ) {
				out_obj.underline = ' u="sng"';
			}

			if ( text_info.font_size ) {
				out_obj.font_size = ' sz="' + text_info.font_size + '00"';
			}

			if ( text_info.char_spacing ) {
				out_obj.char_spacing = ' spc="' + (text_info.char_spacing * 100) + '"';
				// must also disable kerning; otherwise text won't actually expand
				out_obj.char_spacing += ' kern="0"';
			}

			if ( text_info.color ) {
				out_obj.rpr_info += genXmlColorSelection( text_info.color );
			}
			else if ( slide_obj && slide_obj.color ) {
				out_obj.rpr_info += genXmlColorSelection( slide_obj.color );
			}

			if ( text_info.font_face ) {
				out_obj.rpr_info += '<a:latin typeface="' + text_info.font_face + '" pitchFamily="34" charset="0"/><a:cs typeface="' + text_info.font_face + '" pitchFamily="34" charset="0"/>';
			}
		}
		else {
			if ( slide_obj && slide_obj.color ) out_obj.rpr_info += genXmlColorSelection( slide_obj.color );
		}

		if ( out_obj.rpr_info != '' ) out_obj.rpr_info += '</a:rPr>';

		return out_obj;
	}

	function genXmlColorSelection( color_info, back_info ) {
		var colorVal;
		var fillType = 'solid';
		var internalElements = '';
		var outText = '';

		if ( back_info && typeof back_info === 'string' ) {
			outText += '<p:bg><p:bgPr>';
			outText += genXmlColorSelection( back_info.replace('#',''), false );
			outText += '<a:effectLst/>';
			outText += '</p:bgPr></p:bg>';
		}

		if ( color_info ) {
			if ( typeof color_info == 'string' ) colorVal = color_info;
			else {
				if ( color_info.type  ) fillType = color_info.type;
				if ( color_info.color ) colorVal = color_info.color;
				if ( color_info.alpha ) internalElements += '<a:alpha val="' + (100 - color_info.alpha) + '000"/>';
			}

			switch ( fillType ) {
				case 'solid':
					outText += '<a:solidFill><a:srgbClr val="' + colorVal + '">' + internalElements + '</a:srgbClr></a:solidFill>';
					break;
			}
		}

		return outText;
	}

	// XML-GEN: First 6 functions create the base /ppt files

	function makeXmlContTypes() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF;
		strXml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
		strXml += ' <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		strXml += ' <Default Extension="xml" ContentType="application/xml"/>';
		strXml += ' <Default Extension="jpeg" ContentType="image/jpeg"/>';
		strXml += ' <Default Extension="png" ContentType="image/png"/>';
		strXml += ' <Default Extension="gif" ContentType="image/gif"/>';
		strXml += ' <Default Extension="m4v" ContentType="video/mp4"/>'; // hard-coded as extn!=type
		strXml += ' <Default Extension="mp4" ContentType="video/mp4"/>'; // same here, we have to add as it wont be added in loop below
		gObjPptx.slides.forEach(function(slide,idx){
			slide.rels.forEach(function(rel,idy){
				if ( rel.type != 'image' && rel.type != 'online' && rel.extn != 'm4v' && strXml.indexOf(rel.type) == -1 )
					strXml += ' <Default Extension="'+ rel.extn +'" ContentType="'+ rel.type +'"/>';
			});
		});
		strXml += ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
		strXml += ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
		strXml += ' <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>';
		strXml += ' <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>';
		strXml += ' <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>';
		strXml += ' <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
		strXml += ' <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>';
		gObjPptx.slides.forEach(function(slide,idx){
			strXml += '<Override PartName="/ppt/slideMasters/slideMaster'+ (idx+1) +'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>';
			strXml += '<Override PartName="/ppt/slideLayouts/slideLayout'+ (idx+1) +'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
			strXml += '<Override PartName="/ppt/slides/slide'            + (idx+1) +'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
		});
		strXml += '</Types>';

		return strXml;
	}

	function makeXmlRootRels() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
					+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
					+ '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
					+ '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
					+ '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
					+ '</Relationships>';
		return strXml;
	}

	function makeXmlApp() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n\
					<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\
						<TotalTime>0</TotalTime>\
						<Words>0</Words>\
						<Application>Microsoft Office PowerPoint</Application>\
						<PresentationFormat>On-screen Show (4:3)</PresentationFormat>\
						<Paragraphs>0</Paragraphs>\
						<Slides>'+ gObjPptx.slides.length +'</Slides>\
						<Notes>0</Notes>\
						<HiddenSlides>0</HiddenSlides>\
						<MMClips>0</MMClips>\
						<ScaleCrop>false</ScaleCrop>\
						<HeadingPairs>\
						  <vt:vector size="4" baseType="variant">\
						    <vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>\
						    <vt:variant><vt:i4>1</vt:i4></vt:variant>\
						    <vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>\
						    <vt:variant><vt:i4>'+ gObjPptx.slides.length +'</vt:i4></vt:variant>\
						  </vt:vector>\
						</HeadingPairs>\
						<TitlesOfParts>';
		strXml += '<vt:vector size="'+ (gObjPptx.slides.length+1) +'" baseType="lpstr">';
		strXml += '<vt:lpstr>Office Theme</vt:lpstr>';
		$.each(gObjPptx.slides, function(idx,slideObj){ strXml += '<vt:lpstr>Slide '+ (idx+1) +'</vt:lpstr>'; });
		strXml += ' </vt:vector>\
						</TitlesOfParts>\
						<Company>PptxGenJS</Company>\
						<LinksUpToDate>false</LinksUpToDate>\
						<SharedDoc>false</SharedDoc>\
						<HyperlinksChanged>false</HyperlinksChanged>\
						<AppVersion>15.0000</AppVersion>\
					</Properties>';
		return strXml;
	}

	function makeXmlCore() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n\
						<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"\
							 xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"\
							 xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\
							<dc:title>'+ gObjPptx.title +'</dc:title>\
							<dc:creator>PptxGenJS</dc:creator>\
							<cp:lastModifiedBy>PptxGenJS</cp:lastModifiedBy>\
							<cp:revision>1</cp:revision>\
							<dcterms:created xsi:type="dcterms:W3CDTF">'+ new Date().toISOString() +'</dcterms:created>\
							<dcterms:modified xsi:type="dcterms:W3CDTF">'+ new Date().toISOString() +'</dcterms:modified>\
						</cp:coreProperties>';
		return strXml;
	}

	function makeXmlPresentationRels() {
		var intRelNum = 0;
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
					+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

		strXml += '  <Relationship Id="rId1" Target="slideMasters/slideMaster1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
		intRelNum++;

		for ( var idx=1; idx<=gObjPptx.slides.length; idx++ ) {
			intRelNum++;
			strXml += '  <Relationship Id="rId'+ intRelNum +'" Target="slides/slide'+ idx +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>';
		}
		intRelNum++;
		strXml += '  <Relationship Id="rId'+  intRelNum    +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>'
				+ '  <Relationship Id="rId'+ (intRelNum+1) +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>'
				+ '  <Relationship Id="rId'+ (intRelNum+2) +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
				+ '  <Relationship Id="rId'+ (intRelNum+3) +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>'
				+ '</Relationships>';
		return strXml;
	}

	function makeXmlSlideLayout() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
		strXml += '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1">\r\n'
				+ '<p:cSld name="Title Slide">'
				+ '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
				+ '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="2130425"/><a:ext cx="7772400" cy="1470025"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
				+ '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1371600" y="3886200"/><a:ext cx="6400800" cy="1752600"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle>'
				+ '  <a:lvl1pPr marL="0"       indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr>'
				+ '  <a:lvl2pPr marL="457200"  indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl2pPr>'
				+ '  <a:lvl3pPr marL="914400"  indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl3pPr>'
				+ '  <a:lvl4pPr marL="1371600" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl4pPr>'
				+ '  <a:lvl5pPr marL="1828800" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl5pPr>'
				+ '  <a:lvl6pPr marL="2286000" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl6pPr>'
				+ '  <a:lvl7pPr marL="2743200" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl7pPr>'
				+ '  <a:lvl8pPr marL="3200400" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl8pPr>'
				+ '  <a:lvl9pPr marL="3657600" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl9pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master subtitle style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>'
				+ '<p:cNvPr id="4" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{F8166F1F-CE9B-4651-A6AA-CD717754106B}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>01/01/2016</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>'
				+ '<p:cNvPr id="5" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>'
				+ '<p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="'+SLDNUMFLDID+'" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t></a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld>'
				+ '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
		//
		return strXml;
	}

	// XML-GEN: Next 5 functions run 1-N times (once for each Slide)

	/**
	 * Generates the XML slide resource from a Slide object
	 * @param {Object} inSlide - The slide object to transform into XML
	 * @return {string} strSlideXml - Slide OOXML
	*/
	function makeXmlSlide(inSlide) {
		var intTableNum = 1;

		// STEP 1: Start slide XML
		var strSlideXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF;
		strSlideXml += '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">';
		strSlideXml += '<p:cSld name="'+ inSlide.name +'">';

		// STEP 2: Add background color or background image (if any)
		// A: Background color
		if ( inSlide.slide.back ) strSlideXml += genXmlColorSelection(false, inSlide.slide.back);
		// B: Add background image (using Strech) (if any)
		if ( inSlide.slide.bkgdImgRid ) {
			// FIXME: We should be doing this in the slideLayout...
			strSlideXml += '<p:bg>'
						+ '<p:bgPr><a:blipFill dpi="0" rotWithShape="1">'
						+ '<a:blip r:embed="rId'+ inSlide.slide.bkgdImgRid +'"><a:lum/></a:blip>'
						+ '<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
						+ '<a:effectLst/></p:bgPr>'
						+ '</p:bg>';
		}

		// STEP 3: Continue slide by starting spTree node
		strSlideXml += '<p:spTree>';
		strSlideXml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>';
		strSlideXml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>';
		strSlideXml += '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

		// STEP 4: Add slide numbers if selected
		if ( inSlide.slideNumberObj || inSlide.hasSlideNumber ) {
			var numberX = (EMU * 0.3); // default and/or inSlide.hasSlideNumber value
			var numberY = (0.90 * gObjPptx.pptLayout.height); // default and/or inSlide.hasSlideNumber value

			if ( inSlide.slideNumberObj && inSlide.slideNumberObj.x ) numberX = getSmartParseNumber(inSlide.slideNumberObj.x, 'X');
			if ( inSlide.slideNumberObj && inSlide.slideNumberObj.y ) numberY = getSmartParseNumber(inSlide.slideNumberObj.y, 'Y');

			strSlideXml += '<p:sp>'
				+ '  <p:nvSpPr>'
				+ '  <p:cNvPr id="25" name="Shape 25"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr></p:nvSpPr>'
				+ '  <p:spPr>'
				+ '    <a:xfrm><a:off x="'+ numberX +'" y="'+ numberY +'"/><a:ext cx="400000" cy="300000"/></a:xfrm>'
				+ '    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
				+ '    <a:extLst>'
				+ '      <a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">'
				+ '      <ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext>'
				+ '    </a:extLst>'
				+ '  </p:spPr>'
				+ '  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr/><a:fld id="'+SLDNUMFLDID+'" type="slidenum"/></a:p></p:txBody>'
				+ '</p:sp>';
		}

		// STEP 5: Loop over all Slide.data objects and add them to this slide ===============================
		$.each(inSlide.data, function(idx,slideObj){
			var x = 0, y = 0, cx = (EMU*10), cy = 0;
			var moreStyles = '', moreStylesAttr = '', outStyles = '', styleData = '', locationAttr = '';
			var shapeType = null;

			// A: Set option vars
			if ( slideObj.options ) {
				if ( slideObj.options.w  || slideObj.options.w  == 0 ) slideObj.options.cx = slideObj.options.w;
				if ( slideObj.options.h  || slideObj.options.h  == 0 ) slideObj.options.cy = slideObj.options.h;
				//
				if ( slideObj.options.x  || slideObj.options.x  == 0 )  x = getSmartParseNumber( slideObj.options.x , 'X' );
				if ( slideObj.options.y  || slideObj.options.y  == 0 )  y = getSmartParseNumber( slideObj.options.y , 'Y' );
				if ( slideObj.options.cx || slideObj.options.cx == 0 ) cx = getSmartParseNumber( slideObj.options.cx, 'X' );
				if ( slideObj.options.cy || slideObj.options.cy == 0 ) cy = getSmartParseNumber( slideObj.options.cy, 'Y' );
				//
				if ( slideObj.options.shape  ) shapeType = getShapeInfo( slideObj.options.shape );
				//
				if ( slideObj.options.flipH  ) locationAttr += ' flipH="1"';
				if ( slideObj.options.flipV  ) locationAttr += ' flipV="1"';
				if ( slideObj.options.rotate ) locationAttr += ' rot="' + ( (slideObj.options.rotate > 360 ? (slideObj.options.rotate - 360) : slideObj.options.rotate) * 60000 ) + '"';
			}

			// B: Add TABLE / TEXT / IMAGE / MEDIA to current Slide ----------------------------
			switch ( slideObj.type ) {
				case 'table':
					// FIRST: Ensure we have rows - otherwise, bail!
					if ( !slideObj.arrTabRows || (Array.isArray(slideObj.arrTabRows) && slideObj.arrTabRows.length == 0) ) break;

					// Set table vars
					var objTableGrid = {};
					var arrTabRows = slideObj.arrTabRows;
					var objTabOpts = slideObj.options;
					var intColCnt = 0, intColW = 0;

					// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
					// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
					for (var tmp=0; tmp<arrTabRows[0].length; tmp++) {
						intColCnt += ( arrTabRows[0][tmp] && arrTabRows[0][tmp].opts && arrTabRows[0][tmp].opts.colspan ) ? Number(arrTabRows[0][tmp].opts.colspan) : 1;
					}

					// STEP 1: Start Table XML =============================
					// NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
					var strXml = '<p:graphicFrame>'
							+ '  <p:nvGraphicFramePr>'
							+ '    <p:cNvPr id="'+ (intTableNum*inSlide.numb + 1) +'" name="Table '+ (intTableNum*inSlide.numb) +'"/>'
							+ '    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
							+ '    <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>'
							+ '  </p:nvGraphicFramePr>'
							+ '  <p:xfrm>'
							+ '    <a:off  x="'+ (x  || EMU) +'"  y="'+ (y  || EMU) +'"/>'
							+ '    <a:ext cx="'+ (cx || EMU) +'" cy="'+ (cy || EMU) +'"/>'
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
					if ( Array.isArray(objTabOpts.colW) ) {
						strXml += '<a:tblGrid>';
						for ( var col=0; col<intColCnt; col++ ) {
							strXml += '  <a:gridCol w="'+ Math.round(inch2Emu(objTabOpts.colW[col]) || (slideObj.options.cx/intColCnt)) +'"/>';
						}
						strXml += '</a:tblGrid>';
					}
					// B: Table Width provided without colW? Then distribute cols
					else {
						intColW = ( objTabOpts.colW ? objTabOpts.colW : EMU );
						if ( slideObj.options.cx && !objTabOpts.colW ) intColW = Math.round( slideObj.options.cx / intColCnt ); // FIX: Issue#12
						strXml += '<a:tblGrid>';
						for ( var col=0; col<intColCnt; col++ ) { strXml += '<a:gridCol w="'+ intColW +'"/>'; }
						strXml += '</a:tblGrid>';
					}

					// STEP 3: Build our row arrays into an actual grid to match the XML we will be building next (ISSUE #36)
					// Note row arrays can arrive "lopsided" as in row1:[1,2,3] row2:[3] when first two cols rowspan!,
					// so a simple loop below in XML building wont suffice to build table right.
					// We have to build an actual grid now
					/*
						EX: (A0:rowspan=3, B1:rowspan=2, C1:colspan=2)

						/------|------|------|------\
						|  A0  |  B0  |  C0  |  D0  |
						|      |  B1  |  C1  |      |
						|      |      |  C2  |  D2  |
						\------|------|------|------/
					*/
 					$.each(arrTabRows, function(rIdx,row){
						// A: Create row if needed (recall one may be created in loop below for rowspans, so dont assume we need to create one each iteration)
						if ( !objTableGrid[rIdx] ) objTableGrid[rIdx] = {};

						// B: Loop over all cells
						$(row).each(function(cIdx,cell){
							// DESIGN: NOTE: Row cell arrays can be "uneven" (diff cell count in each) due to rowspan/colspan
							// Therefore, for each cell we run 0->colCount to determien the correct slot for it to reside
							// as the uneven/mixed nature of the data means we cannot use the cIdx value alone.
							// E.g.: the 2nd element in the row array may actually go into the 5th table grid row cell b/c of colspans!
							for (var idx=0; (cIdx+idx)<intColCnt; idx++) {
								var currColIdx = (cIdx + idx);

								if ( !objTableGrid[rIdx][currColIdx] ) {
									// A: Set this cell
									objTableGrid[rIdx][currColIdx] = cell;

									// B: Handle `colspan` or `rowspan` (a {cell} cant have both! FIXME: FUTURE: ROWSPAN & COLSPAN in same cell)
									if ( cell && cell.opts && cell.opts.colspan && !isNaN(Number(cell.opts.colspan)) ) {
										for (var idy=1; idy<Number(cell.opts.colspan); idy++) {
											objTableGrid[rIdx][currColIdx+idy] = {"hmerge":true, text:"hmerge"};
										}
									}
									else if ( cell && cell.opts && cell.opts.rowspan && !isNaN(Number(cell.opts.rowspan)) ) {
										for (var idz=1; idz<Number(cell.opts.rowspan); idz++) {
											if ( !objTableGrid[rIdx+idz] ) objTableGrid[rIdx+idz] = {};
											objTableGrid[rIdx+idz][currColIdx] = {"vmerge":true, text:"vmerge"};
										}
									}

									// C: Break out of colCnt loop now that slot has been filled
									break;
								}
							}
						});
					});
					if ( objTabOpts.debug ) {
						console.table(objTableGrid);
						var arrText = [];
						$.each(objTableGrid, function(i,row){ var arrRow = []; $.each(row,function(i,cell){ arrRow.push(cell.text); }); arrText.push(arrRow); });
						console.table( arrText );
					}

					// STEP 4: Build table rows/cells ============================
					$.each(objTableGrid, function(rIdx,rowObj){
						// A: Table Height provided without rowH? Then distribute rows
						var intRowH = 0; // IMPORTANT: Default must be zero for auto-sizing to work
						if ( Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx] ) intRowH = inch2Emu(Number(objTabOpts.rowH[rIdx]));
						else if ( objTabOpts.rowH && !isNaN(Number(objTabOpts.rowH)) ) intRowH = inch2Emu(Number(objTabOpts.rowH));
						else if ( slideObj.options.cy || slideObj.options.h ) intRowH = ( slideObj.options.h ? inch2Emu(slideObj.options.h) : slideObj.options.cy) / arrTabRows.length;

						// B: Start row
						strXml += '<a:tr h="'+ intRowH +'">';

						// C: Loop over each CELL
						$.each(rowObj, function(cIdx,cell){
							// FIRST: "hmerge" cells are just place-holders in the table grid - skip those and go to next cell
							if ( cell.hmerge ) return;

							// 1: OPTIONS: Build/set cell options (blocked for code folding)
							{
								// A: Create cell if needed (handle [null] and other manner of junk values)
								// IMPORTANT: MS-PPTX PROBLEM: using '' will cause PPT to use its own default font/size! (Arial/18 in US)
								// SOLN: Pass a space instead to cement formatting options (Issue #20)
								if ( !cell ) cell = { text:' ' };

								// B: Load/Create options
								var cellOpts = cell.options || cell.opts || {};

								// C: Do Important/Override Opts
								// Feature: TabOpts Default Values (tabOpts being used when cellOpts dont exist):
								// SEE: http://officeopenxml.com/drwTableCellProperties-alignment.php
								$.each(['align','bold','border','color','fill','font_face','font_size','underline','valign'], function(i,name){
									if ( objTabOpts[name] && ! cellOpts[name]) cellOpts[name] = objTabOpts[name];
								});

								var cellB       = (cellOpts.bold)       ? ' b="1"' : ''; // [0,1] or [false,true]
								var cellU       = (cellOpts.underline)  ? ' u="sng"' : ''; // [none,sng (single), et al.]
								var cellFont    = (cellOpts.font_face)  ? ' <a:latin typeface="'+ cellOpts.font_face +'"/>' : '';
								var cellFontPt  = (cellOpts.font_size)  ? ' sz="'+ cellOpts.font_size +'00"' : '';
								var cellAlign   = (cellOpts.align)      ? ' algn="'+ cellOpts.align.replace(/^c$/i,'ctr').replace('center','ctr').replace('left','l').replace('right','r') +'"' : '';
								var cellValign  = (cellOpts.valign)     ? ' anchor="'+ cellOpts.valign.replace(/^c$/i,'ctr').replace(/^m$/i,'ctr').replace('center','ctr').replace('middle','ctr').replace('top','t').replace('btm','b').replace('bottom','b') +'"' : '';
								var cellColspan = (cellOpts.colspan)    ? ' gridSpan="'+ cellOpts.colspan +'"' : '';
								var cellRowspan = (cellOpts.rowspan)    ? ' rowSpan="'+ cellOpts.rowspan +'"' : '';
								var cellFontClr = ((cell.optImp && cell.optImp.color) || cellOpts.color) ? ' <a:solidFill><a:srgbClr val="'+ ((cell.optImp && cell.optImp.color) || cellOpts.color.replace('#','')) +'"/></a:solidFill>' : '';
								var cellFill    = ((cell.optImp && cell.optImp.fill)  || cellOpts.fill ) ? ' <a:solidFill><a:srgbClr val="'+ ((cell.optImp && cell.optImp.fill) || cellOpts.fill.replace('#','')) +'"/></a:solidFill>' : '';
								if ( cellOpts.margin == 0 ) cellOpts.margin = [0, 0, 0, 0]; // Allow 0 (zero) as a margin - otherwise our fancy short-circuit eval doesnt allow for thhis condition - solve here for doc and sanity sake
								var cellMargin  = (cellOpts.margin && Array.isArray(cellOpts.margin) ? cellOpts.margin : (cellOpts.margin || '') );

								// Margin/Padding:
								if ( cellMargin ) {
									var arrMargin = [];
									if ( Array.isArray(cellMargin) ) arrMargin = cellMargin;
									else if ( !isNaN(Number(cellMargin)) ) arrMargin = [Number(cellMargin), Number(cellMargin), Number(cellMargin), Number(cellMargin)];
									else arrMargin = [0, 0, 0, 0];
									cellMargin = ' marL="'+ arrMargin[3]*ONEPT +'" marR="'+ arrMargin[1]*ONEPT +'" marT="'+ arrMargin[0]*ONEPT +'" marB="'+ arrMargin[2]*ONEPT +'"';
								}
							}

							// 2: Cell Content: Either the text element or the cell itself (for when users just pass a string - no object or options)
							var strCellText = ( typeof cell === 'object' ? cell.text : cell );

							// FIXME: Cell NOWRAP property (text wrap: add to a:tcPr (horzOverflow="overflow" or whatev opts exist)

							// 3: ROWSPAN: Add dummy cells for any active rowspan
							if ( cell.vmerge ) {
								strXml += '<a:tc vMerge="1"><a:tcPr/></a:tc>';
								return;
							}

							// 4: Start Table Cell, add Align, add Text content =========================
							strXml += '<a:tc'+ cellColspan + cellRowspan +'>'
									+ '  <a:txBody>'
									+ '    <a:bodyPr/>'
									+ '    <a:lstStyle/>'
									+ '    <a:p>'
									+ '      <a:pPr'+ cellAlign +'/>'
									+ '      <a:r>'
									+ '        <a:rPr lang="en-US" dirty="0" smtClean="0"'+ cellFontPt + cellB + cellU +'>'+ cellFontClr + cellFont +'</a:rPr>'
									+ '        <a:t>'+ decodeXmlEntities(strCellText) +'</a:t>'
									+ '      </a:r>'
									+ '      <a:endParaRPr lang="en-US" dirty="0"/>'
									+ '    </a:p>'
									+ '  </a:txBody>'
									+ '  <a:tcPr'+ cellMargin + cellValign +'>';

							// 5: Borders: Add any borders
							if ( cellOpts.border && typeof cellOpts.border === 'string' ) {
								strXml += '  <a:lnL w="'+ ONEPT +'" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="'+ cellOpts.border +'"/></a:solidFill></a:lnL>';
								strXml += '  <a:lnR w="'+ ONEPT +'" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="'+ cellOpts.border +'"/></a:solidFill></a:lnR>';
								strXml += '  <a:lnT w="'+ ONEPT +'" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="'+ cellOpts.border +'"/></a:solidFill></a:lnT>';
								strXml += '  <a:lnB w="'+ ONEPT +'" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="'+ cellOpts.border +'"/></a:solidFill></a:lnB>';
							}
							else if ( cellOpts.border && Array.isArray(cellOpts.border) ) {
								$.each([ {idx:3,name:'lnL'}, {idx:1,name:'lnR'}, {idx:0,name:'lnT'}, {idx:2,name:'lnB'} ], function(i,obj){
									if ( cellOpts.border[obj.idx] ) {
										var strC = '<a:solidFill><a:srgbClr val="'+ ((cellOpts.border[obj.idx].color) ? cellOpts.border[obj.idx].color : '666666') +'"/></a:solidFill>';
										var intW = (cellOpts.border[obj.idx] && (cellOpts.border[obj.idx].pt || cellOpts.border[obj.idx].pt == 0)) ? (ONEPT * Number(cellOpts.border[obj.idx].pt)) : ONEPT;
										strXml += '<a:'+ obj.name +' w="'+ intW +'" cap="flat" cmpd="sng" algn="ctr">'+ strC +'</a:'+ obj.name +'>';
									}
									else strXml += '<a:'+ obj.name +' w="0"><a:miter lim="400000" /></a:'+ obj.name +'>';
								});
							}
							else if ( cellOpts.border && typeof cellOpts.border === 'object' ) {
								var intW = (cellOpts.border && (cellOpts.border.pt || cellOpts.border.pt == 0) ) ? (ONEPT * Number(cellOpts.border.pt)) : ONEPT;
								var strClr = '<a:solidFill><a:srgbClr val="'+ ((cellOpts.border.color) ? cellOpts.border.color.replace('#','') : '666666') +'"/></a:solidFill>';
								var strAttr = '<a:prstDash val="';
								strAttr += ((cellOpts.border.type && cellOpts.border.type.toLowerCase().indexOf('dash') > -1) ? "sysDash" : "solid" );
								strAttr += '"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>';
								// *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
								strXml += '<a:lnL w="'+ intW +'" cap="flat" cmpd="sng" algn="ctr">'+ strClr + strAttr +'</a:lnL>';
								strXml += '<a:lnR w="'+ intW +'" cap="flat" cmpd="sng" algn="ctr">'+ strClr + strAttr +'</a:lnR>';
								strXml += '<a:lnT w="'+ intW +'" cap="flat" cmpd="sng" algn="ctr">'+ strClr + strAttr +'</a:lnT>';
								strXml += '<a:lnB w="'+ intW +'" cap="flat" cmpd="sng" algn="ctr">'+ strClr + strAttr +'</a:lnB>';
								// *** IMPORTANT! *** LRTB order matters!
							}

							// 6: Close cell Properties & Cell
							strXml += cellFill
									+ '  </a:tcPr>'
									+ ' </a:tc>';

							// LAST: COLSPAN: Add a 'merged' col for each column being merged (SEE: http://officeopenxml.com/drwTableGrid.php)
							if ( cellOpts.colspan ) {
								for (var tmp=1; tmp<Number(cellOpts.colspan); tmp++) { strXml += '<a:tc hMerge="1"><a:tcPr/></a:tc>'; }
							}
						});

						// D: Complete row
						strXml += '</a:tr>';
					});

					// STEP 5: Complete table
					strXml += '      </a:tbl>'
							+ '    </a:graphicData>'
							+ '  </a:graphic>'
							+ '</p:graphicFrame>';

					// STEP 6: Set table XML
					strSlideXml += strXml;

					// LAST: Increment counter
					intTableNum++;
					break;

				case 'text':
					// Lines can have zero cy, but text should not
					if ( !slideObj.options.line && cy == 0 ) cy = (EMU * 0.3);

					// Margin/Padding/Inset for textboxes
					if ( slideObj.options.margin && Array.isArray(slideObj.options.margin) ) {
						slideObj.options.bodyProp.lIns = (slideObj.options.margin[0] * ONEPT || 0);
						slideObj.options.bodyProp.rIns = (slideObj.options.margin[1] * ONEPT || 0);
						slideObj.options.bodyProp.bIns = (slideObj.options.margin[2] * ONEPT || 0);
						slideObj.options.bodyProp.tIns = (slideObj.options.margin[3] * ONEPT || 0);
					}
					else if ( (slideObj.options.margin || slideObj.options.margin == 0) && Number.isInteger(slideObj.options.margin) ) {
						slideObj.options.bodyProp.lIns = (slideObj.options.margin * ONEPT);
						slideObj.options.bodyProp.rIns = (slideObj.options.margin * ONEPT);
						slideObj.options.bodyProp.bIns = (slideObj.options.margin * ONEPT);
						slideObj.options.bodyProp.tIns = (slideObj.options.margin * ONEPT);
					}

					var effectsList = '';
					if ( shapeType == null ) shapeType = getShapeInfo(null);

					// A: Start Shape
					strSlideXml += '<p:sp>';

					// B: The addition of the "txBox" attribute is the sole determiner of if an object is a Shape or Textbox
					strSlideXml += '<p:nvSpPr><p:cNvPr id="'+ (idx+2) +'" name="Object '+ (idx+1) +'"/>';
					strSlideXml += '<p:cNvSpPr' + ((slideObj.options && slideObj.options.isTextBox) ? ' txBox="1"/><p:nvPr/>' : '/><p:nvPr/>');
					strSlideXml += '</p:nvSpPr>';
					strSlideXml += '<p:spPr><a:xfrm' + locationAttr + '>';
					strSlideXml += '<a:off x="'  + x  + '" y="'  + y  + '"/>';
					strSlideXml += '<a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>';
					strSlideXml += '<a:prstGeom prst="' + shapeType.name + '"><a:avLst/></a:prstGeom>';

					if ( slideObj.options ) {
						( slideObj.options.fill ) ? strSlideXml += genXmlColorSelection(slideObj.options.fill) : strSlideXml += '<a:noFill/>';

						if ( slideObj.options.line ) {
							var lineAttr = '';
							if ( slideObj.options.line_size ) lineAttr += ' w="' + (slideObj.options.line_size * ONEPT) + '"';
							strSlideXml += '<a:ln' + lineAttr + '>';
							strSlideXml += genXmlColorSelection( slideObj.options.line );
							if ( slideObj.options.line_head ) strSlideXml += '<a:headEnd type="' + slideObj.options.line_head + '"/>';
							if ( slideObj.options.line_tail ) strSlideXml += '<a:tailEnd type="' + slideObj.options.line_tail + '"/>';
							strSlideXml += '</a:ln>';
						}
					}
					else {
						strSlideXml += '<a:noFill/>';
					}

					// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
					if ( slideObj.options.shadow ) {
						slideObj.options.shadow.type    = ( slideObj.options.shadow.type    || 'outer' );
						slideObj.options.shadow.blur    = ( slideObj.options.shadow.blur    || 8 ) * ONEPT;
						slideObj.options.shadow.offset  = ( slideObj.options.shadow.offset  || 4 ) * ONEPT;
						slideObj.options.shadow.angle   = ( slideObj.options.shadow.angle   || 270 ) * 60000;
						slideObj.options.shadow.color   = ( slideObj.options.shadow.color   || '000000' );
						slideObj.options.shadow.opacity = ( slideObj.options.shadow.opacity || 0.75 ) * 100000;

						strSlideXml += '<a:effectLst>';
						strSlideXml += '<a:'+ slideObj.options.shadow.type +'Shdw sx="100000" sy="100000" kx="0" ky="0" ';
						strSlideXml += ' algn="bl" rotWithShape="0" blurRad="'+ slideObj.options.shadow.blur +'" ';
						strSlideXml += ' dist="'+ slideObj.options.shadow.offset +'" dir="'+ slideObj.options.shadow.angle +'">';
						strSlideXml += '<a:srgbClr val="'+ slideObj.options.shadow.color +'">';
						strSlideXml += '<a:alpha val="'+ slideObj.options.shadow.opacity +'"/></a:srgbClr>'
						strSlideXml += '</a:outerShdw>';
						strSlideXml += '</a:effectLst>';
					}

					// FIXME: FUTURE: Text wrapping (copied from MS-PPTX export)
					/*
					// Commented out b/c i'm not even sure this works - current code produces text that wraps in shapes and textboxes, so...
					if ( slideObj.options.textWrap ) {
						strSlideXml += '<a:extLst>'
									+ '<a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">'
									+ '<ma14:wrappingTextBoxFlag xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main" val="1" />'
									+ '</a:ext>'
									+ '</a:extLst>';
					}
					*/

					// B: Close Shape
					strSlideXml += '</p:spPr>';

					// OPTION: align
					if ( slideObj.options &&
						( slideObj.options.align || (Array.isArray(slideObj.text) && slideObj.text[0].options && slideObj.text[0].options.align) )
						) {
						var align = (Array.isArray(slideObj.text) && slideObj.text[0].options && slideObj.text[0].options.align ? slideObj.text[0].options.align : slideObj.options.align);

						switch ( align ) {
							case 'r':
							case 'right':
								moreStylesAttr += ' algn="r"';
								break;
							case 'c':
							case 'ctr':
							case 'center':
								moreStylesAttr += ' algn="ctr"';
								break;
							case 'justify':
								moreStylesAttr += ' algn="just"';
								break;
						}
					}

					// OPTION: indent
					if ( slideObj.options && slideObj.options.indentLevel > 0 ) {
						moreStylesAttr += ' lvl="' + slideObj.options.indentLevel + '"';
					}

					// OPTION: bullet
					if ( slideObj.options.bullet ) {
						moreStylesAttr = ' marL="228600" indent="-228600"';
						moreStyles = '<a:buSzPct val="100000"/><a:buChar char="&#x2022;"/>';
					}

					// Add properties
					if ( moreStyles ) outStyles = '<a:pPr' + moreStylesAttr + '>' + moreStyles + '</a:pPr>';
					else if ( moreStylesAttr ) outStyles = '<a:pPr' + moreStylesAttr + '/>';

					// Add style
					if ( styleData ) strSlideXml += '<p:style>' + styleData + '</p:style>';

					// NOTE: MULTI-LINE: SPEC: Each "line" needs a `p`-aragraph section
					/* <a:p>
						<a:pPr marL="228600" indent="-228600"><a:buSzPct val="100000"/><a:buChar char="&#x2022;"/></a:pPr>
						<a:r><a:t>bullet one</a:t></a:r>
					</a:p> */

					if ( slideObj.options.bullet
						&& typeof slideObj.text == 'string' && slideObj.text.split('\n').length > 0
						&& slideObj.text.split('\n')[1] && slideObj.text.split('\n')[1].length > 0 )
					{
						strSlideXml += '<p:txBody>' + genXmlBodyProperties( slideObj.options ) + '<a:lstStyle/>';
						// IMPORTANT: Split text here (even though `genXmlTextCommand` handles \n) b/c we need `outStyles` for bullets to build correctly
						slideObj.text.toString().split('\n').forEach(function(line,i){
							if ( i > 0 ) strSlideXml += '</a:p>';
							strSlideXml += '<a:p>' + outStyles;
							strSlideXml += genXmlTextCommand( slideObj.options, line, inSlide.slide, inSlide.slide.getPageNumber() );
						});
					}
					else if ( typeof slideObj.text == 'string' || typeof slideObj.text == 'number' ) {
						strSlideXml += '<p:txBody>' + genXmlBodyProperties( slideObj.options ) + '<a:lstStyle/><a:p>' + outStyles;
						strSlideXml += genXmlTextCommand( slideObj.options, slideObj.text.toString(), inSlide.slide, inSlide.slide.getPageNumber() );
					}
					else if ( slideObj.text && Array.isArray(slideObj.text) ) {
						// DESC: This is an array of text objects: [{},{}]
						// EX: slide.addText([ {text:'hello', options:{color:'0088CC'}}, {text:'world', options:{color:'CC8800'}} ]);
						strSlideXml += '<p:txBody>' + genXmlBodyProperties( slideObj.options ) + '<a:lstStyle/><a:p>' + outStyles;
						slideObj.text.forEach(function(textObj,idx){
							textObj.options = textObj.options || {};
							textObj.options.lineIdx = idx;
							strSlideXml += genXmlTextCommand( (textObj.options || slideObj.options), textObj.text.toString(), inSlide.slide, inSlide.slide.getPageNumber() );
						});
					}
					/*
					else if ( slideObj.text && typeof slideObj.text == 'object' && slideObj.text.field ) {
						strSlideXml += '<p:txBody>' + genXmlBodyProperties( slideObj.options ) + '<a:lstStyle/><a:p>' + outStyles;
						strSlideXml += genXmlTextCommand( slideObj.options, slideObj.text, inSlide.slide, inSlide.slide.getPageNumber() );
					}
					*/

					// LAST: End of every paragraph
					if ( slideObj.text ) {
						strSlideXml += '<a:endParaRPr lang="en-US" '+ ( slideObj.options && slideObj.options.font_size ? ' sz="'+ slideObj.options.font_size +'00"' : '') +' dirty="0"/>';
						strSlideXml += '</a:p>';
						strSlideXml += '</p:txBody>';
					}

					strSlideXml += (slideObj.type == 'cxn') ? '</p:cxnSp>' : '</p:sp>';
					break;

				case 'image':
			        strSlideXml += '<p:pic>\r\n';
					strSlideXml += '  <p:nvPicPr><p:cNvPr id="'+ (idx + 2) +'" name="Object '+ (idx + 1) +'" descr="'+ slideObj.image +'"/>\r\n';
			        strSlideXml += '  <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>';
					strSlideXml += '<p:blipFill><a:blip r:embed="rId' + slideObj.imageRid + '" cstate="print"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>';
					strSlideXml += '<p:spPr>'
					strSlideXml += ' <a:xfrm' + locationAttr + '>'
					strSlideXml += '  <a:off  x="' + x  + '"  y="' + y  + '"/>'
					strSlideXml += '  <a:ext cx="' + cx + '" cy="' + cy + '"/>'
					strSlideXml += ' </a:xfrm>'
					strSlideXml += ' <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
					strSlideXml += '</p:spPr>\r\n';
					strSlideXml += '</p:pic>\r\n';
					break;

				case 'media':
					if ( slideObj.mtype == 'online' ) {
						strSlideXml += '<p:pic>';
						strSlideXml += ' <p:nvPicPr>';
						// IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preiew image rId, PowerPoint throws error!
						strSlideXml += ' <p:cNvPr id="'+ (slideObj.mediaRid+2) +'" name="Picture'+ (idx + 1) +'"/>';
						strSlideXml += ' <p:cNvPicPr/>';
						strSlideXml += ' <p:nvPr>';
						strSlideXml += '  <a:videoFile r:link="rId'+ slideObj.mediaRid +'"/>';
						strSlideXml += ' </p:nvPr>';
						strSlideXml += ' </p:nvPicPr>';
						strSlideXml += ' <p:blipFill><a:blip r:embed="rId'+ (slideObj.mediaRid+2) +'"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'; // NOTE: Preview image is required!
						strSlideXml += ' <p:spPr>';
						strSlideXml += '  <a:xfrm' + locationAttr + '>';
						strSlideXml += '   <a:off x="'  + x  + '" y="'  + y  + '"/>';
						strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>';
						strSlideXml += '  </a:xfrm>';
						strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
						strSlideXml += ' </p:spPr>';
						strSlideXml += '</p:pic>';
					}
					else {
						strSlideXml += '<p:pic>\r\n';
						strSlideXml += ' <p:nvPicPr>';
						// IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preiew image rId, PowerPoint throws error!
						strSlideXml += ' <p:cNvPr id="'+ (slideObj.mediaRid+2) +'" name="'+ slideObj.media.split('/').pop().split('.').shift() +'"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>';
						strSlideXml += ' <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>';
						strSlideXml += ' <p:nvPr>';
						strSlideXml += '  <a:videoFile r:link="rId'+ slideObj.mediaRid +'"/>';
						strSlideXml += '  <p:extLst>';
						strSlideXml += '   <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">';
						strSlideXml += '    <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId' + (slideObj.mediaRid+1) + '"/>';
						strSlideXml += '   </p:ext>';
						strSlideXml += '  </p:extLst>';
						strSlideXml += ' </p:nvPr>';
						strSlideXml += ' </p:nvPicPr>';
						strSlideXml += ' <p:blipFill><a:blip r:embed="rId'+ (slideObj.mediaRid+2) +'"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'; // NOTE: Preview image is required!
						strSlideXml += ' <p:spPr>';
						strSlideXml += '  <a:xfrm' + locationAttr + '>';
						strSlideXml += '   <a:off x="'  + x  + '" y="'  + y  + '"/>';
						strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>';
						strSlideXml += '  </a:xfrm>';
						strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
						strSlideXml += ' </p:spPr>';
						strSlideXml += '</p:pic>';
					}
					break;
			}
		});

		// STEP 6: Close spTree and finalize slide XML
		strSlideXml += '</p:spTree>';
		strSlideXml += '<p:extLst>';
		strSlideXml += ' <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">';
		strSlideXml += '  <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1544976994"/>';
		strSlideXml += ' </p:ext>';
		strSlideXml += '</p:extLst>';
		strSlideXml += '</p:cSld>';
		strSlideXml += '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>';
		strSlideXml += '</p:sld>';

		// LAST: Return
		return strSlideXml;
	}

	function makeXmlSlideLayoutRel(inSlideNum) {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
			strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n';
			//?strXml += '  <Relationship Id="rId'+ inSlideNum +'" Target="../slideMasters/slideMaster'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
			//strXml += '  <Relationship Id="rId1" Target="../slideMasters/slideMaster'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
			strXml += '  <Relationship Id="rId1" Target="../slideMasters/slideMaster1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>\r\n';
			strXml += '</Relationships>';
		//
		return strXml;
	}

	function makeXmlSlideRel(inSlideNum) {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
			+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n'
			+ ' <Relationship Id="rId1" Target="../slideLayouts/slideLayout'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"/>\r\n';

		// Add any IMAGEs for this Slide
		gObjPptx.slides[inSlideNum-1].rels.forEach(function(rel,idx){
			if      ( rel.type.toLowerCase().indexOf('image')  > -1 ) {
				strXml += ' <Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>\r\n';
			}
			else if ( rel.type.toLowerCase().indexOf('audio')  > -1 ) {
				// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
				if ( strXml.indexOf(' Target="'+ rel.Target +'"') > -1 )
					strXml += ' <Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>\r\n';
				else
					strXml += ' <Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"/>\r\n';
			}
			else if ( rel.type.toLowerCase().indexOf('video')  > -1 ) {
				// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
				if ( strXml.indexOf(' Target="'+ rel.Target +'"') > -1 )
					strXml += ' <Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>'+CRLF;
				else
					strXml += ' <Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>'+CRLF;
			}
			else if ( rel.type.toLowerCase().indexOf('online') > -1 ) {
				// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
				if ( strXml.indexOf(' Target="'+ rel.Target +'"') > -1 )
					strXml += ' <Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.microsoft.com/office/2007/relationships/image"/>'+CRLF;
				else
					strXml += ' <Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>'+CRLF;
			}
		});

		strXml += '</Relationships>';
		//
		return strXml;
	}

	function makeXmlSlideMaster() {
		var intSlideLayoutId = 2147483649;
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
					+ '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">\r\n'
					+ '  <p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr>\r\n'
					+ '<p:cNvPr id="2" name="Title Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>\r\n'
					+ '<p:cNvPr id="3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>\r\n'
					+ '<p:cNvPr id="4" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="6356350"/><a:ext cx="2133600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{F8166F1F-CE9B-4651-A6AA-CD717754106B}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>12/25/2015</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>\r\n'
					+ '<p:cNvPr id="5" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3124200" y="6356350"/><a:ext cx="2895600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>\r\n'
					+ '<p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="6553200" y="6356350"/><a:ext cx="2133600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="'+SLDNUMFLDID+'" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t></a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>\r\n'
					+ '<p:sldLayoutIdLst>\r\n';
		// Create a sldLayout for each SLIDE
		for ( var idx=1; idx<=gObjPptx.slides.length; idx++ ) {
			strXml += ' <p:sldLayoutId id="'+ intSlideLayoutId +'" r:id="rId'+ idx +'"/>\r\n';
			intSlideLayoutId++;
		}
		strXml += '</p:sldLayoutIdLst>\r\n'
					+ '<p:txStyles>\r\n'
					+ ' <p:titleStyle>\r\n'
					+ '  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>\r\n'
					+ ' </p:titleStyle>'
					+ ' <p:bodyStyle>\r\n'
					+ '  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>'
					+ '  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>'
					+ '  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>'
					+ '  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>'
					+ '  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>'
					+ '  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>'
					+ '  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>'
					+ '  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>'
					+ '  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>'
					+ ' </p:bodyStyle>\r\n'
					+ ' <p:otherStyle>\r\n'
					+ '  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>'
					+ '  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>'
					+ '  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>'
					+ '  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>'
					+ '  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>'
					+ '  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>'
					+ '  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>'
					+ '  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>'
					+ '  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>'
					+ '  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>'
					+ ' </p:otherStyle>\r\n'
					+ '</p:txStyles>\r\n'
					+ '</p:sldMaster>';
		//
		return strXml;
	}

	function makeXmlSlideMasterRel() {
		// FIXME: create a slideLayout for each SLDIE
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
					+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n';
		for ( var idx=1; idx<=gObjPptx.slides.length; idx++ ) {
			strXml += '  <Relationship Id="rId'+ idx +'" Target="../slideLayouts/slideLayout'+ idx +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"/>\r\n';
		}
		strXml += '  <Relationship Id="rId'+ (gObjPptx.slides.length+1) +'" Target="../theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"/>\r\n';
		strXml += '</Relationships>';
		//
		return strXml;
	}

	// XML-GEN: Last 5 functions create root /ppt files

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

	function makeXmlPresentation() {
		var intCurPos = 0;
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
					+ '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">\r\n';

		// STEP 1: Build SLIDE master list
		strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>\r\n';
		strXml += '<p:sldIdLst>\r\n';
		for ( var idx=0; idx<gObjPptx.slides.length; idx++ ) {
			strXml += '<p:sldId id="' + (idx + 256) + '" r:id="rId' + (idx + 2) + '"/>\r\n';
		}
		strXml += '</p:sldIdLst>\r\n';

		// STEP 2: Build SLIDE text styles
		strXml += '<p:sldSz cx="'+ gObjPptx.pptLayout.width +'" cy="'+ gObjPptx.pptLayout.height +'" type="'+ gObjPptx.pptLayout.name +'"/>\r\n'
				+ '<p:notesSz cx="'+ gObjPptx.pptLayout.height +'" cy="' + gObjPptx.pptLayout.width + '"/>'
				+ '<p:defaultTextStyle>';
				+ '  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>';
		for ( var idx=1; idx<10; idx++ ) {
			strXml += '  <a:lvl' + idx + 'pPr marL="' + intCurPos + '" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
					+ '    <a:defRPr sz="1800" kern="1200">'
					+ '      <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
					+ '      <a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>'
					+ '    </a:defRPr>'
					+ '  </a:lvl' + idx + 'pPr>';
			intCurPos += 457200;
		}
		strXml += '</p:defaultTextStyle>\r\n';

		strXml += '<p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext></p:extLst>\r\n'
				+ '</p:presentation>';
		//
		return strXml;
	}

	function makeXmlPresProps() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
					+ '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">\r\n'
					+ '  <p:extLst>\r\n'
					+ '    <p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="0"/></p:ext>\r\n'
					+ '    <p:ext uri="{D31A062A-798A-4329-ABDD-BBA856620510}"><p14:defaultImageDpi xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="220"/></p:ext>\r\n'
					+ '    <p:ext uri="{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}"><p15:chartTrackingRefBased xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" val="1"/></p:ext>\r\n'
					+ '  </p:extLst>\r\n'
					+ '</p:presentationPr>';
		return strXml;
	}

	function makeXmlTableStyles() {
		// SEE: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
					+ '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>';
		return strXml;
	}

	function makeXmlViewProps() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
					+ '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
					+ '<p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>'
					+ '<p:slideViewPr>'
					+ '  <p:cSldViewPr>'
					+ '    <p:cViewPr varScale="1"><p:scale><a:sx n="64" d="100"/><a:sy n="64" d="100"/></p:scale><p:origin x="-1392" y="-96"/></p:cViewPr>'
					+ '    <p:guideLst><p:guide orient="horz" pos="2160"/><p:guide pos="2880"/></p:guideLst>'
					+ '  </p:cSldViewPr>'
					+ '</p:slideViewPr>'
					+ '<p:notesTextViewPr>'
					+ '  <p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr>'
					+ '</p:notesTextViewPr>'
					+ '<p:gridSpacing cx="78028800" cy="78028800"/>'
					+ '</p:viewPr>';
		return strXml;
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

	// Expose a couple private helper functions from above
	this.inch2Emu = inch2Emu;
	this.rgbToHex = rgbToHex;

	/**
	 * Gets the version of this library
	 */
	this.getVersion = function getVersion() {
		return APP_VER;
	};

	/**
	 * Sets the Presentation's Title
	 */
	this.setTitle = function setTitle(inStrTitle) {
		gObjPptx.title = inStrTitle || 'PptxGenJs Presentation';
	};

	/**
	 * Sets the Presentation's Slide Layout {object}: [screen4x3, screen16x9, widescreen]
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 * @param {string} inLayout - a const name from LAYOUTS variable
	 * @param {object} inLayout - an object with user-defined w/h
	 */
	this.setLayout = function setLayout(inLayout) {
		// Allow custom slide size (inches) [ISSUE #29]
		if ( typeof inLayout === 'object' && inLayout.width && inLayout.height ) {
			LAYOUTS['LAYOUT_USER'].width  = Number(inLayout.width ) * EMU;
			LAYOUTS['LAYOUT_USER'].height = Number(inLayout.height) * EMU;

			gObjPptx.pptLayout = LAYOUTS['LAYOUT_USER'];
		}
		else if ( $.inArray(inLayout, Object.keys(LAYOUTS)) > -1 ) {
			gObjPptx.pptLayout = LAYOUTS[inLayout];
		}
		else {
			try { console.warn('UNKNOWN LAYOUT! Valid values = ' + Object.keys(LAYOUTS)); } catch(ex){}
		}
	}

	/**
	 * Gets the Presentation's Slide Layout {object}: [screen4x3, screen16x9, widescreen]
	 */
	this.getLayout = function getLayout() {
		return gObjPptx.pptLayout;
	};

	/**
	 * Export the Presentation to an .pptx file
	 * @param {string} [inStrExportName] - Filename to use for the export
	 */
	this.save = function save(inStrExportName, callback) {
		var intRels = 0, arrRelsDone = [];

		// STEP 1: Set export title (if any)
		if ( inStrExportName ) gObjPptx.fileName = inStrExportName;

		// STEP 2: Read/Encode Images
		// B: Total all physical rels across the Presentation
		// PERF: Only send unique paths for encoding (encoding func will find and fill *ALL* matching paths across the Presentation)
		gObjPptx.slides.forEach(function(slide,idx){
			slide.rels.forEach(function(rel,idy){
				if ( rel.type != 'online' && !rel.data && $.inArray(rel.path, arrRelsDone) == -1 ) {
					// Node encoding is syncronous, so we can load all images here, then call export with a callback (if any)
					if ( NODEJS ) {
						try {
							var bitmap = fs.readFileSync(rel.path);
							rel.data = new Buffer(bitmap).toString('base64');
						}
						catch(ex) {
							console.error('ERROR: Unable to read image: '+rel.path);
							rel.data = IMG_BROKEN;
						}
					}
					else {
						intRels++;
						convertImgToDataURLviaCanvas(rel, callbackImgToDataURLDone);
						arrRelsDone.push(rel.path);
					}
				}
			});
		});

		// STEP 3: Export now if there's no images to encode (otherwise, last async imgConvert call above will call exportFile)
		if ( intRels == 0 ) doExportPresentation(callback);
	};

	/**
	 * Add a new Slide to the Presentation
	 * @returns {Object[]} slideObj - The new Slide object
	 */
	this.addNewSlide = function addNewSlide(inMaster, inMasterOpts) {
		var slideObj = {};
		var slideNum = gObjPptx.slides.length;
		var slideObjNum = 0;
		var pageNum  = (slideNum + 1);

		var inMasterOpts = inMasterOpts || {};
		if ( inMaster && inMasterOpts && inMasterOpts.bkgd ) inMaster.bkgd = inMasterOpts.bkgd; // ISSUE#7: Allow bkgd image/color override on a per-slide basic

		// A: Add this SLIDE to PRESENTATION, Add default values as well
		gObjPptx.slides[slideNum] = {};
		gObjPptx.slides[slideNum].slide = slideObj;
		gObjPptx.slides[slideNum].name = 'Slide ' + pageNum;
		gObjPptx.slides[slideNum].numb = pageNum;
		gObjPptx.slides[slideNum].data = [];
		gObjPptx.slides[slideNum].rels = [];
		gObjPptx.slides[slideNum].slideNumberObj = null;
		gObjPptx.slides[slideNum].hasSlideNumber = false; // DEPRECATED

		// ==========================================================================
		// SLIDE METHODS:
		// ==========================================================================

		slideObj.getPageNumber = function() {
			return pageNum;
		};

		// WARN: DEPRECATED: (leaves in 1.5 or 2.0 at latest)
		slideObj.hasSlideNumber = function( inBool ) {
			if ( inBool ) gObjPptx.slides[slideNum].hasSlideNumber = inBool;
			else return gObjPptx.slides[slideNum].hasSlideNumber;
		};
		slideObj.slideNumber = function( inObj ) {
			if ( inObj && typeof inObj === 'object' ) gObjPptx.slides[slideNum].slideNumberObj = inObj;
			else return gObjPptx.slides[slideNum].slideNumberObj;
		};

		// WARN: DEPRECATED: Will soon take a single {object} as argument (per current docs 20161120)
		// FUTURE: slideObj.addImage = function(opt){
		slideObj.addImage = function( strImagePath, intPosX, intPosY, intSizeX, intSizeY, strImageData ) {
			var intRels = 1;

			// FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
			// FIXME: FUTURE: DEPRECATED: Only allow object param in 1.5 or 2.0
			if ( typeof strImagePath === 'object' ) {
				intPosX = (strImagePath.x || 0);
				intPosY = (strImagePath.y || 0);
				intSizeX = (strImagePath.cx || strImagePath.w || 0);
				intSizeY = (strImagePath.cy || strImagePath.h || 0);
				strImageData = (strImagePath.data || '');
				strImagePath = (strImagePath.path || ''); // This line must be last as were about to ovewrite ourself!
			}

			// REALITY-CHECK:
			if ( !strImagePath && !strImageData ) {
				console.error("ERROR: `addImage()` requires either 'data' or 'path' parameter!");
				return null;
			}
			else if ( strImageData && strImageData.toLowerCase().indexOf('base64,') == -1 ) {
				console.error("ERROR: Image `data` value lacks a base64 header! Ex: 'image/png;base64,NMP[...]')");
				return null;
			}

			// STEP 2: Set vars for this Slide
			var slideObjNum = gObjPptx.slides[slideNum].data.length;
			var slideObjRels = gObjPptx.slides[slideNum].rels;
			// Every image encoded via canvas>base64 is png (as of early 2017 no browser will produce other mime types)
			var strImgExtn = 'png';
			// However, pre-encoded images can be whatever mime-type they want (and good for them!)
			if ( strImageData && /image\/(\w+)\;/.exec(strImageData) && /image\/(\w+)\;/.exec(strImageData).length > 0 ) {
				strImgExtn = /image\/(\w+)\;/.exec(strImageData)[1];
			}
			// Node.js can read/base64-encode any image, so take at face value
			if ( NODEJS && strImagePath.indexOf('.') > -1 ) strImgExtn = strImagePath.split('.').pop();

			gObjPptx.slides[slideNum].data[slideObjNum]       = {};
			gObjPptx.slides[slideNum].data[slideObjNum].type  = 'image';
			gObjPptx.slides[slideNum].data[slideObjNum].image = (strImagePath || 'preencoded.png');

			// STEP 3: Set image properties & options
			// FIXME: Measure actual image when no intSizeX/intSizeY params passed
			// ....: This is an async process: we need to make getSizeFromImage use callback, then set H/W...
			// if ( !intSizeX || !intSizeY ) { var imgObj = getSizeFromImage(strImagePath);
			var imgObj = { width:1, height:1 };
			gObjPptx.slides[slideNum].data[slideObjNum].options    = {};
			gObjPptx.slides[slideNum].data[slideObjNum].options.x  = (intPosX  || 0);
			gObjPptx.slides[slideNum].data[slideObjNum].options.y  = (intPosY  || 0);
			gObjPptx.slides[slideNum].data[slideObjNum].options.cx = (intSizeX || imgObj.width );
			gObjPptx.slides[slideNum].data[slideObjNum].options.cy = (intSizeY || imgObj.height);

			// STEP 4: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
			// NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
			$.each(gObjPptx.slides, function(i,slide){ intRels += slide.rels.length; });
			slideObjRels.push({
				path: (strImagePath || 'preencoded'+strImgExtn),
				type: 'image/'+strImgExtn,
				extn: strImgExtn,
				data: (strImageData || ''),
				rId:  (intRels+1),
				Target: '../media/image' + intRels + '.' + strImgExtn
			});
			gObjPptx.slides[slideNum].data[slideObjNum].imageRid = slideObjRels[slideObjRels.length-1].rId;

			// LAST: Return this Slide
			return this;
		};

		slideObj.addMedia = function( opt ) {
			var intRels  = 1;
			var intPosX  = (opt.x || 0);
			var intPosY  = (opt.y || 0);
			var intSizeX = (opt.w || 2);
			var intSizeY = (opt.h || 2);
			var strData  = (opt.data || '');
			var strLink  = (opt.link || '');
			var strPath  = (opt.path || '');
			var strType  = (opt.type || "audio");
			var strExtn  = "mp3";

			// STEP 1: REALITY-CHECK
			if ( !strPath && !strData && strType != 'online'  ) {
				console.error("ERROR: `addMedia()` requires either 'data' or 'path' values!");
				return null;
			}
			else if ( strData && strData.toLowerCase().indexOf('base64,') == -1 ) {
				console.error("ERROR: Media `data` value lacks a base64 header! Ex: 'video/mpeg;base64,NMP[...]')");
				return null;
			}
			// Online Video: requires `link`
			if ( strType == 'online' && !strLink ) {
				console.error('ERROR: online videos require `link` value')
				return null;
			}
			// Client-Browser: Cant base64 anything except png basically!
			if ( typeof window !== 'undefined' && window.location.href.indexOf('file:') == 0 && !strData && strType != 'online' ) {
				console.error('ERROR: Client browsers cannot encode media - use pre-encoded base64 `data` or use Node.js')
				return null;
			}

			// STEP 2: Set vars for this Slide
			var slideObjNum = gObjPptx.slides[slideNum].data.length;
			var slideObjRels = gObjPptx.slides[slideNum].rels;

			strType = ( strData ? strData.split(';')[0].split('/')[0] : strType );
			strExtn = ( strData ? strData.split(';')[0].split('/')[1] : strPath.split('.').pop() );

			gObjPptx.slides[slideNum].data[slideObjNum]       = {};
			gObjPptx.slides[slideNum].data[slideObjNum].type  = 'media';
			gObjPptx.slides[slideNum].data[slideObjNum].mtype = strType;
			gObjPptx.slides[slideNum].data[slideObjNum].media = (strPath || 'preencoded.mov');

			// STEP 3: Set image properties & options
			gObjPptx.slides[slideNum].data[slideObjNum].options    = {};
			gObjPptx.slides[slideNum].data[slideObjNum].options.x  = intPosX;
			gObjPptx.slides[slideNum].data[slideObjNum].options.y  = intPosY;
			gObjPptx.slides[slideNum].data[slideObjNum].options.cx = intSizeX;
			gObjPptx.slides[slideNum].data[slideObjNum].options.cy = intSizeY;

			// STEP 4: Add this image to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
			// NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
			$.each(gObjPptx.slides, function(i,slide){ intRels += slide.rels.length; });

			if ( strType == 'online' ) {
				slideObjRels.push({
					path: (strPath || 'preencoded'+strExtn),
					type: 'online',
					extn: strExtn,
					data: 'dummy',
					rId:  (intRels+1),
					Target: strLink
				});
				gObjPptx.slides[slideNum].data[slideObjNum].mediaRid = slideObjRels[slideObjRels.length-1].rId;
				// Add preview/overlay image
				slideObjRels.push({
					data: IMG_PLAYBTN,
					path: 'preencoded.png',
					type: 'image/png',
					extn: 'png',
					rId:  (intRels+2),
					Target: '../media/image' + intRels + '.png'
				});
			}
			else {
				// Audio/Video files consume *TWO* rId's:
				// <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
			    // <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
				slideObjRels.push({
					path: (strPath || 'preencoded'+strExtn),
					type: strType+'/'+strExtn,
					extn: strExtn,
					data: (strData || ''),
					rId:  (intRels+1),
					Target: '../media/media' + intRels + '.' + strExtn
				});
				gObjPptx.slides[slideNum].data[slideObjNum].mediaRid = slideObjRels[slideObjRels.length-1].rId;
				slideObjRels.push({
					path: (strPath || 'preencoded'+strExtn),
					type: strType+'/'+strExtn,
					extn: strExtn,
					data: (strData || ''),
					rId:  (intRels+2),
					Target: '../media/media' + intRels + '.' + strExtn
				});
				// Add preview/overlay image
				slideObjRels.push({
					data: IMG_PLAYBTN,
					path: 'preencoded.png',
					type: 'image/png',
					extn: 'png',
					rId:  (intRels+3),
					Target: '../media/image' + intRels + '.png'
				});
			}

			// LAST: Return this Slide
			return this;
		}

		slideObj.addShape = function( shape, opt ) {
			// STEP 1: Grab Slide object count
			slideObjNum = gObjPptx.slides[slideNum].data.length;

			// STEP 2: Set props
			gObjPptx.slides[slideNum].data[slideObjNum] = {};
			gObjPptx.slides[slideNum].data[slideObjNum].type = 'text';
			gObjPptx.slides[slideNum].data[slideObjNum].options = (typeof opt === 'object') ? opt : {};
			gObjPptx.slides[slideNum].data[slideObjNum].options.shape = shape;

			// LAST: Return
			return this;
		};

		// WARN: DEPRECATED: Will soon combine 2nd and 3rd arguments into single {object} (20161216-v1.1.2) (1.5 or 2.0 at the latest)
		// FUTURE: slideObj.addTable = function(arrTabRows, inOpt){
		slideObj.addTable = function( arrTabRows, inOpt, tabOpt ) {
			var opt = ( inOpt && typeof inOpt === 'object' ? inOpt : {} );
			for (var attr in tabOpt) { opt[attr] = tabOpt[attr]; } // FIXME: DEPRECATED: merge opts for now for non-breaking fix (20161216)

			// STEP 1: REALITY-CHECK
			if ( arrTabRows == null || arrTabRows.length == 0 || !Array.isArray(arrTabRows) ) {
				try { console.warn('[warn] addTable: Array expected! USAGE: slide.addTable( [rows], {options} );'); } catch(ex){}
				return null;
			}

			// STEP 2: Set default options if needed
			opt.x         = getSmartParseNumber( (opt.x || (EMU/2)), 'X' );
			opt.y         = getSmartParseNumber( (opt.y || EMU), 'Y' );
			opt.cx        = getSmartParseNumber( (opt.w || opt.cx || (gObjPptx.pptLayout.width - (EMU*2))), 'X' );
			// NOTE: Dont set default `cy` - leaving it null triggers auto-rowH in `makeXMLSlide()`
			opt.cy        = opt.h || opt.cy;
			if ( opt.cy ) opt.cy = getSmartParseNumber( opt.cy, 'Y' );
			opt.w         = opt.cx;
			opt.h         = opt.cy;
			opt.font_size = opt.font_size || 12;
			opt.margin    = opt.marginPt || opt.margin || 0;
			opt.autoPage  = (opt.autoPage == false ? false : true);

			// STEP 3: Convert units to EMU now (we use different logic in makeSlide - smartCalc is not used)
			if ( opt.x  < 20 ) opt.x  = inch2Emu(opt.x);
			if ( opt.y  < 20 ) opt.y  = inch2Emu(opt.y);
			if ( opt.cx < 20 ) opt.cx = inch2Emu(opt.cx);
			if ( opt.cy && opt.cy < 20 ) opt.cy = inch2Emu(opt.cy);

			// STEP 4: Row setup: Handle case where user passed in a simple array
			var arrRows = $.extend(true,[],arrTabRows);
			if ( !Array.isArray(arrRows[0]) ) arrRows = [ $.extend(true,[],arrTabRows) ];

			// STEP 5: Auto-Paging: (via {options} and used internally)
			// (used internally by `addSlidesForTable()` to not engage recursion - we've already paged the table data, just add this one)
			if ( opt && opt.autoPage == false ) {
				// A: Grab Slide object count
				var slideObjNum = gObjPptx.slides[slideNum].data.length;

				// B: Add data (NOTE: Use `extend` to avoid mutation)
				gObjPptx.slides[slideNum].data[slideObjNum] = {
					type:       'table',
					arrTabRows: arrRows,
					options:    $.extend(true,{},opt)
				};
			}
			else {
				// STEP 5: Loop over rows and create one+ tables as needed (ISSUE#21)
				getSlidesForTableRows(arrRows,opt).forEach(function(arrRows,idx){
					// A: We've got a the current Slide already, so add first slides' worth of rows to that,
					// BUT, create new ones going forward!
					if ( !gObjPptx.slides[slideNum+idx] ) {
						gObjPptx.slides[slideNum+idx] = $.extend(true,{},gObjPptx.slides[slideNum]);
						gObjPptx.slides[slideNum+idx].data = [];
					}

					// B: Grab Slide object count
					var slideObjNum = gObjPptx.slides[slideNum+idx].data.length;

					// C: Add data (NOTE: Use `extend` to avoid mutation)
					gObjPptx.slides[slideNum+idx].data[slideObjNum] = {
						type:       'table',
						arrTabRows: arrRows,
						options:    $.extend(true,{},opt)
					};
				});
			}

			// LAST: Return this Slide object
			return this;
		};

		slideObj.addText = function( text, options ) {
			var opt = ( options && typeof options === 'object' ? options : {} );

			// STEP 1: Grab Slide object count
			slideObjNum = gObjPptx.slides[slideNum].data.length;

			// ROBUST: Convert attr values that will likely be passed by users to valid OOXML values
			if ( opt.valign ) opt.valign = opt.valign.toLowerCase().replace(/^c.*/i,'ctr').replace(/^m.*/i,'ctr').replace(/^t.*/i,'t').replace(/^b.*/i,'b');
			if ( opt.align  ) opt.align  = opt.align.toLowerCase().replace(/^c.*/i,'center').replace(/^m.*/i,'center').replace(/^l.*/i,'left').replace(/^r.*/i,'right');

			// STEP 2: Set props
			gObjPptx.slides[slideNum].data[slideObjNum] = {};
			gObjPptx.slides[slideNum].data[slideObjNum].type = 'text';
			gObjPptx.slides[slideNum].data[slideObjNum].text = text;

			gObjPptx.slides[slideNum].data[slideObjNum].options = (typeof opt === 'object') ? opt : {}; // CAPTURE everything passed
			gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp = {};
			gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.autoFit = (opt.autoFit || false); // If true, shape will collapse to text size (Fit To Shape)
			gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.anchor = (opt.valign || 'ctr'); // VALS: [t,ctr,b]
			if ( (opt.inset && !isNaN(Number(opt.inset))) || opt.inset == 0 ) {
				gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.lIns = inch2Emu(opt.inset);
				gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.rIns = inch2Emu(opt.inset);
				gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.tIns = inch2Emu(opt.inset);
				gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.bIns = inch2Emu(opt.inset);
			}

			// ROBUST: Set rational values for some shadow props if needed
			if ( opt.shadow ) {
				// OPT: `type`
				if ( opt.shadow.type != 'outer' && opt.shadow.type != 'inner' ) {
					console.warn('Warning: shadow.type options are `outer` or `inner`.');
					gObjPptx.slides[slideNum].data[slideObjNum].options.type = 'outer';
				}

				// OPT: `angle`
				if ( opt.shadow.angle ) {
					// A: REALITY-CHECK
					if ( isNaN(Number(opt.shadow.angle)) || opt.shadow.angle < 0 || opt.shadow.angle > 359 ) {
						console.warn('Warning: shadow.angle can only be 0-359');
						opt.shadow.angle = 270;
					}

					// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
					gObjPptx.slides[slideNum].data[slideObjNum].options.angle = Math.round(Number(opt.shadow.angle));
				}

				// OPT: `opacity`
				if ( opt.shadow.opacity ) {
					// A: REALITY-CHECK
					if ( isNaN(Number(opt.shadow.opacity)) || opt.shadow.opacity < 0 || opt.shadow.opacity > 1 ) {
						console.warn('Warning: shadow.opacity can only be 0-1');
						opt.shadow.opacity = 0.75;
					}

					// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
					gObjPptx.slides[slideNum].data[slideObjNum].options.opacity = Number(opt.shadow.opacity)
				}
			}

			// LAST: Return
			return this;
		};

		// ==========================================================================
		// POST-METHODS:
		// ==========================================================================

		// C: Add 'Master Slide' attr to Slide if a valid master was provided
		if ( inMaster && this.masters ) {
			// A: Add images (do this before adding slide bkgd)
			if ( inMaster.images && inMaster.images.length > 0 ) {
				$.each(inMaster.images, function(i,image){
					slideObj.addImage({
						data: (image.data || ''),
						path: (image.path || image.src || ''),
						x: inch2Emu(image.x),
						y: inch2Emu(image.y),
						w: inch2Emu(image.w || image.cx),
						h: inch2Emu(image.h || image.cy)
					});
				});
			}

			// B: Add any Slide Background: Image or Fill
			if ( inMaster.bkgd && (inMaster.bkgd.src || inMaster.bkgd.data) ) {
				// Allow the use of only the data key (no src reqd)
				if (!inMaster.bkgd.src) inMaster.bkgd.src = 'preencoded.png';
				var slideObjRels = gObjPptx.slides[slideNum].rels;
				var strImgExtn = inMaster.bkgd.src.substring( inMaster.bkgd.src.indexOf('.')+1 ).toLowerCase();
				if ( strImgExtn == 'jpg' ) strImgExtn = 'jpeg';
				if ( strImgExtn == 'gif' ) strImgExtn = 'png'; // MS-PPT: canvas.toDataURL for gif comes out image/png, and PPT will show "needs repair" unless we do this
				// FIXME: The next few lines are copies from .addImage above. A bad idea thats already bit me once! So of course it's makred as future :)
				var intRels = 1;
				for ( var idx=0; idx<gObjPptx.slides.length; idx++ ) { intRels += gObjPptx.slides[idx].rels.length; }
				slideObjRels.push({
					path: inMaster.bkgd.src,
					type: 'image/'+strImgExtn,
					extn: strImgExtn,
					data: (inMaster.bkgd.data || ''),
					rId: (intRels+1),
					Target: '../media/image' + intRels + '.' + strImgExtn
				});
				slideObj.bkgdImgRid = slideObjRels[slideObjRels.length-1].rId;
			}
			else if ( inMaster.bkgd ) {
				slideObj.back = inMaster.bkgd;
			}

			// C: Add shapes
			if ( inMaster.shapes && inMaster.shapes.length > 0 ) {
				$.each(inMaster.shapes, function(i,shape){
					// 1: Grab all options (x, y, color, etc.)
					var objOpts = {};
					$.each(Object.keys(shape), function(i,key){ if ( shape[key] != 'type' ) objOpts[key] = shape[key]; });
					// 2: Create object using 'type'
					if ( shape.type == 'text' ) slideObj.addText(shape.text, objOpts);
					else if ( shape.type == 'line' ) slideObj.addShape(gObjPptxShapes.LINE, objOpts);
					else if ( shape.type == 'rectangle' ) slideObj.addShape(gObjPptxShapes.RECTANGLE, objOpts);
				});
			}

			// D: Slide numbers
			if ( typeof inMaster.isNumbered !== 'undefined' ) slideObj.hasSlideNumber(inMaster.isNumbered); // DEPRECATED
			if ( inMaster.slideNumber ) slideObj.slideNumber(inMaster.slideNumber);
		}

		// LAST: Return this Slide to allow command chaining
		return slideObj;
	};

	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * "Auto-Paging is the future!" --Elon Musk
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
		if ( $('#'+tabEleId).length == 0 ) { console.error('Table "'+tabEleId+'" does not exist!'); return; }

		// NOTE: Look for opts.margin first as user can override Slide Master settings if they want
		var arrInchMargins = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
		opts.margin = opts.marginPt || opts.margin || 0.5;
		if ( opts && opts.margin ) {
			if ( Array.isArray(opts.margin) ) arrInchMargins = opts.margin;
			else if ( !isNaN(opts.margin) ) arrInchMargins = [opts.margin, opts.margin, opts.margin, opts.margin];
		}
		else if ( opts && opts.master && opts.master.margin && gObjPptxMasters) {
			if ( Array.isArray(opts.master.margin) ) arrInchMargins = opts.master.margin;
			else if ( !isNaN(opts.master.margin) ) arrInchMargins = [opts.master.margin, opts.master.margin, opts.master.margin, opts.master.margin];
			opts.margin = arrInchMargins;
		}
		var emuSlideTabW = ( opts.w ? inch2Emu(opts.w) : (gObjPptx.pptLayout.width  - inch2Emu(arrInchMargins[1] + arrInchMargins[3])) );
		var emuSlideTabH = ( opts.h ? inch2Emu(opts.h) : (gObjPptx.pptLayout.height - inch2Emu(arrInchMargins[0] + arrInchMargins[2])) );

		// STEP 1: Grab overall table style/col widths
		$.each(['thead','tbody','tfoot'], function(i,val){
			if ( $('#'+tabEleId+' > '+val+' > tr').length > 0 ) {
				$('#'+tabEleId+' > '+val+' > tr:first-child').find('> th, > td').each(function(i,cell){
					// FIXME: This is a hack - guessing at col widths when colspan
					if ( $(this).attr('colspan') ) {
						for (var idx=0; idx<$(this).attr('colspan'); idx++ ) {
							arrTabColW.push( Math.round($(this).outerWidth()/$(this).attr('colspan')) );
						}
					}
					else {
						arrTabColW.push( $(this).outerWidth() );
					}
				});
				return false; // break out of .each loop
			}
		});
		$.each(arrTabColW, function(i,colW){ intTabW += colW; });

		// STEP 2: Calc/Set column widths by using same column width percent from HTML table
		$.each(arrTabColW, function(i,colW){
			( $('#'+tabEleId+' thead tr:first-child th:nth-child('+ (i+1) +')').data('pptx-min-width') )
				? arrColW.push( $('#'+tabEleId+' thead tr:first-child th:nth-child('+ (i+1) +')').data('pptx-min-width') )
				: arrColW.push( Math.round( (emuSlideTabW * (colW / intTabW * 100) ) / 100 / EMU ) );
		});

		// STEP 3: Iterate over each table element and create data arrays (text and opts)
		// NOTE: We create 3 arrays instead of one so we can loop over body then show header/footer rows on first and last page
		$.each(['thead','tbody','tfoot'], function(i,val){
			$('#'+tabEleId+' > '+val+' > tr').each(function(i,row){
				var arrObjTabCells = [];
				$(row).find('> th, > td').each(function(i,cell){
					// A: Covert colors from RGB to Hex
					var arrRGB1 = [];
					var arrRGB2 = [];
					arrRGB1 = $(cell).css('color').replace(/\s+/gi,'').replace('rgba(','').replace('rgb(','').replace(')','').split(',');
					arrRGB2 = $(cell).css('background-color').replace(/\s+/gi,'').replace('rgba(','').replace('rgb(','').replace(')','').split(',');

					// B: Create option object
					var objOpts = {
						font_size: $(cell).css('font-size').replace(/\D/gi,''),
						bold:      (( $(cell).css('font-weight') == "bold" || Number($(cell).css('font-weight')) >= 500 ) ? true : false),
						color:     rgbToHex( Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2]) ),
						fill:      rgbToHex( Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2]) )
					};
					if ( $.inArray($(cell).css('text-align'), ['left','center','right','start','end']) > -1 ) objOpts.align = $(cell).css('text-align').replace('start','left').replace('end','right');
					if ( $.inArray($(cell).css('vertical-align'), ['top','middle','bottom']) > -1 ) objOpts.valign = $(cell).css('vertical-align');

					// C: Add padding [margin] (if any)
					// NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
					if ( $(cell).css('padding-left') ) {
						objOpts.margin = [];
						$.each(['padding-top', 'padding-right', 'padding-bottom', 'padding-left'],function(i,val){
							objOpts.margin.push( Math.round($(cell).css(val).replace(/\D/gi,'')) );
						});
					}

					// D: Add colspan (if any)
					if ( $(cell).attr('colspan') ) objOpts.colspan = $(cell).attr('colspan');

					// E: Add border (if any)
					if ( $(cell).css('border-top-width') || $(cell).css('border-right-width') || $(cell).css('border-bottom-width') || $(cell).css('border-left-width') ) {
						objOpts.border = [];
						$.each(['top','right','bottom','left'], function(i,val){
							var intBorderW = Math.round( Number($(cell).css('border-'+val+'-width').replace('px','')) );
							var arrRGB = [];
							arrRGB = $(cell).css('border-'+val+'-color').replace(/\s+/gi,'').replace('rgba(','').replace('rgb(','').replace(')','').split(',');
							var strBorderC = rgbToHex( Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2]) );
							objOpts.border.push( {pt:intBorderW, color:strBorderC} );
						});
					}

					// F: Massage cell text so we honor linebreak tag as a line break during line parsing
					var $cell2 = $(cell).clone();
					$cell2.html( $(cell).html().replace(/<br[^>]*>/gi,'\n') );

					// LAST: Add cell
					arrObjTabCells.push({
						text: $.trim( $cell2.text() ),
						opts: objOpts
					});
				});
				switch (val) {
					case 'thead': arrObjTabHeadRows.push( arrObjTabCells ); break;
					case 'tbody': arrObjTabBodyRows.push( arrObjTabCells ); break;
					case 'tfoot': arrObjTabFootRows.push( arrObjTabCells ); break;
					default:
				}
			});
		});

		// STEP 4: NOTE: `margin` is "cell margin (pt)" everywhere else tables are used, so explicitly convert to "slide margin" here
		if (opts.margin) {
			opts.slideMargin = opts.margin;
			delete(opts.margin);
		}
		// STEP 5: Break table into Slides as needed
		// Pass head-rows as there is an option to add to each table and the parse func needs this daa to fulfill that option
		opts.arrObjTabHeadRows = arrObjTabHeadRows || '';
		opts.colW = arrColW;

		getSlidesForTableRows( arrObjTabHeadRows.concat(arrObjTabBodyRows).concat(arrObjTabFootRows), opts )
		.forEach(function(arrTabRows,i){
			// A: Create new Slide
			var newSlide = ( opts.master && gObjPptxMasters ? api.addNewSlide(opts.master) : api.addNewSlide() );

			// B: Add table to Slide
			newSlide.addTable(arrTabRows, {x:(opts.x || arrInchMargins[3]), y:(opts.y || arrInchMargins[0]), cx:(emuSlideTabW/EMU), colW:arrColW, autoPage:false});

			// C: Add any additional objects
			if ( opts.addImage ) newSlide.addImage({ path:opts.addImage.url, x:opts.addImage.x, y:opts.addImage.y, w:opts.addImage.w, h:opts.addImage.h });
			if ( opts.addShape ) newSlide.addShape( opts.addShape.shape, (opts.addShape.opts || {}) );
			if ( opts.addTable ) newSlide.addTable( opts.addTable.rows,  (opts.addTable.opts || {}) );
			if ( opts.addText  ) newSlide.addText(  opts.addText.text,   (opts.addText.opts  || {}) );
		});
	}
};

// [Node.js] support
if ( NODEJS ) {
	// A: Load 2 depdendencies
	var fs = require("fs");
	var $ = require("jquery-node");
	var JSZip = require("jszip");
	var sizeOf = require("image-size");

	// B: Export module
	module.exports = new PptxGenJS();
}
