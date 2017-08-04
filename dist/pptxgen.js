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
	var APP_VER = "1.7.0-beta";
	var APP_REL = "20170803";
	//
	var MASTER_OBJECTS = {
		'chart': { name:'chart' },
		'image': { name:'image' },
		'line' : { name:'line'  },
		'rect' : { name:'rect'  },
		'text' : { name:'text'  }
	}
	var LAYOUTS = {
		'LAYOUT_4x3'  : { name: 'screen4x3',   width:  9144000, height: 6858000 },
		'LAYOUT_16x9' : { name: 'screen16x9',  width:  9144000, height: 5143500 },
		'LAYOUT_16x10': { name: 'screen16x10', width:  9144000, height: 5715000 },
		'LAYOUT_WIDE' : { name: 'custom',      width: 12191996, height: 6858000 },
		'LAYOUT_USER' : { name: 'custom',      width: 12191996, height: 6858000 }
	};
	var BASE_SHAPES = {
		'RECTANGLE': { 'displayName': 'Rectangle', 'name': 'rect', 'avLst': {} },
		'LINE'     : { 'displayName': 'Line',      'name': 'line', 'avLst': {} }
	};
	// NOTE: 20170304: Only default is used so far. I'd like to combine the two peices of code that use these before implementing these as options
	// Since we close <p> within the text object bullets, its slightly more difficult than combining into a func and calling to get the paraProp
	// and i'm not sure if anyone will even use these... so, skipping for now.
	var BULLET_TYPES = {
		'DEFAULT' : "&#x2022;",
		'CHECK'   : "&#x2713;",
		'STAR'    : "&#x2605;",
		'TRIANGLE': "&#x25B6;"
	};
	var CHART_TYPES = {
		'AREA': { 'displayName':'Area Chart', 'name':'area' },
		'BAR' : { 'displayName':'Bar Chart' , 'name':'bar'  },
		'LINE': { 'displayName':'Line Chart', 'name':'line' },
		'PIE' : { 'displayName':'Pie Chart' , 'name':'pie'  }
	}
	//var RAINBOW_COLORS = ['8A56E2','CF56E2','E256AE','E25668','E28956','E2CF56','AEE256','68E256','56E289','56E2CF','56AEE2','5668E2'];
	var PIECHART_COLORS = ['5DA5DA','FAA43A','60BD68','F17CB0','B2912F','B276B2','DECF3F','F15854','A7A7A7', '5DA5DA','FAA43A','60BD68','F17CB0','B2912F','B276B2','DECF3F','F15854','A7A7A7'];
	var BARCHART_COLORS = ['C0504D','4F81BD','9BBB59','8064A2','4BACC6','F79646','628FC6','C86360', 'C0504D','4F81BD','9BBB59','8064A2','4BACC6','F79646','628FC6','C86360'];
	//
	var SLDNUMFLDID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}';
	{
		var IMG_BROKEN  = 'data:image/png;base64,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';
		var IMG_PLAYBTN = 'data:image/png;base64,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';
	}
	//
	var LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
	var CRLF = '\r\n'; // AKA: Chr(13) & Chr(10)
	var EMU = 914400;  // One (1) inch (OfficeXML measures in EMU (English Metric Units))
	var ONEPT = 12700; // One (1) point (pt)
	//
	var DEF_CELL_MARGIN_PT = [3, 3, 3, 3]; // TRBL-style
	var DEF_FONT_SIZE = 12;
	var DEF_SLIDE_MARGIN_IN = [0.5, 0.5, 0.5, 0.5]; // TRBL-style

	// A: Create internal pptx object
	var gObjPptx = {};

	// B: Set Presentation Property Defaults
	gObjPptx.author = 'PptxGenJS';
	gObjPptx.company = 'PptxGenJS';
	gObjPptx.revision = '1';
	gObjPptx.subject = 'PptxGenJS Presentation';
	gObjPptx.title = 'PptxGenJS Presentation';
	gObjPptx.fileName = 'Presentation';
	gObjPptx.fileExtn = '.pptx';
	gObjPptx.pptLayout = LAYOUTS['LAYOUT_16x9'];
	gObjPptx.rtlMode = false;
	gObjPptx.slides = [];

	// C: Expose shape library to clients
	this.charts  = CHART_TYPES;
	this.masters = ( typeof gObjPptxMasters !== 'undefined' ? gObjPptxMasters : {} );
	this.shapes  = ( typeof gObjPptxShapes  !== 'undefined' ? gObjPptxShapes  : BASE_SHAPES );

	// D: Fall back to base shapes if shapes file was not linked
	gObjPptxShapes = ( gObjPptxShapes || this.shapes );

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
		var arrChartPromises = [];
		var intSlideNum = 0, intRels = 0;

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
		//
		zip.file("[Content_Types].xml", makeXmlContTypes());
		zip.file("_rels/.rels", makeXmlRootRels());
		zip.file("docProps/app.xml", makeXmlApp());
		zip.file("docProps/core.xml", makeXmlCore());
		zip.file("ppt/_rels/presentation.xml.rels", makeXmlPresentationRels());
		//
		zip.file("ppt/theme/theme1.xml", makeXmlTheme());
		zip.file("ppt/presentation.xml", makeXmlPresentation());
		zip.file("ppt/presProps.xml",    makeXmlPresProps());
		zip.file("ppt/tableStyles.xml",  makeXmlTableStyles());
		zip.file("ppt/viewProps.xml",    makeXmlViewProps());

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

		// Create all Rels (images, media, chart data)
		gObjPptx.slides.forEach(function(slide,idx){
			slide.rels.forEach(function(rel,idy){
				if ( rel.type == 'chart' ) {
					arrChartPromises.push(new Promise(function(resolve,reject){
						var zipExcel = new JSZip();
						//
						zipExcel.folder("_rels");
						zipExcel.folder("docProps");
						zipExcel.folder("xl/_rels");
						zipExcel.folder("xl/theme");
						zipExcel.folder("xl/worksheets");
						//
						zipExcel.file("[Content_Types].xml",
							'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="jpeg" ContentType="image/jpg"/><Default Extension="png" ContentType="image/png"/>'
							+ '<Default Extension="bmp" ContentType="image/bmp"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="tif" ContentType="image/tif"/><Default Extension="pdf" ContentType="application/pdf"/><Default Extension="mov" ContentType="application/movie"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
							+ '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
							+ '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
							+ '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
							+ '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>\n'
						);
						zipExcel.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
							+ '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
							+ '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
							+ '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
							+ '</Relationships>\n');
						zipExcel.file("docProps/app.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>\n');
						zipExcel.file("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>\n');
						zipExcel.file("xl/_rels/workbook.xml.rels",
							'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
							+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
							+ '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
							+ '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
							+ '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
							+ '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
							+ '</Relationships>\n'
						);
						zipExcel.file("xl/styles.xml",
							'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="0" formatCode="General"/></numFmts><fonts count="4"><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="10"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="18"/><color indexed="8"/>'
							+ '<name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><dxfs count="0"/><tableStyles count="0"/><colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ffff0000"/><rgbColor rgb="ff00ff00"/><rgbColor rgb="ff0000ff"/>'
							+ '<rgbColor rgb="ffffff00"/><rgbColor rgb="ffff00ff"/><rgbColor rgb="ff00ffff"/><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ff878787"/><rgbColor rgb="fff9f9f9"/></indexedColors></colors></styleSheet>\n'
						);
						zipExcel.file("xl/theme/theme1.xml",
							'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Office Theme"><a:themeElements><a:clrScheme name="Office Theme"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="A7A7A7"/></a:dk2><a:lt2><a:srgbClr val="535353"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="FF00FF"/></a:folHlink></a:clrScheme><a:fontScheme name="Office Theme"><a:majorFont><a:latin typeface="Helvetica"/><a:ea typeface="Helvetica"/><a:cs typeface="Helvetica"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office Theme"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="129999"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="104999"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="b" rotWithShape="0" blurRad="38100" dist="23000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="b" rotWithShape="0" blurRad="38100" dist="23000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="b" rotWithShape="0" blurRad="38100" dist="20000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults><a:spDef><a:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="25400" cap="flat"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln><a:effectLst><a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="b" rotWithShape="0" blurRad="38100" dist="23000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:sp3d/></a:spPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="45719" tIns="45719" rIns="45719" bIns="45719" numCol="1" spcCol="38100" rtlCol="0" anchor="ctr" upright="0"><a:spAutoFit/></a:bodyPr><a:lstStyle><a:defPPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="0" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/><a:sym typeface="Arial"/></a:defRPr></a:defPPr><a:lvl1pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl9pPr></a:lstStyle><a:style><a:lnRef idx="0"/><a:fillRef idx="0"/><a:effectRef idx="0"/><a:fontRef idx="none"/></a:style></a:spDef><a:lnDef><a:spPr><a:noFill/><a:ln w="25400" cap="flat"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln><a:effectLst><a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="b" rotWithShape="0" blurRad="38100" dist="20000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst><a:sp3d/></a:spPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="91439" tIns="45719" rIns="91439" bIns="45719" numCol="1" spcCol="38100" rtlCol="0" anchor="t" upright="0"><a:noAutofit/></a:bodyPr><a:lstStyle><a:defPPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:defPPr><a:lvl1pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl9pPr></a:lstStyle><a:style><a:lnRef idx="0"/><a:fillRef idx="0"/><a:effectRef idx="0"/><a:fontRef idx="none"/></a:style></a:lnDef><a:txDef><a:spPr><a:noFill/><a:ln w="12700" cap="flat"><a:noFill/><a:miter lim="400000"/></a:ln><a:effectLst/><a:sp3d/></a:spPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="45719" tIns="45719" rIns="45719" bIns="45719" numCol="1" spcCol="38100" rtlCol="0" anchor="t" upright="0"><a:spAutoFit/></a:bodyPr><a:lstStyle><a:defPPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="0" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/><a:sym typeface="Arial"/></a:defRPr></a:defPPr><a:lvl1pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="0" marR="0" indent="0" algn="l" defTabSz="914400" rtl="0" fontAlgn="auto" latinLnBrk="1" hangingPunct="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buNone/><a:tabLst/><a:defRPr b="0" baseline="0" cap="none" i="0" spc="0" strike="noStrike" sz="1800" u="none" kumimoji="0" normalizeH="0"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/><a:uFillTx/></a:defRPr></a:lvl9pPr></a:lstStyle><a:style><a:lnRef idx="0"/><a:fillRef idx="0"/><a:effectRef idx="0"/><a:fontRef idx="none"/></a:style></a:txDef></a:objectDefaults></a:theme>\n'
						);
						zipExcel.file("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><workbookPr date1904="1"/><bookViews><workbookView xWindow="0" yWindow="40" windowWidth="15960" windowHeight="18080"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId4"/></sheets></workbook>\n');

						// `sharedStrings.xml`
						var strSharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
						strSharedStrings += '<sst uniqueCount="'+ (rel.data[0].labels.length + rel.data.length) +'" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si/>';
						{
							// A: Pie name comes before labels
							if ( rel.opts.type == 'pie' ) {
								strSharedStrings += '<si><t>'+ (rel.data[0].name || ' ') +'</t></si>';
							}

							// B: Lables
							rel.data[0].labels.forEach(function(label,idx){ strSharedStrings += '<si><t>'+ label +'</t></si>'; });

							// C: Bar series names come after labels
							if ( rel.opts.type == 'bar' ) {
								rel.data.forEach(function(objData,idx){ strSharedStrings += '<si><t>'+ (objData.name || ' ') +'</t></si>'; });
							}
						}
						strSharedStrings += '</sst>\n';
						zipExcel.file("xl/sharedStrings.xml", strSharedStrings);

						var strSheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
						strSheetXml += '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
						strSheetXml += '<sheetViews><sheetView workbookViewId="0" defaultGridColor="1"/></sheetViews>';
						strSheetXml += '<sheetFormatPr defaultColWidth="8" defaultRowHeight="12.75" customHeight="1" outlineLevelRow="0" outlineLevelCol="0"/>';
						strSheetXml += '<sheetData>';

						if ( rel.opts.type == 'area' || rel.opts.type == 'line' ) {
							/* EX: INPUT: `rel.data`
							[
								{ name:'Red', labels:['Jan..May-17'], values:[11,13,14,15,16] },
								{ name:'Amb', labels:['Jan..May-17'], values:[22, 6, 7, 8, 9] },
								{ name:'Grn', labels:['Jan..May-17'], values:[33,32,42,53,63] }
							];
							*/
							/* EX: OUTPUT: barChart Worksheet:
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
							strSheetXml += '<row r="1">';
							for (var idx=1; idx<=rel.data.length; idx++) {
								// TODO: FIXME: Max cols is 52
								strSheetXml += '<c r="'+ ( idx < 26 ? LETTERS[idx] : 'A'+LETTERS[idx%LETTERS.length] ) +'1" t="s">';
								strSheetXml += '<v>'+ idx +'</v>';
								strSheetXml += '</c>';
							}
							strSheetXml += '</row>';

							// B: Add data row(s)
							for (var idx=2; idx<rel.data[0].labels.length+2; idx++) {
								// Leading col is reserved for the label, so hard-code it, then loop over col values
								strSheetXml += '<row r="'+ idx +'">';
								strSheetXml += '<c r="A'+ idx +'" t="s">';
								strSheetXml += '<v>'+ (rel.data[0].values.length + idx - 1) +'</v>';
								strSheetXml += '</c>';
								rel.data.forEach(function(col,idy){
									// TODO: FIXME: Max cols is 52
									strSheetXml += '<c r="'+ ( (idy+1) < 26 ? LETTERS[(idy+1)] : 'A'+LETTERS[(idy+1)%LETTERS.length] ) +''+ idx +'" t="s">';
									strSheetXml += '<v>'+ rel.data[idy].values[idx] +'</v>';
									strSheetXml += '</c>';
								});
								strSheetXml += '</row>';
							}
						}
						else if ( rel.opts.type == 'bar' ) {
							/* EX: INPUT: `rel.data`
								[
									{
										name: 'Income',
										labels: ['2005', '2006', '2007', '2008', '2009'],
										values: [23.5, 26.2, 30.1, 29.5, 24.6]
									},
									{
										name: 'Expense',
										labels: ['2005', '2006', '2007', '2008', '2009'],
										values: [18.1, 22.8, 23.9, 25.1, 25]
									}
								]
							*/
							/* EX: OUTPUT: barChart Worksheet:
								-|---A---|--B--|--C--|--D--|--E--|
								1|       |April|  May| June| July|
								2|Income |   17|   26|   53|   96|
								3|Expense|   55|   43|   70|   58|
								-|-------|-----|-----|-----|-----|
							*/

							// A: Create header row first
							strSheetXml += '<row r="1">';
							for (var i=1; i<=rel.data[0].values.length; i++) {
								// TODO: FIXME: Max cols is 52
								strSheetXml += '<c r="'+ ( i < 26 ? LETTERS[i] : 'A'+LETTERS[i%LETTERS.length] ) +'1" t="s">';
								strSheetXml += '<v>'+ i +'</v>';
								strSheetXml += '</c>';
							}
							strSheetXml += '</row>';

							// B: Add data row(s)
							rel.data.forEach(function(row,idx){
								// Leading col is reserved for the label, so hard-code it, then loop over col values
								strSheetXml += '<row r="'+ (idx+2) +'">';
								strSheetXml += '<c r="A'+ (idx+2) +'" t="s">';
								strSheetXml += '<v>'+ (rel.data[0].values.length + idx + 1) +'</v>';
								strSheetXml += '</c>';
								row.values.forEach(function(val,idy){
									strSheetXml += '<c r="'+ LETTERS[(idy+1)] +''+ (idx+2) +'">';
									strSheetXml += '<v>'+ val +'</v>';
									strSheetXml += '</c>';
								});
								strSheetXml += '</row>';
							});
						}
						else if ( rel.opts.type == 'pie' ) {
							/* EX: INPUT: `rel.data`
								[
									{
										name: 'Project Status',
										labels: ['Red', 'Amber', 'Green', 'Unknown'],
										values: [10, 20, 38, 2]
									}
								]
							*/
							/* EX: OUTPUT: pieChart Worksheet:
								-|---A---|---B---|
								1|       | Status|
								2|Red    |     10|
								3|Amber  |     20|
								4|Green  |     38|
								5|Unknown|      2|
								-|-------|-------|
							*/

							// A: Create header row first
							strSheetXml += '<row r="1"><c r="B1" t="s"><v>1</v></c></row>';

							// B: Add data row(s)
							for (var idx=0; idx<rel.data[0].labels.length; idx++) {
								strSheetXml += '<row r="'+ (idx+2) +'">';
								strSheetXml += '  <c r="A'+ (idx+2) +'" t="s">';
								strSheetXml += '    <v>'+ (idx+2) +'</v>';
								strSheetXml += '  </c>';
								strSheetXml += '  <c r="B'+ (idx+2) +'">';
								strSheetXml += '    <v>'+ rel.data[0].values[idx] +'</v>';
								strSheetXml += '  </c>';
								strSheetXml += '</row>';
							}
						}

						strSheetXml += '</sheetData>';
						strSheetXml += '</worksheet>\n';
						zipExcel.file("xl/worksheets/sheet1.xml", strSheetXml);

						// C: Add XLSX to PPTX export
						zipExcel.generateAsync({type:'base64'})
						.then(function(content){
							// 1: Create the embedded Excel worksheet with labels and data
							zip.file( "ppt/embeddings/Microsoft_Excel_Sheet"+ (rel.rId-1) +".xlsx", content, {base64:true} );

							// 2: Create the chart.xml and rels files
							zip.file("ppt/charts/_rels/"+ rel.fileName +".rels",
								'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
								+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
								+ '<Relationship Id="rId'+ (rel.rId-1) +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Sheet'+ (rel.rId-1) +'.xlsx"/>'
								+ '</Relationships>'
							);
							zip.file("ppt/charts/"+rel.fileName, makeXmlCharts(rel));

							// 3: Done
							resolve();
						})
						.catch(function(strErr){
							reject(strErr);
						});
					}) );
				}
				else if ( rel.type != 'online' && rel.type != 'hyperlink' ) {
					// A: Loop vars
					var data = rel.data;

					// B: Users will undoubtedly pass various string formats, so modify as needed
					if      ( data.indexOf(',') == -1 && data.indexOf(';') == -1 ) data = 'image/png;base64,' + data;
					else if ( data.indexOf(',') == -1                            ) data = 'image/png;base64,' + data;
					else if ( data.indexOf(';') == -1                            ) data = 'image/png;' + data;

					// C: Add media
					zip.file( rel.Target.replace('..','ppt'), data.split(',').pop(), {base64:true} );
				}
			});
		});

		// STEP 3: Wait for Promises (if any) then push the PPTX file to client-browser
		Promise.all( arrChartPromises )
		.then(function(arrResults){
			var strExportName = ((gObjPptx.fileName.toLowerCase().indexOf('.ppt') > -1) ? gObjPptx.fileName : gObjPptx.fileName+gObjPptx.fileExtn);
			if ( NODEJS ) {
				if ( callback ) {
					if ( strExportName.indexOf('http') == 0 ) {
						zip.generateAsync({type:'nodebuffer'}).then(function(content){ callback(content); });
					}
					else {
						zip.generateAsync({type:'nodebuffer'}).then(function(content){ fs.writeFile(strExportName, content, callback(strExportName)); });
					}
				}
				else
					zip.generateAsync({type:'nodebuffer'}).then(function(content){ fs.writeFile(strExportName, content); });
			}
			else {
				zip.generateAsync({type:'blob'}).then(function(content){ writeFileToBrowser(strExportName, content, callback); });
			}
		})
		.catch(function(strErr){
			console.error(strErr);
		});
	}

	function writeFileToBrowser(strExportName, content, callback) {
		// STEP 1: Create element
		var a = document.createElement("a");
		document.body.appendChild(a);
		a.style = "display: none";

		// STEP 2: Download file to browser
		// DESIGN: Use `createObjectURL()` (or MS-specific func for IE11) to D/L files in client browsers (FYI: synchronously executed)
		if ( window.navigator.msSaveOrOpenBlob ) {
			// REF: https://docs.microsoft.com/en-us/microsoft-edge/dev-guide/html5/file-api/blob
			blobObject = new Blob([content]);
			$(a).click(function(){
				window.navigator.msSaveOrOpenBlob(blobObject, strExportName);
			});
			a.click();

			// Clean-up
			document.body.removeChild(a);

			// LAST: Callback (if any)
			if ( callback ) callback(strExportName);
		}
		else if ( window.URL.createObjectURL ) {
			var blob = new Blob([content], {type: "octet/stream"});
			var url = window.URL.createObjectURL(blob);
			a.href = url;
			a.download = strExportName;
			a.click();

			// Clean-up (NOTE: Add a slight delay before removing to avoid 'blob:null' error in Firefox Issue#81)
			setTimeout(function(){
				window.URL.revokeObjectURL(url);
				document.body.removeChild(a);
			}, 100);

			// LAST: Callback (if any)
			if ( callback ) callback(strExportName);
		}
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

	function getShapeInfo(shapeName) {
		if ( !shapeName ) return gObjPptxShapes.RECTANGLE;

		if ( typeof shapeName == 'object' && shapeName.name && shapeName.displayName && shapeName.avLst ) return shapeName;

		if ( gObjPptxShapes[shapeName] ) return gObjPptxShapes[shapeName];

		var objShape = gObjPptxShapes.filter(function(obj){ return obj.name == shapeName || obj.displayName; })[0];
		if ( typeof objShape !== 'undefined' && objShape != null ) return objShape;

		return gObjPptxShapes.RECTANGLE;
	}

	function getSmartParseNumber(inVal, inDir) {
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

	/**
	 * DESC: Replace special XML characters with HTML-encoded strings
	 */
	function decodeXmlEntities(inStr) {
		// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
		if ( typeof inStr === 'undefined' || inStr == null ) return "";
		return inStr.toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/\'/g,'&apos;');
	}

	function createHyperlinkRels(inText, slideRels) {
		var arrTextObjects = [];

		// Only text objects can have hyperlinks, so return if this is plain text/number
		if ( typeof inText === 'string' || typeof inText === 'number' ) return;
		// IMPORTANT: Check for isArray before typeof=object, or we'll exhaust recursion!
		else if ( Array.isArray(inText) ) arrTextObjects = inText;
		else if ( typeof inText === 'object' ) arrTextObjects = [inText];

		arrTextObjects.forEach(function(text,idx){
			// `text` can be an array of other `text` objects (table cell word-level formatting), so use recursion
			if ( Array.isArray(text) ) createHyperlinkRels(text, slideRels);
			else if ( text && typeof text === 'object' && text.options && text.options.hyperlink && !text.options.hyperlink.rId ) {
				if ( typeof text.options.hyperlink !== 'object' ) console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:'https://github.com'}` ");
				else if ( !text.options.hyperlink.url || typeof text.options.hyperlink.url !== 'string' ) console.log("ERROR: 'hyperlink.url is required and/or should be a string'");
				else {
					var intRels = 1;
					gObjPptx.slides.forEach(function(slide,idx){ intRels += slide.rels.length; });
					var intRelId = intRels+1;

					slideRels.push({
						type: 'hyperlink',
						data: 'dummy',
						rId:  intRelId,
						Target: text.options.hyperlink.url
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
		var CPL = (inWidth*EMU / ( (cell.opts.font_size || DEF_FONT_SIZE)/CHAR )); // Chars-Per-Line
		var arrLines = [];
		var strCurrLine = '';

		// Allow a single space/whitespace as cell text
		if ( cell.text && cell.text.trim() == '' ) return [' '];

		// A: Remove leading/trailing space
		var inStr = (cell.text || '').toString().trim();

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
			strCurrLine = '';
		});

		// C: Remove trailing linebreak
		arrLines[(arrLines.length-1)] = $.trim(arrLines[(arrLines.length-1)]);

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
		var emuSlideTabW = EMU*1, emuSlideTabH = EMU*1;
		var arrObjTabHeadRows = opts.arrObjTabHeadRows || '';
		var numCols = 0;

		if (opts.debug) console.log('------------------------------------');
		if (opts.debug) console.log('opts.w ............. = '+ (opts.w||'').toString());
		if (opts.debug) console.log('opts.colW .......... = '+ (opts.colW||'').toString());
		if (opts.debug) console.log('opts.slideMargin ... = '+ (opts.slideMargin||'').toString());

		// NOTE: Use default size as zero cell margin is causing our tables to be too large and touch bottom of slide!
		if ( !opts.slideMargin && opts.slideMargin != 0 ) opts.slideMargin = DEF_SLIDE_MARGIN_IN[0];

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
			if (!cell) cell = {};
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

		// STEP 2: Calc usable space/table size now that we have usable space calc'd
		emuSlideTabW = ( opts.w ? inch2Emu(opts.w) : (gObjPptx.pptLayout.width - inch2Emu((opts.x || arrInchMargins[1]) + arrInchMargins[3])) );
		if (opts.debug) console.log('emuSlideTabW (in) ........ = '+ (emuSlideTabW/EMU).toFixed(1) );
		if (opts.debug) console.log('gObjPptx.pptLayout.h ..... = '+ (gObjPptx.pptLayout.height/EMU));

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

		// STEP 4: Iterate over each line and perform magic =========================
		// NOTE: inArrRows will be an array of {text:'', opts{}} whether from `addSlidesForTable()` or `.addTable()`
		inArrRows.forEach(function(row,iRow){
			// A: Reset ROW variables
			var arrCellsLines = [], arrCellsLineHeights = [], emuRowH = 0, intMaxLineCnt = 0, intMaxColIdx = 0;

			// B: Calc usable vertical space/table height
			// NOTE: Use margins after the first Slide (dont re-use opt.y - it could've been halfway down the page!) (ISSUE#43,ISSUE#47,ISSUE#48)
			if ( arrObjSlides.length > 0 ) {
				emuSlideTabH = ( gObjPptx.pptLayout.height - inch2Emu( (opts.y/EMU < arrInchMargins[0] ? opts.y/EMU : arrInchMargins[0]) + arrInchMargins[2]) );
				// Use whichever is greater: area between margins or the table H provided (dont shrink usable area - the whole point of over-riding X on paging is to *increarse* usable space)
				if ( emuSlideTabH < opts.h ) emuSlideTabH = opts.h;
			}
			else emuSlideTabH = ( opts.h ? opts.h : (gObjPptx.pptLayout.height - inch2Emu((opts.y/EMU || arrInchMargins[0]) + arrInchMargins[2])) );
			if (opts.debug) console.log('* Slide '+arrObjSlides.length+': emuSlideTabH (in) ........ = '+ (emuSlideTabH/EMU).toFixed(1));

			// C: Parse and store each cell's text into line array (**MAGIC HAPPENS HERE**)
			row.forEach(function(cell,iCell){
				// FIRST: REALITY-CHECK:
				if (!cell) cell = {};

				// DESIGN: Cells are henceforth {objects} with `text` and `opts`
				var lines = [];

				// 1: Cleanse data
				if ( !isNaN(cell) || typeof cell === 'string' ) {
					// Grab table formatting `opts` to use here so text style/format inherits as it should
					cell = { text:cell.toString(), opts:opts };
				}
				else if ( typeof cell === 'object' ) {
					// ARG0: `text`
					if ( typeof cell.text === 'number' ) cell.text = cell.text.toString();
					else if ( typeof cell.text === 'undefined' || cell.text == null ) cell.text = "";

					// ARG1: `options`
					var opt = cell.options || cell.opts || {}; // Legacy support for `opts` (<= v1.2.0)
					cell.opts = opt; // This odd soln is needed until `opts` can be safely discarded (DEPRECATED)
				}
				// Capture some table options for use in other functions
				cell.opts.lineWeight = opts.lineWeight;

				// 2: Create a cell object for each table column
				currRow.push({ text:'', opts:cell.opts });

				// 3: Parse cell contents into lines (**MAGIC HAPPENSS HERE**)
				var lines = parseTextToLines(cell, (opts.colW[iCell]/ONEPT));
				arrCellsLines.push( lines );
				//if (opts.debug) console.log('Cell:'+iCell+' - lines:'+lines.length);

				// 4: Keep track of max line count within all row cells
				if ( lines.length > intMaxLineCnt ) { intMaxLineCnt = lines.length; intMaxColIdx = iCell; }
				var lineHeight = inch2Emu((cell.opts.font_size || opts.font_size || DEF_FONT_SIZE) * LINEH_MODIFIER / 100);
				// NOTE: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
				if ( cell.opts && cell.opts.rowspan ) lineHeight = 0;

				// 5: Add cell margins to lineHeight (if any)
				if ( cell.opts.margin ) {
					if ( cell.opts.margin[0] ) lineHeight += (cell.opts.margin[0]*ONEPT) / intMaxLineCnt;
					if ( cell.opts.margin[2] ) lineHeight += (cell.opts.margin[2]*ONEPT) / intMaxLineCnt;
				}

				// Add to array
				arrCellsLineHeights.push( Math.round(lineHeight) );
			});

			// D: AUTO-PAGING: Add text one-line-a-time to this row's cells until: lines are exhausted OR table H limit is hit
			for (var idx=0; idx<intMaxLineCnt; idx++) {
				// 1: Add the current line to cell
				for (var col=0; col<arrCellsLines.length; col++) {
					// A: Commit this slide to Presenation if table Height limit is hit
					if ( emuTabCurrH + arrCellsLineHeights[intMaxColIdx] > emuSlideTabH ) {
						if (opts.debug) console.log('--------------- New Slide Created ---------------');
						if (opts.debug) console.log(' (calc) '+ (emuTabCurrH/EMU).toFixed(1) +'+'+ (arrCellsLineHeights[intMaxColIdx]/EMU).toFixed(1) +' > '+ emuSlideTabH/EMU.toFixed(1));
						if (opts.debug) console.log('--------------- New Slide Created ---------------');
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
								var lines = parseTextToLines(cell,(opts.colW[iCell]/ONEPT));
								if ( lines.length > intMaxLineCnt ) { intMaxLineCnt = lines.length; intMaxColIdx = iCell; }
							});
							arrRows.push( $.extend(true, [], headRow) );
						}
					}

					// B: Add next line of text to this cell
					if ( arrCellsLines[col][idx] ) currRow[col].text += arrCellsLines[col][idx];
				}

				// 2: Add this new rows H to overall (use cell with the most lines as the determiner for overall row Height)
				emuTabCurrH += arrCellsLineHeights[intMaxColIdx];
			}

			if (opts.debug) console.log('-> '+iRow+ ' row done!');
			if (opts.debug) console.log('-> emuTabCurrH (in) . = '+ (emuTabCurrH/EMU).toFixed(1));

			// E: Flush row buffer - Add the current row to table, then truncate row cell array
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

	/**
	* @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
	*/
	function makeXmlCharts(rel) {
		function getExcelColName(length) {
			var strName = '';

			if ( length <= 26 ) {
				strName = LETTERS[length];
			}
			else {
				strName += LETTERS[ Math.floor(length/LETTERS.length)-1 ];
				strName += LETTERS[ (length % LETTERS.length) ];
			}

			return strName;
		}

		// STEP 1: Create chart
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
		// CHARTSPACE: BEGIN vvv
		strXml += '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
		strXml += '<c:chart>';

		// OPTION: Title
		if ( rel.opts.showTitle ) {
			strXml += '<c:title>';
			strXml += ' <c:tx>';
			strXml += '  <c:rich>';
			strXml += '  <a:bodyPr rot="0"/>';
			strXml += '  <a:lstStyle/>';
			strXml += '  <a:p>';
			strXml += '    <a:pPr>';
			strXml += '      <a:defRPr b="0" i="0" strike="noStrike" sz="'+ (rel.opts.titleFontSize || DEF_FONT_SIZE) +'00" u="none">';
			strXml += '        <a:solidFill><a:srgbClr val="'+ (rel.opts.titleColor || '000000') +'"/></a:solidFill>';
			strXml += '        <a:latin typeface="'+ (rel.opts.titleFontFace || 'Arial') +'"/>';
			strXml += '      </a:defRPr>';
			strXml += '    </a:pPr>';
			strXml += '    <a:r>';
			strXml += '      <a:rPr b="0" i="0" strike="noStrike" sz="'+ (rel.opts.titleFontSize || DEF_FONT_SIZE) +'00" u="none">';
			strXml += '        <a:solidFill><a:srgbClr val="'+ (rel.opts.titleColor || '000000') +'"/></a:solidFill>';
			strXml += '        <a:latin typeface="'+ (rel.opts.titleFontFace || 'Arial') +'"/>';
			strXml += '      </a:rPr>';
			strXml += '      <a:t>'+ (decodeXmlEntities(rel.opts.title) || '') +'</a:t>';
			strXml += '    </a:r>';
			strXml += '  </a:p>';
			strXml += '  </c:rich>';
			strXml += ' </c:tx>';
			strXml += ' <c:layout/>';
			strXml += ' <c:overlay val="0"/>';
			strXml += '</c:title>';
			strXml += '<c:autoTitleDeleted val="0"/>';
		}

		strXml += '<c:plotArea>';
		// IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
		strXml += '<c:layout/>';

		// A: CHART TYPES -----------------------------------------------------------
		switch ( rel.opts.type ) {
			case 'area':
			case 'bar':
			case 'line':
				strXml += '<c:'+ rel.opts.type +'Chart>';
				if ( rel.opts.type == 'bar' ) strXml += '  <c:barDir val="'+ rel.opts.barDir +'"/>';
				strXml += '  <c:grouping val="'+ rel.opts.barGrouping + '"/>';
				strXml += '  <c:varyColors val="0"/>';

				// A: "Series" block for every data row
				/* EX:
					data: [
				     {
				       name: 'Region 1',
				       labels: ['April', 'May', 'June', 'July'],
				       values: [17, 26, 53, 96]
				     },
				     {
				       name: 'Region 2',
				       labels: ['April', 'May', 'June', 'July'],
				       values: [55, 43, 70, 58]
				     }
				    ]
				*/
				rel.data.forEach(function(obj,idx){
					strXml += '<c:ser>';
					strXml += '  <c:idx val="'+ idx +'"/>';
					strXml += '  <c:order val="'+ idx +'"/>';
					strXml += '  <c:tx>';
					strXml += '    <c:strRef>';
					strXml += '      <c:f>Sheet1!$A$'+ (idx+2) +'</c:f>';
					strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>'+ obj.name +'</c:v></c:pt></c:strCache>';
					strXml += '    </c:strRef>';
					strXml += '  </c:tx>';

					// Fill and Border
					var strSerColor = rel.opts.chartColors[(idx+1 > rel.opts.chartColors.length ? (Math.floor(Math.random() * rel.opts.chartColors.length)) : idx)];
					strXml += '  <c:spPr>';

					if ( rel.opts.chartColorsOpacity ) {
						strXml += '    <a:solidFill><a:srgbClr val="'+ strSerColor +'"><a:alpha val="50000"/></a:srgbClr></a:solidFill>';
					}
					else {
						strXml += '    <a:solidFill><a:srgbClr val="'+ strSerColor +'"/></a:solidFill>';
					}

					if ( rel.opts.type == 'line' ) {
						strXml += '<a:ln w="'+ (rel.opts.lineSize * ONEPT) +'" cap="flat"><a:solidFill><a:srgbClr val="'+ strSerColor +'"/></a:solidFill>';
						strXml += '<a:prstDash val="' + (rel.opts.line_dash || "solid") + '"/><a:round/></a:ln>';
					}
					else if ( rel.opts.dataBorder ) {
						strXml += '<a:ln w="'+ (rel.opts.dataBorder.pt * ONEPT) +'" cap="flat"><a:solidFill><a:srgbClr val="'+ rel.opts.dataBorder.color +'"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
					}
					strXml += '    <a:effectLst>';
					strXml += '      <a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="tl" rotWithShape="1" blurRad="38100" dist="23000" dir="5400000">';
					strXml += '        <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>';
					strXml += '      </a:outerShdw>';
					strXml += '    </a:effectLst>';
					strXml += '  </c:spPr>';

					// LINE CHART ONLY: `marker`
					if ( rel.opts.type == 'line' ) {
						strXml += '<c:marker>';
						strXml += '  <c:symbol val="'+ rel.opts.lineDataSymbol +'"/>';
						if ( rel.opts.lineDataSymbolSize ) strXml += '  <c:size val="'+ rel.opts.lineDataSymbolSize +'"/>'; // Defaults to "auto" otherwise (but this is usually too small, so there is a default)
						strXml += '  <c:spPr>';
	  					strXml += '    <a:solidFill><a:srgbClr val="'+ rel.opts.chartColors[(idx+1 > rel.opts.chartColors.length ? (Math.floor(Math.random() * rel.opts.chartColors.length)) : idx)] +'"/></a:solidFill>';
						strXml += '    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="'+ strSerColor +'"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
						strXml += '    <a:effectLst/>';
						strXml += '  </c:spPr>';
						strXml += '</c:marker>';

					}

					// 1: "Data Labels"
					strXml += '  <c:dLbls>';
					strXml += '    <c:numFmt formatCode="'+ rel.opts.dataLabelFormatCode +'" sourceLinked="0"/>';
					strXml += '    <c:txPr>';
					strXml += '      <a:bodyPr/>';
					strXml += '      <a:lstStyle/>';
					strXml += '      <a:p><a:pPr>';
					strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="'+ (rel.opts.dataLabelFontSize || DEF_FONT_SIZE) +'00" u="none">';
					strXml += '          <a:solidFill><a:srgbClr val="'+ (rel.opts.dataLabelColor || '000000') +'"/></a:solidFill>';
					strXml += '          <a:latin typeface="'+ (rel.opts.dataLabelFontFace || 'Arial') +'"/>';
					strXml += '        </a:defRPr>';
					strXml += '      </a:pPr></a:p>';
					strXml += '    </c:txPr>';
					if ( rel.opts.type != 'area' ) strXml += '    <c:dLblPos val="'+ (rel.opts.dataLabelPosition || 'outEnd') +'"/>';
					strXml += '    <c:showLegendKey val="0"/>';
					strXml += '    <c:showVal val="'+ (rel.opts.showValue ? '1' : '0') +'"/>';
					strXml += '    <c:showCatName val="0"/>';
					strXml += '    <c:showSerName val="0"/>';
					strXml += '    <c:showPercent val="0"/>';
					strXml += '    <c:showBubbleSize val="0"/>';
					strXml += '    <c:showLeaderLines val="0"/>';
					strXml += '  </c:dLbls>';

					// 2: "Categories"
					strXml += '<c:cat>';
					strXml += '  <c:strRef>';
					strXml += '    <c:f>Sheet1!'+ '$B$1:$'+ getExcelColName(obj.labels.length) +'$1' +'</c:f>';
					strXml += '    <c:strCache>';
					strXml += '	     <c:ptCount val="'+ obj.labels.length +'"/>';
					obj.labels.forEach(function(label,idx){ strXml += '<c:pt idx="'+ idx +'"><c:v>'+ label +'</c:v></c:pt>'; });
					strXml += '    </c:strCache>';
					strXml += '  </c:strRef>';
					strXml += '</c:cat>';

					// 3: "Values"
					strXml += '  <c:val>';
					strXml += '    <c:numRef>';
					strXml += '      <c:f>Sheet1!'+ '$B$'+ (idx+2) +':$'+ getExcelColName(obj.labels.length) +'$'+ (idx+2) +'</c:f>';
					strXml += '      <c:numCache>';
					strXml += '	       <c:ptCount val="'+ obj.labels.length +'"/>';
					obj.values.forEach(function(value,idx){ strXml += '<c:pt idx="'+ idx +'"><c:v>'+ value +'</c:v></c:pt>'; });
					strXml += '      </c:numCache>';
					strXml += '    </c:numRef>';
					strXml += '  </c:val>';

					// LINE CHART ONLY: `smooth`
					if ( rel.opts.type == 'line' ) strXml += '<c:smooth val="'+ (rel.opts.lineSmooth ? "1" : "0" ) +'"/>';

					// 4: Close "SERIES"
					strXml += '</c:ser>';
				});
				//
				if ( rel.opts.type == 'bar' ) {
					strXml += '  <c:gapWidth val="'+ rel.opts.barGapWidthPct +'"/>';
					strXml += '  <c:overlap val="'+ (rel.opts.barGrouping.indexOf('tacked') > -1 ? 100 : 0) +'"/>';
				}
				else if ( rel.opts.type == 'line' ) {
					strXml += '  <c:marker val="1"/>';
				}
				strXml += '  <c:axId val="2094734552"/>';
				strXml += '  <c:axId val="2094734553"/>';
				strXml += '</c:'+ rel.opts.type +'Chart>';

				// B: "Category Axis"
				{
					strXml += '<c:catAx>';
					strXml += '  <c:axId val="2094734552"/>';
					strXml += '  <c:scaling><c:orientation val="'+ (rel.opts.catAxisOrientation || (rel.opts.barDir == 'col' ? 'minMax' : 'minMax')) +'"/></c:scaling>';
					strXml += '  <c:delete val="0"/>';
					strXml += '  <c:axPos val="'+ (rel.opts.barDir == 'col' ? 'b' : 'l') +'"/>';
					strXml += '  <c:numFmt formatCode="General" sourceLinked="0"/>';
					strXml += '  <c:majorTickMark val="out"/>';
					strXml += '  <c:minorTickMark val="none"/>';
					strXml += '  <c:tickLblPos val="'+ (rel.opts.barDir == 'col' ? 'low' : 'nextTo') +'"/>';
					strXml += '  <c:spPr>';
					strXml += '    <a:ln w="12700" cap="flat"><a:solidFill><a:srgbClr val="888888"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
					strXml += '  </c:spPr>';
					strXml += '  <c:txPr>';
					strXml += '    <a:bodyPr rot="0"/>';
					strXml += '    <a:lstStyle/>';
					strXml += '    <a:p>';
					strXml += '    <a:pPr>';
					strXml += '<a:defRPr b="0" i="0" strike="noStrike" sz="'+ (rel.opts.catAxisLabelFontSize || DEF_FONT_SIZE) +'00" u="none">';
					strXml += '<a:solidFill><a:srgbClr val="'+ (rel.opts.catAxisLabelColor || '000000') +'"/></a:solidFill>';
					strXml += '<a:latin typeface="'+ (rel.opts.catAxisLabelFontFace || 'Arial') +'"/>';
					strXml += '   </a:defRPr>';
					strXml += '  </a:pPr>';
					strXml += '  </a:p>';
					strXml += ' </c:txPr>';
					strXml += ' <c:crossAx val="2094734553"/>';
					strXml += ' <c:crosses val="autoZero"/>';
					strXml += ' <c:auto val="1"/>';
					strXml += ' <c:lblAlgn val="ctr"/>';
					strXml += ' <c:noMultiLvlLbl val="1"/>';
					strXml += '</c:catAx>';
				}

				// C: "Value Axis"
				{
					strXml += '<c:valAx>';
					strXml += '  <c:axId val="2094734553"/>';
					strXml += '  <c:scaling>';
					strXml += '    <c:orientation val="'+ (rel.opts.valAxisOrientation || (rel.opts.barDir == 'col' ? 'minMax' : 'minMax')) +'"/>';
					if (rel.opts.valAxisMaxVal) strXml += '<c:max val="'+ rel.opts.valAxisMaxVal +'"/>';
					strXml += '  </c:scaling>';
					strXml += '  <c:delete val="0"/>';
					strXml += '  <c:axPos val="'+ (rel.opts.barDir == 'col' ? 'l' : 'b') +'"/>';

					// TESTING: 20170527 - omitting this 'strXml' works fine in Keynote/PPT-Online!!
					// TODO: gridLine options! (colors/style(solid/dashed) -OR- NONE (eg:omit this section)
					strXml += ' <c:majorGridlines>\
								<c:spPr>\
								  <a:ln w="12700" cap="flat">\
								    <a:solidFill><a:srgbClr val="888888"/></a:solidFill>\
								    <a:prstDash val="solid"/><a:round/>\
								  </a:ln>\
								</c:spPr>\
								</c:majorGridlines>';
					strXml += ' <c:numFmt formatCode="General" sourceLinked="0"/>';
					strXml += ' <c:majorTickMark val="out"/>';
					strXml += ' <c:minorTickMark val="none"/>';
					strXml += ' <c:tickLblPos val="'+ (rel.opts.barDir == 'col' ? 'nextTo' : 'low') +'"/>';
					strXml += ' <c:spPr>';
					strXml += '  <a:ln w="12700" cap="flat"><a:solidFill><a:srgbClr val="888888"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
					strXml += ' </c:spPr>';
					strXml += ' <c:txPr>';
					strXml += '  <a:bodyPr rot="0"/>';
					strXml += '  <a:lstStyle/>';
					strXml += '  <a:p>';
					strXml += '    <a:pPr>';
					strXml += '      <a:defRPr b="0" i="0" strike="noStrike" sz="'+ (rel.opts.valAxisLabelFontSize || DEF_FONT_SIZE) +'00" u="none">';
					strXml += '        <a:solidFill><a:srgbClr val="'+ (rel.opts.valAxisLabelColor || '000000') +'"/></a:solidFill>';
					strXml += '        <a:latin typeface="'+ (rel.opts.valAxisLabelFontFace || 'Arial') +'"/>';
					strXml += '      </a:defRPr>';
					strXml += '    </a:pPr>';
					strXml += '  </a:p>';
					strXml += ' </c:txPr>';
					strXml += ' <c:crossAx val="2094734552"/>';
					strXml += ' <c:crosses val="autoZero"/>';
					strXml += ' <c:crossBetween val="'+ ( rel.opts.type == 'area' ? 'midCat' : 'between' ) +'"/>';
					//strXml += ' <c:majorUnit val="25"/>'; // NOTE: Not Required.	// TODO: Add OPTION?
					//strXml += ' <c:minorUnit val="12.5"/>'; // NOTE: Not Required.	// TODO: Add OPTION?
					strXml += '</c:valAx>';
				}

				// Done with CHART.BAR/LINE
				break;

			case 'pie':
				// Use the same var name so code blocks from barChart are interchangeable
				var obj = rel.data[0];

				/* EX:
					data: [
					 {
					   name: 'Project Status',
					   labels: ['Red', 'Amber', 'Green', 'Unknown'],
					   values: [10, 20, 38, 2]
					 }
					]
				*/

				// 1: Start pieChart
				strXml += '<c:pieChart>';
				strXml += '  <c:varyColors val="0"/>';
				strXml += '<c:ser>';
				strXml += '  <c:idx val="0"/>';
				strXml += '  <c:order val="0"/>';
				strXml += '  <c:tx>';
				strXml += '    <c:strRef>';
				strXml += '      <c:f>Sheet1!$A$2</c:f>';
				strXml += '      <c:strCache>';
				strXml += '        <c:ptCount val="1"/>';
				strXml += '        <c:pt idx="0"><c:v>'+ decodeXmlEntities(obj.name) +'</c:v></c:pt>';
				strXml += '      </c:strCache>';
				strXml += '    </c:strRef>';
				strXml += '  </c:tx>';
				strXml += '  <c:spPr>';
				strXml += '    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>';
				strXml += '    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
				strXml += '    <a:effectLst>';
				strXml += '      <a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="tl" rotWithShape="1" blurRad="38100" dist="23000" dir="5400000">';
				strXml += '        <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>';
				strXml += '      </a:outerShdw>';
				strXml += '    </a:effectLst>';
				strXml += '  </c:spPr>';
				strXml += '<c:explosion val="0"/>';

				// 2: "Data Point" block for every data row
				obj.labels.forEach(function(label,idx){
					strXml += '<c:dPt>';
					strXml += '  <c:idx val="'+ idx +'"/>';
					strXml += '  <c:explosion val="0"/>';
					strXml += '  <c:spPr>';
					strXml += '    <a:solidFill><a:srgbClr val="'+ rel.opts.chartColors[(idx+1 > rel.opts.chartColors.length ? (Math.floor(Math.random() * rel.opts.chartColors.length)) : idx)] +'"/></a:solidFill>';
					if ( rel.opts.dataBorder ) {
						strXml += '<a:ln w="'+ (rel.opts.dataBorder.pt * ONEPT) +'" cap="flat"><a:solidFill><a:srgbClr val="'+ rel.opts.dataBorder.color +'"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
					}
					strXml += '    <a:effectLst>';
					strXml += '      <a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="tl" rotWithShape="1" blurRad="38100" dist="23000" dir="5400000">';
					strXml += '        <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>';
					strXml += '      </a:outerShdw>';
					strXml += '    </a:effectLst>';
					strXml += '  </c:spPr>';
					strXml += '</c:dPt>';
				});

				// 3: "Data Label" block for every data Label
				strXml += '<c:dLbls>';
				obj.labels.forEach(function(label,idx){
					strXml += '<c:dLbl>';
					strXml += '  <c:idx val="'+ idx +'"/>';
					strXml += '    <c:numFmt formatCode="'+ rel.opts.dataLabelFormatCode +'" sourceLinked="0"/>';
					strXml += '    <c:txPr>';
					strXml += '      <a:bodyPr/><a:lstStyle/>';
					strXml += '      <a:p><a:pPr>';
					strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="'+ (rel.opts.dataLabelFontSize || DEF_FONT_SIZE) +'00" u="none">';
					strXml += '          <a:solidFill><a:srgbClr val="'+ (rel.opts.dataLabelColor || '000000') +'"/></a:solidFill>';
					strXml += '          <a:latin typeface="'+ (rel.opts.dataLabelFontFace || 'Arial') +'"/>';
					strXml += '        </a:defRPr>';
					strXml += '      </a:pPr></a:p>';
					strXml += '    </c:txPr>';
					strXml += '    <c:dLblPos val="'+ (rel.opts.dataLabelPosition || 'inEnd') +'"/>';
					strXml += '    <c:showLegendKey val="0"/>';
					strXml += '    <c:showVal val="'+ (rel.opts.showValue ? "1" : "0") +'"/>';
					strXml += '    <c:showCatName val="'+ (rel.opts.showLabel ? "1" : "0") +'"/>';
					strXml += '    <c:showSerName val="0"/>';
					strXml += '    <c:showPercent val="'+ (rel.opts.showPercent ? "1" : "0") +'"/>';
					strXml += '    <c:showBubbleSize val="0"/>';
					strXml += '  </c:dLbl>';
				});
				strXml += '<c:numFmt formatCode="'+ rel.opts.dataLabelFormatCode +'" sourceLinked="0"/>\
		            <c:txPr>\
		              <a:bodyPr/>\
		              <a:lstStyle/>\
		              <a:p>\
		                <a:pPr>\
		                  <a:defRPr b="0" i="0" strike="noStrike" sz="1800" u="none">\
		                    <a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/>\
		                  </a:defRPr>\
		                </a:pPr>\
		              </a:p>\
		            </c:txPr>\
		            <c:dLblPos val="ctr"/>\
		            <c:showLegendKey val="0"/>\
		            <c:showVal val="0"/>\
		            <c:showCatName val="1"/>\
		            <c:showSerName val="0"/>\
		            <c:showPercent val="1"/>\
		            <c:showBubbleSize val="0"/>\
		            <c:showLeaderLines val="0"/>';
				strXml += '</c:dLbls>';

				// 2: "Categories"
				strXml += '<c:cat>';
				strXml += '  <c:strRef>';
				strXml += '    <c:f>Sheet1!'+ '$B$1:$'+ getExcelColName(obj.labels.length) +'$1' +'</c:f>';
				strXml += '    <c:strCache>';
				strXml += '	     <c:ptCount val="'+ obj.labels.length +'"/>';
				obj.labels.forEach(function(label,idx){ strXml += '<c:pt idx="'+ idx +'"><c:v>'+ label +'</c:v></c:pt>'; });
				strXml += '    </c:strCache>';
				strXml += '  </c:strRef>';
				strXml += '</c:cat>';

				// 3: Create vals
				strXml += '  <c:val>';
				strXml += '    <c:numRef>';
				strXml += '      <c:f>Sheet1!'+ '$B$2:$'+ getExcelColName(obj.labels.length) +'$'+ 2 +'</c:f>';
				strXml += '      <c:numCache>';
				strXml += '	       <c:ptCount val="'+ obj.labels.length +'"/>';
				obj.values.forEach(function(value,idx){ strXml += '<c:pt idx="'+ idx +'"><c:v>'+ value +'</c:v></c:pt>'; });
				strXml += '      </c:numCache>';
				strXml += '    </c:numRef>';
				strXml += '  </c:val>';

				// 4: Close "SERIES"
				strXml += '  </c:ser>';
				strXml += '  <c:firstSliceAng val="0"/>';
				strXml += '</c:pieChart>';

				// Done with CHART.BAR
				break;
		}

		// B: Chart Properties + Options: Fill, Border, Legend
		{
			strXml += '  <c:spPr>';

			// OPTION: Fill
			strXml += ( rel.opts.fill ? genXmlColorSelection(rel.opts.fill) : '<a:noFill/>' );

			// OPTION: Border
			strXml += ( rel.opts.border ? '<a:ln w="'+ (rel.opts.border.pt * ONEPT) +'"'+' cap="flat">'+ genXmlColorSelection( rel.opts.border.color ) +'</a:ln>' : '<a:ln><a:noFill/></a:ln>' );

			// Close shapeProp/plotArea before Legend
			strXml += '    <a:effectLst/>';
			strXml += '  </c:spPr>';
			strXml += '</c:plotArea>';

			// OPTION: Legend
			// IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
			if ( rel.opts.showLegend ) strXml += '<c:legend><c:legendPos val="'+ rel.opts.legendPos +'"/><c:layout/><c:overlay val="0"/></c:legend>';
		}

		strXml += '  <c:plotVisOnly val="1"/>';
		strXml += '  <c:dispBlanksAs val="gap"/>';
		strXml += '</c:chart>';

		// C: CHARTSPACE SHAPE PROPS
		strXml += '<c:spPr>';
		strXml += '  <a:noFill/>';
		strXml += '  <a:ln w="12700" cap="flat"><a:noFill/><a:miter lim="400000"/></a:ln>';
		strXml += '  <a:effectLst/>';
		strXml += '</c:spPr>';

		// D: DATA (Add relID)
		strXml += '<c:externalData r:id="rId'+ (rel.rId-1) +'"><c:autoUpdate val="0"/></c:externalData>';

		// LAST: chartSpace end
		strXml += '</c:chartSpace>';

		return strXml;
	}

	/**
	* DESC: Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
	* EX:
		<p:txBody>
			<a:bodyPr wrap="none" lIns="50800" tIns="50800" rIns="50800" bIns="50800" anchor="ctr">
			</a:bodyPr>
			<a:lstStyle/>
			<a:p>
			  <a:pPr marL="228600" indent="-228600"><a:buSzPct val="100000"/><a:buChar char="&#x2022;"/></a:pPr>
			  <a:r>
				<a:t>bullet 1 </a:t>
			  </a:r>
			  <a:r>
				<a:rPr>
				  <a:solidFill><a:srgbClr val="7B2CD6"/></a:solidFill>
				</a:rPr>
				<a:t>colored text</a:t>
			  </a:r>
			</a:p>
		  </p:txBody>
	* NOTES:
	* - PPT text lines [lines followed by line-breaks] are createing using <p>-aragraph's
	* - Bullets are a paragprah-level formatting device
	*
	* @param slideObj (object) - slideObj -OR- table `cell` object
	* @returns XML string containing the param object's text and formatting
	*/
	function genXmlTextBody(slideObj) {
		// FIRST: Shapes without text, etc. may be sent here during buidl, but have no text to render so return empty string
		if ( !slideObj.text ) return '';

		// Create options if needed
		if ( !slideObj.options ) slideObj.options = {};

		// Vars
		var arrTextObjects = [];
		var tagStart = ( slideObj.options.isTableCell ? '<a:txBody>'  : '<p:txBody>' );
		var tagClose = ( slideObj.options.isTableCell ? '</a:txBody>' : '</p:txBody>' );
		var strSlideXml = tagStart;
		var strXmlBullet = '', strXmlLnSpc = '';
		var bulletLvl0Margin = 342900;
		var paragraphPropXml = '<a:pPr ';

		// STEP 1: Modify slideObj to be consistent array of `{ text:'', options:{} }`
		/* CASES:
			addText( 'string' )
			addText( 'line1\n line2' )
			addText( ['barry','allen'] )
			addText( [{text'word1'}, {text:'word2'}] )
			addText( [{text'line1\n line2'}, {text:'end word'}] )
		*/
		// A: Handle string/number
		if ( typeof slideObj.text === 'string' || typeof slideObj.text === 'number' ) {
			slideObj.text = [ {text:slideObj.text.toString(), options:(slideObj.options || {})} ];
		}

		// Grab options, format line-breaks, etc.
		if ( Array.isArray(slideObj.text) ) {
			slideObj.text.forEach(function(obj,idx){
				// A: Set options
				obj.options = obj.options || slideObj.options || {};
				if ( idx == 0 && !obj.options.bullet && slideObj.options.bullet ) obj.options.bullet = slideObj.options.bullet;

				// B: Cast to text-object and fix line-breaks (if needed)
				if ( typeof obj.text === 'string' || typeof obj.text === 'number' ) {
					obj.text = obj.text.toString().replace(/\r*\n/g, CRLF);
					// Plain strings like "hello \n world" need to have lineBreaks set to break as intended
					if ( obj.text.indexOf(CRLF) > -1 ) obj.options.breakLine = true;
				}

				// C: If text string has line-breaks, then create a separate text-object for each (much easier than dealing with split inside a loop below)
				if ( obj.text.split(CRLF).length > 0 ) {
					obj.text.toString().split(CRLF).forEach(function(line,idx){
						// Add line-breaks if not bullets/aligned (we add CRLF for those below in STEP 2)
						line += ( obj.options.breakLine && !obj.options.bullet && !obj.options.align ? CRLF : '' );
						arrTextObjects.push( {text:line, options:obj.options} );
					});
				}
				else {
					// NOTE: The replace used here is for non-textObjects (plain strings) eg:'hello\nworld'
					arrTextObjects.push( obj );
				}
			});
		}

		// STEP 2: Loop over each text object and create paragraph props, text run, etc.
		arrTextObjects.forEach(function(textObj,idx){
			// Clear/Increment loop vars
			paragraphPropXml = '<a:pPr '+ (textObj.options.rtlMode ? ' rtl="1" ' : '');
			strXmlBullet = '';
			textObj.options.lineIdx = idx;

			// A: Build paragraphProperties
			{
				// OPTION: align
				if ( textObj.options.align ) {
					switch ( textObj.options.align ) {
						case 'r':
						case 'right':
							paragraphPropXml += 'algn="r"';
							break;
						case 'c':
						case 'ctr':
						case 'center':
							paragraphPropXml += 'algn="ctr"';
							break;
						case 'justify':
							paragraphPropXml += 'algn="just"';
							break;
					}
				}

				if ( textObj.options.lineSpacing ) {
					strXmlLnSpc = '<a:lnSpc><a:spcPts val="' + textObj.options.lineSpacing + '00"/></a:lnSpc>';
				}

				// OPTION: indent
				if ( textObj.options.indentLevel && !isNaN(Number(textObj.options.indentLevel)) && textObj.options.indentLevel > 0 ) {
					paragraphPropXml += ' lvl="' + textObj.options.indentLevel + '"';
				}

				// Set core XML for use below
				paraPropXmlCore = paragraphPropXml;

				// OPTION: bullet
				// NOTE: OOXML uses the unicode character set for Bullets
				// EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
				if ( typeof textObj.options.bullet === 'object' ) {
					if ( textObj.options.bullet.type ) {
						if ( textObj.options.bullet.type.toString().toLowerCase() == "number" ) {
							paragraphPropXml += ' marL="'+ (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin+(bulletLvl0Margin*textObj.options.indentLevel) : bulletLvl0Margin) +'" indent="-'+bulletLvl0Margin+'"';
							strXmlBullet = '<a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>';
						}
					}
					else if ( textObj.options.bullet.code ) {
						var bulletCode = '&#x'+ textObj.options.bullet.code +';';

						// Check value for hex-ness (s/b 4 char hex)
						if ( /^[0-9A-Fa-f]{4}$/.test(textObj.options.bullet.code) == false ) {
							console.warn('Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!');
							bulletCode = BULLET_TYPES['DEFAULT'];
						}

						paragraphPropXml += ' marL="'+ (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin+(bulletLvl0Margin*textObj.options.indentLevel) : bulletLvl0Margin) +'" indent="-'+bulletLvl0Margin+'"';
						strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="'+ bulletCode +'"/>';
					}
				}
				// DEPRECATED: old bool value (FIXME:Drop in 2.0)
				else if ( textObj.options.bullet == true ) {
					paragraphPropXml += ' marL="'+ (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin+(bulletLvl0Margin*textObj.options.indentLevel) : bulletLvl0Margin) +'" indent="-'+bulletLvl0Margin+'"';
					strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="'+ BULLET_TYPES['DEFAULT'] +'"/>';
				}

				// Close Paragraph-Properties --------------------
				// IMPORTANT: strXmlLnSpc must precede strXmlBullet for bullet lineSpacing to work (PPT-Online)
				paragraphPropXml += '>'+ strXmlLnSpc + strXmlBullet +'</a:pPr>';
			}

			// B: Start paragraph if this is the first text obj, or if current textObj is about to be bulleted or aligned
			if ( idx == 0 ) {
				// TODO: FIXME: WIP: ISSUE#79
				/*
				// DESIGN: `addText()` can take an array of text objects, but has options like valign, and we need to pass those
				textObj.options.bodyProp = ( textObj.options.bodyProp || {} );
				textObj.options.bodyProp.anchor = ( textObj.options.valign || slideObj.options.valign );
				//console.log(textObj.options.bodyProp.anchor);
				*/
				// ISSUE#69: Adding bodyProps more than once inside <p:txBody> causes "corrupt presentation" errors in PPT 2007, PPT 2010.
				strSlideXml += genXmlBodyProperties(textObj.options);
				// NOTE: Shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
				// FIXME: LINE align doesnt work (code diff is substantial!)
				if ( textObj.options.h == 0 && textObj.options.line && textObj.options.align ) {
					strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>';
				}
				else {
					strSlideXml += '<a:lstStyle/>';
				}
				strSlideXml += '<a:p>' + paragraphPropXml;
			}
			else if ( idx > 0 && (typeof textObj.options.bullet !== 'undefined' || typeof textObj.options.align !== 'undefined') ) {
				strSlideXml += '</a:p><a:p>' + paragraphPropXml;
			}

			// C: Inherit any main options (color, font_size, etc.)
			// We only pass the text.options to genXmlTextRun (not the Slide.options),
			// so the run building function cant just fallback to Slide.color, therefore, we need to do that here before passing options below.
			$.each(slideObj.options, function(key,val){
				// NOTE: This loop will pick up unecessary keys (`x`, etc.), but it doesnt hurt anything
				if ( key != 'bullet' && !textObj.options[key] ) textObj.options[key] = val;
			});

			// D: Add formatted textrun
			strSlideXml += genXmlTextRun(textObj.options, textObj.text);
		});

		// STEP 3: Close paragraphProperties and the current open paragraph
		strSlideXml += '</a:p>';

		// STEP 4: Close the textBody
		strSlideXml += tagClose;

		// LAST: Return XML
		return strSlideXml;
	}

	/**
	<a:r>
	  <a:rPr lang="en-US" sz="2800" dirty="0" smtClean="0">
		<a:solidFill>
		  <a:srgbClr val="00FF00">
		  </a:srgbClr>
		</a:solidFill>
		<a:latin typeface="Courier New" pitchFamily="34" charset="0"/>
		<a:cs typeface="Courier New" pitchFamily="34" charset="0"/>
	  </a:rPr>
	  <a:t>Misc font/color, size = 28</a:t>
	</a:r>
	*/
	function genXmlTextRun(opts, text_string) {
		var xmlTextRun = '';
		var paraProp = '';
		var parsedText;

		// BEGIN runProperties
		var startInfo = '<a:rPr lang="en-US" ';
		startInfo += ( opts.bold      ? ' b="1"' : '' );
		startInfo += ( opts.font_size ? ' sz="'+ Math.round(opts.font_size) +'00"' : '' ); // NOTE: Use round so sizes like '7.5' wont cause corrupt pres.
		startInfo += ( opts.italic    ? ' i="1"' : '' );
		startInfo += ( opts.underline || opts.hyperlink ? ' u="sng"' : '' );
		startInfo += ( opts.subscript ? ' baseline="-40000"' : (opts.superscript  ? ' baseline="30000"' : '') );
		// not doc in API yet: startInfo += ( opts.char_spacing ? ' spc="' + (text_info.char_spacing * 100) + '" kern="0"' : '' ); // IMPORTANT: Also disable kerning; otherwise text won't actually expand
		startInfo += ' dirty="0" smtClean="0">';
		// Color and Font are children of <a:rPr>, so add them now before closing the runProperties tag
		if ( opts.color || opts.font_face ) {
			if ( opts.color     ) startInfo += genXmlColorSelection( opts.color );
			if ( opts.font_face ) startInfo += '<a:latin typeface="' + opts.font_face + '" pitchFamily="34" charset="0"/><a:cs typeface="' + opts.font_face + '" pitchFamily="34" charset="0"/>';
		}

		// Hyperlink support
		if ( opts.hyperlink ) {
			if ( typeof opts.hyperlink !== 'object' ) console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` ");
			else if ( !opts.hyperlink.url || typeof opts.hyperlink.url !== 'string' ) console.log("ERROR: 'hyperlink.url is required and/or should be a string'");
			else if ( opts.hyperlink.url ) {
				// FIXME-20170410: FUTURE-FEATURE: color (link is always blue in Keynote and PPT online, so usual text run above isnt honored for links..?)
				//startInfo += '<a:uFill>'+ genXmlColorSelection('0000FF') +'</a:uFill>'; // Breaks PPT2010! (Issue#74)
				startInfo += '<a:hlinkClick r:id="rId'+ opts.hyperlink.rId +'" invalidUrl="" action="" tgtFrame="" tooltip="'+ (opts.hyperlink.tooltip ? decodeXmlEntities(opts.hyperlink.tooltip) : '') +'" history="1" highlightClick="0" endSnd="0"/>';
			}
		}

		// END runProperties
		startInfo += '</a:rPr>';

		// LINE-BREAKS/MULTI-LINE: Split text into multi-p:
		parsedText = text_string.split(CRLF);
		if ( parsedText.length > 1 ) {
			var outTextData = '';
			for ( var i = 0, total_size_i = parsedText.length; i < total_size_i; i++ ) {
				outTextData += '<a:r>' + startInfo+ '<a:t>' + decodeXmlEntities(parsedText[i]);
				// Stop/Start <p>aragraph as long as there is more lines ahead (otherwise its closed at the end of this function)
				if ( (i + 1) < total_size_i ) outTextData += (opts.breakLine ? CRLF : '') + '</a:t></a:r>';
			}
			xmlTextRun = outTextData;
		}
		else {
			// Handle cases where addText `text` was an array of objects - if a text object doesnt contain a '\n' it still need alignment!
			// The first pPr-align is done in makeXml - use line countr to ensure we only add subsequently as needed
			xmlTextRun = ( (opts.align && opts.lineIdx > 0) ? paraProp : '') + '<a:r>' + startInfo+ '<a:t>' + decodeXmlEntities(text_string);
		}

		// Return paragraph with text run
		return xmlTextRun + '</a:t></a:r>';
	}

	/**
	* DESC: Builds <a:bodyPr></a:bodyPr> tag
	*/
	function genXmlBodyProperties(objOptions) {
		var bodyProperties = '<a:bodyPr';

		if ( objOptions && objOptions.bodyProp ) {
			// A: Enable or disable textwrapping none or square:
			( objOptions.bodyProp.wrap ) ? bodyProperties += ' wrap="' + objOptions.bodyProp.wrap + '" rtlCol="0"' : bodyProperties += ' wrap="square" rtlCol="0"';

			// B: Set anchorPoints ('t','ctr',b'):
			if ( objOptions.bodyProp.anchor ) bodyProperties += ' anchor="' + objOptions.bodyProp.anchor + '"';
			//if ( objOptions.bodyProp.anchorCtr ) bodyProperties += ' anchorCtr="' + objOptions.bodyProp.anchorCtr + '"';

			// C: Textbox margins [padding]:
			if ( objOptions.bodyProp.bIns || objOptions.bodyProp.bIns == 0 ) bodyProperties += ' bIns="' + objOptions.bodyProp.bIns + '"';
			if ( objOptions.bodyProp.lIns || objOptions.bodyProp.lIns == 0 ) bodyProperties += ' lIns="' + objOptions.bodyProp.lIns + '"';
			if ( objOptions.bodyProp.rIns || objOptions.bodyProp.rIns == 0 ) bodyProperties += ' rIns="' + objOptions.bodyProp.rIns + '"';
			if ( objOptions.bodyProp.tIns || objOptions.bodyProp.tIns == 0 ) bodyProperties += ' tIns="' + objOptions.bodyProp.tIns + '"';

			// D: Close <a:bodyPr element
			bodyProperties += '>';

			// E: NEW: Add autofit type tags
			if ( objOptions.shrinkText ) bodyProperties += '<a:normAutofit fontScale="85000" lnSpcReduction="20000" />'; // MS-PPT > Format Shape > Text Options: "Shrink text on overflow"
			// MS-PPT > Format Shape > Text Options: "Resize shape to fit text" [spAutoFit]
			// NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
			bodyProperties += ( objOptions.bodyProp.autoFit !== false ? '<a:spAutoFit/>' : '' );

			// LAST: Close bodyProp
			bodyProperties += '</a:bodyPr>';
		}
		else {
			// DEFAULT:
			bodyProperties += ' wrap="square" rtlCol="0">';
			bodyProperties += '</a:bodyPr>';
		}

		// LAST: Return Close bodyProp
		return ( objOptions.isTableCell ? '<a:bodyPr/>' : bodyProperties );
	}

	function genXmlColorSelection(color_info, back_info) {
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
		strXml += ' <Default Extension="xml" ContentType="application/xml"/>';
		strXml += ' <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		strXml += ' <Default Extension="jpeg" ContentType="image/jpeg"/>';
		strXml += ' <Default Extension="png" ContentType="image/png"/>';
		strXml += ' <Default Extension="gif" ContentType="image/gif"/>';
		strXml += ' <Default Extension="m4v" ContentType="video/mp4"/>'; // hard-coded as extn!=type
		strXml += ' <Default Extension="mp4" ContentType="video/mp4"/>'; // same here, we have to add as it wont be added in loop below
		gObjPptx.slides.forEach(function(slide,idx){
			slide.rels.forEach(function(rel,idy){
				if ( rel.type != 'image' && rel.type != 'online' && rel.type != 'chart' && rel.extn != 'm4v' && strXml.indexOf(rel.type) == -1 )
					strXml += ' <Default Extension="'+ rel.extn +'" ContentType="'+ rel.type +'"/>';
			});
		});
		strXml += ' <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
		strXml += ' <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>';

		strXml += ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
		strXml += ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
		strXml += ' <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>';
		strXml += ' <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>';
		strXml += ' <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>';
		strXml += ' <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
		strXml += ' <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>';
		gObjPptx.slides.forEach(function(slide,idx){
			strXml += '<Override PartName="/ppt/slideMasters/slideMaster'+ (idx+1) +'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>';
			strXml += '<Override PartName="/ppt/slideLayouts/slideLayout'+ (idx+1) +'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
			strXml += '<Override PartName="/ppt/slides/slide'            + (idx+1) +'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
		});

		// Add charts (if any)
		gObjPptx.slides.forEach(function(slide,idx){
			slide.rels.forEach(function(rel,idy){
				if ( rel.type == 'chart' ) {
					strXml += ' <Override PartName="'+ rel.Target +'" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
				}
			});
		});

		strXml += '</Types>';

		return strXml;
	}

	function makeXmlRootRels() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF
					+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
					+ '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
					+ '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
					+ '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
					+ '</Relationships>';
		return strXml;
	}

	function makeXmlApp() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF;
		strXml += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
		strXml += '<TotalTime>0</TotalTime>';
		strXml += '<Words>0</Words>';
		strXml += '<Application>Microsoft Office PowerPoint</Application>';
		strXml += '<PresentationFormat>On-screen Show</PresentationFormat>';
		strXml += '<Paragraphs>0</Paragraphs>';
		strXml += '<Slides>'+ gObjPptx.slides.length +'</Slides>';
		strXml += '<Notes>0</Notes>';
		strXml += '<HiddenSlides>0</HiddenSlides>';
		strXml += '<MMClips>0</MMClips>';
		strXml += '<ScaleCrop>false</ScaleCrop>';
		strXml += '<HeadingPairs>';
		strXml += '  <vt:vector size="4" baseType="variant">';
		strXml += '    <vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>';
		strXml += '    <vt:variant><vt:i4>1</vt:i4></vt:variant>';
		strXml += '    <vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>';
		strXml += '    <vt:variant><vt:i4>'+ gObjPptx.slides.length +'</vt:i4></vt:variant>';
		strXml += '  </vt:vector>';
		strXml += '</HeadingPairs>';
		strXml += '<TitlesOfParts>';
		strXml += '<vt:vector size="'+ (gObjPptx.slides.length+1) +'" baseType="lpstr">';
		strXml += '<vt:lpstr>Office Theme</vt:lpstr>';
		$.each(gObjPptx.slides, function(idx,slideObj){ strXml += '<vt:lpstr>Slide '+ (idx+1) +'</vt:lpstr>'; });
		strXml += '</vt:vector>';
		strXml += '</TitlesOfParts>';
		strXml += '<Company>'+gObjPptx.company+'</Company>';
		strXml += '<LinksUpToDate>false</LinksUpToDate>';
		strXml += '<SharedDoc>false</SharedDoc>';
		strXml += '<HyperlinksChanged>false</HyperlinksChanged>';
		strXml += '<AppVersion>15.0000</AppVersion>';
		strXml += '</Properties>';

		return strXml;
	}

	function makeXmlCore() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF;
		strXml += '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
		strXml += '<dc:title>'+ decodeXmlEntities(gObjPptx.title) +'</dc:title>';
		strXml += '<dc:subject>'+ decodeXmlEntities(gObjPptx.subject) +'</dc:subject>';
		strXml += '<dc:creator>'+ decodeXmlEntities(gObjPptx.author) +'</dc:creator>';
		strXml += '<cp:lastModifiedBy>'+ decodeXmlEntities(gObjPptx.author) +'</cp:lastModifiedBy>';
		strXml += '<cp:revision>'+ gObjPptx.revision +'</cp:revision>';
		strXml += '<dcterms:created xsi:type="dcterms:W3CDTF">'+ new Date().toISOString() +'</dcterms:created>';
		strXml += '<dcterms:modified xsi:type="dcterms:W3CDTF">'+ new Date().toISOString() +'</dcterms:modified>';
		strXml += '</cp:coreProperties>';
		return strXml;
	}

	function makeXmlPresentationRels() {
		var intRelNum = 0;
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF;
		strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
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
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF;
		strXml += '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1">'+CRLF
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

		// STEP 4: Loop over all Slide.data objects and add them to this slide ===============================
		$.each(inSlide.data, function(idx,slideObj){
			var x = 0, y = 0, cx = getSmartParseNumber('75%','X'), cy = 0;
			var locationAttr = "", shapeType = null;

			// A: Set option vars
			slideObj.options = slideObj.options || {};

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

			// B: Add OBJECT to current Slide ----------------------------
			switch ( slideObj.type ) {
				case 'table':
					// FIRST: Ensure we have rows - otherwise, bail!
					if ( !slideObj.arrTabRows || (Array.isArray(slideObj.arrTabRows) && slideObj.arrTabRows.length == 0) ) break;

					// Set table vars
					var objTableGrid = {};
					var arrTabRows = slideObj.arrTabRows;
					var objTabOpts = slideObj.options;
					var intColCnt = 0, intColW = 0;

					// Calc number of columns
					// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
					// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
					arrTabRows[0].forEach(function(cell,idx){
						var cellOpts = cell.options || cell.opts || null; // DEPRECATED (`opts`)
						intColCnt += ( cellOpts && cellOpts.colspan ? cellOpts.colspan : 1 );
					});

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
					/* Only useful for rowspan/colspan testing
					if ( objTabOpts.debug ) {
						console.table(objTableGrid);
						var arrText = [];
						$.each(objTableGrid, function(i,row){ var arrRow = []; $.each(row,function(i,cell){ arrRow.push(cell.text); }); arrText.push(arrRow); });
						console.table( arrText );
					}
					*/

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
							// FIRST: Create cell if needed (handle [null] and other manner of junk values)
							// IMPORTANT: MS-PPTX PROBLEM: using '' will cause PPT to use its own default font/size! (Arial/18 in US)
							// SOLN: Pass a space instead to cement formatting options (Issue #20)
							if ( typeof cell === 'undefined' || cell == null ) cell = { text:' ', options:{} };

							// 1: "hmerge" cells are just place-holders in the table grid - skip those and go to next cell
							if ( cell.hmerge ) return;

							// 2: OPTIONS: Build/set cell options (blocked for code folding) ===========================
							{
								var cellOpts = cell.options || cell.opts || {};
								if ( typeof cell === 'number' || typeof cell === 'string' ) cell = { text:cell.toString() };
								cellOpts.isTableCell = true; // Used to create textBody XML
								cell.options = cellOpts;

								// B: Apply default values (tabOpts being used when cellOpts dont exist):
								// SEE: http://officeopenxml.com/drwTableCellProperties-alignment.php
								['align','bold','border','color','fill','font_face','font_size','margin','marginPt','underline','valign']
								.forEach(function(name,idx){
									if ( objTabOpts[name] && !cellOpts[name] && cellOpts[name] != 0 ) cellOpts[name] = objTabOpts[name];
								});

								var cellValign  = (cellOpts.valign)     ? ' anchor="'+ cellOpts.valign.replace(/^c$/i,'ctr').replace(/^m$/i,'ctr').replace('center','ctr').replace('middle','ctr').replace('top','t').replace('btm','b').replace('bottom','b') +'"' : '';
								var cellColspan = (cellOpts.colspan)    ? ' gridSpan="'+ cellOpts.colspan +'"' : '';
								var cellRowspan = (cellOpts.rowspan)    ? ' rowSpan="'+ cellOpts.rowspan +'"' : '';
								var cellFill    = ((cell.optImp && cell.optImp.fill)  || cellOpts.fill ) ? ' <a:solidFill><a:srgbClr val="'+ ((cell.optImp && cell.optImp.fill) || cellOpts.fill.replace('#','')) +'"/></a:solidFill>' : '';
								var cellMargin  = ( cellOpts.margin == 0 || cellOpts.margin ? cellOpts.margin : (cellOpts.marginPt || DEF_CELL_MARGIN_PT) );
								if ( !Array.isArray(cellMargin) && typeof cellMargin === 'number' ) cellMargin = [cellMargin,cellMargin,cellMargin,cellMargin];
								cellMargin = ' marL="'+ cellMargin[3]*ONEPT +'" marR="'+ cellMargin[1]*ONEPT +'" marT="'+ cellMargin[0]*ONEPT +'" marB="'+ cellMargin[2]*ONEPT +'"';
							}

							// FIXME: Cell NOWRAP property (text wrap: add to a:tcPr (horzOverflow="overflow" or whatev opts exist)

							// 3: ROWSPAN: Add dummy cells for any active rowspan
							if ( cell.vmerge ) {
								strXml += '<a:tc vMerge="1"><a:tcPr/></a:tc>';
								return;
							}

							// 4: Set CELL content and properties ==================================
							strXml += '<a:tc'+ cellColspan + cellRowspan +'>' + genXmlTextBody(cell) + '<a:tcPr'+ cellMargin + cellValign +'>';

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
							strXml += cellFill;
							strXml += '  </a:tcPr>';
							strXml += ' </a:tc>';

							// LAST: COLSPAN: Add a 'merged' col for each column being merged (SEE: http://officeopenxml.com/drwTableGrid.php)
							if ( cellOpts.colspan ) {
								for (var tmp=1; tmp<Number(cellOpts.colspan); tmp++) { strXml += '<a:tc hMerge="1"><a:tcPr/></a:tc>'; }
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

					// A: Start SHAPE =======================================================
					strSlideXml += '<p:sp>';

					// B: The addition of the "txBox" attribute is the sole determiner of if an object is a Shape or Textbox
					strSlideXml += '<p:nvSpPr><p:cNvPr id="'+ (idx+2) +'" name="Object '+ (idx+1) +'"/>';
					strSlideXml += '<p:cNvSpPr' + ((slideObj.options && slideObj.options.isTextBox) ? ' txBox="1"/><p:nvPr/>' : '/><p:nvPr/>');
					strSlideXml += '</p:nvSpPr>';
					strSlideXml += '<p:spPr><a:xfrm' + locationAttr + '>';
					strSlideXml += '<a:off x="'  + x  + '" y="'  + y  + '"/>';
					strSlideXml += '<a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>';
					strSlideXml += '<a:prstGeom prst="' + shapeType.name + '"><a:avLst>'
						+ (slideObj.options.rectRadius ? '<a:gd name="adj" fmla="val '+ Math.round(slideObj.options.rectRadius * EMU * 100000 / Math.min(cx, cy)) +'" />' : '')
						+ '</a:avLst></a:prstGeom>';

					// Option: FILL
					strSlideXml += ( slideObj.options.fill ? genXmlColorSelection(slideObj.options.fill) : '<a:noFill/>' );

					// Shape Type: LINE: line color
					if ( slideObj.options.line ) {
						strSlideXml += '<a:ln'+ ( slideObj.options.line_size ? ' w="'+ (slideObj.options.line_size*ONEPT) +'"' : '') +'>';
						strSlideXml += genXmlColorSelection( slideObj.options.line );
						if ( slideObj.options.line_dash ) strSlideXml += '<a:prstDash val="' + slideObj.options.line_dash + '"/>';
						if ( slideObj.options.line_head ) strSlideXml += '<a:headEnd type="' + slideObj.options.line_head + '"/>';
						if ( slideObj.options.line_tail ) strSlideXml += '<a:tailEnd type="' + slideObj.options.line_tail + '"/>';
						strSlideXml += '</a:ln>';
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

					/* FIXME: FUTURE: Text wrapping (copied from MS-PPTX export)
					// Commented out b/c i'm not even sure this works - current code produces text that wraps in shapes and textboxes, so...
					if ( slideObj.options.textWrap ) {
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
					strSlideXml += genXmlTextBody(slideObj);

					// LAST: Close SHAPE =======================================================
					strSlideXml += '</p:sp>';
					break;

				case 'image':
			        strSlideXml += '<p:pic>';
					strSlideXml += '  <p:nvPicPr>'
					strSlideXml += '    <p:cNvPr id="'+ (idx + 2) +'" name="Object '+ (idx + 1) +'" descr="'+ slideObj.image +'">';
					if ( slideObj.hyperlink ) strSlideXml += '<a:hlinkClick r:id="rId'+ slideObj.hyperlink.rId +'" tooltip="'+ (slideObj.hyperlink.tooltip ? decodeXmlEntities(slideObj.hyperlink.tooltip) : '') +'"/>';
					strSlideXml += '    </p:cNvPr>';
			        strSlideXml += '    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/>';
					strSlideXml += '  </p:nvPicPr>';
					strSlideXml += '<p:blipFill><a:blip r:embed="rId' + slideObj.imageRid + '" cstate="print"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>';
					strSlideXml += '<p:spPr>'
					strSlideXml += ' <a:xfrm' + locationAttr + '>'
					strSlideXml += '  <a:off  x="' + x  + '"  y="' + y  + '"/>'
					strSlideXml += '  <a:ext cx="' + cx + '" cy="' + cy + '"/>'
					strSlideXml += ' </a:xfrm>'
					strSlideXml += ' <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
					strSlideXml += '</p:spPr>';
					strSlideXml += '</p:pic>';
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
						strSlideXml += '<p:pic>';
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

				case 'chart':
					strSlideXml += '<p:graphicFrame>';
					strSlideXml += ' <p:nvGraphicFramePr>';
					strSlideXml += '   <p:cNvPr id="'+ (idx + 2) +'" name="Chart '+ (idx + 1) +'"/>';
					strSlideXml += '   <p:cNvGraphicFramePr/>';
					strSlideXml += '   <p:nvPr/>';
					strSlideXml += ' </p:nvGraphicFramePr>';
					strSlideXml += ' <p:xfrm>'
					strSlideXml += '  <a:off  x="' + x  + '"  y="' + y  + '"/>'
					strSlideXml += '  <a:ext cx="' + cx + '" cy="' + cy + '"/>'
					strSlideXml += ' </p:xfrm>'
					strSlideXml += ' <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
					strSlideXml += '  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">';
					strSlideXml += '   <c:chart r:id="rId'+ (slideObj.chartRid) +'" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>';
					strSlideXml += '  </a:graphicData>';
					strSlideXml += ' </a:graphic>';
					strSlideXml += '</p:graphicFrame>';
					break;
			}
		});

		// STEP 5: Add slide numbers last (if any)
		if ( inSlide.slideNumberObj || inSlide.hasSlideNumber ) {
			if ( !inSlide.slideNumberObj ) inSlide.slideNumberObj = { x:0.3, y:'90%' }

			// TODO: FIXME: page numbers over 99 wrap in PPT-2013 (Desktop)
			strSlideXml += '<p:sp>'
				+ '  <p:nvSpPr>'
				+ '  <p:cNvPr id="25" name="Shape 25"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr></p:nvSpPr>'
				+ '  <p:spPr>'
				+ '    <a:xfrm><a:off x="'+ getSmartParseNumber(inSlide.slideNumberObj.x, 'X') +'" y="'+ getSmartParseNumber(inSlide.slideNumberObj.y, 'Y') +'"/><a:ext cx="400000" cy="300000"/></a:xfrm>'
				+ '    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
				+ '    <a:extLst>'
				+ '      <a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}"><ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext>'
				+ '    </a:extLst>'
				+ '  </p:spPr>';
			// ISSUE #68: "Page number styling"
			strSlideXml += '<p:txBody>';
			strSlideXml += '  <a:bodyPr/>';
			strSlideXml += '  <a:lstStyle><a:lvl1pPr>';
			if ( inSlide.slideNumberObj.fontFace || inSlide.slideNumberObj.fontSize || inSlide.slideNumberObj.color ) {
				strSlideXml += '<a:defRPr sz="'+ (inSlide.slideNumberObj.fontSize || '12') +'00">';
				if ( inSlide.slideNumberObj.color ) strSlideXml += genXmlColorSelection(inSlide.slideNumberObj.color);
				if ( inSlide.slideNumberObj.fontFace ) strSlideXml += '<a:latin typeface="'+ inSlide.slideNumberObj.fontFace +'"/><a:cs typeface="'+ inSlide.slideNumberObj.fontFace +'"/>';
				strSlideXml += '</a:defRPr>';
			}
			strSlideXml += '</a:lvl1pPr></a:lstStyle>';
			strSlideXml += '<a:p><a:pPr/><a:fld id="'+SLDNUMFLDID+'" type="slidenum"/></a:p></p:txBody>'
			strSlideXml += '</p:sp>';
		}

		// STEP 6: Close spTree and finalize slide XML
		strSlideXml += '</p:spTree>';
		/* FIXME: Remove this in 2.0.0 (commented for 1.6.0)
		strSlideXml += '<p:extLst>';
		strSlideXml += ' <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">';
		strSlideXml += '  <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1544976994"/>';
		strSlideXml += ' </p:ext>';
		strSlideXml += '</p:extLst>';
		*/
		strSlideXml += '</p:cSld>';
		strSlideXml += '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>';
		strSlideXml += '</p:sld>';

		// LAST: Return
		return strSlideXml;
	}

	function makeXmlSlideLayoutRel(inSlideNum) {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF;
			strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
			//?strXml += '  <Relationship Id="rId'+ inSlideNum +'" Target="../slideMasters/slideMaster'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
			//strXml += '  <Relationship Id="rId1" Target="../slideMasters/slideMaster'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
			strXml += '  <Relationship Id="rId1" Target="../slideMasters/slideMaster1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
			strXml += '</Relationships>';
		//
		return strXml;
	}

	function makeXmlSlideRel(inSlideNum) {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF
			+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
			+ ' <Relationship Id="rId1" Target="../slideLayouts/slideLayout'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"/>';

		// Add any rels for this Slide (image/audio/video/youtube/chart)
		gObjPptx.slides[inSlideNum-1].rels.forEach(function(rel,idx){
			if      ( rel.type.toLowerCase().indexOf('image')  > -1 ) {
				strXml += '<Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>';
			}
			else if ( rel.type.toLowerCase().indexOf('chart')  > -1 ) {
				strXml += '<Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>';
			}
			else if ( rel.type.toLowerCase().indexOf('audio')  > -1 ) {
				// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
				if ( strXml.indexOf(' Target="'+ rel.Target +'"') > -1 )
					strXml += '<Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>';
				else
					strXml += '<Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"/>';
			}
			else if ( rel.type.toLowerCase().indexOf('video')  > -1 ) {
				// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
				if ( strXml.indexOf(' Target="'+ rel.Target +'"') > -1 )
					strXml += '<Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>';
				else
					strXml += '<Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>';
			}
			else if ( rel.type.toLowerCase().indexOf('online') > -1 ) {
				// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
				if ( strXml.indexOf(' Target="'+ rel.Target +'"') > -1 )
					strXml += '<Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" Type="http://schemas.microsoft.com/office/2007/relationships/image"/>';
				else
					strXml += '<Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>';
			}
			else if ( rel.type.toLowerCase().indexOf('hyperlink') > -1 ) {
				strXml += '<Relationship Id="rId'+ rel.rId +'" Target="'+ rel.Target +'" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"/>';
			}
		});

		strXml += '</Relationships>';
		//
		return strXml;
	}

	function makeXmlSlideMaster() {
		var intSlideLayoutId = 2147483649;
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF
					+ '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
					+ '  <p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr>'
					+ '<p:cNvPr id="2" name="Title Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>'
					+ '<p:cNvPr id="3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>'
					+ '<p:cNvPr id="4" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="6356350"/><a:ext cx="2133600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{F8166F1F-CE9B-4651-A6AA-CD717754106B}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>12/25/2015</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>'
					+ '<p:cNvPr id="5" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3124200" y="6356350"/><a:ext cx="2895600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr>'
					+ '<p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="6553200" y="6356350"/><a:ext cx="2133600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="'+SLDNUMFLDID+'" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t></a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
					+ '<p:sldLayoutIdLst>';
		// Create a sldLayout for each SLIDE
		for ( var idx=1; idx<=gObjPptx.slides.length; idx++ ) {
			strXml += ' <p:sldLayoutId id="'+ intSlideLayoutId +'" r:id="rId'+ idx +'"/>';
			intSlideLayoutId++;
		}
		strXml += '</p:sldLayoutIdLst>'
					+ '<p:txStyles>'
					+ ' <p:titleStyle>'
					+ '  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>'
					+ ' </p:titleStyle>'
					+ ' <p:bodyStyle>'
					+ '  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>'
					+ '  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>'
					+ '  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>'
					+ '  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>'
					+ '  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>'
					+ '  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>'
					+ '  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>'
					+ '  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>'
					+ '  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="?"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>'
					+ ' </p:bodyStyle>'
					+ ' <p:otherStyle>'
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
					+ ' </p:otherStyle>'
					+ '</p:txStyles>'
					+ '</p:sldMaster>';
		//
		return strXml;
	}

	function makeXmlSlideMasterRel() {
		// FIXME: create a slideLayout for each SLDIE
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF
					+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		for ( var idx=1; idx<=gObjPptx.slides.length; idx++ ) {
			strXml += '  <Relationship Id="rId'+ idx +'" Target="../slideLayouts/slideLayout'+ idx +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"/>';
		}
		strXml += '  <Relationship Id="rId'+ (gObjPptx.slides.length+1) +'" Target="../theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"/>';
		strXml += '</Relationships>';
		//
		return strXml;
	}

	// XML-GEN: Last 5 functions create root /ppt files

	function makeXmlTheme() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF;
		strXml += '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Office Theme">\
						<a:themeElements>\
						  <a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>\
						  <a:dk2><a:srgbClr val="A7A7A7"/></a:dk2>\
						  <a:lt2><a:srgbClr val="535353"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3>\
						  <a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5>\
						  <a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/>\
						  </a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/>\
						</a:theme>';
		return strXml;
	}

	function makeXmlPresentation() {
		var intCurPos = 0;
		// REF: http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF
			+ '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
			+ 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
			+ 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '+ (gObjPptx.rtlMode ? 'rtl="1"' : '') +' saveSubsetFonts="1">';

		// STEP 1: Build SLIDE master list
		strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>';
		strXml += '<p:sldIdLst>';
		for ( var idx=0; idx<gObjPptx.slides.length; idx++ ) {
			strXml += '<p:sldId id="' + (idx + 256) + '" r:id="rId' + (idx + 2) + '"/>';
		}
		strXml += '</p:sldIdLst>';

		// STEP 2: Build SLIDE text styles
		strXml += '<p:sldSz cx="'+ gObjPptx.pptLayout.width +'" cy="'+ gObjPptx.pptLayout.height +'" type="'+ gObjPptx.pptLayout.name +'"/>'
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
		strXml += '</p:defaultTextStyle>';

		strXml += '<p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext></p:extLst>'
				+ '</p:presentation>';

		return strXml;
	}

	function makeXmlPresProps() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF
					+ '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
					+ '  <p:extLst>'
					+ '    <p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="0"/></p:ext>'
					+ '    <p:ext uri="{D31A062A-798A-4329-ABDD-BBA856620510}"><p14:defaultImageDpi xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="220"/></p:ext>'
					+ '    <p:ext uri="{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}"><p15:chartTrackingRefBased xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" val="1"/></p:ext>'
					+ '  </p:extLst>'
					+ '</p:presentationPr>';
		return strXml;
	}

	function makeXmlTableStyles() {
		// SEE: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF
					+ '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>';
		return strXml;
	}

	function makeXmlViewProps() {
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+CRLF
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

	/**
	 * Expose a couple private helper functions from above
	 */
	this.inch2Emu = inch2Emu;
	this.rgbToHex = rgbToHex;

	/**
	 * Gets the version of this library
	 */
	this.getVersion = function getVersion() {
		return APP_VER;
	};

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
		if ( typeof inBool !== 'boolean' ) return;
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
		if ( typeof inLayout === 'object' && inLayout.width && inLayout.height ) {
			LAYOUTS['LAYOUT_USER'].width  = Math.round(Number(inLayout.width ) * EMU);
			LAYOUTS['LAYOUT_USER'].height = Math.round(Number(inLayout.height) * EMU);

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
	 * Sets the Presentation's Title
	 */
	this.setTitle = function setTitle(inStrTitle) {
		gObjPptx.title = inStrTitle || 'PptxGenJS Presentation';
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
		gObjPptx.revision = gObjPptx.revision.replace(/[\.\,\-]+/gi,'');
		if ( isNaN(gObjPptx.revision) ) gObjPptx.revision = '1';
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
	this.save = function save(inStrExportName, callback) {
		var intRels = 0, arrRelsDone = [];

		// STEP 1: Set export title (if any)
		if ( inStrExportName ) gObjPptx.fileName = inStrExportName;

		// STEP 2: Read/Encode Images
		// B: Total all physical rels across the Presentation
		// PERF: Only send unique paths for encoding (encoding func will find and fill *ALL* matching paths across the Presentation)
		gObjPptx.slides.forEach(function(slide,idx){
			slide.rels.forEach(function(rel,idy){
				// Read and Encode each image into base64 for use in export
				if ( rel.type != 'online' && rel.type != 'chart' && !rel.data && $.inArray(rel.path, arrRelsDone) == -1 ) {
					// Node encoding is syncronous, so we can load all images here, then call export with a callback (if any)
					if ( NODEJS ) {
						try {
							var bitmap = fs.readFileSync(rel.path);
							rel.data = new Buffer(bitmap).toString('base64');
						}
						catch(ex) {
							console.error('ERROR: Unable to read media: '+rel.path);
							rel.data = IMG_BROKEN;
						}
					}
					else {
						intRels++;
						convertImgToDataURLviaCanvas(rel);
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
		var inMasterOpts = ( inMasterOpts && typeof inMasterOpts === 'object' ? inMasterOpts : {} );
		var slideObj = {};
		var slideNum = gObjPptx.slides.length;
		var slideObjNum = 0;
		var pageNum  = (slideNum + 1);

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
		// PUBLIC METHODS:
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
		 *   	{
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
		 * 	}
		 */
		slideObj.addChart = function ( inType, inData, inOpt ) {
			var intRels = 1;
			var options = ( inOpt && typeof inOpt === 'object' ? inOpt : {} );

			// STEP 1: TODO: check for reqd fields, correct type, etc
			// inType in CHART_TYPES
			// Array.isArray(inData)
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

			// STEP 2: Set vars for this Slide
			var slideObjNum = gObjPptx.slides[slideNum].data.length;
			var slideObjRels = gObjPptx.slides[slideNum].rels;

			// STEP 3: Set default options/decode user options
			// A: Core
			options.type = inType.name;
			options.x = (typeof options.x !== 'undefined' && options.x != null && !isNaN(options.x) ? options.x : 1);
			options.y = (typeof options.y !== 'undefined' && options.y != null && !isNaN(options.y) ? options.y : 1);
			options.w = (options.w || '50%');
			options.h = (options.h || '50%');

			// B: Options: misc
			if ( ['bar','col'].indexOf(options.barDir || '') < 0 ) options.barDir = 'col';
			// IMPORTANT: 'bestFit' will cause issues with PPT-Online in some cases, so defualt to 'ctr'!
			if ( ['bestFit','b','ctr','inBase','inEnd','l','outEnd','r','t'].indexOf(options.dataLabelPosition || '') < 0 ) options.dataLabelPosition = (options.type == 'pie' ? 'bestFit' : 'ctr');
			if ( ['b','l','r','t','tr'].indexOf(options.legendPos || '') < 0 ) options.legendPos = 'r';
			// barGrouping: "21.2.3.17 ST_Grouping (Grouping)"
			if ( ['clustered','standard','stacked','percentStacked'].indexOf(options.barGrouping || '') < 0 ) options.barGrouping = 'standard';
			if ( options.barGrouping.indexOf('tacked') > -1 ) {
				options.dataLabelPosition = 'ctr'; // IMPORTANT: PPT-Online will not open Presentation when 'outEnd' etc is used on stacked!
				if (!options.barGapWidthPct) options.barGapWidthPct = 50;
			}
			// lineDataSymbol: http://www.datypic.com/sc/ooxml/a-val-32.html
			// Spec has [plus,star,x] however neither PPT2013 nor PPT-Online support them
			if ( ['circle','dash','diamond','dot','none','square','triangle'].indexOf(options.lineDataSymbol || '') < 0 ) options.lineDataSymbol = 'circle';
			options.lineDataSymbolSize = ( options.lineDataSymbolSize && !isNaN(options.lineDataSymbolSize ) ? options.lineDataSymbolSize : 6 );

			// C: Options: plotArea
			options.showLabel   = (options.showLabel   == true || options.showLabel   == false ? options.showLabel   : false);
			options.showLegend  = (options.showLegend  == true || options.showLegend  == false ? options.showLegend  : false);
			options.showPercent = (options.showPercent == true || options.showPercent == false ? options.showPercent : true );
			options.showTitle   = (options.showTitle   == true || options.showTitle   == false ? options.showTitle   : false);
			options.showValue   = (options.showValue   == true || options.showValue   == false ? options.showValue   : false);

			// D: Options: chart
			options.barGapWidthPct = (!isNaN(options.barGapWidthPct) && options.barGapWidthPct >= 0 && options.barGapWidthPct <= 1000 ? options.barGapWidthPct : 150);
			options.chartColors = ( Array.isArray(options.chartColors) ? options.chartColors : (options.type == 'pie' ? PIECHART_COLORS : BARCHART_COLORS) );
			options.chartColorsOpacity = ( options.chartColorsOpacity && !isNaN(options.chartColorsOpacity) ? options.chartColorsOpacity : null );
			//
			options.border = ( options.border && typeof options.border === 'object' ? options.border : null );
			if ( options.border && (!options.border.pt || isNaN(options.border.pt)) ) options.border.pt = 1;
			if ( options.border && (!options.border.color || typeof options.border.color !== 'string' || options.border.color.length != 6) ) options.border.color = '363636';
			//
			options.dataBorder = ( options.dataBorder && typeof options.dataBorder === 'object' ? options.dataBorder : null );
			if ( options.dataBorder && (!options.dataBorder.pt || isNaN(options.dataBorder.pt)) ) options.dataBorder.pt = 0.75;
			if ( options.dataBorder && (!options.dataBorder.color || typeof options.dataBorder.color !== 'string' || options.dataBorder.color.length != 6) ) options.dataBorder.color = 'F9F9F9';
			//
			options.dataLabelFormatCode = ( options.dataLabelFormatCode && typeof options.dataLabelFormatCode === 'string' ? options.dataLabelFormatCode : (options.type == 'pie' ? '0%' : '#,##0') );
			//
			options.lineSize = ( typeof options.lineSize === 'number' ? options.lineSize : 2 );

			// STEP 4: Set props
			gObjPptx.slides[slideNum].data[slideObjNum] = {};
			gObjPptx.slides[slideNum].data[slideObjNum].type    = 'chart';
			gObjPptx.slides[slideNum].data[slideObjNum].options = options;

			// STEP 5: Add this chart to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
			// NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
			gObjPptx.slides.forEach(function(slide,idx){ intRels += slide.rels.length; });
			slideObjRels.push({
				rId:     (intRels+1),
				data:    inData,
				opts:    options,
				type:    'chart',
				fileName:'chart'+ intRels +'.xml',
				Target:  '/ppt/charts/chart'+ intRels +'.xml'
			});
			gObjPptx.slides[slideNum].data[slideObjNum].chartRid = slideObjRels[slideObjRels.length-1].rId;

			// LAST: Return
			return this;
		}

		// WARN: DEPRECATED: Will soon take a single {object} as argument (per current docs 20161120)
		// FUTURE: slideObj.addImage = function(opt){
		// NOTE: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
		slideObj.addImage = function( strImagePath, intPosX, intPosY, intSizeX, intSizeY, strImageData ) {
			var intRels = 1;

			// FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
			// FIXME: FUTURE: DEPRECATED: Only allow object param in 1.5 or 2.0
			if ( typeof strImagePath === 'object' ) {
				intPosX = (strImagePath.x || 0);
				intPosY = (strImagePath.y || 0);
				intSizeX = (strImagePath.cx || strImagePath.w || 0);
				intSizeY = (strImagePath.cy || strImagePath.h || 0);
				objHyperlink = (strImagePath.hyperlink || '');
				strImageData = (strImagePath.data || '');
				strImagePath = (strImagePath.path || ''); // IMPORTANT: This line must be last as were about to ovewrite ourself!
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
				Target: '../media/image'+ intRels +'.'+ strImgExtn
			});
			gObjPptx.slides[slideNum].data[slideObjNum].imageRid = slideObjRels[slideObjRels.length-1].rId;

			// STEP 5: (Issue#77) Hyperlink support
			if ( typeof objHyperlink === 'object' ) {
				if ( !objHyperlink.url || typeof objHyperlink.url !== 'string' ) console.log("ERROR: 'hyperlink.url is required and/or should be a string'");
				else {
					var intRelId = intRels+2;

					slideObjRels.push({
						type: 'hyperlink',
						data: 'dummy',
						rId:  intRelId,
						Target: objHyperlink.url
					});

					objHyperlink.rId = intRelId;
					gObjPptx.slides[slideNum].data[slideObjNum].hyperlink = objHyperlink;
				}
			}

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
			if ( !strPath && !strData && strType != 'online' ) {
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

			// STEP 3: Set media properties & options
			gObjPptx.slides[slideNum].data[slideObjNum].options    = {};
			gObjPptx.slides[slideNum].data[slideObjNum].options.x  = intPosX;
			gObjPptx.slides[slideNum].data[slideObjNum].options.y  = intPosY;
			gObjPptx.slides[slideNum].data[slideObjNum].options.cx = intSizeX;
			gObjPptx.slides[slideNum].data[slideObjNum].options.cy = intSizeY;

			// STEP 4: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
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
			var options = ( typeof opt === 'object' ? opt : {} );

			if ( !shape || typeof shape !== 'object' ) {
				console.log("ERROR: Missing/Invalid shape parameter! Example: `addShape(pptx.shapes.LINE, {x:1, y:1, w:1, h:1});` ");
				return;
			}

			// STEP 1: Grab Slide object count
			slideObjNum = gObjPptx.slides[slideNum].data.length;

			// STEP 2: Set props
			gObjPptx.slides[slideNum].data[slideObjNum] = {};
			gObjPptx.slides[slideNum].data[slideObjNum].type = 'text';
			gObjPptx.slides[slideNum].data[slideObjNum].options = options;

			// STEP 3: Set option defaults
			options.shape     = shape;
			options.x         = ( options.x || (options.x == 0 ? 0 : 1) );
			options.y         = ( options.y || (options.y == 0 ? 0 : 1) );
			options.w         = ( options.w || 1.0 );
			options.h         = ( options.h || (shape.name == 'line' ? 0 : 1.0) );
			options.line      = ( options.line || (shape.name == 'line' ? '333333' : null) );
			options.line_size = ( options.line_size || (shape.name == 'line' ? 1 : null) );
			if ( ['dash','dashDot','lgDash','lgDashDot','lgDashDotDot','solid','sysDash','sysDot'].indexOf(options.line_dash || '') < 0 ) options.line_dash = 'solid';

			// LAST: Return
			return this;
		};

		// RECURSIVE: (sometimes)
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

			// STEP 2: Row setup: Handle case where user passed in a simple 1-row array. EX: `["cell 1", "cell 2"]`
			//var arrRows = $.extend(true,[],arrTabRows);
			//if ( !Array.isArray(arrRows[0]) ) arrRows = [ $.extend(true,[],arrTabRows) ];
			var arrRows = arrTabRows;
			if ( !Array.isArray(arrRows[0]) ) arrRows = [ arrTabRows ];

			// STEP 3: Set options
			opt.x          = getSmartParseNumber( (opt.x || (EMU/2)), 'X' );
			opt.y          = getSmartParseNumber( (opt.y || EMU), 'Y' );
			opt.cy         = opt.h || opt.cy; // NOTE: Dont set default `cy` - leaving it null triggers auto-rowH in `makeXMLSlide()`
			if ( opt.cy ) opt.cy = getSmartParseNumber( opt.cy, 'Y' );
			opt.h          = opt.cy;
			opt.autoPage   = ( opt.autoPage == false ? false : true );
			opt.font_size  = opt.font_size || DEF_FONT_SIZE;
			opt.lineWeight = ( typeof opt.lineWeight !== 'undefined' && !isNaN(Number(opt.lineWeight)) ? Number(opt.lineWeight) : 0 );
			opt.margin     = (opt.marginPt || opt.margin); // (Legacy Support/DEPRECATED)
			opt.margin     = (opt.margin == 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_PT);
			if ( !isNaN(opt.margin) ) opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)]
			if ( opt.lineWeight > 1 ) opt.lineWeight = 1;
			else if ( opt.lineWeight < -1 ) opt.lineWeight = -1;
			// Set default color if needed (table option > inherit from Slide > default to black)
			if ( !opt.color ) opt.color = opt.color || this.color || '000000';

			// Set/Calc table width
			// Get slide margins - start with default values, then adjust if master or slide margins exist
			var arrTableMargin = DEF_SLIDE_MARGIN_IN;
			// Case 1: Master margins
			if ( inMaster && typeof inMaster.margin !== 'undefined' ) {
				if ( Array.isArray(inMaster.margin) ) arrTableMargin = inMaster.margin;
				else if ( !isNaN(Number(inMaster.margin)) ) arrTableMargin = [Number(inMaster.margin), Number(inMaster.margin), Number(inMaster.margin), Number(inMaster.margin)];
			}
			// Case 2: Table margins
			/* FIXME: add `margin` option to slide options
				else if ( slideObj.margin ) {
					if ( Array.isArray(slideObj.margin) ) arrTableMargin = slideObj.margin;
					else if ( !isNaN(Number(slideObj.margin)) ) arrTableMargin = [Number(slideObj.margin), Number(slideObj.margin), Number(slideObj.margin), Number(slideObj.margin)];
				}
			*/

			// Calc table width depending upon what data we have - several scenarios exist (including bad data, eg: colW doesnt match col count)
			if ( opt.w || opt.cx ) {
				opt.cx = getSmartParseNumber( (opt.w || opt.cx), 'X' );
				opt.w = opt.cx;
			}
			else if ( opt.colW ) {
				if ( typeof opt.colW === 'string' || typeof opt.colW === 'number' ) {
					opt.cx = Math.floor(Number(opt.colW) * arrRows[0].length);
					opt.w = opt.cx;
				}
				else if ( opt.colW && Array.isArray(opt.colW) && opt.colW.length != arrRows[0].length ) {
					console.warn('addTable: colW.length != data.length! Defaulting to evenly distributed col widths.');

					var numColWidth = Math.floor( ( (gObjPptx.pptLayout.width/EMU) - arrTableMargin[1] - arrTableMargin[3] ) / arrRows[0].length );
					opt.colW = [];
					for (var idx=0; idx<arrRows[0].length; idx++) { opt.colW.push( numColWidth ); }
					opt.cx = Math.floor(numColWidth * arrRows[0].length);
					opt.w = opt.cx;
				}
			}
			else {
				var numTabWidth = ( (gObjPptx.pptLayout.width/EMU) - arrTableMargin[1] - arrTableMargin[3] );
				opt.cx = Math.floor(numTabWidth);
				opt.w = opt.cx;
			}

			// STEP 4: Convert units to EMU now (we use different logic in makeSlide->table - smartCalc is not used)
			if ( opt.x            < 20 ) opt.x  = inch2Emu(opt.x);
			if ( opt.y            < 20 ) opt.y  = inch2Emu(opt.y);
			if ( opt.cx           < 20 ) opt.cx = inch2Emu(opt.cx);
			if ( opt.cy && opt.cy < 20 ) opt.cy = inch2Emu(opt.cy);

			// STEP 5: Check for fine-grained formatting, disable auto-page when found
			// Since genXmlTextBody already checks for text array ( text:[{},..{}] ) we're done!
			// Text in individual cells will be formatted as they are added by calls to genXmlTextBody within table builder
			arrRows.forEach(function(row,rIdx){
				row.forEach(function(cell,cIdx){
					if ( cell && cell.text && Array.isArray(cell.text) ) opt.autoPage = false;
				});
			});

			// STEP 6: Create hyperlink rels
			createHyperlinkRels(arrRows, gObjPptx.slides[slideNum].rels);

			// STEP 7: Auto-Paging: (via {options} and used internally)
			// (used internally by `addSlidesForTable()` to not engage recursion - we've already paged the table data, just add this one)
			if ( opt && opt.autoPage == false ) {
				// Add data (NOTE: Use `extend` to avoid mutation)
				gObjPptx.slides[slideNum].data[gObjPptx.slides[slideNum].data.length] = {
					type:       'table',
					arrTabRows: arrRows,
					options:    $.extend(true,{},opt)
				};
			}
			else {
				// Loop over rows and create 1-N tables as needed (ISSUE#21)
				getSlidesForTableRows(arrRows,opt).forEach(function(arrRows,idx){
					// A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
					var currSlide = ( !gObjPptx.slides[slideNum+idx] ? addNewSlide(inMaster, inMasterOpts) : gObjPptx.slides[slideNum+idx].slide );

					// B: Reset opt.y to `option`/`margin` after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
					if ( idx > 0 ) opt.y = inch2Emu( opt.newPageStartY || arrTableMargin[0] );

					// C: Add this table to new Slide
					opt.autoPage = false;
					currSlide.addTable(arrRows, $.extend(true,{},opt));
				});
			}

			// LAST: Return this Slide
			return this;
		};

		slideObj.addText = function( inText, options ) {
			var opt = ( options && typeof options === 'object' ? options : {} );
			var text = ( inText || '' );
			if ( Array.isArray(text) && text.length == 0 ) text = '';

			// STEP 1: Grab Slide object count
			slideObjNum = gObjPptx.slides[slideNum].data.length;

			// STEP 2: Set some options
			// Set color (options > inherit from Slide > default to black)
			opt.color = ( opt.color || this.color || '000000' );

			// ROBUST: Convert attr values that will likely be passed by users to valid OOXML values
			if ( opt.valign ) opt.valign = opt.valign.toLowerCase().replace(/^c.*/i,'ctr').replace(/^m.*/i,'ctr').replace(/^t.*/i,'t').replace(/^b.*/i,'b');
			if ( opt.align  ) opt.align  = opt.align.toLowerCase().replace(/^c.*/i,'center').replace(/^m.*/i,'center').replace(/^l.*/i,'left').replace(/^r.*/i,'right');

			// ROBUST: Set rational values for some shadow props if needed
			if ( opt.shadow ) {
				// OPT: `type`
				if ( opt.shadow.type != 'outer' && opt.shadow.type != 'inner' ) {
					console.warn('Warning: shadow.type options are `outer` or `inner`.');
					opt.type = 'outer';
				}

				// OPT: `angle`
				if ( opt.shadow.angle ) {
					// A: REALITY-CHECK
					if ( isNaN(Number(opt.shadow.angle)) || opt.shadow.angle < 0 || opt.shadow.angle > 359 ) {
						console.warn('Warning: shadow.angle can only be 0-359');
						opt.shadow.angle = 270;
					}

					// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
					opt.angle = Math.round(Number(opt.shadow.angle));
				}

				// OPT: `opacity`
				if ( opt.shadow.opacity ) {
					// A: REALITY-CHECK
					if ( isNaN(Number(opt.shadow.opacity)) || opt.shadow.opacity < 0 || opt.shadow.opacity > 1 ) {
						console.warn('Warning: shadow.opacity can only be 0-1');
						opt.shadow.opacity = 0.75;
					}

					// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
					opt.opacity = Number(opt.shadow.opacity)
				}
			}

			// STEP 3: Set props
			gObjPptx.slides[slideNum].data[slideObjNum] = {};
			gObjPptx.slides[slideNum].data[slideObjNum].type = 'text';
			gObjPptx.slides[slideNum].data[slideObjNum].text = text;

			// STEP 4: Set options
			gObjPptx.slides[slideNum].data[slideObjNum].options = opt;
			if ( opt.shape && opt.shape.name == 'line' ) {
				opt.line      = (opt.line      || '333333' );
				opt.line_size = (opt.line_size || 1 );
			}
			gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp = {};
			gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.autoFit = (opt.autoFit || false); // If true, shape will collapse to text size (Fit To Shape)
			gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.anchor = (opt.valign || 'ctr'); // VALS: [t,ctr,b]
			gObjPptx.slides[slideNum].data[slideObjNum].options.lineSpacing = (opt.lineSpacing && !isNaN(opt.lineSpacing) ? opt.lineSpacing : null);

			if ( (opt.inset && !isNaN(Number(opt.inset))) || opt.inset == 0 ) {
				gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.lIns = inch2Emu(opt.inset);
				gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.rIns = inch2Emu(opt.inset);
				gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.tIns = inch2Emu(opt.inset);
				gObjPptx.slides[slideNum].data[slideObjNum].options.bodyProp.bIns = inch2Emu(opt.inset);
			}

			// STEP 5: Create hyperlink rels
			createHyperlinkRels(text, gObjPptx.slides[slideNum].rels);

			// LAST: Return
			return this;
		};

		// ==========================================================================
		// POST-METHODS:
		// ==========================================================================

		// Add Master-Slide objects (if any)
		if ( inMaster && typeof inMaster === 'object' ) {
			// Add Slide Master objects in order
			$.each(inMaster, function(key,val){
				// ISSUE#7: Allow bkgd image/color override on Slide-level
				if ( key == 'bkgd' && inMasterOpts.bkgd ) val = inMasterOpts.bkgd;

				// Background color/image
				// DEPRECATED `src` is replaced by `path` in v1.5.0
				if ( key == 'bkgd' && typeof val === 'object' && (val.src || val.path || val.data) ) {
					// Allow the use of only the data key (no src reqd)
					val.src = val.src || val.path || null;
					if (!val.src) val.src = 'preencoded.png';
					var slideObjRels = gObjPptx.slides[slideNum].rels;
					var strImgExtn = val.src.substring( val.src.indexOf('.')+1 ).toLowerCase();
					if ( strImgExtn == 'jpg' ) strImgExtn = 'jpeg';
					if ( strImgExtn == 'gif' ) strImgExtn = 'png'; // MS-PPT: canvas.toDataURL for gif comes out image/png, and PPT will show "needs repair" unless we do this
					// FIXME: The next few lines are copies from .addImage above. A bad idea thats already bit me once! So of course it's makred as future :)
					var intRels = 1;
					for ( var idx=0; idx<gObjPptx.slides.length; idx++ ) { intRels += gObjPptx.slides[idx].rels.length; }
					slideObjRels.push({
						path: val.src,
						type: 'image/'+strImgExtn,
						extn: strImgExtn,
						data: (val.data || ''),
						rId: (intRels+1),
						Target: '../media/image' + intRels + '.' + strImgExtn
					});
					slideObj.bkgdImgRid = slideObjRels[slideObjRels.length-1].rId;
				}
				else if ( key == 'bkgd' && val && typeof val === 'string' ) {
					slideObj.back = val;
				}

				// Images **DEPRECATED** (v1.5.0)
				if ( key == "images" && Array.isArray(val) && val.length > 0 ) {
					$.each(val, function(i,image){
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

				// Shapes **DEPRECATED** (v1.5.0)
				if ( key == "shapes" && Array.isArray(val) && val.length > 0 ) {
					$.each(val, function(i,shape){
						// 1: Grab all options (x, y, color, etc.)
						var objOpts = {};
						$.each(Object.keys(shape), function(i,key){ if ( shape[key] != 'type' ) objOpts[key] = shape[key]; });
						// 2: Create object using 'type'
						if      ( shape.type == 'text'      ) slideObj.addText(shape.text, objOpts);
						else if ( shape.type == 'line'      ) slideObj.addShape(gObjPptxShapes.LINE, objOpts);
						else if ( shape.type == 'rectangle' ) slideObj.addShape(gObjPptxShapes.RECTANGLE, objOpts);
					});
				}

				// Add all Slide Master objects in the order they were given (Issue#53)
				if ( key == "objects" && Array.isArray(val) && val.length > 0 ) {
					val.forEach(function(object,idx){
						var key = Object.keys(object)[0];
						if      ( MASTER_OBJECTS[key] && key == 'chart' ) slideObj.addChart( CHART_TYPES[(object.chart.type||'').toUpperCase()], object.chart.data, object.chart.opts );
						else if ( MASTER_OBJECTS[key] && key == 'image' ) slideObj.addImage(object[key]);
						else if ( MASTER_OBJECTS[key] && key == 'line'  ) slideObj.addShape(gObjPptxShapes.LINE, object[key]);
						else if ( MASTER_OBJECTS[key] && key == 'rect'  ) slideObj.addShape(gObjPptxShapes.RECTANGLE, object[key]);
						else if ( MASTER_OBJECTS[key] && key == 'text'  ) slideObj.addText(object[key].text, object[key].options);
					});
				}
			});

			// Add Slide Numbers
			if ( typeof inMaster.isNumbered !== 'undefined' ) slideObj.hasSlideNumber(inMaster.isNumbered); // DEPRECATED
			if ( inMaster.slideNumber && typeof inMaster.slideNumber === 'object' ) slideObj.slideNumber(inMaster.slideNumber);
			else if ( inMasterOpts.slideNumber && typeof inMasterOpts.slideNumber === 'object' ) slideObj.slideNumber(inMasterOpts.slideNumber);
		}

		// LAST: Return this Slide
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

		var arrInchMargins = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
		opts.margin = (opts.marginPt || opts.margin); // (Legacy Support/DEPRECATED)
		opts.margin = (opts.margin || opts.margin == 0 ? opts.margin : 0.5);
		if ( opts && opts.master && opts.master.margin && gObjPptxMasters) {
			if ( Array.isArray(opts.master.margin) ) arrInchMargins = opts.master.margin;
			else if ( !isNaN(opts.master.margin) ) arrInchMargins = [opts.master.margin, opts.master.margin, opts.master.margin, opts.master.margin];
			opts.margin = arrInchMargins;
		}
		else if ( opts && opts.margin ) {
			if ( Array.isArray(opts.margin) ) arrInchMargins = opts.margin;
			else if ( !isNaN(opts.margin) ) arrInchMargins = [opts.margin, opts.margin, opts.margin, opts.margin];
		}
		var emuSlideTabW = ( opts.w ? inch2Emu(opts.w) : (gObjPptx.pptLayout.width  - inch2Emu(arrInchMargins[1] + arrInchMargins[3])) );
		var emuSlideTabH = ( opts.h ? inch2Emu(opts.h) : (gObjPptx.pptLayout.height - inch2Emu(arrInchMargins[0] + arrInchMargins[2])) );

		// STEP 1: Grab table col widths
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
				: arrColW.push( Number(((emuSlideTabW * (colW / intTabW * 100) ) / 100 / EMU).toFixed(2)) );
		});

		// STEP 3: Iterate over each table element and create data arrays (text and opts)
		// NOTE: We create 3 arrays instead of one so we can loop over body then show header/footer rows on first and last page
		$.each(['thead','tbody','tfoot'], function(i,val){
			$('#'+tabEleId+' > '+val+' > tr').each(function(i,row){
				var arrObjTabCells = [];
				$(row).find('> th, > td').each(function(i,cell){
					// A: Get RGB text/bkgd colors
					var arrRGB1 = [];
					var arrRGB2 = [];
					arrRGB1 = $(cell).css('color').replace(/\s+/gi,'').replace('rgba(','').replace('rgb(','').replace(')','').split(',');
					arrRGB2 = $(cell).css('background-color').replace(/\s+/gi,'').replace('rgba(','').replace('rgb(','').replace(')','').split(',');
					// ISSUE#57: jQuery default is this rgba value of below giving unstyled tables a black bkgd, so use white instead (FYI: if cell has `background:#000000` jQuery returns 'rgb(0, 0, 0)', so this soln is pretty solid)
					if ( $(cell).css('background-color') == 'rgba(0, 0, 0, 0)' || $(cell).css('background-color') == 'transparent' ) arrRGB2 = [255,255,255];

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
	// A: Load depdendencies
	var fs = require("fs");
	var $ = require("jquery-node");
	var JSZip = require("jszip");
	var sizeOf = require("image-size");

	// B: Export module
	module.exports = new PptxGenJS();
}
