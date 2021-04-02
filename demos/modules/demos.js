/**
 * NAME: demos.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.5.0
 * BLD.: 20210225
 */

import { COMPRESS, CUST_NAME, gPaths } from "../modules/enums.js";
import { starlabsLogoSml } from "../modules/media.js";
import { genSlides_Chart } from "./demo_chart.js";
import { genSlides_Image } from "./demo_image.js";
import { genSlides_Master } from "./demo_master.js";
import { genSlides_Media } from "./demo_media.js";
import { genSlides_Shape } from "./demo_shape.js";
import { genSlides_Table } from "./demo_table.js";
import { genSlides_Text } from "./demo_text.js";

// Detect Node.js (NODEJS is ultimately used to determine how to save: either `fs` or web-based, so using fs-detection is perfect)
let NODEJS = false;
// NOTE: `NODEJS` determines which network library to use, so using fs-detection is apropos.
if (typeof module !== "undefined" && module.exports && typeof require === "function" && typeof window === "undefined") {
	try {
		require.resolve("fs");
		NODEJS = true;
	} catch (ex) {
		NODEJS = false;
	}
}
//TODO: ? if (NODEJS) { var LOGO_STARLABS; }

// ==================================================================================================================

export function getTimestamp() {
	let dateNow = new Date();
	let dateMM = dateNow.getMonth() + 1;
	let dateDD = dateNow.getDate();
	let h = dateNow.getHours();
	let m = dateNow.getMinutes();
	return (
		dateNow.getFullYear() +
		"" +
		(dateMM <= 9 ? "0" + dateMM : dateMM) +
		"" +
		(dateDD <= 9 ? "0" + dateDD : dateDD) +
		(h <= 9 ? "0" + h : h) +
		(m <= 9 ? "0" + m : m)
	);
}

// ==================================================================================================================

export function runEveryTest() {
	return execGenSlidesFuncs(["Master", "Chart", "Image", "Media", "Shape", "Text", "Table"]);

	// NOTE: Html2Pptx needs table to be visible (otherwise col widths are even and look horrible)
	// ....: Therefore, run it mnaually. // if ( typeof table2slides1 !== 'undefined' ) table2slides1();
}

export function execGenSlidesFuncs(type) {
	// STEP 1: Instantiate new PptxGenJS object
	var pptx;
	if (NODEJS) {
		var PptxGenJsLib;
		var fs = require("fs");
		// TODO: we dont use local anymore as of 3.1
		if (fs.existsSync("../../dist/pptxgen.cjs.js")) {
			PptxGenJsLib = require("../../dist/pptxgen.cjs.js"); // for LOCAL TESTING
		} else {
			PptxGenJsLib = require("pptxgenjs");
		}
		pptx = new PptxGenJsLib();
		var base64Images = require("../common/images/base64Images.js");
		LOGO_STARLABS = base64Images.LOGO_STARLABS();
	} else {
		pptx = new PptxGenJS();
	}

	// STEP 2: Set Presentation props (as QA test only - these are not required)
	pptx.title = "PptxGenJS Test Suite Presentation";
	pptx.subject = "PptxGenJS Test Suite Export";
	pptx.author = "Brent Ely";
	pptx.company = CUST_NAME;
	pptx.revision = "15";

	// STEP 3: Set layout
	pptx.layout = "LAYOUT_WIDE";

	// STEP 4: Create Master Slides (from the old `pptxgen.masters.js` file - `gObjPptxMasters` items)
	{
		var objBkg = { path: NODEJS ? gPaths.starlabsBkgd.path.replace(/http.+\/examples/, "../common") : gPaths.starlabsBkgd.path };
		var objImg = {
			path: NODEJS ? gPaths.starlabsLogo.path.replace(/http.+\/examples/, "../common") : gPaths.starlabsLogo.path,
			x: 4.6,
			y: 3.5,
			w: 4,
			h: 1.8,
		};

		// TITLE_SLIDE
		pptx.defineSlideMaster({
			title: "TITLE_SLIDE",
			background: objBkg,
			//bkgd: objBkg, // TEST: @deprecated
			objects: [
				//{ 'line':  { x:3.5, y:1.0, w:6.0, h:0.0, line:{color:'0088CC'}, lineSize:5 } },
				//{ 'chart': { type:'PIE', data:[{labels:['R','G','B'], values:[10,10,5]}], options:{x:11.3, y:0.0, w:2, h:2, dataLabelFontSize:9} } },
				//{ 'image': { x:11.3, y:6.4, w:1.67, h:0.75, data:starlabsLogoSml } },
				{ rect: { x: 0.0, y: 5.7, w: "100%", h: 0.75, fill: { color: "F1F1F1" } } },
				{
					text: {
						text: "Global IT & Services :: Status Report",
						options: {
							x: 0.0,
							y: 5.7,
							w: "100%",
							h: 0.75,
							fontFace: "Arial",
							color: "363636",
							fontSize: 20,
							align: "center",
							valign: "middle",
							margin: 0,
						},
					},
				},
			],
		});

		// MASTER_PLAIN
		pptx.defineSlideMaster({
			title: "MASTER_PLAIN",
			background: { fill: "FFFFFF" },
			margin: [0.5, 0.25, 1.0, 0.25],
			objects: [
				{ rect: { x: 0.0, y: 6.9, w: "100%", h: 0.6, fill: { color: "003b75" } } },
				{ image: { x: 11.45, y: 5.95, w: 1.67, h: 0.75, data: starlabsLogoSml } },
				{
					text: {
						options: { x: 0, y: 6.9, w: "100%", h: 0.6, align: "center", valign: "middle", color: "FFFFFF", fontSize: 12 },
						text: "S.T.A.R. Laboratories - Confidential",
					},
				},
			],
			slideNumber: { x: 0.6, y: 7.1, color: "FFFFFF", fontFace: "Arial", fontSize: 10 },
		});

		// MASTER_SLIDE (MASTER_PLACEHOLDER)
		pptx.defineSlideMaster({
			title: "MASTER_SLIDE",
			background: { fill: "F1F1F1" },
			margin: [0.5, 0.25, 1.0, 0.25],
			slideNumber: { x: 0.6, y: 7.1, color: "FFFFFF", fontFace: "Arial", fontSize: 10 },
			objects: [
				{ rect: { x: 0.0, y: 6.9, w: "100%", h: 0.6, fill: { color: "003b75" } } },
				//{ 'image': { x:11.45, y:5.95, w:1.67, h:0.75, data:starlabsLogoSml } },
				{
					text: {
						options: { x: 0, y: 6.9, w: "100%", h: 0.6, align: "center", valign: "middle", color: "FFFFFF", fontSize: 12 },
						text: "S.T.A.R. Laboratories - Confidential",
					},
				},
				{
					placeholder: {
						options: { name: "title", type: "title", x: 0.6, y: 0.2, w: 12, h: 1.0 },
						text: "",
					},
				},
				{
					placeholder: {
						options: { name: "body", type: "body", x: 0.6, y: 1.5, w: 12, h: 5.25 },
						text: "(supports custom placeholder text!)",
					},
				},
			],
		});

		// THANKS_SLIDE (THANKS_PLACEHOLDER)
		pptx.defineSlideMaster({
			title: "THANKS_SLIDE",
			bkgd: "36ABFF", // BACKWARDS-COMPAT/DEPRECATED CHECK (`bkgd` will be removed in v4.x)
			objects: [
				{ rect: { x: 0.0, y: 3.4, w: "100%", h: 2.0, fill: { color: "FFFFFF" } } },
				{ image: objImg },
				{
					placeholder: {
						options: {
							name: "thanksText",
							type: "title",
							x: 0.0,
							y: 0.9,
							w: "100%",
							h: 1,
							fontFace: "Arial",
							color: "FFFFFF",
							fontSize: 60,
							align: "center",
						},
					},
				},
				{
					placeholder: {
						options: {
							name: "body",
							type: "body",
							x: 0.0,
							y: 6.45,
							w: "100%",
							h: 1,
							fontFace: "Courier",
							color: "FFFFFF",
							fontSize: 32,
							align: "center",
						},
						text: "(add homepage URL)",
					},
				},
			],
		});

		// PLACEHOLDER_SLIDE
		/* FUTURE: ISSUE#599
		pptx.defineSlideMaster({
		  title : 'PLACEHOLDER_SLIDE',
		  margin: [0.5, 0.25, 1.00, 0.25],
		  bkgd  : 'FFFFFF',
		  objects: [
			  { 'placeholder':
			  	{
					options: {type:'body'},
					image: {x:11.45, y:5.95, w:1.67, h:0.75, data:starlabsLogoSml}
				}
			},
			  { 'placeholder':
				  {
					  options: { name:'body', type:'body', x:0.6, y:1.5, w:12, h:5.25 },
					  text: '(supports custom placeholder text!)'
				  }
			  }
		  ],
		  slideNumber: { x:1.0, y:7.0, color:'FFFFFF' }
	  });*/

		// MISC: Only used for Issues, ad-hoc slides etc (for screencaps)
		pptx.defineSlideMaster({
			title: "DEMO_SLIDE",
			objects: [
				{ rect: { x: 0.0, y: 7.1, w: "100%", h: 0.4, fill: { color: "F1F1F1" } } },
				{
					text: {
						text: "PptxGenJS - JavaScript PowerPoint Library - (github.com/gitbrent/PptxGenJS)",
						options: { x: 0.0, y: 7.1, w: "100%", h: 0.4, color: "6c6c6c", fontSize: 10, align: "center" },
					},
				},
			],
		});
	}

	// STEP 5: Run requested test
	var arrTypes = typeof type === "string" ? [type] : type;
	arrTypes.forEach((type) => {
		//if (console.time) console.time(type);
		if (type === "Chart") genSlides_Chart(pptx);
		else if (type === "Image") genSlides_Image(pptx);
		else if (type === "Master") genSlides_Master(pptx);
		else if (type === "Media") genSlides_Media(pptx);
		else if (type === "Shape") genSlides_Shape(pptx);
		else if (type === "Table") genSlides_Table(pptx);
		else if (type === "Text") genSlides_Text(pptx);
		//if (console.timeEnd) console.timeEnd(type);
	});

	// LAST: Export Presentation
	if (NODEJS) {
		return pptx.writeFile({ fileName: "PptxGenJS_Demo_Node_" + type + "_" + getTimestamp() });
	} else {
		return pptx.writeFile({ fileName: "PptxGenJS_Demo_Browser_" + type + "_" + getTimestamp(), compression: COMPRESS });
	}
}
