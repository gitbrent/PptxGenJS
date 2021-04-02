/**
 * NAME: demos.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.5.0
 * BLD.: 20210225
 */

import { starlabsLogoSml } from "../modules/media.js";
import {
	COLOR_AMB,
	COLOR_GRN,
	COLOR_RED,
	COLOR_UNK,
	COMPRESS,
	CUST_NAME,
	TESTMODE,
	gOptsTabOpts,
	gOptsTextL,
	gOptsTextR,
	gPaths,
	gStrLoremEnglish,
} from "../modules/enums.js";
import { genSlides_Image } from "./demo_images.js";
import { genSlides_Media } from "./demo_media.js";
import { genSlides_Shape } from "./demo_shapes.js";
import { genSlides_Table } from "./demo_tables.js";
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
	arrTypes.forEach(function (type) {
		//if (console.time) console.time(type);
		if (type === "image") genSlides_Image(pptx);
		else if (type === "media") genSlides_Media(pptx);
		else if (type === "shape") genSlides_Shape(pptx);
		else if (type === "table") genSlides_Table(pptx);
		else if (type === "text") genSlides_Text(pptx);
		else eval("genSlides_" + type + "(pptx)");
		//if (console.timeEnd) console.timeEnd(type);
	});

	// LAST: Export Presentation
	if (NODEJS) {
		return pptx.writeFile({ fileName: "PptxGenJS_Demo_Node_" + type + "_" + getTimestamp() });
	} else {
		return pptx.writeFile({ fileName: "PptxGenJS_Demo_Browser_" + type + "_" + getTimestamp(), compression: COMPRESS });
	}
}

// ==================================================================================================================

function genSlides_Chart(pptx) {
	var LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
	var MONS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
	var QTRS = ["Q1", "Q2", "Q3", "Q4"];

	var dataChartPieStat = [
		{
			name: "Project Status",
			labels: ["Red", "Amber", "Green", "Complete", "Cancelled", "Unknown"],
			values: [25, 5, 5, 5, 5, 5],
		},
	];
	var dataChartPieLocs = [
		{
			name: "Location",
			labels: ["CN", "DE", "GB", "MX", "JP", "IN", "US"],
			values: [69, 35, 40, 85, 38, 99, 101],
		},
	];
	var arrDataLineStat = [];
	{
		var tmpObjRed = { name: "Red", labels: QTRS, values: [] };
		var tmpObjAmb = { name: "Amb", labels: QTRS, values: [] };
		var tmpObjGrn = { name: "Grn", labels: QTRS, values: [] };
		var tmpObjUnk = { name: "Unk", labels: QTRS, values: [] };

		for (var idy = 0; idy < QTRS.length; idy++) {
			tmpObjRed.values.push(Math.floor(Math.random() * 30) + 1);
			tmpObjAmb.values.push(Math.floor(Math.random() * 50) + 1);
			tmpObjGrn.values.push(Math.floor(Math.random() * 80) + 1);
			tmpObjUnk.values.push(Math.floor(Math.random() * 10) + 1);
		}

		arrDataLineStat.push(tmpObjRed);
		arrDataLineStat.push(tmpObjAmb);
		arrDataLineStat.push(tmpObjGrn);
		arrDataLineStat.push(tmpObjUnk);
	}
	// Create a gap for testing `displayBlanksAs` in line charts (2.3.0)
	arrDataLineStat[2].values = [55, null, null, 55];

	pptx.addSection({ title: "Charts" });

	// SLIDE 1: Bar Chart ------------------------------------------------------------------
	function slide1() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Bar Chart", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var arrDataRegions = [
			{
				name: "Region 1",
				labels: ["May", "June", "July", "August"],
				values: [26, 53, 100, 75],
			},
			{
				name: "Region 2",
				labels: ["May", "June", "July", "August"],
				values: [43.5, 70.3, 90.1, 80.05],
			},
		];
		var arrDataHighVals = [
			{
				name: "California",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [2000, 2800, 3200, 4000, 5000],
			},
			{
				name: "Texas",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [1400, 2000, 2500, 3000, 3800],
			},
		];

		// TOP-LEFT: H/bar
		var optsChartBar1 = {
			x: 0.5,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			barDir: "bar",
			border: { pt: "3", color: "00EE00" },
			fill: "F1F1F1",

			catAxisLabelColor: "CC0000",
			catAxisLabelFontFace: "Helvetica Neue",
			catAxisLabelFontSize: 14,
			catAxisOrientation: "maxMin",
			catAxisMajorTickMark: "in",
			catAxisMinorTickMark: "cross",

			// valAxisCrossesAt: 10,
			valAxisMajorTickMark: "cross",
			valAxisMinorTickMark: "out",

			titleColor: "33CF22",
			titleFontFace: "Helvetica Neue",
			titleFontSize: 24,
		};
		slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar1);

		// TOP-RIGHT: V/col
		var optsChartBar2 = {
			x: 7.0,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			barDir: "col",

			catAxisLabelColor: "0000CC",
			catAxisLabelFontFace: "Courier",
			catAxisLabelFontSize: 12,
			catAxisOrientation: "minMax",
			catAxisMajorTickMark: "none",
			catAxisMinorTickMark: "none",

			dataBorder: { pt: "1", color: "F1F1F1" },
			dataLabelColor: "696969",
			dataLabelFontFace: "Arial",
			dataLabelFontSize: 11,
			dataLabelPosition: "outEnd",
			dataLabelFormatCode: "#.0",
			showValue: true,

			valAxisOrientation: "maxMin",
			valAxisMajorTickMark: "none",
			valAxisMinorTickMark: "none",
			//valAxisLogScaleBase: '25',

			showLegend: false,
			showTitle: false,
		};
		slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar2);

		// BTM-LEFT: H/bar - TITLE and LEGEND
		slide.addText(".", { x: 0.5, y: 3.8, w: 6.0, h: 3.5, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		var optsChartBar3 = {
			x: 0.5,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			barDir: "bar",

			border: { pt: "3", color: "CF0909" },
			fill: "F1C1C1",

			catAxisLabelColor: "CC0000",
			catAxisLabelFontFace: "Helvetica Neue",
			catAxisLabelFontSize: 14,
			catAxisOrientation: "minMax",

			titleColor: "33CF22",
			titleFontFace: "Helvetica Neue",
			titleFontSize: 16,

			showTitle: true,
			title: "Sales by Region",
		};
		slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChartBar3);

		// BTM-RIGHT: V/col - TITLE and LEGEND
		slide.addText(".", { x: 7.0, y: 3.8, w: 6.0, h: 3.5, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		var optsChartBar4 = {
			x: 7.0,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			barDir: "col",
			barGapWidthPct: 25,
			chartColors: ["0088CC", "99FFCC"],
			chartColorsOpacity: 50,
			valAxisMaxVal: 5000,

			catAxisLabelColor: "0000CC",
			catAxisLabelFontFace: "Times",
			catAxisLabelFontSize: 11,
			catAxisOrientation: "minMax",

			dataBorder: { pt: "1", color: "F1F1F1" },
			dataLabelColor: "FFFFFF",
			dataLabelFontFace: "Arial",
			dataLabelFontSize: 10,
			dataLabelPosition: "ctr",
			showValue: true,

			showLegend: true,
			legendPos: "t",
			legendColor: "FF0000",
			showTitle: true,
			titleColor: "FF0000",
			title: "Red Title and Legend",
		};
		slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChartBar4);
	}

	// SLIDE 2: Bar Chart Grid/Axis Options ------------------------------------------------
	function slide2() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Bar Chart Grid/Axis Options", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var arrDataRegions = [
			{
				name: "Region 1",
				labels: ["May", "June", "July", "August"],
				values: [26, 53, 100, 75],
			},
			{
				name: "Region 2",
				labels: ["May", "June", "July", "August"],
				values: [43.5, 70.3, 90.1, 80.05],
			},
		];
		var arrDataHighVals = [
			{
				name: "California",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [2000, 2800, 3200, 4000, 5000],
			},
			{
				name: "Texas",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [1400, 2000, 2500, 3000, 3800],
			},
		];

		// TOP-LEFT: H/bar
		var optsChartBar1 = {
			x: 0.5,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			barDir: "bar",
			fill: "F1F1F1",

			catAxisLabelColor: "CC0000",
			catAxisLabelFontFace: "Helvetica Neue",
			catAxisLabelFontSize: 14,
			catGridLine: { style: "none" },
			catAxisHidden: true,

			valGridLine: { color: "cc6699", style: "dash", size: 1 },
			valAxisLineColor: "44AA66",
			valAxisLineSize: 1,
			valAxisLineStyle: "dash",

			showLegend: true,
			showTitle: true,
			title: "catAxisHidden:true, valGridLine/valAxisLine:dash",
			titleColor: "a9a9a9",
			titleFontFace: "Helvetica Neue",
			titleFontSize: 14,
		};
		slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar1);

		// TOP-RIGHT: V/col
		var optsChartBar2 = {
			x: 7.0,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			barDir: "col",
			fill: "E1F1FF",

			dataBorder: { pt: "1", color: "F1F1F1" },
			dataLabelColor: "696969",
			dataLabelFontFace: "Arial",
			dataLabelFontSize: 11,
			dataLabelPosition: "outEnd",
			dataLabelFormatCode: "#.0",
			showValue: true,

			catAxisHidden: true,
			catGridLine: { style: "none" },
			valAxisHidden: true,
			valAxisDisplayUnitLabel: true,
			valGridLine: { style: "none" },

			showLegend: true,
			legendPos: "b",
			showTitle: false,
		};
		slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar2);

		// BTM-LEFT: H/bar - TITLE and LEGEND
		slide.addText(".", { x: 0.5, y: 3.8, w: 6.0, h: 3.5, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		var optsChartBar3 = {
			x: 0.5,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			barDir: "bar",

			border: { pt: "3", color: "CF0909" },
			fill: "F1C1C1",

			catAxisLabelColor: "CC0000",
			catAxisLabelFontFace: "Helvetica Neue",
			catAxisLabelFontSize: 14,
			catAxisOrientation: "maxMin",
			catAxisTitle: "Housing Type",
			catAxisTitleColor: "428442",
			catAxisTitleFontSize: 14,
			showCatAxisTitle: true,

			valAxisOrientation: "maxMin",
			valGridLine: { style: "none" },
			valAxisHidden: true,
			valAxisDisplayUnitLabel: true,
			catGridLine: { color: "cc6699", style: "dash", size: 1 },

			titleColor: "33CF22",
			titleFontFace: "Helvetica Neue",
			titleFontSize: 16,

			showTitle: true,
			title: "Sales by Region",
		};
		slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChartBar3);

		// BTM-RIGHT: V/col - TITLE and LEGEND
		slide.addText(".", { x: 7.0, y: 3.8, w: 6.0, h: 3.5, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		var optsChartBar4 = {
			x: 7.0,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			barDir: "col",
			barGapWidthPct: 25,
			chartColors: ["0088CC", "99FFCC"],
			chartColorsOpacity: 50,
			valAxisMinVal: 1000,
			valAxisMaxVal: 5000,

			catAxisLabelColor: "0000CC",
			catAxisLabelFontFace: "Times",
			catAxisLabelFontSize: 11,
			catAxisLabelFrequency: 1,
			catAxisOrientation: "minMax",

			dataBorder: { pt: "1", color: "F1F1F1" },
			dataLabelColor: "FFFFFF",
			dataLabelFontFace: "Arial",
			dataLabelFontSize: 10,
			dataLabelPosition: "ctr",
			showValue: true,

			valAxisHidden: true,
			catAxisTitle: "Housing Type",
			showCatAxisTitle: true,

			showLegend: false,
			showTitle: false,
		};
		slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChartBar4);
	}

	// SLIDE 3: Stacked Bar Chart ----------------------------------------------------------
	function slide3() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable(
			[[{ text: "Chart Examples: Bar Chart: Stacked/PercentStacked and Data Table", options: gOptsTextL }, gOptsTextR]],
			gOptsTabOpts
		);

		var arrDataRegions = [
			{
				name: "Region 3",
				labels: ["April", "May", "June", "July", "August"],
				values: [17, 26, 53, 100, 75],
			},
			{
				name: "Region 4",
				labels: ["April", "May", "June", "July", "August"],
				values: [55, 43, 70, 90, 80],
			},
		];
		var arrDataHighVals = [
			{
				name: "California",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [2000, 2800, 3200, 4000, 5000],
			},
			{
				name: "Texas",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [1400, 2000, 2500, 3000, 3800],
			},
		];

		// TOP-LEFT: H/bar
		var optsChartBar1 = {
			x: 0.5,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			barDir: "bar",
			barGrouping: "stacked",

			catAxisOrientation: "maxMin",
			catAxisLabelColor: "CC0000",
			catAxisLabelFontFace: "Helvetica Neue",
			catAxisLabelFontSize: 14,
			catAxisLabelFontBold: true,
			valAxisLabelFontBold: true,

			dataLabelColor: "FFFFFF",
			showValue: true,

			titleColor: "33CF22",
			titleFontFace: "Helvetica Neue",
			titleFontSize: 24,
		};
		slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar1);

		// TOP-RIGHT: V/col
		var optsChartBar2 = {
			x: 7.0,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			barDir: "col",
			barGrouping: "stacked",

			dataLabelColor: "FFFFFF",
			dataLabelFontFace: "Arial",
			dataLabelFontSize: 12,
			dataLabelFontBold: true,
			showValue: true,

			catAxisLabelColor: "0000CC",
			catAxisLabelFontFace: "Courier",
			catAxisLabelFontSize: 12,
			catAxisOrientation: "minMax",

			showLegend: false,
			showTitle: false,
		};
		slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar2);

		// BTM-LEFT: H/bar - 100% layout without axis labels
		var optsChartBar3 = {
			x: 0.5,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			barDir: "bar",
			barGrouping: "percentStacked",
			dataBorder: { pt: "1", color: "F1F1F1" },
			catAxisHidden: true,
			valAxisHidden: true,
			showTitle: false,
			layout: { x: 0.1, y: 0.1, w: 1, h: 1 },
			showDataTable: true,
			showDataTableKeys: true,
			showDataTableHorzBorder: false,
			showDataTableVertBorder: false,
			showDataTableOutline: false,
			dataTableFontSize: 10,
		};
		slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar3);

		// BTM-RIGHT: V/col - TITLE and LEGEND
		slide.addText(".", { x: 7.0, y: 3.8, w: 6.0, h: 3.5, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		var optsChartBar4 = {
			x: 7.0,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			barDir: "col",
			barGrouping: "percentStacked",

			catAxisLabelColor: "0000CC",
			catAxisLabelFontFace: "Times",
			catAxisLabelFontSize: 12,
			catAxisOrientation: "minMax",
			chartColors: ["5DA5DA", "FAA43A"],
			showLegend: true,
			legendPos: "t",
			showDataTable: true,
			showDataTableKeys: false,
			dataTableFormatCode: "$#",
			//dataTableFormatCode: '0.00%' // @since v3.3.0
			//dataTableFormatCode: '$0.00' // @since v3.3.0
		};
		slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChartBar4);
	}

	// SLIDE 4: Bar Chart - Lots of Bars ---------------------------------------------------
	function slide4() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Lots of Bars (>26 letters)", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var arrDataHighVals = [
			{
				name: "Single Data Set",
				labels: LETTERS.concat(["AA", "AB", "AC", "AD"]),
				values: [-5, -3, 0, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30],
			},
		];

		var optsChart = {
			x: 0.5,
			y: 0.5,
			w: "90%",
			h: "90%",
			barDir: "col",
			title: "Chart With >26 Cols",
			showTitle: true,
			titleFontSize: 20,
			titleRotate: 10,
			showCatAxisTitle: true,
			catAxisTitle: "Letters",
			catAxisTitleColor: "4286f4",
			catAxisTitleFontSize: 14,

			showLegend: true,
			chartColors: ["154384"],
			invertedColors: ["0088CC"],

			showValAxisTitle: true,
			valAxisTitle: "Column Index",
			valAxisTitleColor: "c11c13",
			valAxisTitleFontSize: 16,
		};

		// TEST `getExcelColName()` to ensure Excel Column names are generated correctly above >26 chars/cols
		slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChart);
	}

	// SLIDE 5: Bar Chart: Data Series Colors, majorUnits, and valAxisLabelFormatCode ------
	function slide5() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable(
			[
				[
					{
						text:
							"Chart Examples: Multi-Color Bars, `catLabelFormatCode`, `valAxisDisplayUnit`, `valAxisMajorUnit`, `valAxisLabelFormatCode`",
						options: gOptsTextL,
					},
					gOptsTextR,
				],
			],
			gOptsTabOpts
		);

		// TOP-LEFT
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name: "Excel Date Values",
					labels: [37987, 38018, 38047, 38078, 38108, 38139],
					values: [20, 30, 10, 25, 15, 5],
				},
			],
			{
				x: 0.5,
				y: 0.6,
				w: "45%",
				h: 3,
				barDir: "bar",
				chartColors: ["0077BF", "4E9D2D", "ECAA00", "5FC4E3", "DE4216", "154384"],
				catLabelFormatCode: "yyyy-mm",
				valAxisMajorUnit: 15,
				valAxisDisplayUnit: "hundreds",
				valAxisMaxVal: 45,
				valLabelFormatCode: "$0", // @since v3.3.0
				showTitle: true,
				titleFontSize: 14,
				titleColor: "0088CC",
				title: "Bar Charts Can Be Multi-Color",
			}
		);

		// TOP-RIGHT
		// NOTE: Labels are ppt/excel dates (days past 1900)
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name: "Too Many Colors Series",
					labels: [37987, 38018, 38047, 38078, 38108, 38139],
					values: [0.2, 0.3, 0.1, 0.25, 0.15, 0.05],
				},
			],
			{
				x: 7,
				y: 0.6,
				w: "45%",
				h: 3,
				valAxisMaxVal: 1,
				barDir: "bar",
				catAxisLineShow: false,
				valAxisLineShow: false,
				showValue: true,
				catLabelFormatCode: "mmm-yy",
				dataLabelPosition: "outEnd",
				dataLabelFormatCode: "#%",
				valAxisLabelFormatCode: "#%",
				valAxisMajorUnit: 0.2,
				chartColors: ["0077BF", "4E9D2D", "ECAA00", "5FC4E3", "DE4216", "154384", "7D666A", "A3C961", "EF907B", "9BA0A3"],
				barGapWidthPct: 25,
			}
		);

		// BOTTOM-LEFT
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name: "Two Color Series",
					labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
					values: [0.2, -0.3, -0.1, 0.25, 0.15, 0.05],
				},
			],
			{
				x: 0.5,
				y: 4.0,
				w: "45%",
				h: 3,
				barDir: "col", // `col`(vert) | `bar`(horiz)
				showValue: true,
				dataLabelPosition: "outEnd",
				dataLabelFormatCode: "#%",
				valAxisLabelFormatCode: "0.#0",
				chartColors: ["0077BF", "4E9D2D", "ECAA00", "5FC4E3", "DE4216", "154384", "7D666A", "A3C961", "EF907B", "9BA0A3"],
				valAxisMaxVal: 0.4,
				barGapWidthPct: 50,
				showLegend: true,
				legendPos: "r",
			}
		);

		// BOTTOM-RIGHT
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name: "Escaped XML Chars",
					labels: ["Es", "cap", "ed", "XML", "Chars", "'", '"', "&", "<", ">"],
					values: [1.2, 2.3, 3.1, 4.25, 2.15, 6.05, 8.01, 2.02, 9.9, 0.9],
				},
			],
			{
				x: 7,
				y: 4,
				w: "45%",
				h: 3,
				barDir: "bar",
				showValue: true,
				dataLabelPosition: "outEnd",
				chartColors: ["0077BF", "4E9D2D", "ECAA00", "5FC4E3", "DE4216", "154384", "7D666A", "A3C961", "EF907B", "9BA0A3"],
				barGapWidthPct: 25,
				catAxisOrientation: "maxMin",
				valAxisOrientation: "maxMin",
				valAxisMaxVal: 10,
				valAxisMajorUnit: 1,
			}
		);
	}

	// SLIDE 6: 3D Bar Chart ---------------------------------------------------------------
	function slide6() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: 3D Bar Chart", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var arrDataRegions = [
			{
				name: "Region 1",
				labels: ["May", "June", "July", "August"],
				values: [26, 53, 100, 75],
			},
			{
				name: "Region 2",
				labels: ["May", "June", "July", "August"],
				values: [43.5, 70.3, 90.1, 80.05],
			},
		];
		var arrDataHighVals = [
			{
				name: "California",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [2000, 2800, 3200, 4000, 5000],
			},
			{
				name: "Texas",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [1400, 2000, 2500, 3000, 3800],
			},
		];

		// TOP-LEFT: H/bar
		var optsChartBar1 = {
			x: 0.5,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			barDir: "bar",
			fill: "F1F1F1",

			catAxisLabelColor: "CC0000",
			catAxisLabelFontFace: "Arial",
			catAxisLabelFontSize: 10,
			catAxisOrientation: "maxMin",

			serAxisLabelColor: "00EE00",
			serAxisLabelFontFace: "Arial",
			serAxisLabelFontSize: 10,
		};
		slide.addChart(pptx.charts.BAR3D, arrDataRegions, optsChartBar1);

		// TOP-RIGHT: V/col
		var optsChartBar2 = {
			x: 7.0,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			barDir: "col",
			bar3DShape: "cylinder",
			catAxisLabelColor: "0000CC",
			catAxisLabelFontFace: "Courier",
			catAxisLabelFontSize: 12,

			dataLabelColor: "000000",
			dataLabelFontFace: "Arial",
			dataLabelFontSize: 11,
			dataLabelPosition: "outEnd",
			dataLabelFormatCode: "#.0",
			dataLabelBkgrdColors: true,
			showValue: true,
		};
		slide.addChart(pptx.charts.BAR3D, arrDataRegions, optsChartBar2);

		// BTM-LEFT: H/bar - TITLE and LEGEND
		slide.addText(".", { x: 0.5, y: 3.8, w: 6.0, h: 3.5, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		var optsChartBar3 = {
			x: 0.5,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			barDir: "col",
			bar3DShape: "pyramid",
			barGrouping: "stacked",

			catAxisLabelColor: "CC0000",
			catAxisLabelFontFace: "Arial",
			catAxisLabelFontSize: 10,

			showValue: true,
			dataLabelBkgrdColors: true,

			showTitle: true,
			title: "Sales by Region",
		};
		slide.addChart(pptx.charts.BAR3D, arrDataHighVals, optsChartBar3);

		// BTM-RIGHT: V/col - TITLE and LEGEND
		slide.addText(".", { x: 7.0, y: 3.8, w: 6.0, h: 3.5, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		var optsChartBar4 = {
			x: 7.0,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			barDir: "col",
			bar3DShape: "coneToMax",
			chartColors: ["0088CC", "99FFCC"],

			catAxisLabelColor: "0000CC",
			catAxisLabelFontFace: "Times",
			catAxisLabelFontSize: 11,
			catAxisOrientation: "minMax",

			dataBorder: { pt: "1", color: "F1F1F1" },
			dataLabelColor: "000000",
			dataLabelFontFace: "Arial",
			dataLabelFontSize: 10,
			dataLabelPosition: "ctr",

			showLegend: true,
			legendPos: "t",
			legendColor: "FF0000",
			showTitle: true,
			titleColor: "FF0000",
			title: "Red Title and Legend",
		};
		slide.addChart(pptx.charts.BAR3D, arrDataHighVals, optsChartBar4);
	}

	// SLIDE 7: Tornado Chart --------------------------------------------------------------
	function slide7() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Tornado Chart - Grid and Axis Formatting", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name: "High",
					labels: ["London", "Munich", "Tokyo"],
					values: [0.2, 0.32, 0.41],
				},
				{
					name: "Low",
					labels: ["London", "Munich", "Tokyo"],
					values: [-0.11, -0.22, -0.29],
				},
			],
			{
				x: 0.5,
				y: 0.5,
				w: "90%",
				h: "90%",
				valAxisMaxVal: 1,
				barDir: "bar",
				axisLabelFormatCode: "#%",
				catGridLine: { color: "D8D8D8", style: "dash", size: 1 },
				valGridLine: { color: "D8D8D8", style: "dash", size: 1 },
				catAxisLineShow: false,
				valAxisLineShow: false,
				barGrouping: "stacked",
				catAxisLabelPos: "low",
				valueBarColors: true,
				shadow: { type: "none" },
				chartColors: ["0077BF", "4E9D2D", "ECAA00", "5FC4E3", "DE4216", "154384", "7D666A", "A3C961", "EF907B", "9BA0A3"],
				invertedColors: ["0065A2", "428526", "C99100", "51A7C1", "BD3813", "123970", "6A575A", "8BAB52", "CB7A69", "84888B"],
				barGapWidthPct: 25,
				valAxisMajorUnit: 0.2,
			}
		);
	}

	// SLIDE 8: Line Chart: Line Smoothing, Line Size, Symbol Size -------------------------
	function slide8() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable(
			[[{ text: "Chart Examples: Line Smoothing, Line Size, Line Shadow, Symbol Size", options: gOptsTextL }, gOptsTextR]],
			gOptsTabOpts
		);

		slide.addText("..", { x: 0.5, y: 0.6, w: 6.0, h: 3.0, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		var optsChartLine1 = {
			x: 0.5,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			chartColors: [COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK],
			lineSize: 8,
			lineSmooth: true,
			showLegend: true,
			legendPos: "t",
			catAxisLabelPos: "high",
		};
		slide.addChart(pptx.charts.LINE, arrDataLineStat, optsChartLine1);

		var optsChartLine2 = {
			x: 7.0,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			chartColors: [COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK],
			lineSize: 16,
			lineSmooth: true,
			showLegend: true,
			legendPos: "r",
		};
		slide.addChart(pptx.charts.LINE, arrDataLineStat, optsChartLine2);

		var optsChartLine1 = {
			x: 0.5,
			y: 4.0,
			w: 6.0,
			h: 3.0,
			chartColors: [COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK],
			lineDataSymbolSize: 10,
			shadow: { type: "none" },
			//displayBlanksAs: 'gap', //uncomment only for test - looks broken otherwise!
			showLegend: true,
			legendPos: "l",
		};
		slide.addChart(pptx.charts.LINE, arrDataLineStat, optsChartLine1);

		// QA: DEMO: Test shadow option
		var shadowOpts = { type: "outer", color: "cd0011", blur: 3, offset: 12, angle: 75, opacity: 0.8 };
		var optsChartLine2 = {
			x: 7.0,
			y: 4.0,
			w: 6.0,
			h: 3.0,
			chartColors: [COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK],
			lineDataSymbolSize: 20,
			shadow: shadowOpts,
			showLegend: true,
			legendPos: "b",
		};
		slide.addChart(pptx.charts.LINE, arrDataLineStat, optsChartLine2);
	}

	// SLIDE 9: Line Chart: TEST: `lineDataSymbol` + `lineDataSymbolSize` ------------------
	function slide9() {
		var intWgap = 4.25;
		var opts_lineDataSymbol = ["circle", "dash", "diamond", "dot", "none", "square", "triangle"];
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Line Chart: lineDataSymbol option test", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		opts_lineDataSymbol.forEach(function (opt, idx) {
			slide.addChart(pptx.charts.LINE, arrDataLineStat, {
				x: idx < 3 ? idx * intWgap : idx < 6 ? (idx - 3) * intWgap : (idx - 6) * intWgap,
				y: idx < 3 ? 0.5 : idx < 6 ? 2.75 : 5,
				w: 4.25,
				h: 2.25,
				lineDataSymbol: opt,
				title: opt,
				showTitle: true,
				lineDataSymbolSize: idx == 5 ? 9 : idx == 6 ? 12 : null,
			});
		});
	}

	// SLIDE 10: Line Chart: Lots of Cats --------------------------------------------------
	function slide10() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Line Chart: Lots of Lines", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var MAXVAL = 20000;

		var arrDataTimeline = [];
		for (var idx = 0; idx < 15; idx++) {
			var tmpObj = {
				name: "Series" + idx,
				labels: MONS,
				values: [],
			};

			for (var idy = 0; idy < MONS.length; idy++) {
				tmpObj.values.push(Math.floor(Math.random() * MAXVAL) + 1);
			}

			arrDataTimeline.push(tmpObj);
		}

		// FULL SLIDE:
		var optsChartLine1 = {
			x: 0.5,
			y: 0.6,
			w: "95%",
			h: "85%",
			fill: "F2F9FC",

			valAxisMaxVal: MAXVAL,

			showLegend: true,
			legendPos: "r",
		};
		slide.addChart(pptx.charts.LINE, arrDataTimeline, optsChartLine1);
	}

	// SLIDE 11: Area Chart: Misc ----------------------------------------------------------
	function slide11() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Area Chart, Stacked Area Chart", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var arrDataAreaSm = [
			{
				name: "Small Samples",
				labels: ["Q1", "Q2", "Q3", "Q4"],
				values: [15, 46, 31, 85],
			},
		];
		var arrDataTimeline2ser = [
			{
				name: "Actual Sales",
				labels: MONS,
				values: [1500, 4600, 5156, 3167, 8510, 8009, 6006, 7855, 12102, 12789, 10123, 15121],
			},
			{
				name: "Proj Sales",
				labels: MONS,
				values: [1000, 2600, 3456, 4567, 5010, 6009, 7006, 8855, 9102, 10789, 11123, 12121],
			},
		];

		// TOP-LEFT
		var optsChartLine1 = {
			x: 0.5,
			y: 0.6,
			w: "45%",
			h: 3,
			catAxisLabelRotate: 45,
			fill: "D1E1F1",
			chartColors: ["0088CC"],
			chartColorsOpacity: 25,
			dataBorder: { pt: 2, color: "FFFFFF" },
			showValue: true,
		};
		slide.addChart(pptx.charts.AREA, arrDataAreaSm, optsChartLine1);

		// TOP-RIGHT (stacked area chart)
		var optsChartLine2 = {
			x: 7,
			y: 0.6,
			w: "45%",
			h: 3,
			chartColors: ["0088CC", "99FFCC"],
			chartColorsOpacity: 25,
			valAxisLabelRotate: 5,
			dataBorder: { pt: 2, color: "FFFFFF" },
			showValue: false,
			fill: "D1E1F1",
			barGrouping: "stacked",
		};
		slide.addChart(pptx.charts.AREA, arrDataTimeline2ser, optsChartLine2);

		// BOTTOM-LEFT
		var optsChartLine3 = {
			x: 0.5,
			y: 4.0,
			w: "45%",
			h: 3,
			chartColors: ["0088CC", "99FFCC"],
			chartColorsOpacity: 50,
			valAxisLabelFormatCode: "#,K",
		};
		slide.addChart(pptx.charts.AREA, arrDataTimeline2ser, optsChartLine3);

		// BOTTOM-RIGHT
		var optsChartLine4 = { x: 7, y: 4.0, w: "45%", h: 3, chartColors: ["CC8833", "CCFF69"], chartColorsOpacity: 75 };
		slide.addChart(pptx.charts.AREA, arrDataTimeline2ser, optsChartLine4);
	}

	// SLIDE 12: Pie Charts: All 4 Legend Options ------------------------------------------
	function slide12() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Pie Charts: Legends", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		// [TEST][INTERNAL USE]: Not visible to user (its behind a chart): Used for ensuring ref counting works across obj types (eg: `rId` check/test)
		if (TESTMODE)
			slide.addImage({
				path: NODEJS ? gPaths.ccCopyRemix.path.replace(/http.+\/examples/, "../common") : gPaths.ccCopyRemix.path,
				x: 0.5,
				y: 1.0,
				w: 1.2,
				h: 1.2,
			});

		// TOP-LEFT
		slide.addText(".", { x: 0.5, y: 0.5, w: 4.2, h: 3.2, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		slide.addChart(pptx.charts.PIE, dataChartPieStat, {
			x: 0.5,
			y: 0.5,
			w: 4.2,
			h: 3.2,
			legendPos: "left",
			legendFontFace: "Courier New",
			showLegend: true,
			showLeaderLines: true,
			showPercent: false,
			showValue: true,
			chartColors: ["FC0000", "FFCC00", "009900", "0088CC", "696969", "6600CC"],
			dataBorder: { pt: "2", color: "F1F1F1" },
			dataLabelColor: "FFFFFF",
			dataLabelFontSize: 14,
			dataLabelPosition: "bestFit", // 'bestFit' | 'outEnd' | 'inEnd' | 'ctr'
		});

		// TOP-MIDDLE
		slide.addText(".", { x: 5.6, y: 0.5, w: 3.2, h: 3.2, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		slide.addChart(pptx.charts.PIE, dataChartPieStat, { x: 5.6, y: 0.5, w: 3.2, h: 3.2, showLegend: true, legendPos: "t" });

		// BTM-LEFT
		slide.addText(".", { x: 0.5, y: 4.0, w: 4.2, h: 3.2, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		slide.addChart(pptx.charts.PIE, dataChartPieLocs, { x: 0.5, y: 4.0, w: 4.2, h: 3.2, showLegend: true, legendPos: "r" });

		// BTM-MIDDLE
		slide.addText(".", { x: 5.6, y: 4.0, w: 3.2, h: 3.2, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		slide.addChart(pptx.charts.PIE, dataChartPieLocs, { x: 5.6, y: 4.0, w: 3.2, h: 3.2, showLegend: true, legendPos: "b" });

		// BOTH: TOP-RIGHT
		// DEMO: `legendFontSize`, `titleAlign`, `titlePos`
		slide.addText(".", { x: 9.8, y: 0.5, w: 3.2, h: 3.2, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		slide.addChart(pptx.charts.PIE, dataChartPieLocs, {
			x: 9.8,
			y: 0.5,
			w: 3.2,
			h: 3.2,
			dataBorder: { pt: "1", color: "F1F1F1" },
			showLegend: true,
			legendPos: "t",
			legendFontSize: 14,
			showLeaderLines: true,
			showTitle: true,
			title: "Right Title & Large Legend",
			titleAlign: "right",
			titlePos: { x: 0, y: 0 },
		});

		// BOTH: BTM-RIGHT
		slide.addText(".", { x: 9.8, y: 4.0, w: 3.2, h: 3.2, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		slide.addChart(pptx.charts.PIE, dataChartPieLocs, {
			x: 9.8,
			y: 4.0,
			w: 3.2,
			h: 3.2,
			dataBorder: { pt: "1", color: "F1F1F1" },
			showLegend: true,
			legendPos: "b",
			showTitle: true,
			title: "Title & Legend",
			firstSliceAng: 90,
		});
	}

	// SLIDE 13: Doughnut Chart ------------------------------------------------------------
	function slide13() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Doughnut Chart", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var optsChartPie1 = {
			x: 0.5,
			y: 1.0,
			w: 6.0,
			h: 6.0,
			chartColors: ["FC0000", "FFCC00", "009900", "0088CC", "696969", "6600CC"],
			dataBorder: { pt: "2", color: "F1F1F1" },
			dataLabelColor: "FFFFFF",
			dataLabelFontSize: 14,

			legendPos: "r",

			showLabel: false,
			showValue: false,
			showPercent: true,
			showLegend: true,
			showTitle: false,

			holeSize: 70,

			title: "Project Status",
			titleColor: "33CF22",
			titleFontFace: "Helvetica Neue",
			titleFontSize: 24,
		};
		slide.addText(".", { x: 0.5, y: 1.0, w: 6.0, h: 6.0, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		slide.addChart(pptx.charts.DOUGHNUT, dataChartPieStat, optsChartPie1);

		var optsChartPie2 = {
			x: 7.0,
			y: 1.0,
			w: 6,
			h: 6,
			dataBorder: { pt: "3", color: "F1F1F1" },
			dataLabelColor: "FFFFFF",
			showLabel: true,
			showValue: true,
			showPercent: true,
			showLegend: false,
			showTitle: false,
			title: "Resource Totals by Location",
			shadow: {
				type: "inner",
				offset: 20,
				blur: 20,
			},
		};
		slide.addChart(pptx.charts.DOUGHNUT, dataChartPieLocs, optsChartPie2);
	}

	// SLIDE 14: XY Scatter Chart ----------------------------------------------------------
	function slide14() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: XY Scatter Chart", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var arrDataScatter1 = [
			{ name: "X-Axis", values: [0, 1, 2, 3, 4, 5] },
			{ name: "Y-Value 1", values: [90, 80, 70, 85, 75, 92], labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"] },
			{ name: "Y-Value 2", values: [21, 32, 40, 49, 31, 29], labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"] },
		];
		var arrDataScatter2 = [
			{ name: "X-Axis", values: [1, 2, 3, 4, 5, 6] },
			{ name: "Airplane", values: [33, 20, 51, 65, 71, 75] },
			{ name: "Train", values: [99, 88, 77, 89, 99, 99] },
			{ name: "Bus", values: [21, 22, 25, 49, 59, 69] },
		];
		var arrDataScatterLabels = [
			{ name: "X-Axis", values: [1, 10, 20, 30, 40, 50] },
			{ name: "Y-Value 1", values: [11, 23, 31, 45, 47, 35], labels: ["Red 1", "Red 2", "Red 3", "Red 4", "Red 5", "Red 6"] },
			{ name: "Y-Value 2", values: [21, 38, 47, 59, 51, 25], labels: ["Blue 1", "Blue 2", "Blue 3", "Blue 4", "Blue 5", "Blue 6"] },
		];

		// TOP-LEFT
		var optsChartScat1 = {
			x: 0.5,
			y: 0.6,
			w: "45%",
			h: 3,
			valAxisTitle: "Renters",
			valAxisTitleColor: "428442",
			valAxisTitleFontSize: 14,
			showValAxisTitle: true,
			lineSize: 0,
			catAxisTitle: "Last 6 Months",
			catAxisTitleColor: "428442",
			catAxisTitleFontSize: 14,
			showCatAxisTitle: true,
			showLabel: true, // Must be set to true or labels will not be shown
			dataLabelPosition: "b", // Options: 't'|'b'|'l'|'r'|'ctr'
		};
		slide.addChart(pptx.charts.SCATTER, arrDataScatter1, optsChartScat1);

		// TOP-RIGHT
		var optsChartScat2 = {
			x: 7.0,
			y: 0.6,
			w: "45%",
			h: 3,
			fill: "f1f1f1",
			showLegend: true,
			legendPos: "b",

			lineSize: 8,
			lineSmooth: true,
			lineDataSymbolSize: 12,
			lineDataSymbolLineColor: "FFFFFF",

			chartColors: [COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK],
			chartColorsOpacity: 25,
		};
		slide.addChart(pptx.charts.SCATTER, arrDataScatter2, optsChartScat2);

		// BOTTOM-LEFT: (Labels)
		var optsChartScat3 = {
			x: 0.5,
			y: 4.0,
			w: "45%",
			h: 3,
			fill: "f2f9fc",
			//catAxisOrientation: 'maxMin',
			//valAxisOrientation: 'maxMin',
			showLegend: true,
			chartColors: ["FF0000", "0088CC"],

			showValAxisTitle: false,
			lineSize: 0,

			catAxisTitle: "Data Point Labels",
			catAxisTitleColor: "0088CC",
			catAxisTitleFontSize: 14,
			showCatAxisTitle: false,

			// Data Labels
			showLabel: true, // Must be set to true or labels will not be shown
			dataLabelPosition: "r", // Options: 't'|'b'|'l'|'r'|'ctr'
			dataLabelFormatScatter: "custom", // Can be set to `custom` (default), `customXY`, or `XY`.
		};
		slide.addChart(pptx.charts.SCATTER, arrDataScatterLabels, optsChartScat3);

		// BOTTOM-RIGHT
		var optsChartScat4 = { x: 7.0, y: 4.0, w: "45%", h: 3 };
		slide.addChart(pptx.charts.SCATTER, arrDataScatter2, optsChartScat4);
	}

	// SLIDE 15: Bubble Charts -------------------------------------------------------------
	function slide15() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Bubble Charts", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var arrDataBubble1 = [
			{ name: "X-Axis", values: [0.3, 0.6, 0.9, 1.2, 1.5, 1.7] },
			{ name: "Y-Value 1", values: [1.3, 9, 7.5, 2.5, 7.5, 5], sizes: [1, 4, 2, 3, 7, 4] },
			{ name: "Y-Value 2", values: [5, 3, 2, 7, 2, 10], sizes: [9, 7, 9, 2, 4, 8] },
		];
		var arrDataBubble2 = [
			{ name: "X-Axis", values: [1, 2, 3, 4, 5, 6] },
			{ name: "Airplane", values: [33, 20, 51, 65, 71, 75], sizes: [10, 10, 12, 12, 15, 20] },
			{ name: "Train", values: [99, 88, 77, 89, 99, 99], sizes: [20, 20, 22, 22, 25, 30] },
			{ name: "Bus", values: [21, 22, 25, 49, 59, 69], sizes: [11, 11, 13, 13, 16, 21] },
		];

		// TOP-LEFT
		var optsChartBubble1 = {
			x: 0.5,
			y: 0.6,
			w: "45%",
			h: 3,
			chartColors: ["4477CC", "ED7D31"],
			chartColorsOpacity: 40,
			dataBorder: { pt: 1, color: "FFFFFF" },
		};
		slide.addText(".", { x: 0.5, y: 0.6, w: 6.0, h: 3.0, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		slide.addChart(pptx.charts.BUBBLE, arrDataBubble1, optsChartBubble1);

		// TOP-RIGHT
		var optsChartBubble2 = {
			x: 7.0,
			y: 0.6,
			w: "45%",
			h: 3,
			fill: "f1f1f1",
			showLegend: true,
			legendPos: "b",

			lineSize: 8,
			lineSmooth: true,
			lineDataSymbolSize: 12,
			lineDataSymbolLineColor: "FFFFFF",

			chartColors: [COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK],
			chartColorsOpacity: 25,
		};
		slide.addChart(pptx.charts.BUBBLE, arrDataBubble2, optsChartBubble2);

		// BOTTOM-LEFT
		var optsChartBubble3 = {
			x: 0.5,
			y: 4.0,
			w: "45%",
			h: 3,
			fill: "f2f9fc",
			catAxisOrientation: "maxMin",
			valAxisOrientation: "maxMin",
			showCatAxisTitle: false,
			showValAxisTitle: false,
			valAxisMinVal: 0,
			dataBorder: { pt: 2, color: "FFFFFF" },
			dataLabelColor: "FFFFFF",
			showValue: true,
		};
		slide.addChart(pptx.charts.BUBBLE, arrDataBubble1, optsChartBubble3);

		// BOTTOM-RIGHT
		var optsChartBubble4 = { x: 7.0, y: 4.0, w: "45%", h: 3, lineSize: 0 };
		slide.addChart(pptx.charts.BUBBLE, arrDataBubble2, optsChartBubble4);
	}

	// SLIDE 15: Radar Chart ---------------------------------------------------------------
	function slide16() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Radar Chart", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var arrDataRegions = [
			{
				name: "Region 1",
				labels: ["May", "June", "July", "August", "September"],
				values: [26, 53, 100, 75, 41],
			},
		];
		var arrDataHighVals = [
			{
				name: "California",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [2000, 2800, 3200, 4000, 5000],
			},
			{
				name: "Texas",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [1400, 2000, 2500, 3000, 3800],
			},
		];

		// TOP-LEFT: Standard
		var optsChartRadar1 = { x: 0.5, y: 0.6, w: 6.0, h: 3.0, radarStyle: "standard", lineDataSymbol: "none", fill: "F1F1F1" };
		slide.addChart(pptx.charts.RADAR, arrDataRegions, optsChartRadar1);

		// TOP-RIGHT: Marker
		var optsChartRadar2 = {
			x: 7.0,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			radarStyle: "marker",
			catAxisLabelColor: "0000CC",
			catAxisLabelFontFace: "Courier",
			catAxisLabelFontSize: 12,
		};
		slide.addChart(pptx.charts.RADAR, arrDataRegions, optsChartRadar2);

		// BTM-LEFT: Filled - TITLE and LEGEND
		slide.addText(".", { x: 0.5, y: 3.8, w: 6.0, h: 3.5, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		var optsChartRadar3 = {
			x: 0.5,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			radarStyle: "filled",
			catAxisLabelColor: "CC0000",
			catAxisLabelFontFace: "Helvetica Neue",
			catAxisLabelFontSize: 14,

			showTitle: true,
			titleColor: "33CF22",
			titleFontFace: "Helvetica Neue",
			titleFontSize: 16,
			title: "Sales by Region",

			showLegend: true,
		};
		slide.addChart(pptx.charts.RADAR, arrDataHighVals, optsChartRadar3);

		// BTM-RIGHT: TITLE and LEGEND
		slide.addText(".", { x: 7.0, y: 3.8, w: 6.0, h: 3.5, fill: { color: "F1F1F1" }, color: "F1F1F1" });
		var optsChartRadar4 = {
			x: 7.0,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			radarStyle: "filled",
			chartColors: ["0088CC", "99FFCC"],

			catAxisLabelColor: "0000CC",
			catAxisLabelFontFace: "Times",
			catAxisLabelFontSize: 11,
			catAxisLineShow: false,

			showLegend: true,
			legendPos: "t",
			legendColor: "FF0000",
			showTitle: true,
			titleColor: "FF0000",
			title: "Red Title and Legend",
		};
		slide.addChart(pptx.charts.RADAR, arrDataHighVals, optsChartRadar4);
	}

	// SLIDE 16: Multi-Type Charts ---------------------------------------------------------
	function slide17() {
		// powerpoint 2016 add secondary category axis labels
		// https://peltiertech.com/chart-with-a-dual-category-axis/

		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Examples: Multi-Type Charts", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		function doStackedLine() {
			// TOP-RIGHT:
			var opts = {
				x: 7.0,
				y: 0.6,
				w: 6.0,
				h: 3.0,
				barDir: "col",
				barGrouping: "stacked",
				catAxisLabelColor: "0000CC",
				catAxisLabelFontFace: "Arial",
				catAxisLabelFontSize: 12,
				catAxisOrientation: "minMax",
				showLegend: false,
				showTitle: false,
				valAxisMaxVal: 100,
				valAxisMajorUnit: 10,
			};

			var labels = ["Mon", "Tue", "Wed", "Thu", "Fri"];
			var chartTypes = [
				{
					type: pptx.charts.BAR,
					data: [
						{
							name: "Bottom",
							labels: labels,
							values: [17, 26, 53, 10, 4],
						},
						{
							name: "Middle",
							labels: labels,
							values: [55, 40, 20, 30, 15],
						},
						{
							name: "Top",
							labels: labels,
							values: [10, 22, 25, 35, 70],
						},
					],
					options: {
						barGrouping: "stacked",
					},
				},
				{
					type: pptx.charts.LINE,
					data: [
						{
							name: "Current",
							labels: labels,
							values: [25, 35, 55, 10, 5],
						},
					],
					options: {
						barGrouping: "standard",
					},
				},
			];
			slide.addChart(chartTypes, opts);
		}

		function doColumnAreaLine() {
			var opts = {
				x: 0.6,
				y: 0.6,
				w: 6.0,
				h: 3.0,
				barDir: "col",
				catAxisLabelColor: "666666",
				catAxisLabelFontFace: "Arial",
				catAxisLabelFontSize: 12,
				catAxisOrientation: "minMax",
				showLegend: false,
				showTitle: false,
				valAxisMaxVal: 100,
				valAxisMajorUnit: 10,

				valAxes: [
					{
						showValAxisTitle: true,
						valAxisTitle: "Primary Value Axis",
					},
					{
						showValAxisTitle: true,
						valAxisTitle: "Secondary Value Axis",
						valGridLine: { style: "none" },
					},
				],

				catAxes: [
					{
						catAxisTitle: "Primary Category Axis",
					},
					{
						catAxisHidden: true,
					},
				],
			};

			var labels = ["April", "May", "June", "July", "August"];
			var chartTypes = [
				{
					type: pptx.charts.AREA,
					data: [
						{
							name: "Current",
							labels: labels,
							values: [1, 4, 7, 2, 3],
						},
					],
					options: {
						chartColors: ["00FFFF"],
						barGrouping: "standard",
						secondaryValAxis: !!opts.valAxes,
						secondaryCatAxis: !!opts.catAxes,
					},
				},
				{
					type: pptx.charts.BAR,
					data: [
						{
							name: "Bottom",
							labels: labels,
							values: [17, 26, 53, 10, 4],
						},
					],
					options: {
						chartColors: ["0000FF"],
						barGrouping: "stacked",
					},
				},
				{
					type: pptx.charts.LINE,
					data: [
						{
							name: "Current",
							labels: labels,
							values: [5, 3, 2, 4, 7],
						},
					],
					options: {
						barGrouping: "standard",
						secondaryValAxis: !!opts.valAxes,
						secondaryCatAxis: !!opts.catAxes,
					},
				},
			];
			slide.addChart(chartTypes, opts);
		}

		function doStackedDot() {
			// BOT-LEFT:
			var opts = {
				x: 0.6,
				y: 4.0,
				w: 6.0,
				h: 3.0,
				barDir: "col",
				barGrouping: "stacked",
				catAxisLabelColor: "999999",
				catAxisLabelFontFace: "Arial",
				catAxisLabelFontSize: 14,
				catAxisOrientation: "minMax",
				showLegend: false,
				showTitle: false,
				valAxisMaxVal: 100,
				valAxisMinVal: 0,
				valAxisMajorUnit: 20,

				lineSize: 0,
				lineDataSymbolSize: 20,
				lineDataSymbolLineSize: 2,
				lineDataSymbolLineColor: "FF0000",

				//dataNoEffects: true,

				valAxes: [
					{
						showValAxisTitle: true,
						valAxisTitle: "Primary Value Axis",
					},
					{
						showValAxisTitle: true,
						valAxisTitle: "Secondary Value Axis",
						catAxisOrientation: "maxMin",
						valAxisMajorUnit: 1,
						valAxisMaxVal: 10,
						valAxisMinVal: 1,
						valGridLine: { style: "none" },
					},
				],
				catAxes: [
					{
						catAxisTitle: "Primary Category Axis",
					},
					{
						catAxisHidden: true,
					},
				],
			};

			var labels = ["Q1", "Q2", "Q3", "Q4", "OT"];
			var chartTypes = [
				{
					type: pptx.charts.BAR,
					data: [
						{
							name: "Bottom",
							labels: labels,
							values: [17, 26, 53, 10, 4],
						},
						{
							name: "Middle",
							labels: labels,
							values: [55, 40, 20, 30, 15],
						},
						{
							name: "Top",
							labels: labels,
							values: [10, 22, 25, 35, 70],
						},
					],
					options: {
						barGrouping: "stacked",
					},
				},
				{
					type: pptx.charts.LINE,
					data: [
						{
							name: "Current",
							labels: labels,
							values: [5, 3, 2, 4, 7],
						},
					],
					options: {
						barGrouping: "standard",
						secondaryValAxis: !!opts.valAxes,
						secondaryCatAxis: !!opts.catAxes,
						chartColors: ["FFFF00"],
					},
				},
			];
			slide.addChart(chartTypes, opts);
		}

		function doBarCol() {
			// BOT-RGT:
			var opts = {
				x: 7,
				y: 4.0,
				w: 6.0,
				h: 3.0,
				barDir: "col",
				barGrouping: "stacked",
				catAxisLabelColor: "999999",
				catAxisLabelFontFace: "Arial",
				catAxisLabelFontSize: 14,
				catAxisOrientation: "minMax",
				showLegend: false,
				showTitle: false,
				valAxisMaxVal: 100,
				valAxisMinVal: 0,
				valAxisMajorUnit: 20,
				valAxes: [
					{
						showValAxisTitle: true,
						valAxisTitle: "Primary Value Axis",
					},
					{
						showValAxisTitle: true,
						valAxisTitle: "Secondary Value Axis",
						catAxisOrientation: "maxMin",
						valAxisMajorUnit: 1,
						valAxisMaxVal: 10,
						valAxisMinVal: 1,
						valGridLine: { style: "none" },
					},
				],
				catAxes: [
					{
						catAxisTitle: "Primary Category Axis",
					},
					{
						catAxisHidden: true,
					},
				],
			};

			var labels = ["Q1", "Q2", "Q3", "Q4", "OT"];
			var chartTypes = [
				{
					type: pptx.charts.BAR,
					data: [
						{
							name: "Bottom",
							labels: labels,
							values: [17, 26, 53, 10, 4],
						},
						{
							name: "Middle",
							labels: labels,
							values: [55, 40, 20, 30, 15],
						},
						{
							name: "Top",
							labels: labels,
							values: [10, 22, 25, 35, 70],
						},
					],
					options: {
						barGrouping: "stacked",
					},
				},
				{
					type: pptx.charts.BAR,
					data: [
						{
							name: "Current",
							labels: labels,
							values: [5, 3, 2, 4, 7],
						},
					],
					options: {
						barDir: "bar",
						barGrouping: "standard",
						secondaryValAxis: !!opts.valAxes,
						secondaryCatAxis: !!opts.catAxes,
					},
				},
			];
			slide.addChart(chartTypes, opts);
		}

		function readmeExample() {
			// for testing - not rendered in demo
			var labels = ["Q1", "Q2", "Q3", "Q4", "OT"];
			var chartTypes = [
				{
					type: pptx.charts.BAR,
					data: [
						{
							name: "Projected",
							labels: labels,
							values: [17, 26, 53, 10, 4],
						},
					],
					options: {
						barDir: "col",
					},
				},
				{
					type: pptx.charts.LINE,
					data: [
						{
							name: "Current",
							labels: labels,
							values: [5, 3, 2, 4, 7],
						},
					],
					options: {
						secondaryValAxis: true,
						secondaryCatAxis: true,
					},
				},
			];
			var multiOpts = {
				x: 1.0,
				y: 1.0,
				w: 6,
				h: 6,
				valAxisMaxVal: 100,
				valAxisMinVal: 0,
				valAxisMajorUnit: 20,
				valAxes: [
					{
						showValAxisTitle: true,
						valAxisTitle: "Primary Value Axis",
					},
					{
						showValAxisTitle: true,
						valAxisTitle: "Secondary Value Axis",
						valAxisMajorUnit: 1,
						valAxisMaxVal: 10,
						valAxisMinVal: 1,
						valGridLine: { style: "none" },
					},
				],
				catAxes: [
					{
						catAxisTitle: "Primary Category Axis",
					},
					{
						catAxisHidden: true,
					},
				],
			};

			slide.addChart(chartTypes, multiOpts);
		}

		doBarCol();
		doStackedDot();
		doColumnAreaLine();
		doStackedLine();
		//readmeExample();
	}

	// SLIDE 17: Charts Options: Shadow, Transparent Colors --------------------------------
	function slide18() {
		var slide = pptx.addSlide({ sectionTitle: "Charts" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
		slide.addTable([[{ text: "Chart Options: Shadow, Transparent Colors", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

		var arrDataRegions = [
			{
				name: "Region 2",
				labels: ["April", "May", "June", "July", "August"],
				values: [0, 30, 53, 10, 25],
			},
			{
				name: "Region 3",
				labels: ["April", "May", "June", "July", "August"],
				values: [17, 26, 53, 100, 75],
			},
			{
				name: "Region 4",
				labels: ["April", "May", "June", "July", "August"],
				values: [55, 43, 70, 90, 80],
			},
			{
				name: "Region 5",
				labels: ["April", "May", "June", "July", "August"],
				values: [55, 43, 70, 90, 80],
			},
		];
		var arrDataHighVals = [
			{
				name: "California",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [2000, 2800, 3200, 4000, 5000],
			},
			{
				name: "Texas",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [1400, 2000, 2500, 3000, 3800],
			},
		];
		var single = [
			{
				name: "Texas",
				labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
				values: [1400, 2000, 2500, 3000, 3800],
			},
		];

		// TOP-LEFT: H/bar
		var optsChartBar1 = {
			x: 0.5,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			showTitle: true,
			title: "Large blue shadow",
			barDir: "bar",
			barGrouping: "standard",
			dataLabelColor: "FFFFFF",
			showValue: true,
			shadow: {
				type: "outer",
				blur: 10,
				offset: 5,
				angle: 45,
				color: "0059B1",
				opacity: 1,
			},
		};

		var pieOptions = {
			x: 7.0,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			showTitle: true,
			title: "Rotated cyan shadow",
			dataLabelColor: "FFFFFF",
			shadow: {
				type: "outer",
				blur: 10,
				offset: 5,
				angle: 180,
				color: "00FFFF",
				opacity: 1,
			},
		};

		// BTM-LEFT: H/bar - 100% layout without axis labels
		var optsChartBar3 = {
			x: 0.5,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			showTitle: true,
			title: "No shadow, transparent colors",
			barDir: "bar",
			barGrouping: "stacked",
			chartColors: ["transparent", "5DA5DA", "transparent", "FAA43A"],
			shadow: { type: "none" },
		};

		// BTM-RIGHT: V/col - TITLE and LEGEND
		var optsChartBar4 = {
			x: 7.0,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			barDir: "col",
			barGrouping: "stacked",
			showTitle: true,
			title: "Red glowing shadow",
			titleBold: true,
			titleFontFace: "Times",
			catAxisLabelColor: "0000CC",
			catAxisLabelFontFace: "Times",
			catAxisLabelFontSize: 12,
			catAxisOrientation: "minMax",
			chartColors: ["5DA5DA", "FAA43A"],
			shadow: {
				type: "outer",
				blur: 20,
				offset: 1,
				angle: 90,
				color: "A70000",
				opacity: 1,
			},
		};

		slide.addChart(pptx.charts.BAR, single, optsChartBar1);
		slide.addChart(pptx.charts.PIE, dataChartPieStat, pieOptions);
		slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar3);
		slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChartBar4);
	}

	// RUN ALL SLIDE DEMOS -----
	slide1();
	slide2();
	slide3();
	slide4();
	slide5();
	slide6();
	slide7();
	slide8();
	slide9();
	slide10();
	slide11();
	slide12();
	slide13();
	slide14();
	slide15();
	slide16();
	slide17();
	slide18();
}

function genSlides_Master(pptx) {
	pptx.addSection({ title: "Masters" });

	var slide1 = pptx.addSlide({ masterName: "TITLE_SLIDE", sectionTitle: "Masters" });
	//var slide1 = pptx.addSlide({masterName:'TITLE_SLIDE', sectionTitle:'FAILTEST'}); // TEST: Should show console warning ("title not found")
	slide1.addNotes("Master name: `TITLE_SLIDE`\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");

	var slide2 = pptx.addSlide({ masterName: "MASTER_SLIDE", sectionTitle: "Masters" });
	slide2.addNotes("Master name: `MASTER_SLIDE`\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
	slide2.addText("", { placeholder: "title" });

	var slide3 = pptx.addSlide({ masterName: "MASTER_SLIDE", sectionTitle: "Masters" });
	slide3.addNotes("Master name: `MASTER_SLIDE` using pre-filled placeholders\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
	slide3.addText("Text Placeholder", { placeholder: "title" });
	slide3.addText(
		[
			{ text: "Pre-filled placeholder bullets", options: { bullet: true, valign: "top" } },
			{ text: "Add any text, charts, whatever", options: { bullet: true, indentLevel: 1, color: "0000AB" } },
			{ text: "Check out the online API docs for more", options: { bullet: true, indentLevel: 2, color: "0000AB" } },
		],
		{ placeholder: "body", valign: "top" }
	);

	var slide4 = pptx.addSlide({ masterName: "MASTER_SLIDE", sectionTitle: "Masters" });
	slide4.addNotes("Master name: `MASTER_SLIDE` using pre-filled placeholders\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
	slide4.addText("Image Placeholder", { placeholder: "title" });
	slide4.addImage({
		placeholder: "body",
		path: NODEJS ? gPaths.starlabsBkgd.path.replace(/http.+\/examples/, "../common") : gPaths.starlabsBkgd.path,
	});

	var dataChartPieLocs = [
		{
			name: "Location",
			labels: ["CN", "DE", "GB", "MX", "JP", "IN", "US"],
			values: [69, 35, 40, 85, 38, 99, 101],
		},
	];
	var slide5 = pptx.addSlide({ masterName: "MASTER_SLIDE", sectionTitle: "Masters" });
	slide5.addNotes("Master name: `MASTER_SLIDE` using pre-filled placeholders\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
	slide5.addText("Chart Placeholder", { placeholder: "title" });
	slide5.addChart(pptx.charts.PIE, dataChartPieLocs, { showLegend: true, legendPos: "l", placeholder: "body" });

	var slide6 = pptx.addSlide({ masterName: "THANKS_SLIDE", sectionTitle: "Masters" });
	slide6.addNotes("Master name: `THANKS_SLIDE`\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
	slide6.addText("Thank You!", { placeholder: "thanksText" });
	//slide6.addText('github.com/gitbrent', { placeholder:'body' });

	//var slide7 = pptx.addSlide('PLACEHOLDER_SLIDE');

	// LEGACY-TEST-ONLY: To check deprecated functionality
	/*
	if ( pptx.masters && Object.keys(pptx.masters).length > 0 ) {
		var slide1 = pptx.addSlide( pptx.masters.TITLE_SLIDE  );
		var slide2 = pptx.addSlide( pptx.masters.MASTER_SLIDE );
		var slide3 = pptx.addSlide( pptx.masters.THANKS_SLIDE );

		var slide4 = pptx.addSlide( pptx.masters.TITLE_SLIDE,  { bkgd:'0088CC', slideNumber:{x:'50%', y:'90%', color:'0088CC'} } );
		var slide5 = pptx.addSlide( pptx.masters.MASTER_SLIDE, { bkgd:{ path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/images/title_bkgd_alt.jpg' } } );
		var slide6 = pptx.addSlide( pptx.masters.THANKS_SLIDE, { bkgd:'ffab33' } );
		//var slide7 = pptx.addSlide( pptx.masters.LEGACY_TEST_ONLY );
	}
	*/
}
