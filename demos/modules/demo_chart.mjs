/**
 * NAME: demo_chart.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.6.0
 * BLD.: 20210426
 */

import { IMAGE_PATHS, BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R, COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK, TESTMODE } from "./enums.mjs";

const LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
const MONS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const QTRS = ["Q1", "Q2", "Q3", "Q4"];

const dataChartPieStat = [
	{
		name: "Project Status",
		labels: ["Red", "Amber", "Green", "Complete", "Cancelled", "Unknown"],
		values: [25, 5, 5, 5, 5, 5],
	},
];
const dataChartPieLocs = [
	{
		name: "Location",
		labels: ["CN", "DE", "GB", "MX", "JP", "IN", "US"],
		values: [69, 35, 40, 85, 38, 99, 101],
	},
];
let arrDataLineStat = [];

export function genSlides_Chart(pptx) {
	initTestData();

	pptx.addSection({ title: "Charts" });

	genSlide01(pptx);
	genSlide02(pptx);
	genSlide03(pptx);
	genSlide04(pptx);
	genSlide05(pptx);
	genSlide06(pptx);
	genSlide07(pptx);
	genSlide08(pptx);
	genSlide09(pptx);
	genSlide10(pptx);
	genSlide11(pptx);
	genSlide12(pptx);
	genSlide13(pptx);
	genSlide14(pptx);
	genSlide15(pptx);
	genSlide16(pptx);
	genSlide17(pptx);
	genSlide18(pptx);
}

function initTestData() {
	let tmpObjRed = { name: "Red", labels: QTRS, values: [] };
	let tmpObjAmb = { name: "Amb", labels: QTRS, values: [] };
	let tmpObjGrn = { name: "Grn", labels: QTRS, values: [] };
	let tmpObjUnk = { name: "Unk", labels: QTRS, values: [] };

	for (let idy = 0; idy < QTRS.length; idy++) {
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

// SLIDE 1: Bar Chart
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Bar Chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataRegions = [
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
	let arrDataHighVals = [
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
	let optsChartBar1 = {
		x: 0.5,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		altText: "this is the alt text content",

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
	let optsChartBar2 = {
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
	let optsChartBar3 = {
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
	let optsChartBar4 = {
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

// SLIDE 2: Bar Chart Grid/Axis Options
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Bar Chart Grid/Axis Options", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataRegions = [
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
	let arrDataHighVals = [
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
	let optsChartBar1 = {
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
	let optsChartBar2 = {
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
	let optsChartBar3 = {
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
	let optsChartBar4 = {
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

// SLIDE 3: Stacked Bar Chart
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable(
		[[{ text: "Chart Examples: Bar Chart: Stacked/PercentStacked and Data Table", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
		BASE_TABLE_OPTS
	);

	let arrDataRegions = [
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
	let arrDataHighVals = [
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
	let optsChartBar1 = {
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
	let optsChartBar2 = {
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
	let optsChartBar3 = {
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
	let optsChartBar4 = {
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

// SLIDE 4: Bar Chart - Lots of Bars
function genSlide04(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Lots of Bars (>26 letters)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataHighVals = [
		{
			name: "Single Data Set",
			labels: LETTERS.concat(["AA", "AB", "AC", "AD"]),
			values: [-5, -3, 0, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30],
		},
	];

	let optsChart = {
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

// SLIDE 5: Bar Chart: Data Series Colors, majorUnits, and valAxisLabelFormatCode
function genSlide05(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable(
		[
			[
				{
					text:
						"Chart Examples: Multi-Color Bars, `catLabelFormatCode`, `valAxisDisplayUnit`, `valAxisMajorUnit`, `valAxisLabelFormatCode`",
					options: BASE_TEXT_OPTS_L,
				},
				BASE_TEXT_OPTS_R,
			],
		],
		BASE_TABLE_OPTS
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

// SLIDE 6: 3D Bar Chart
function genSlide06(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: 3D Bar Chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataRegions = [
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
	let arrDataHighVals = [
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
	let optsChartBar1 = {
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
	let optsChartBar2 = {
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
	let optsChartBar3 = {
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
	let optsChartBar4 = {
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

// SLIDE 7: Tornado Chart
function genSlide07(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Tornado Chart - Grid and Axis Formatting", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

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

// SLIDE 8: Line Chart: Line Smoothing, Line Size, Symbol Size
function genSlide08(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable(
		[[{ text: "Chart Examples: Line Smoothing, Line Size, Line Shadow, Symbol Size", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
		BASE_TABLE_OPTS
	);

	slide.addText("..", { x: 0.5, y: 0.6, w: 6.0, h: 3.0, fill: { color: "F1F1F1" }, color: "F1F1F1" });
	let optsChartLine1 = {
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

	let optsChartLine2 = {
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

	// Create a gap for testing `displayBlanksAs` in line charts (2.3.0)
	//arrDataLineStat[2].values = [55, null, null, 55]; // NOTE: uncomment only for test - looks broken otherwise!

	let optsChartLine3 = {
		x: 0.5,
		y: 4.0,
		w: 6.0,
		h: 3.0,
		chartColors: [COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK],
		lineDataSymbolSize: 10,
		shadow: { type: "none" },
		//displayBlanksAs: 'gap', // NOTE: uncomment only for test - looks broken otherwise!
		showLegend: true,
		legendPos: "l",
	};
	slide.addChart(pptx.charts.LINE, arrDataLineStat, optsChartLine3);

	// QA: DEMO: Test shadow option
	let shadowOpts = { type: "outer", color: "cd0011", blur: 3, offset: 12, angle: 75, opacity: 0.8 };
	let optsChartLine4 = {
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
	slide.addChart(pptx.charts.LINE, arrDataLineStat, optsChartLine4);
}

// SLIDE 9: Line Chart: TEST: `lineDataSymbol` + `lineDataSymbolSize`
function genSlide09(pptx) {
	let intWgap = 4.25;
	let opts_lineDataSymbol = ["circle", "dash", "diamond", "dot", "none", "square", "triangle"];
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable(
		[[{ text: "Chart Examples: Line Chart: lineDataSymbol option test", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
		BASE_TABLE_OPTS
	);

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

// SLIDE 10: Line Chart: Lots of Cats
function genSlide10(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Line Chart: Lots of Lines", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let MAXVAL = 20000;

	let arrDataTimeline = [];
	for (let idx = 0; idx < 15; idx++) {
		let tmpObj = {
			name: "Series" + idx,
			labels: MONS,
			values: [],
		};

		for (let idy = 0; idy < MONS.length; idy++) {
			tmpObj.values.push(Math.floor(Math.random() * MAXVAL) + 1);
		}

		arrDataTimeline.push(tmpObj);
	}

	// FULL SLIDE:
	let optsChartLine1 = {
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

// SLIDE 11: Area Chart: Misc
function genSlide11(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Area Chart, Stacked Area Chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataAreaSm = [
		{
			name: "Small Samples",
			labels: ["Q1", "Q2", "Q3", "Q4"],
			values: [15, 46, 31, 85],
		},
	];
	let arrDataTimeline2ser = [
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
	let optsChartLine1 = {
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
	let optsChartLine2 = {
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
	let optsChartLine3 = {
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
	let optsChartLine4 = { x: 7, y: 4.0, w: "45%", h: 3, chartColors: ["CC8833", "CCFF69"], chartColorsOpacity: 75 };
	slide.addChart(pptx.charts.AREA, arrDataTimeline2ser, optsChartLine4);
}

// SLIDE 12: Pie Charts: All 4 Legend Options
function genSlide12(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Pie Charts: Legends", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	// [TEST][INTERNAL USE]: Not visible to user (its behind a chart): Used for ensuring ref counting works across obj types (eg: `rId` check/test)
	if (TESTMODE)
		slide.addImage({
			path: NODEJS ? IMAGE_PATHS.ccCopyRemix.path.replace(/http.+\/examples/, "../common") : IMAGE_PATHS.ccCopyRemix.path,
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

// SLIDE 13: Doughnut Chart
function genSlide13(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Doughnut Chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let optsChartPie1 = {
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

	let optsChartPie2 = {
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

// SLIDE 14: XY Scatter Chart
function genSlide14(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: XY Scatter Chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataScatter1 = [
		{ name: "X-Axis", values: [0, 1, 2, 3, 4, 5] },
		{ name: "Y-Value 1", values: [90, 80, 70, 85, 75, 92], labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"] },
		{ name: "Y-Value 2", values: [21, 32, 40, 49, 31, 29], labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"] },
	];
	let arrDataScatter2 = [
		{ name: "X-Axis", values: [1, 2, 3, 4, 5, 6] },
		{ name: "Airplane", values: [33, 20, 51, 65, 71, 75] },
		{ name: "Train", values: [99, 88, 77, 89, 99, 99] },
		{ name: "Bus", values: [21, 22, 25, 49, 59, 69] },
	];
	let arrDataScatterLabels = [
		{ name: "X-Axis", values: [1, 10, 20, 30, 40, 50] },
		{ name: "Y-Value 1", values: [11, 23, 31, 45, 47, 35], labels: ["Red 1", "Red 2", "Red 3", "Red 4", "Red 5", "Red 6"] },
		{ name: "Y-Value 2", values: [21, 38, 47, 59, 51, 25], labels: ["Blue 1", "Blue 2", "Blue 3", "Blue 4", "Blue 5", "Blue 6"] },
	];

	// TOP-LEFT
	let optsChartScat1 = {
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
	let optsChartScat2 = {
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
	let optsChartScat3 = {
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
	let optsChartScat4 = { x: 7.0, y: 4.0, w: "45%", h: 3 };
	slide.addChart(pptx.charts.SCATTER, arrDataScatter2, optsChartScat4);
}

// SLIDE 15: Bubble Charts
function genSlide15(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Bubble Charts", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataBubble1 = [
		{ name: "X-Axis", values: [0.3, 0.6, 0.9, 1.2, 1.5, 1.7] },
		{ name: "Y-Value 1", values: [1.3, 9, 7.5, 2.5, 7.5, 5], sizes: [1, 4, 2, 3, 7, 4] },
		{ name: "Y-Value 2", values: [5, 3, 2, 7, 2, 10], sizes: [9, 7, 9, 2, 4, 8] },
	];
	let arrDataBubble2 = [
		{ name: "X-Axis", values: [1, 2, 3, 4, 5, 6] },
		{ name: "Airplane", values: [33, 20, 51, 65, 71, 75], sizes: [10, 10, 12, 12, 15, 20] },
		{ name: "Train", values: [99, 88, 77, 89, 99, 99], sizes: [20, 20, 22, 22, 25, 30] },
		{ name: "Bus", values: [21, 22, 25, 49, 59, 69], sizes: [11, 11, 13, 13, 16, 21] },
	];

	// TOP-LEFT
	let optsChartBubble1 = {
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
	let optsChartBubble2 = {
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
	let optsChartBubble3 = {
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
	let optsChartBubble4 = { x: 7.0, y: 4.0, w: "45%", h: 3, lineSize: 0 };
	slide.addChart(pptx.charts.BUBBLE, arrDataBubble2, optsChartBubble4);
}

// SLIDE 16: Radar Chart
function genSlide16(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Radar Chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataRegions = [
		{
			name: "Region 1",
			labels: ["May", "June", "July", "August", "September"],
			values: [26, 53, 100, 75, 41],
		},
	];
	let arrDataHighVals = [
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
	let optsChartRadar1 = { x: 0.5, y: 0.6, w: 6.0, h: 3.0, radarStyle: "standard", lineDataSymbol: "none", fill: "F1F1F1" };
	slide.addChart(pptx.charts.RADAR, arrDataRegions, optsChartRadar1);

	// TOP-RIGHT: Marker
	let optsChartRadar2 = {
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
	let optsChartRadar3 = {
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
	let optsChartRadar4 = {
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

// SLIDE 17: Multi-Type Charts
function genSlide17(pptx) {
	// powerpoint 2016 add secondary category axis labels
	// https://peltiertech.com/chart-with-a-dual-category-axis/

	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Multi-Type Charts", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	function doStackedLine() {
		// TOP-RIGHT:
		let opts = {
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

		let labels = ["Mon", "Tue", "Wed", "Thu", "Fri"];
		let chartTypes = [
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
		let opts = {
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

		let labels = ["April", "May", "June", "July", "August"];
		let chartTypes = [
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
		let opts = {
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

		let labels = ["Q1", "Q2", "Q3", "Q4", "OT"];
		let chartTypes = [
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
		let opts = {
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

		let labels = ["Q1", "Q2", "Q3", "Q4", "OT"];
		let chartTypes = [
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
		let labels = ["Q1", "Q2", "Q3", "Q4", "OT"];
		let chartTypes = [
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
		let multiOpts = {
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

// SLIDE 18: Charts Options: Shadow, Transparent Colors
function genSlide18(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Options: Shadow, Transparent Colors", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataRegions = [
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
	let arrDataHighVals = [
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
	let single = [
		{
			name: "Texas",
			labels: ["Apartment", "Townhome", "Duplex", "House", "Big House"],
			values: [1400, 2000, 2500, 3000, 3800],
		},
	];

	// TOP-LEFT: H/bar
	let optsChartBar1 = {
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

	let pieOptions = {
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
	let optsChartBar3 = {
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
	let optsChartBar4 = {
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
