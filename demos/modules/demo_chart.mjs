/**
 * NAME: demo_chart.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.12.0
 * BLD.: 20230116
 */

import { BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R, FOOTER_TEXT_OPTS, IMAGE_PATHS, TESTMODE } from "./enums.mjs";
import {
	COLORS_ACCENT,
	CHART_DATA,
	COLORS_CHART,
	COLORS_RYGU,
	COLORS_SPECTRUM,
	COLORS_VIVID,
	LETTERS,
	MONS,
	arrDataLineStat,
	dataChartBar3Series,
	dataChartBar8Series,
	dataChartPieLocs,
	dataChartPieStat,
} from "./enums_charts.mjs";

export function genSlides_Chart(pptx) {
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
	genSlide19(pptx);
	genSlide20(pptx);
	genSlide21(pptx);

	if (TESTMODE) {
		pptx.addSection({ title: "Charts-DevTest" });
		devSlide01(pptx);
		devSlide02(pptx);
		devSlide03(pptx);
		devSlide04(pptx);
		devSlide05(pptx);
		devSlide06(pptx);
		devSlide07(pptx);
	}
}

// SLIDE 1: Bar Chart: Chart Title, Cat/Val Axis Title
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Options: Chart Title, Cat/Val Axis Title", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let optsChart = {
		x: 0.5,
		y: 0.5,
		w: "90%",
		h: "90%",
		barDir: "col",
		barGrouping: "stacked",
		chartColors: COLORS_CHART,
		invertedColors: ["C0504D"],
		showLegend: true,
		//
		showTitle: true,
		title: "Chart Title",
		titleFontFace: "Helvetica Neue Thin",
		titleFontSize: 24,
		titleColor: COLORS_ACCENT[0],
		titlePos: { x: 1.5, y: 0 },
		//titleRotate: 10,
		//
		showCatAxisTitle: true,
		catAxisLabelColor: COLORS_ACCENT[1],
		catAxisTitleColor: COLORS_ACCENT[1],
		catAxisTitle: "Cat Axis Title",
		catAxisTitleFontSize: 14,
		//
		showValAxisTitle: true,
		valAxisLabelColor: COLORS_ACCENT[2],
		valAxisTitleColor: COLORS_ACCENT[2],
		valAxisTitle: "Val Axis Title",
		valAxisTitleFontSize: 14,
	};

	// TEST `getExcelColName()` to ensure Excel Column names are generated correctly above >26 chars/cols
	slide.addChart(pptx.charts.BAR, dataChartBar8Series, optsChart);
}

// SLIDE 2: Bar Chart: Various Designs
function genSlide02(pptx) {
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
	let arrDataSersCats = [
		{ name: "Series 1", labels: ["Category 1", "Category 2", "Category 3", "Category 4"], values: [4.3, 2.5, 3.5, 4.5] },
		{ name: "Series 2", labels: ["Category 1", "Category 2", "Category 3", "Category 4"], values: [2.4, 4.4, 1.8, 2.8] },
		{ name: "Series 3", labels: ["Category 1", "Category 2", "Category 3", "Category 4"], values: [2, 2, 3, 5] },
	];

	// TOP-LEFT: H/bar
	let optsChartBar1 = {
		x: 0.5,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { border: { color: COLORS_SPECTRUM[0], pt: 1 } },
		//chartArea: { fill: { color: pptx.colors.BACKGROUND2 }, border: { color: pptx.colors.BACKGROUND2, pt: 1 }  },
		plotArea: { fill: { color: "DAE3F3" } },
		chartColors: COLORS_SPECTRUM,

		objectName: "bar chart (top L)",
		altText: "this is the alt text content",
		barDir: "bar",

		catAxisLabelColor: COLORS_ACCENT[0],
		catAxisLabelFontFace: "Helvetica Neue",
		catAxisLabelFontSize: 12,
		catAxisOrientation: "maxMin",
		catAxisMajorTickMark: "in",
		catAxisMinorTickMark: "cross",

		valAxisMajorTickMark: "cross",
		valAxisMinorTickMark: "out",
		//valAxisLabelColor: COLORS_ACCENT[0],
		//valAxisCrossesAt: 100,
	};
	slide.addChart(pptx.charts.BAR, arrDataSersCats, optsChartBar1);

	// TOP-RIGHT: V/col
	let optsChartBar2 = {
		x: 7.0,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { border: { color: COLORS_SPECTRUM[0], pt: 1 } },
		//chartArea: { fill: { color: pptx.colors.BACKGROUND2 } },
		plotArea: { fill: { color: "DAE3F3" } },
		//plotArea: { fill: { color: pptx.colors.BACKGROUND1 }, border: { color: pptx.colors.BACKGROUND2, pt: 1 } },
		chartColors: COLORS_SPECTRUM,

		objectName: "bar chart (top R)",
		barDir: "col",

		catAxisLabelColor: COLORS_ACCENT[0],
		catAxisLabelFontFace: "Arial",
		catAxisLabelFontSize: 11,
		catAxisOrientation: "minMax",
		catAxisMajorTickMark: "none",
		catAxisMinorTickMark: "none",

		dataBorder: { pt: 1, color: "F1F1F1" },
		dataLabelColor: COLORS_ACCENT[0],
		dataLabelFontFace: "Courier",
		dataLabelFontSize: 10,
		dataLabelPosition: "outEnd",
		dataLabelFormatCode: "#.0",
		showValue: true,

		valAxisLabelColor: COLORS_ACCENT[0],
		valAxisOrientation: "maxMin",
		valAxisMajorTickMark: "none",
		valAxisMinorTickMark: "none",
		//valAxisLogScaleBase: '25',

		showLegend: false,
		showTitle: false,
	};
	slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar2);

	// BTM-LEFT: H/bar - TITLE and LEGEND
	let optsChartBar3 = {
		x: 0.5,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		barDir: "bar",

		chartArea: { fill: { color: pptx.colors.BACKGROUND2 }, border: { color: pptx.colors.ACCENT3, pt: 2 } },
		//chartArea: { fill: { color: pptx.colors.BACKGROUND2 } },
		//chartArea: { fill: { color: "F1F1F1", transparency: 75 } },
		//chartArea: { fill: { color: "F1F1F1" } },
		plotArea: { fill: { color: "F2F9FC" } },
		//plotArea: { border: { pt: 3, color: "CF0909" }, fill: { color: "F1C1C1" } },
		//plotArea: { border: { pt: 3, color: "CF0909" }, fill: { color: "F1C1C1", transparency: 10 } },

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
	slide.addChart(pptx.charts.BAR, dataChartBar3Series, optsChartBar3);

	// BTM-RIGHT: V/col - TITLE and LEGEND
	let optsChartBar4 = {
		x: 7.0,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		chartArea: { fill: { color: "F1F1F1" } },
		plotArea: { fill: { color: "404040" } },
		//
		barDir: "col",
		barGapWidthPct: 25,
		chartColors: COLORS_ACCENT,
		chartColorsOpacity: 50,
		//
		catAxisLabelColor: COLORS_ACCENT[0],
		catAxisLabelFontFace: "Calibri",
		catAxisLabelFontSize: 11,
		catAxisOrientation: "maxMin",
		//
		valAxisMaxVal: 5000,
		valAxisLabelColor: COLORS_ACCENT[0],
		//
		dataBorder: { pt: 1, color: "F1F1F1" },
		dataLabelColor: "FFFFFF",
		dataLabelFontFace: "Arial",
		dataLabelFontSize: 10,
		dataLabelPosition: "inEnd",
		showValue: true,
		//
		showLegend: false,
		legendPos: "b",
		legendColor: COLORS_ACCENT[1],
		//
		showTitle: true,
		title: "Device Prices",
		titleColor: COLORS_ACCENT[0],
	};
	slide.addChart(pptx.charts.BAR, dataChartBar3Series, optsChartBar4);
}

// SLIDE 3: Bar Chart Options: Axis, DataLabel, Grid
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Bar Chart Options: Axis, DataLabel, Grid", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

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
		plotArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_ACCENT,

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
		titleFontSize: 11,
	};
	slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar1);

	// TOP-RIGHT: V/col
	let optsChartBar2 = {
		x: 7.0,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		barDir: "col",
		plotArea: { fill: { color: "E1F1FF" } },

		dataBorder: { pt: 1, color: "F1F1F1" },
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
	let optsChartBar3 = {
		x: 0.5,
		y: 3.8,
		w: 6.0,
		h: 3.5,

		chartArea: { fill: { color: "F1F1F1" } },
		plotArea: { border: { pt: 3, color: "CF0909" }, fill: { color: "F1C1C1" } },

		barDir: "bar",
		barOverlapPct: -50,

		catAxisLabelColor: "CC0000",
		catAxisLabelFontFace: "Helvetica Neue",
		catAxisLabelFontSize: 10,
		catAxisOrientation: "maxMin",
		catAxisTitle: "Housing Type",
		catAxisTitleColor: "696969",
		catAxisTitleFontSize: 10,
		showCatAxisTitle: true,

		catGridLine: { color: "cc6699", style: "dash", size: 1 },
		valGridLine: { style: "none" },
		valAxisOrientation: "maxMin",
		valAxisHidden: true,
		valAxisDisplayUnitLabel: true,

		titleColor: "33CF22",
		titleFontFace: "Helvetica Neue",
		titleFontSize: 16,

		showTitle: true,
		title: "Sales by Region",
	};
	slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChartBar3);

	// BTM-RIGHT: V/col - TITLE and LEGEND
	let optsChartBar4 = {
		x: 7.0,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		chartArea: { fill: { color: "F1F1F1" } },
		plotArea: { fill: { color: "FFFFFF" } },

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

		dataBorder: { pt: 1, color: "F1F1F1" },
		dataLabelColor: "696969",
		dataLabelFontFace: "Arial",
		dataLabelFontSize: 10,
		dataLabelPosition: "inEnd",
		showValue: true,

		valAxisHidden: true,
		catAxisTitle: "Housing Type",
		showCatAxisTitle: true,

		showLegend: false,
		showTitle: false,
	};
	slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChartBar4);
}

// SLIDE 4: Bar Chart: Stacked
function genSlide04(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable(
		[[{ text: "Chart Examples: Bar Chart: Stacked/PercentStacked and DataTable", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
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

	// TOP-LEFT: H/bar
	let optsChartBar1 = {
		x: 0.5,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "404040" } },
		plotArea: { fill: { color: "0d0d0d" } },
		barDir: "bar",
		barGrouping: "stacked",
		chartColors: ["F2AF00", "4472C4"],

		catAxisOrientation: "maxMin",
		catAxisLabelColor: "4472C4",
		catAxisLabelFontFace: "Helvetica Neue",
		catAxisLabelFontSize: 14,
		//catAxisLabelFontBold: true,
		valAxisLabelColor: "F2AF00",
		valAxisLabelFontFace: "Helvetica Neue",
		valAxisLabelFontSize: 14,
		//valAxisLabelFontBold: true,
		dataLabelColor: "FFFFFF",
		showValue: true,
	};
	slide.addChart(pptx.charts.BAR, arrDataRegions, optsChartBar1);

	// TOP-RIGHT: V/col
	let optsChartBar2 = {
		x: 7.0,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "0d0d0d" } },
		plotArea: { fill: { color: "4d4d4d" } },
		chartColors: COLORS_VIVID,
		valGridLine: { color: "141414" },
		valAxisLabelColor: "F1F1F1",
		catAxisLabelColor: "F1F1F1",
		dataLabelColor: "F1F1F1",

		barDir: "col",
		barGrouping: "stacked",

		dataLabelFontFace: "Arial",
		dataLabelFontSize: 12,
		dataLabelFontBold: true,
		showValue: true,

		catAxisLabelFontFace: "Courier",
		catAxisLabelFontSize: 12,
		catAxisOrientation: "minMax",

		showLegend: false,
		showTitle: false,
	};
	slide.addChart(pptx.charts.BAR, dataChartBar3Series, optsChartBar2);

	// BTM-LEFT: H/bar - 100% layout without axis labels
	let optsChartBar3 = {
		x: 0.5,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		barDir: "bar",
		barGrouping: "percentStacked",
		chartColors: ["F2AF00", "4472C4"],
		dataBorder: { pt: 1, color: "F1F1F1" },
		catAxisHidden: true,
		valAxisHidden: true,
		valGridLine: { style: "none" },
		showTitle: false,
		//
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
	let optsChartBar4 = {
		x: 7.0,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		chartArea: { fill: { color: "f1f1f1" } },
		plotArea: { fill: { color: "ffffff" } },
		chartColors: COLORS_VIVID,
		//
		barDir: "col",
		barGrouping: "percentStacked",
		catAxisLabelFontFace: "Times",
		catAxisLabelFontSize: 12,
		catAxisOrientation: "minMax",
		showLegend: true,
		legendPos: "t",
		showDataTable: true,
		showDataTableKeys: false,
		dataTableFormatCode: "$#",
		//dataTableFormatCode: '0.00%' // @since v3.3.0
		//dataTableFormatCode: '$0.00' // @since v3.3.0
	};
	slide.addChart(pptx.charts.BAR, dataChartBar3Series, optsChartBar4);
}

// SLIDE 5: Bar Chart: Data Series Colors, majorUnits, and valAxisLabelFormatCode
function genSlide05(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable(
		[
			[
				{
					text: "Chart Examples: Multi-Color Bars, `catLabelFormatCode`, `valAxisDisplayUnit`, `valAxisMajorUnit`, `valAxisLabelFormatCode`",
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
			chartArea: { fill: { color: "404040" } },
			barDir: "bar",
			chartColors: ["0077BF", "4E9D2D", "ECAA00", "5FC4E3", "DE4216", "154384"],
			//
			catAxisLabelColor: "F1F1F1",
			catLabelFormatCode: "yyyy-mm",
			/*
			valAxisLabelColor: "F1F1F1",
			valAxisMajorUnit: 15,
			valAxisDisplayUnit: "hundreds",
			valAxisMaxVal: 45,
			valLabelFormatCode: "$0", // @since v3.3.0
			*/
			valAxisHidden: true,
			//
			showTitle: true,
			title: "Categories can be Multi-Color",
			titleColor: "0088CC",
			titleFontSize: 14,
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
			chartArea: { fill: { color: "404040" } },
			catAxisLabelColor: "F1F1F1",
			valAxisLabelColor: "F1F1F1",
			valAxisLineColor: "7F7F7F",
			valGridLine: { color: "7F7F7F" },
			dataLabelColor: "B7B7B7",

			valAxisMaxVal: 1,
			barDir: "bar",
			catAxisLineShow: false,
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
			chartArea: { fill: { color: "404040" } },
			plotArea: { fill: { color: "202020" } },
			catAxisLabelColor: "F1F1F1",
			valAxisLabelColor: "F1F1F1",
			valAxisLineColor: "7F7F7F",
			valGridLine: { color: "7F7F7F" },
			dataLabelColor: "B7B7B7",
			valAxisHidden: true,
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
			legendColor: "F1F1F1",
		}
	);

	// BOTTOM-RIGHT
	slide.addChart(
		pptx.charts.BAR,
		[
			{
				name: "EV",
				labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
				values: [102, 103, 121, 125, 135, 155],
			},
			{
				name: "ICE",
				labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
				values: [150, 153, 151, 125, 115, 105],
			},
		],
		{
			x: 7,
			y: 4,
			w: "45%",
			h: 3,
			chartArea: { fill: { color: "202020" } },
			barDir: "bar",
			catAxisLabelColor: "F1F1F1",
			valAxisLabelColor: "F1F1F1",
			valAxisLineColor: "7F7F7F",
			valGridLine: { color: "7F7F7F" },
			dataLabelColor: "B7B7B7",
			chartColorsOpacity: 75,
			//showValue: true,
			//dataLabelPosition: "outEnd",
			chartColors: ["0077BF", "4E9D2D", "ECAA00", "5FC4E3", "DE4216", "154384", "7D666A", "A3C961", "EF907B", "9BA0A3"],
			barGapWidthPct: 25,
			catAxisOrientation: "maxMin",
			valAxisOrientation: "maxMin",
			valAxisMaxVal: 200,
			valAxisMajorUnit: 25,
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
			labels: ["Q1", "Q2", "Q3", "Q4"],
			values: [26, 53, 80, 75],
		},
		{
			name: "Region 2",
			labels: ["Q1", "Q2", "Q3", "Q4"],
			values: [43.5, 70.3, 90.01, 80.05],
		},
	];

	// TOP-LEFT: H/bar
	let optsChartBar1 = {
		x: 0.5,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1", transparency: 50 } },
		barDir: "bar",
		barGapWidthPct: 25,
		chartColors: COLORS_SPECTRUM,
		chartColorsOpacity: 80,
		//
		v3DRotX: 20,
		v3DRotY: 10,
		v3DRAngAx: false,
		//
		catAxisLabelColor: COLORS_SPECTRUM[1],
		catAxisLineColor: COLORS_SPECTRUM[1],
		catAxisLabelFontFace: "Arial",
		catAxisLabelFontSize: 10,
		catAxisOrientation: "maxMin",
		//
		serAxisLabelFontFace: "Arial",
		serAxisLabelFontSize: 10,
		serAxisLabelColor: COLORS_SPECTRUM[1],
		serAxisLineColor: COLORS_SPECTRUM[1],
		//serAxisLineColor: pptx.colors.ACCENT6,
		//
		//valAxisHidden: true,
		valAxisLabelColor: COLORS_SPECTRUM[0],
		valAxisLineColor: COLORS_SPECTRUM[0],
		valAxisLabelFontSize: 10,
	};
	slide.addChart(pptx.charts.BAR3D, arrDataRegions, optsChartBar1);

	// TOP-RIGHT: V/col
	let optsChartBar2 = {
		x: 7.0,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1", transparency: 50 } },
		chartColors: COLORS_SPECTRUM,
		barDir: "col",
		bar3DShape: "cylinder",
		//
		v3DRotX: 10,
		v3DRotY: 20,
		v3DRAngAx: false,
		//
		catAxisLabelColor: "0000CC",
		catAxisLabelFontFace: "Courier",
		catAxisLabelFontSize: 12,
		//
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
	let optsChartBar3 = {
		x: 0.5,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		chartArea: { fill: { color: "F1F1F1", transparency: 50 } },
		chartColors: COLORS_ACCENT,
		//
		barDir: "col",
		bar3DShape: "pyramid",
		barGrouping: "stacked",
		v3DRAngAx: true,
		//
		catAxisLabelFontFace: "Arial",
		catAxisLabelFontSize: 10,
		//
		showValue: true,
		dataLabelBkgrdColors: true,
		//
		showTitle: true,
		title: "Sales by Region",
		titleFontFace: "Helvetica Neue Thin",
		titleFontSize: 18,
		titleColor: COLORS_ACCENT[0],
	};
	slide.addChart(pptx.charts.BAR3D, arrDataRegions, optsChartBar3);

	// BTM-RIGHT: V/col - TITLE and LEGEND
	let optsChartBar4 = {
		x: 7.0,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		chartArea: { fill: { color: "F1F1F1", transparency: 50 } },
		//
		chartColors: COLORS_ACCENT,
		barDir: "col",
		bar3DShape: "coneToMax",
		v3DRAngAx: true,
		//
		catAxisLabelColor: COLORS_ACCENT[0],
		catAxisLabelFontSize: 11,
		catAxisOrientation: "minMax",
		//
		serAxisLabelFontFace: "Helvetica Neue Thin",
		serAxisLabelColor: COLORS_ACCENT[0],
		//
		dataBorder: { pt: 1, color: "F1F1F1" },
		dataLabelColor: "696969",
		dataLabelFontFace: "Arial",
		dataLabelFontSize: 10,
		dataLabelPosition: "ctr",
	};
	slide.addChart(pptx.charts.BAR3D, arrDataRegions, optsChartBar4);
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
			chartArea: { fill: { color: "F1F1F1", transparency: 50 } },

			valAxisMaxVal: 1,
			barDir: "bar",
			axisLabelFormatCode: "#%",
			catGridLine: { color: "D8D8D8", style: "dash", size: 1, cap: "round" },
			valGridLine: { color: "D8D8D8", style: "dash", size: 1, cap: "square" },
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

// SLIDE 8: Line Chart
function genSlide08(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Line Chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addText(`(${CHART_DATA.LongTermIntRates.sourceUrl})`, FOOTER_TEXT_OPTS);

	// FULL SLIDE:
	const OPTS_CHART = {
		x: 0.5,
		y: 0.6,
		w: "95%",
		h: "85%",
		plotArea: { fill: { color: "F2F9FC" } },
		//
		showLegend: true,
		legendPos: "r",
		//
		showTitle: true,
		lineDataSymbol: "none",
		title: CHART_DATA.LongTermIntRates.chartTitle,
		titleColor: "0088CC",
		titleFontFace: "Arial",
		titleFontSize: 18,
	};
	slide.addChart(pptx.charts.LINE, CHART_DATA.LongTermIntRates.chartData, OPTS_CHART);
}

// SLIDE 9: Line Chart: Line Smoothing, Line Size, Symbol Size
function genSlide09(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable(
		[[{ text: "Chart Examples: Line Smoothing, Line Size, Line Shadow, Symbol Size", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
		BASE_TABLE_OPTS
	);

	let optsChartLine1 = {
		x: 0.5,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_RYGU,
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
		chartArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_RYGU,
		lineSize: 16,
		lineSmooth: true,
		showLegend: true,
		legendPos: "r",
	};
	slide.addChart(pptx.charts.LINE, arrDataLineStat, optsChartLine2);

	let optsChartLine3 = {
		x: 0.5,
		y: 4.0,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_RYGU,
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
		chartArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_RYGU,
		lineDataSymbolSize: 20,
		shadow: shadowOpts,
		showLegend: true,
		legendPos: "b",
	};
	slide.addChart(pptx.charts.LINE, arrDataLineStat, optsChartLine4);
}

// SLIDE 10: Line Chart: `lineDataSymbol` and `lineDataSymbolSize`
function genSlide10(pptx) {
	const intWgap = 4.25;
	const opts_lineDataSymbol = ["circle", "dash", "diamond", "dot", "none", "square", "triangle"];
	const slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Line Chart: lineDataSymbol options", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	opts_lineDataSymbol.forEach((opt, idx) => {
		slide.addChart(pptx.charts.LINE, arrDataLineStat, {
			x: (idx < 3 ? idx * intWgap : idx < 6 ? (idx - 3) * intWgap : (idx - 6) * intWgap) + 0.3,
			y: idx < 3 ? 0.5 : idx < 6 ? 2.85 : 5.1,
			w: 4.25,
			h: 2.25,
			lineCap: 'round',
			lineDataSymbol: opt,
			lineDataSymbolSize: idx == 5 ? 9 : idx == 6 ? 12 : null,
			chartColors: COLORS_VIVID,
			title: opt,
			showTitle: true,
		});
	});
}

// SLIDE 11: Area Chart
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
		chartArea: { fill: { color: "e9e9e9" } },
		plotArea: { fill: { color: "f2f9fc" } },
		chartColors: COLORS_VIVID,
		dataBorder: { pt: 1, color: "F1F1F1" },
		showTitle: true,
		title: CHART_DATA.CeoPayRatio_Comp.chartTitle,
		titleFontSize: 11,
		titleColor: "fc0000",
		valAxisLabelFormatCode: "#-1",
		valAxisLabelFontSize: 10,
		valAxisLabelColor: "494949",
		catAxisLabelFontSize: 10,
		catAxisLabelColor: "494949",
		catAxisLabelRotate: 45,
		chartColors: ["EF423E"],
		chartColorsOpacity: 25,
		//showValue: true,
	};
	slide.addChart(pptx.charts.AREA, CHART_DATA.CeoPayRatio_Comp.chartData, optsChartLine1);

	// TOP-RIGHT (stacked area chart)
	let optsChartLine2 = {
		x: 7,
		y: 0.6,
		w: "45%",
		h: 3,
		plotArea: { fill: { color: "D1E1F1" } },

		chartColors: ["0088CC", "99FFCC"],
		chartColorsOpacity: 25,
		valAxisLabelRotate: 5,
		dataBorder: { pt: 2, color: "FFFFFF" },
		showValue: false,
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

// SLIDE 12: Pie Chart
function genSlide12(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Pie Charts: Legends", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	// TOP-LEFT
	slide.addChart(pptx.charts.PIE, dataChartPieStat, {
		x: 0.5,
		y: 0.6,
		w: 4.0,
		h: 3.2,
		chartArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_RYGU,
		dataBorder: { pt: 2, color: "F1F1F1" },
		//
		legendPos: "l",
		legendFontFace: "Courier New",
		showLegend: true,
		//
		showLeaderLines: true,
		showPercent: false,
		showValue: true,
		dataLabelColor: "FFFFFF",
		dataLabelFontSize: 14,
		dataLabelPosition: "bestFit", // 'bestFit' | 'outEnd' | 'inEnd' | 'ctr'
	});

	// TOP-MIDDLE
	slide.addChart(pptx.charts.PIE, dataChartPieStat, {
		x: 4.67,
		y: 0.6,
		w: 4.0,
		h: 3.2,
		chartArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_SPECTRUM,
		dataBorder: { pt: 1, color: "404040" },
		dataLabelColor: "f2f9fc",
		showPercent: true,
		showLegend: true,
		legendPos: "t",
	});

	// TOP-RIGHT (DEMO: `legendFontSize`, `titleAlign`, `titlePos`)
	slide.addChart(pptx.charts.PIE, dataChartPieLocs, {
		x: 8.83,
		y: 0.6,
		w: 4.0,
		h: 3.2,
		chartArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_SPECTRUM,
		dataBorder: { pt: "1", color: "F1F1F1" },
		showLegend: true,
		showPercent: true,
		legendPos: "t",
		legendFontSize: 14,
		showLeaderLines: true,
		showTitle: true,
		title: "Title Position {0,0}",
		titleAlign: "right",
		titlePos: { x: 0, y: 0 },
	});

	// BTM-LEFT
	slide.addChart(pptx.charts.PIE, dataChartPieLocs, {
		x: 0.5,
		y: 4.0,
		w: 4.0,
		h: 3.2,
		chartArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_CHART,
		dataBorder: { pt: "1", color: "F1F1F1" },
		//
		showValue: true,
		showLabel: true,
		showPercent: true,
		//
		dataLabelColor: "F1F1F1",
		dataLabelFontSize: 10,
	});

	// BTM-MIDDLE
	slide.addChart(pptx.charts.PIE, dataChartPieLocs, {
		x: 4.67,
		y: 4.0,
		w: 4.0,
		h: 3.2,
		chartArea: { fill: { color: "F1F1F1" } },
		dataBorder: { pt: "1", color: "F1F1F1" },
		chartColors: COLORS_SPECTRUM,
		dataLabelColor: "F1F1F1",
		showPercent: true,
		showLegend: true,
		legendPos: "b",
	});

	// BOTH: BTM-RIGHT
	slide.addChart(pptx.charts.PIE, dataChartPieLocs, {
		x: 8.83,
		y: 4.0,
		w: 4.0,
		h: 3.2,
		chartArea: { fill: { color: "F1F1F1" } },
		dataBorder: { pt: "1", color: "F1F1F1" },
		showPercent: true,
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
		y: 0.6,
		w: 6.0,
		h: 6.4,
		chartArea: { fill: { color: "F1F1F1" } },
		holeSize: 70,
		showLabel: false,
		showValue: false,
		showPercent: true,
		showLegend: true,
		legendPos: "b",
		//
		chartColors: COLORS_RYGU,
		dataBorder: { pt: "2", color: "F1F1F1" },
		dataLabelColor: "FFFFFF",
		dataLabelFontSize: 14,
		//
		showTitle: false,
		title: "Project Status",
		titleColor: "33CF22",
		titleFontFace: "Helvetica Neue",
		titleFontSize: 24,
	};
	slide.addChart(pptx.charts.DOUGHNUT, dataChartPieStat, optsChartPie1);

	let optsChartPie2 = {
		x: 6.83,
		y: 0.6,
		w: 6.0,
		h: 6.4,
		chartArea: { fill: { color: "404040" } },
		chartColors: COLORS_VIVID,
		dataBorder: { pt: "1", color: "F1F1F1" },
		dataLabelColor: "FFFFFF",
		showLabel: true,
		showValue: true,
		showPercent: true,
		//
		showLegend: true,
		legendPos: "b",
		legendColor: "F1F1F1",
		legendFontSize: 12,
		//
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
		plotArea: { fill: { color: "F1F1F1" } },

		showLegend: true,
		legendPos: "b",

		lineSize: 8,
		lineSmooth: true,
		lineDataSymbolSize: 12,
		lineDataSymbolLineColor: "FFFFFF",

		chartColors: COLORS_RYGU,
		chartColorsOpacity: 25,
	};
	slide.addChart(pptx.charts.SCATTER, arrDataScatter2, optsChartScat2);

	// BOTTOM-LEFT: (Labels)
	let optsChartScat3 = {
		x: 0.5,
		y: 4.0,
		w: "45%",
		h: 3,
		plotArea: { fill: { color: "F2F9FC" } },

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

// SLIDE 15: Bubble Chart
function genSlide15(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Bubble Charts", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataBubble1 = [
		{ name: "X-Axis", values: [0.3, 0.6, 0.9, 1.2, 1.5, 1.7] },
		{ name: "Y-Value 1", values: [1.3, 9, 7.5, 2.5, 7.5, 3], sizes: [1, 4, 2, 3, 7, 4] },
		{ name: "Y-Value 2", values: [5.0, 3, 2.0, 7.0, 2.0, 9], sizes: [9, 7, 9, 2, 4, 8] },
	];
	let arrDataBubble2 = [
		{ name: "X-Axis", values: [1, 2, 3, 4, 5, 6] },
		{ name: "Airplane", values: [33, 20, 51, 65, 71, 75], sizes: [10, 10, 12, 12, 15, 20] },
		{ name: "Train", values: [99, 88, 77, 89, 99, 99], sizes: [20, 20, 22, 22, 25, 30] },
		{ name: "Bus", values: [21, 25, 32, 49, 59, 69], sizes: [11, 11, 13, 13, 16, 21] },
	];

	// TOP-LEFT
	let optsChartBubble1 = {
		x: 0.5,
		y: 0.6,
		w: "45%",
		h: 3,
		chartArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_ACCENT,
		chartColorsOpacity: 40,
		dataBorder: { pt: 1, color: "FFFFFF" },
		//valAxisCrossesAt: 4,
		//catAxisCrossesAt: 4,
		dataLabelFontFace: "Arial",
		dataLabelFontSize: 10,
		dataLabelColor: "363636",
		dataLabelPosition: "r",
		showSerName: true,
		showLeaderLines: true,
	};
	slide.addChart(pptx.charts.BUBBLE, arrDataBubble1, optsChartBubble1);

	// TOP-RIGHT
	let optsChartBubble2 = {
		x: 7.0,
		y: 0.6,
		w: "45%",
		h: 3,
		plotArea: { fill: { color: "F1F1F1" } },
		chartColors: COLORS_RYGU,
		chartColorsOpacity: 25,

		showLegend: true,
		legendPos: "b",

		lineSize: 8,
		lineSmooth: true,
		lineDataSymbolSize: 12,
		lineDataSymbolLineColor: "FFFFFF",
	};
	slide.addChart(pptx.charts.BUBBLE, arrDataBubble2, optsChartBubble2);

	// BOTTOM-LEFT
	let optsChartBubble3 = {
		x: 0.5,
		y: 4.0,
		w: "45%",
		h: 3,
		chartArea: { fill: { color: "404040" } },
		plotArea: { fill: { color: "202020" } },

		catAxisLabelColor: "F1F1F1",
		catAxisLabelFontSize: 10,
		catAxisOrientation: "maxMin",
		showCatAxisTitle: false,
		//
		valAxisLabelColor: "F1F1F1",
		valAxisLabelFontSize: 10,
		valAxisMinVal: 0,
		valAxisOrientation: "maxMin",
		showValAxisTitle: false,
		//
		dataBorder: { pt: 2, color: "e1e1e1" },
		dataLabelFontFace: "Arial",
		dataLabelFontSize: 10,
		dataLabelColor: "e1e1e1",
		showValue: true,
	};
	slide.addChart(pptx.charts.BUBBLE, arrDataBubble1, optsChartBubble3);

	// BOTTOM-RIGHT
	let optsChartBubble4 = { x: 7.0, y: 4.0, w: "45%", h: 3, lineSize: 0, chartColors: COLORS_RYGU };
	slide.addChart(pptx.charts.BUBBLE3D, arrDataBubble2, optsChartBubble4);
}

// SLIDE 16: Radar Chart
function genSlide16(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Radar Chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	const arrDataRegions = [
		{
			name: "Region 1",
			labels: ["Jun", "Jul", "Aug", "Sep"],
			values: [20, 18, 15, 10],
		},
	];
	const arrDataStudents = [
		{
			name: "Student 1",
			labels: ["Logic", "Coding", "Results", "Comments", "Runtime"],
			values: [3, 1, 3, 3, 4],
		},
		{
			name: "Student 2",
			labels: ["Logic", "Coding", "Results", "Comments", "Runtime"],
			values: [1, 2, 2, 3, 2],
		},
		{
			name: "Student 3",
			labels: ["Logic", "Coding", "Results", "Comments", "Runtime"],
			values: [2, 3, 3, 4, 3],
		},
	];

	// TOP-ROW
	{
		// TOP-L: `{ radar:'normal' }`
		let optsChartRadar1 = {
			x: 0.5,
			y: 0.6,
			w: 4.0,
			h: 3.0,
			chartArea: { fill: { color: "F9F9F9" } },
			//
			radarStyle: "standard",
			//
			showTitle: true,
			titleColor: "7F7F7F",
			titleFontFace: "Segoe UI",
			titleFontSize: 12,
			title: "radarStyle: 'standard'",
			//
			lineDataSymbol: "none",
		};
		slide.addChart(pptx.charts.RADAR, arrDataRegions, optsChartRadar1);

		// TOP-C: `{ radar:'marker' }` Cat Axis options
		let optsChartRadar2 = {
			x: 4.65,
			y: 0.6,
			w: 4.0,
			h: 3.0,
			chartArea: { fill: { color: "F9F9F9" } },
			//
			radarStyle: "marker",
			//
			showTitle: true,
			titleColor: "7F7F7F",
			titleFontFace: "Segoe UI",
			titleFontSize: 12,
			title: "radarStyle: 'marker'",
		};
		slide.addChart(pptx.charts.RADAR, arrDataRegions, optsChartRadar2);

		// TOP-R: `{ radar:'marker' }` Cat Axis options
		let optsChartRadar3 = {
			x: 8.8,
			y: 0.6,
			w: 4.0,
			h: 3.0,
			chartArea: { fill: { color: "F9F9F9" } },
			//
			radarStyle: "filled",
			//
			showTitle: true,
			titleColor: "7F7F7F",
			titleFontFace: "Segoe UI",
			titleFontSize: 12,
			title: "radarStyle: 'filled'",
		};
		slide.addChart(pptx.charts.RADAR, arrDataRegions, optsChartRadar3);
	}

	// BTM-ROW
	{
		// BTM-L: marker/line options
		let optsChartRadar10 = {
			x: 0.5,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			chartArea: { fill: { color: "F1F1F1" } },
			//
			radarStyle: "marker",
			catAxisLabelColor: "0088CC",
			catAxisLabelFontFace: "Courier",
			catAxisLabelFontSize: 11,
			//
			chartColors: COLORS_RYGU, // marker & line color
			lineDataSymbol: "diamond", // marker type ('circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'square' | 'triangle')
			lineDataSymbolLineColor: "0088CC", // marker border color (hex)
			lineDataSymbolLineSize: 2, // marker border size (points)
			lineDataSymbolSize: 12, // marker size (2-72)
			lineSize: 3, // line size
			valAxisLineColor: "D9D9D9", // val axis is the main, center N-S, W-E lines
			valAxisLineSize: 2, // val axis lines size
			//
			showLegend: true,
			legendPos: "l",
			//
			showTitle: true,
			title: "Line/Marker Options",
			titleColor: "7F7F7F",
			titleFontFace: "Helvetica Neue",
			titleFontSize: 12,
		};
		slide.addChart(pptx.charts.RADAR, arrDataStudents, optsChartRadar10);

		// BTM-R: Filled/Axis Options
		let optsChartRadar11 = {
			x: 6.83,
			y: 3.8,
			w: 6.0,
			h: 3.5,
			chartArea: { fill: { color: "F1F1F1" } },
			//
			radarStyle: "filled",
			//
			chartColors: COLORS_RYGU, // marker & line color
			chartColorsOpacity: 25,
			catAxisLabelColor: "404040",
			catAxisLabelFontFace: "Segoe UI",
			catAxisLabelFontSize: 10,
			catAxisLineShow: false,
			//
			lineDataSymbolSize: 2, // marker size (2-72)
			lineSize: 1, // line size
			valAxisLabelFontFace: "Segoe UI",
			valAxisLabelFontSize: 10,
			//
			showLegend: true,
			legendPos: "r",
			legendColor: "404040",
			//
			titleColor: "404040",
			titleFontFace: "Helvetica Neue",
			titleFontSize: 12,
			showTitle: true,
			title: "Filled/Axis Options",
		};
		slide.addChart(pptx.charts.RADAR, arrDataStudents, optsChartRadar11);
	}
}

// SLIDE 17: Multi-Level Category Axes (2 levels)
function genSlide17(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Multi-Level Category Axes", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	const arrDataLabels = [
		["Gear", "Bearing", "Motor", "Switch", "Plug", "Cord", "Fuse", "Bulb", "Pump", "Leak", "Seals"],
		["Mechanical", "", "", "Electrical", "", "", "", "", "Hydraulic", "", ""],
	];
	const arrDataRegions = [
		{
			name: "Mechanical",
			labels: arrDataLabels,
			values: [11, 8, 3, 0, 0, 0, 0, 0, 0, 0, 0],
		},
		{
			name: "Electrical",
			labels: arrDataLabels,
			values: [0, 0, 0, 19, 12, 11, 3, 2, 0, 0, 0],
		},
		{
			name: "Hydraulic",
			labels: arrDataLabels,
			values: [0, 0, 0, 0, 0, 0, 0, 0, 4, 3, 1],
		},
	];

	const opts1 = {
		x: 0.5,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1" } },
		catAxisMultiLevelLabels: true,
		catAxisLabelFontFace: "Helvetica Neue Thin",
	};

	const opts2 = {
		x: 6.8,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1" } },
		catAxisMultiLevelLabels: true,
		catAxisLabelFontFace: "Helvetica Neue Thin",
		barDir: "col",
		barGapWidthPct: 0,
		//catAxisLabelColor: "696969",
	};

	const opts3 = {
		x: 0.5,
		y: 4.0,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1" } },
		catAxisMultiLevelLabels: true,
		barDir: "col",
		v3DRAngAx: true,
	};

	const opts4 = {
		x: 6.8,
		y: 4.0,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1" } },
		catAxisMultiLevelLabels: true,
	};

	slide.addChart(pptx.charts.AREA, arrDataRegions, opts1);
	slide.addChart(pptx.charts.BAR, arrDataRegions, opts2);
	slide.addChart(pptx.charts.BAR3D, arrDataRegions, opts3);
	slide.addChart(pptx.charts.LINE, arrDataRegions, opts4);
}

// SLIDE 18: Multi-Level Category Axes (3 Levels)
function genSlide18(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Multi-Level Category Axes (3 Levels)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	const arrDataRegions = [
		{
			name: "Fruits",
			labels: [
				["Q1", "Q2", "Q3", "Q4", "Q1", "Q2", "Q3", "Q4", "Q1", "Q2", "Q3", "Q4", "Q1", "Q2", "Q3", "Q4"],
				["Apple", "", "", "", "Banana", "", "", "", "Apple", "", "", "", "Banana", "", "", ""],
				["2014", "", "", "", "", "", "", "", "2015", "", "", "", "", "", "", ""],
			],
			values: [734, 465, 656, 176, 434, 165, 613, 359, 279, 660, 307, 270, 539, 142, 554, 405],
		},
	];

	const opts1 = {
		x: 0.5,
		y: 0.6,
		w: 12.3,
		h: 6.5,
		chartArea: { fill: { color: "F1F1F1" }, roundedCorners: false },
		catAxisMultiLevelLabels: true,
		chartColors: ["C0504D", "C0504D", "C0504D", "C0504D", "FFC000", "FFC000", "FFC000", "FFC000"],
	};

	slide.addChart(pptx.charts.BAR, arrDataRegions, opts1);
}

// SLIDE 19: Combo Chart
function genSlide19(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Combo Chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addText(`(${CHART_DATA.EvSales_Vol.sourceUrl})`, FOOTER_TEXT_OPTS);

	const comboProps = {
		x: 0.5,
		y: 0.6,
		w: 12.3,
		h: "85%",
		chartArea: { fill: { color: "F1F1F1" } },
		barDir: "col",
		barGrouping: "stacked",
		//
		catAxisLabelColor: "494949",
		catAxisLabelFontFace: "Arial",
		catAxisLabelFontSize: 10,
		catAxisOrientation: "minMax",
		//
		showLegend: true,
		legendPos: "b",
		//
		showTitle: true,
		titleFontFace: "Calibri Light",
		titleFontSize: 14,
		title: CHART_DATA.EvSales_Vol.chartTitle,
		//
		valAxes: [
			{
				showValAxisTitle: true,
				valAxisTitle: "Cars Produced (m)",
				valAxisMaxVal: 10,
				valAxisTitleColor: "1982c4",
				valAxisLabelColor: "1982c4",
			},
			{
				showValAxisTitle: true,
				valAxisTitle: "Global Market Share (%)",
				valAxisMaxVal: 10,
				valAxisTitleColor: "F38940",
				valAxisLabelColor: "F38940",
				valGridLine: { style: "none" },
			},
		],
		//
		catAxes: [{ catAxisTitle: "Year" }, { catAxisHidden: true }],
	};
	const comboTypes = [
		{
			type: pptx.charts.BAR,
			data: CHART_DATA.EvSales_Vol.chartData,
			options: { chartColors: COLORS_SPECTRUM, barGrouping: "stacked" },
		},
		{
			type: pptx.charts.LINE,
			data: CHART_DATA.EvSales_Pct.chartData,
			options: { chartColors: ["F38940"], secondaryValAxis: true, secondaryCatAxis: true },
		},
	];

	slide.addChart(comboTypes, comboProps);
}

// SLIDE 20: Combo Chart: Various Options
function genSlide20(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Combo Chart Options", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	// TOP-L: charts use same val axis (T-B)
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

	// TOP-R: charts use diff val axis (T-B, L-R)
	function doStackedLine() {
		let opts = {
			x: 6.83,
			y: 0.6,
			w: 6.0,
			h: 3.0,
			chartArea: { fill: { color: "F1F1F1" } },
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

	// BTM-L:
	function doStackedDot() {
		let opts = {
			x: 0.5,
			y: 4.0,
			w: 6.0,
			h: 3.0,
			chartArea: { fill: { color: "F1F1F1" } },

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

	// BTM-R:
	function doBarCol() {
		let opts = {
			x: 6.83,
			y: 4.0,
			w: 6.0,
			h: 3.0,
			chartArea: { fill: { color: "F1F1F1" } },

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
					chartColors: ["0077BF", "4E9D2D", "ECAA00", "5FC4E3", "DE4216", "154384"],
					secondaryValAxis: !!opts.valAxes,
					secondaryCatAxis: !!opts.catAxes,
				},
			},
		];
		slide.addChart(chartTypes, opts);
	}

	doColumnAreaLine();
	doStackedLine();
	doStackedDot();
	doBarCol();
}

// SLIDE 21: Misc Options
function genSlide21(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Misc Options: Shadow, Transparent Colors", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

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
		/* NOTE: following are optional and default to `false`, leavign chart "plain" (without labels, etc.)
		dataLabelFontSize: 9,
		showLabel: true,
		showValue: true,
		showPercent: true,
		*/
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

// --------------------------------------------------------------------------------

// DEV/TEST 01: Bar Chart: Lots of Series
function devSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts-DevTest" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "DEV-TEST: lots-of-bars (>26 letters); negative val check", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataHighVals = [
		{
			name: "Alphabet Letter Value",
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
		showLegend: true,
		chartColors: ["154384"],
		invertedColors: ["0088CC"],
		//
		title: "Chart With >26 Cols",
		showTitle: true,
		titleFontSize: 18,
		//
		showCatAxisTitle: true,
		catAxisTitle: "Letters",
		catAxisTitleColor: "4286f4",
		catAxisTitleFontSize: 14,
		//
		showValAxisTitle: true,
		valAxisTitle: "Column Index",
		valAxisTitleColor: "c11c13",
		valAxisTitleFontSize: 16,
	};

	// TEST `getExcelColName()` to ensure Excel Column names are generated correctly above >26 chars/cols
	slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChart);
}

// DEV/TEST 02: Line Chart: Lots of Series
function devSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts-DevTest" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "DEV-TEST: lots-of-series", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let MAXVAL = 20000;

	let arrDataTimeline = [];
	for (let idx = 0; idx < 15; idx++) {
		let tmpObj = {
			name: `Series ${idx}`,
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
		plotArea: { fill: { color: "F2F9FC" } },
		valAxisMaxVal: MAXVAL,
		showLegend: true,
		legendPos: "r",
	};
	slide.addChart(pptx.charts.LINE, arrDataTimeline, optsChartLine1);
}

// DEV/TEST 03: escaped-XML
function devSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts-DevTest" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "DEV-TEST: escaped-xml", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	slide.addChart(
		pptx.charts.BAR,
		[
			{
				name: "Escaped XML chars",
				labels: ["escaped", "xml", "chars", "'", '"', "&", "<", ">"],
				values: [1.2, 2.3, 2.15, 6.05, 8.01, 2.02, 9.9, 0.9],
			},
		],
		{
			x: 0.5,
			y: 0.6,
			w: "90%",
			h: "90%",
			chartArea: { fill: { color: "404040" } },
			catAxisLabelColor: "F1F1F1",
			valAxisLabelColor: "F1F1F1",
			valAxisLineColor: "7F7F7F",
			valGridLine: { color: "7F7F7F" },
			dataLabelColor: "B7B7B7",
			barDir: "bar",
			showValue: true,
			chartColors: [...COLORS_ACCENT, ...COLORS_ACCENT],
			barGapWidthPct: 25,
			catAxisOrientation: "maxMin",
			valAxisOrientation: "maxMin",
			valAxisMaxVal: 10,
			valAxisMajorUnit: 1,
		}
	);
}

// DEV/TEST 04: Combo Chart
function devSlide04(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts-DevTest" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "DEV-TEST: combo-chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

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
				chartColors: COLORS_VIVID,
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
		x: 0.5,
		y: 0.6,
		w: 12.33,
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

// DEV/TEST 05: ref-check
function devSlide05(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts-DevTest" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "DEV-TEST: ref-test", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let optsChartPie1 = {
		x: 0.5,
		y: 0.6,
		w: 6.0,
		h: 6.4,
		chartArea: { fill: { color: "F1F1F1" } },
		holeSize: 70,
		showLabel: false,
		showValue: false,
		showPercent: true,
		showLegend: true,
		legendPos: "b",
		//
		chartColors: COLORS_SPECTRUM,
		dataBorder: { pt: 2, color: "F1F1F1" },
		dataLabelColor: "FFFFFF",
		dataLabelFontSize: 14,
		//
		showTitle: false,
		title: "Project Status",
		titleColor: "33CF22",
		titleFontFace: "Helvetica Neue",
		titleFontSize: 24,
	};
	slide.addChart(pptx.charts.DOUGHNUT, dataChartPieStat, optsChartPie1);

	// [TEST][INTERNAL]: Used for ensuring ref counting works across mixed object types (eg: `rId` check/test)
	slide.addImage({
		path: IMAGE_PATHS.ccCopyRemix.path,
		x: 6.83,
		y: 0.6,
		w: 6.0,
		h: 6.0,
	});
}

// DEV/TEST 06: legacy-tests
function devSlide06(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts-DevTest" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "DEV-TEST: legacy-tests", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataHighVals = [
		{
			name: "Series With Negative Values",
			labels: ["N2", "N1", "ZERO", "P1", "P2", "P3", "P4", "P5", "P6", "P7"],
			values: [-5, -3, 0, 3, 5, 6, 7, 8, 9, 10],
		},
	];

	let optsChartBar1 = {
		x: 0.5,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: pptx.colors.BACKGROUND2 } },
		plotArea: { fill: { color: pptx.colors.BACKGROUND1 }, border: { color: pptx.colors.BACKGROUND2, pt: 1 } },
		//
		objectName: "bar chart (top L)",
		altText: "this is the alt text content",
		barDir: "bar",
		border: { pt: "3", color: "00CE00" }, // @deprecated - legacy text only (dont use this syntax - use `plotArea`)
		fill: "F1F1F1", // @deprecated - legacy text only (dont use this syntax - use `plotArea`)
		//
		catAxisLabelColor: "CC0000",
		catAxisLabelFontFace: "Helvetica Neue",
		catAxisLabelFontSize: 14,
		catAxisOrientation: "maxMin",
		catAxisMajorTickMark: "in",
		catAxisMinorTickMark: "cross",
		//
		//valAxisCrossesAt: 100,
		valAxisMajorTickMark: "cross",
		valAxisMinorTickMark: "out",
		//
		titleColor: "33CF22",
		titleFontFace: "Helvetica Neue",
		titleFontSize: 24,
	};
	slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChartBar1);
}

// DEV/TEST 07: title-options & inverted-colors
function devSlide07(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts-DevTest" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "DEV-TEST: title-options & inverted-colors", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	let arrDataHighVals = [
		{
			name: "Series With Negative Values",
			labels: ["N2", "N1", "ZERO", "P1", "P2", "P3", "P4", "P5", "P6", "P7"],
			values: [-5, -3, 0, 3, 5, 6, 7, 8, 9, 10],
		},
	];

	let optsChart = {
		x: 0.5,
		y: 0.5,
		w: "90%",
		h: "90%",
		barDir: "col",
		//
		showTitle: true,
		title: "Rotated Title",
		titleFontSize: 20,
		titleRotate: 10,
		//
		showLegend: true,
		chartColors: COLORS_CHART,
		invertedColors: ["C0504D"],
		//
		showCatAxisTitle: true,
		catAxisTitle: "Cat Axis Title",
		catAxisTitleColor: "4286f4",
		catAxisTitleFontSize: 14,
		//
		showValAxisTitle: true,
		valAxisTitle: "Val Axis Title",
		valAxisTitleColor: "c11c13",
		valAxisTitleFontSize: 16,
	};

	slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChart);
}

/**
 * TODO:
 * 	// Create a gap for testing `displayBlanksAs` in line charts (2.3.0)
 *	//arrDataLineStat[2].values = [55, null, null, 55]; // NOTE: uncomment only for test - looks broken otherwise!
 */
