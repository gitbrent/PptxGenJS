/**
 * NAME: demo_chart.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.11.0
 * BLD.: 20220605
 */

import {
	BASE_TABLE_OPTS,
	BASE_TEXT_OPTS_L,
	BASE_TEXT_OPTS_R,
	COLOR_AMB,
	COLOR_BLU,
	COLOR_GRN,
	COLOR_RED,
	COLOR_UNK,
	FOOTER_TEXT_OPTS,
	IMAGE_PATHS,
	TESTMODE,
} from "./enums.mjs";

const LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
const MONS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const QTRS = ["Q1", "Q2", "Q3", "Q4"];
const ACCENT_COLORS = ["4472C4", "ED7D31", "FFC000", "70AD47"]; // 1,2,4,6
const COLORS_SPECTRUM = ["56B4E4", "126CB0", "672C7E", "E92A31", "F06826", "E9AF1F", "51B747", "189247"]; // B-G spectrum wheel
const COLORS_CHART = ["003f5c", "0077b6", "084c61", "177e89", "3066be", "00a9b5", "58508d", "bc5090", "db3a34", "ff6361", "ffa600"];
const COLORS_VIVID = ["ff595e", "F38940", "ffca3a", "8ac926", "1982c4", "5FBDE1", "6a4c93"]; // (R, Y, G, B, P)

const INTRATES_URL = "https://data.oecd.org/interest/long-term-interest-rates.htm";
const INTRATES_TITLE = "Long Term Interest Rates";
const INTRATES_LBLS = ["2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020"];
const INTRATES_DATA = [
	{ name: "Canada", labels: INTRATES_LBLS, values: [4.27, 3.61, 3.23, 3.24, 2.78, 1.87, 2.26, 2.23, 1.52, 1.25, 1.78, 2.28, 1.59, 0.75] },
	{ name: "France", labels: INTRATES_LBLS, values: [4.3, 4.23, 3.65, 3.12, 3.32, 2.54, 2.2, 1.67, 0.84, 0.47, 0.81, 0.78, 0.13, -0.15] },
	{ name: "Germany", labels: INTRATES_LBLS, values: [4.22, 3.98, 3.22, 2.74, 2.61, 1.5, 1.57, 1.16, 0.5, 0.09, 0.32, 0.4, -0.25, -0.51] },
	{ name: "Italy", labels: INTRATES_LBLS, values: [4.49, 4.68, 4.31, 4.04, 5.42, 5.49, 4.32, 2.89, 1.71, 1.49, 2.11, 2.61, 1.95, 1.17] },
	{ name: "Japan", labels: INTRATES_LBLS, values: [1.67, 1.47, 1.33, 1.15, 1.1, 0.84, 0.69, 0.52, 0.35, -0.07, 0.05, 0.07, -0.11, -0.01] },
	{ name: "United Kingdom", labels: INTRATES_LBLS, values: [5.01, 4.59, 3.65, 3.62, 3.14, 1.92, 2.39, 2.57, 1.9, 1.31, 1.24, 1.46, 0.94, 0.37] },
	{ name: "United States", labels: INTRATES_LBLS, values: [4.63, 3.67, 3.26, 3.21, 2.79, 1.8, 2.35, 2.54, 2.14, 1.84, 2.33, 2.91, 2.14, 0.89] },
];

const CEOPAY_URL = "https://www.epi.org/publication/ceo-pay-in-2020/";
const CEOPAY_TITLE = "CEO-to-worker compensation ratio";
const CEOPAY_LBLS = [
	"1965",
	"1966",
	"1967",
	"1968",
	"1969",
	"1970",
	"1971",
	"1972",
	"1973",
	"1974",
	"1975",
	"1976",
	"1977",
	"1978",
	"1979",
	"1980",
	"1981",
	"1982",
	"1983",
	"1984",
	"1985",
	"1986",
	"1987",
	"1988",
	"1989",
	"1990",
	"1991",
	"1992",
	"1993",
	"1994",
	"1995",
	"1996",
	"1997",
	"1998",
	"1999",
	"2000",
	"2001",
	"2002",
	"2003",
	"2004",
	"2005",
	"2006",
	"2007",
	"2008",
	"2009",
	"2010",
	"2011",
	"2012",
	"2013",
	"2014",
	"2015",
	"2016",
	"2017",
	"2018",
	"2019",
	"2020",
];
const CEOPAY_DATA = [
	{
		name: "Realized CEO compensation",
		labels: CEOPAY_LBLS,
		values: [
			21.1, 22.3, 23.5, 24.8, 24.5, 24.3, 24, 23.7, 23.4, 24.9, 26.4, 27.9, 29.6, 31.4, 33.4, 35.5, 37.7, 40.1, 42.6, 45.3, 48.1, 51.1, 54.4,
			57.8, 61.4, 74.3, 90, 109, 108.6, 87.4, 117.6, 150.6, 223.4, 297.4, 266.1, 365.7, 210.6, 186.8, 228.8, 265.7, 318.4, 328.2, 330.9, 206.7,
			177.6, 213.1, 242.4, 371.7, 318.5, 326.6, 318.8, 271.6, 302.1, 293.3, 306.9, 351.1,
		],
	},
	{
		name: "Granted CEO compensation",
		labels: CEOPAY_LBLS,
		values: [
			15.4, 16.3, 17.2, 18.2, 18, 17.8, 17.6, 17.4, 17.2, 18.2, 19.3, 20.5, 21.7, 23, 24.4, 26, 27.6, 29.3, 31.2, 33.1, 35.2, 37.4, 39.8, 42.3,
			45, 54.4, 65.9, 79.8, 99.6, 116.6, 131, 177, 234.1, 293.4, 284.4, 386.1, 322.6, 235, 227, 235, 244.5, 244.1, 242, 219.4, 178.3, 202.8,
			208.1, 201.3, 209.7, 217.2, 209.5, 205.8, 193.2, 212.3, 211.9, 202.7,
		],
	},
];
const CEOPAY_DATA2 = [
	{
		name: "Realized CEO compensation",
		labels: CEOPAY_LBLS,
		values: [
			21.1, 22.3, 23.5, 24.8, 24.5, 24.3, 24, 23.7, 23.4, 24.9, 26.4, 27.9, 29.6, 31.4, 33.4, 35.5, 37.7, 40.1, 42.6, 45.3, 48.1, 51.1, 54.4,
			57.8, 61.4, 74.3, 90, 109, 108.6, 87.4, 117.6, 150.6, 223.4, 297.4, 266.1, 365.7, 210.6, 186.8, 228.8, 265.7, 318.4, 328.2, 330.9, 206.7,
			177.6, 213.1, 242.4, 371.7, 318.5, 326.6, 318.8, 271.6, 302.1, 293.3, 306.9, 351.1,
		],
	},
];

// @source: https://finance.yahoo.com/quote/BTC-USD/history
const BTC_LBLS = [
	"Oct-2014",
	"Nov-2014",
	"Dec-2014",
	"Jan-2015",
	"Feb-2015",
	"Mar-2015",
	"Apr-2015",
	"May-2015",
	"Jun-2015",
	"Jul-2015",
	"Aug-2015",
	"Sep-2015",
	"Oct-2015",
	"Nov-2015",
	"Dec-2015",
	"Jan-2016",
	"Feb-2016",
	"Mar-2016",
	"Apr-2016",
	"May-2016",
	"Jun-2016",
	"Jul-2016",
	"Aug-2016",
	"Sep-2016",
	"Oct-2016",
	"Nov-2016",
	"Dec-2016",
	"Jan-2017",
	"Feb-2017",
	"Mar-2017",
	"Apr-2017",
	"May-2017",
	"Jun-2017",
	"Jul-2017",
	"Aug-2017",
	"Sep-2017",
	"Oct-2017",
	"Nov-2017",
	"Dec-2017",
	"Jan-2018",
	"Feb-2018",
	"Mar-2018",
	"Apr-2018",
	"May-2018",
	"Jun-2018",
	"Jul-2018",
	"Aug-2018",
	"Sep-2018",
	"Oct-2018",
	"Nov-2018",
	"Dec-2018",
	"Jan-2019",
	"Feb-2019",
	"Mar-2019",
	"Apr-2019",
	"May-2019",
	"Jun-2019",
	"Jul-2019",
	"Aug-2019",
	"Sep-2019",
	"Oct-2019",
	"Nov-2019",
	"Dec-2019",
	"Jan-2020",
	"Feb-2020",
	"Mar-2020",
	"Apr-2020",
	"May-2020",
	"Jun-2020",
	"Jul-2020",
	"Aug-2020",
	"Sep-2020",
	"Oct-2020",
	"Nov-2020",
	"Dec-2020",
	"Jan-2021",
	"Feb-2021",
	"Mar-2021",
	"Apr-2021",
	"May-2021",
	"Jun-2021",
	"Jul-2021",
	"Aug-2021",
	"Sep-2021",
	"Oct-2021",
	"Nov-2021",
	"Dec-2021",
	"Jan-2022",
	"Feb-2022",
	"Mar-2022",
	"Apr-2022",
	"May-2022",
];
const BTC_DATA_USD = [
	{
		name: "Close (USD)",
		labels: BTC_LBLS,
		values: [
			338.32, 378.05, 320.19, 217.46, 254.26, 244.22, 236.15, 230.19, 263.07, 284.65, 230.06, 236.06, 314.17, 377.32, 430.57, 368.77, 437.7,
			416.73, 448.32, 531.39, 673.34, 624.68, 575.47, 609.73, 700.97, 745.69, 963.74, 970.4, 1179.97, 1071.79, 1347.89, 2286.41, 2480.84,
			2875.34, 4703.39, 4338.71, 6468.4, 10233.6, 14156.4, 10221.1, 10397.9, 6973.53, 9240.55, 7494.17, 6404, 7780.44, 7037.58, 6625.56,
			6317.61, 4017.27, 3742.7, 3457.79, 3854.79, 4105.4, 5350.73, 8574.5, 10817.16, 10085.63, 9630.66, 8293.87, 9199.58, 7569.63, 7193.6,
			9350.53, 8599.51, 6438.64, 8658.55, 9461.06, 9137.99, 11323.47, 11680.82, 10784.49, 13781, 19625.84, 29001.72, 33114.36, 45137.77,
			58918.83, 57750.18, 37332.86, 35040.84, 41626.2, 47166.69, 43790.89, 61318.96, 57005.43, 46306.45, 38483.13, 43193.23, 45538.68, 37714.88,
			31792.31,
		],
	},
];
const BTC_DATA_VOL = [
	{
		name: "Volume",
		labels: BTC_LBLS,
		values: [
			902994450, 659733360, 553102310, 1098811912, 711518700, 959098300, 672338700, 568122600, 629780200, 999892200, 905192300, 603623900,
			953279500, 2177623396, 2096250000, 1990880304, 1876238692, 2332852776, 1811475204, 2234432796, 4749702740, 3454186204, 2686220180,
			2004401400, 2115443796, 2635773092, 3556763800, 5143971692, 4282761200, 10872455960, 9757448112, 34261856864, 44478140928, 32619956992,
			63548016640, 55700949056, 58009357952, 140735010304, 410336495104, 416247858176, 229717780480, 193751709184, 196550010624, 197611709696,
			130214179584, 141441939792, 132292770000, 129745370000, 118436880000, 158359524484, 168826809069, 167335706864, 199100675597,
			297952790260, 445364556718, 724157870864, 675855385074, 676416326705, 533984971734, 480544963230, 595205134748, 676919523650,
			633790373416, 852872174496, 1163376492768, 1290442059648, 1156127164831, 1286368141507, 650913318680, 545813339109, 708377092130,
			1075949438431, 1050874546086, 1093144913227, 1212259707946, 2153473433571, 2267152936675, 1681184264687, 1844481772417, 1976593438572,
			1189647451707, 819103381204, 1014674184428, 1102139678824, 1153077903534, 1053270271383, 957047184722, 923979037681, 671335993325,
			830943838435, 830115888649, 1105689315990,
		],
	},
];

const EVSALES_URL = "https://www.iea.org/data-and-statistics/charts/global-sales-and-sales-market-share-of-electric-cars-2010-2021";
const EVSALES_TITLE = "Electric Cars Sales and Market Share";
const EVSALES_LBLS = ["2010", "2011", "2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021"];
const EVSALES_DATA = [
	{
		name: "United States",
		labels: EVSALES_LBLS,
		values: [0, 0.02, 0.05, 0.1, 0.12, 0.12, 0.16, 0.2, 0.36, 0.33, 0.3, 0.67],
	},
	{
		name: "Europe",
		labels: EVSALES_LBLS,
		values: [0, 0.01, 0.03, 0.07, 0.1, 0.2, 0.22, 0.31, 0.4, 0.59, 1.4, 2.29],
	},
	{
		name: "China",
		labels: EVSALES_LBLS,
		values: [0, 0.01, 0.01, 0.02, 0.07, 0.22, 0.37, 0.65, 1.17, 1.1, 1.2, 3.35],
	},
	{
		name: "Others",
		labels: EVSALES_LBLS,
		values: [0, 0.02, 0.03, 0.04, 0.04, 0.04, 0.05, 0.11, 0.19, 0.16, 0.17, 0.29],
	},
];
const EVSALES_DATA_PCT = [
	{
		name: "Global Market Share (%)",
		labels: EVSALES_LBLS,
		values: [0.01, 0.07, 0.17, 0.27, 0.41, 0.67, 0.89, 1.36, 2.3, 2.49, 4.11, 8.57],
	},
];

//
//
//

const dataChartPieStat = [
	{
		name: "Project Status",
		labels: ["Red", "Amber", "Green", "Complete", "Cancelled", "Unknown"],
		values: [25, 5, 5, 5, 5, 5],
	},
];
const dataChartPieLocs = [
	{
		name: "Sales by Location",
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
	genSlide19(pptx);
	genSlide20(pptx);

	if (TESTMODE) {
		pptx.addSection({ title: "Charts-DevTest" });
		devSlide01(pptx);
		devSlide02(pptx);
		devSlide03(pptx);
		devSlide04(pptx);
		devSlide05(pptx);
	}
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
		objectName: "bar chart (top L)",
		altText: "this is the alt text content",

		barDir: "bar",
		//plotArea: { border: { pt: "3", color: "00CE00" }, fill: { color: "F1F1F1" } },
		border: { pt: "3", color: "00CE00" }, // @deprecated - legacy text only (dont use this syntax - use `plotArea`)
		fill: "F1F1F1", // @deprecated - legacy text only (dont use this syntax - use `plotArea`)

		catAxisLabelColor: "CC0000",
		catAxisLabelFontFace: "Helvetica Neue",
		catAxisLabelFontSize: 14,
		catAxisOrientation: "maxMin",
		catAxisMajorTickMark: "in",
		catAxisMinorTickMark: "cross",

		//valAxisCrossesAt: 100,
		valAxisMajorTickMark: "cross",
		valAxisMinorTickMark: "out",

		titleColor: "33CF22",
		titleFontFace: "Helvetica Neue",
		titleFontSize: 24,
	};
	slide.addChart(pptx.charts.BAR, arrDataSersCats, optsChartBar1);

	// TOP-RIGHT: V/col
	let optsChartBar2 = {
		x: 7.0,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		objectName: "bar chart (top R)",
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
	let optsChartBar3 = {
		x: 0.5,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		barDir: "bar",

		//chartArea: { fill: { color: "F1F1F1" } },
		//chartArea: { fill: { color: pptx.colors.BACKGROUND2 } },
		chartArea: { fill: { color: pptx.colors.BACKGROUND2 }, border: { color: pptx.colors.ACCENT3, pt: 3 } },
		//chartArea: { fill: { color: "F1F1F1", transparency: 75 } },
		plotArea: { border: { pt: "3", color: "CF0909" }, fill: { color: "F1C1C1" } },
		//plotArea: { border: { pt: "3", color: "CF0909" }, fill: { color: "F1C1C1", transparency: 10 } },

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
	let optsChartBar4 = {
		x: 7.0,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		chartArea: { fill: { color: "F1F1F1" } },

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
		plotArea: { fill: { color: "F1F1F1" } },

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
		plotArea: { fill: { color: "E1F1FF" } },

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
	let optsChartBar3 = {
		x: 0.5,
		y: 3.8,
		w: 6.0,
		h: 3.5,

		chartArea: { fill: { color: "F1F1F1" } },
		plotArea: { border: { pt: "3", color: "CF0909" }, fill: { color: "F1C1C1" } },

		barDir: "bar",
		barOverlapPct: -50,

		catAxisLabelColor: "CC0000",
		catAxisLabelFontFace: "Helvetica Neue",
		catAxisLabelFontSize: 14,
		catAxisOrientation: "maxMin",
		catAxisTitle: "Housing Type",
		catAxisTitleColor: "428442",
		catAxisTitleFontSize: 14,
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
	let optsChartBar4 = {
		x: 7.0,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		chartArea: { fill: { color: "F1F1F1" } },

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

// SLIDE 4: Bar Chart: Title Options, Inverted Colors
function genSlide04(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Title Options; invertedColors", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

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
		chartColors: ["00B050"],
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
			chartArea: { fill: { color: "404040" } },
			barDir: "bar",
			catAxisLabelColor: "F1F1F1",
			valAxisLabelColor: "F1F1F1",
			valAxisLineColor: "7F7F7F",
			valGridLine: { color: "7F7F7F" },
			dataLabelColor: "B7B7B7",
			chartColorsOpacity: 50,
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
		chartArea: { fill: { color: "F1F1F1", transparency: 50 } },
		barDir: "bar",
		//
		catAxisLabelColor: pptx.colors.ACCENT2,
		catAxisLabelFontFace: "Arial",
		catAxisLabelFontSize: 10,
		catAxisOrientation: "maxMin",
		//
		serAxisLabelColor: pptx.colors.ACCENT4,
		serAxisLabelFontFace: "Arial",
		serAxisLabelFontSize: 10,
		serAxisLineColor: pptx.colors.ACCENT6,
		//
		valAxisHidden: true,
	};
	slide.addChart(pptx.charts.BAR3D, arrDataRegions, optsChartBar1);

	// TOP-RIGHT: V/col
	let optsChartBar2 = {
		x: 7.0,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1", transparency: 50 } },

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
	let optsChartBar3 = {
		x: 0.5,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		chartArea: { fill: { color: "F1F1F1", transparency: 50 } },

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
	let optsChartBar4 = {
		x: 7.0,
		y: 3.8,
		w: 6.0,
		h: 3.5,
		chartArea: { fill: { color: "F1F1F1", transparency: 50 } },

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
			chartArea: { fill: { color: "F1F1F1", transparency: 50 } },

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

// SLIDE 8: Line Chart
function genSlide08(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Line Chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addText(`(${INTRATES_URL})`, FOOTER_TEXT_OPTS);

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
		title: "Long-Term Interest Rates",
		titleColor: "0088CC",
		titleFontFace: "Arial",
		titleFontSize: 18,
	};
	slide.addChart(pptx.charts.LINE, INTRATES_DATA, OPTS_CHART);
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

// SLIDE 10: Line Chart: TEST: `lineDataSymbol` + `lineDataSymbolSize`
function genSlide10(pptx) {
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
			x: (idx < 3 ? idx * intWgap : idx < 6 ? (idx - 3) * intWgap : (idx - 6) * intWgap) + 0.3,
			y: idx < 3 ? 0.5 : idx < 6 ? 2.85 : 5.1,
			w: 4.25,
			h: 2.25,
			lineDataSymbol: opt,
			title: opt,
			showTitle: true,
			lineDataSymbolSize: idx == 5 ? 9 : idx == 6 ? 12 : null,
		});
	});
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
		chartArea: { fill: { color: "e9e9e9" } },
		plotArea: { fill: { color: "f2f9fc" } },
		chartColors: COLORS_VIVID,
		dataBorder: { pt: "1", color: "F1F1F1" },
		showTitle: true,
		title: CEOPAY_TITLE,
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
	slide.addChart(pptx.charts.AREA, CEOPAY_DATA2, optsChartLine1);

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

// SLIDE 12: Pie Charts
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
		chartColors: ["FC0000", "FFCC00", "009900", "0088CC", "696969", "6600CC"],
		dataBorder: { pt: "2", color: "F1F1F1" },
		//
		legendPos: "left",
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
		legendPos: "t",
		legendFontSize: 14,
		showLeaderLines: true,
		showTitle: true,
		title: "Right Title & Large Legend",
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
		chartColors: COLORS_SPECTRUM,
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

// SLIDE 15: Bubble Charts
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
		chartColors: ["4477CC", "ED7D31"],
		chartColorsOpacity: 40,
		dataBorder: { pt: 1, color: "FFFFFF" },
	};
	slide.addChart(pptx.charts.BUBBLE, arrDataBubble1, optsChartBubble1);

	// TOP-RIGHT
	let optsChartBubble2 = {
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
	let optsChartBubble4 = { x: 7.0, y: 4.0, w: "45%", h: 3, lineSize: 0 };
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
			catAxisLabelColor: COLOR_BLU,
			catAxisLabelFontFace: "Courier",
			catAxisLabelFontSize: 11,
			//
			chartColors: [COLOR_RED, COLOR_AMB, COLOR_GRN], // marker & line color
			lineDataSymbol: "diamond", // marker type ('circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'square' | 'triangle')
			lineDataSymbolLineColor: COLOR_BLU, // marker border color (hex)
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
			chartColors: [COLOR_RED, COLOR_AMB, COLOR_GRN], // marker & line color
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

// SLIDE 17: Multi-Level Category Axes
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
	};

	const opts2 = {
		x: 6.8,
		y: 0.6,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1" } },
		catAxisMultiLevelLabels: true,
		barDir: "col",
	};

	const opts3 = {
		x: 0.5,
		y: 4.0,
		w: 6.0,
		h: 3.0,
		chartArea: { fill: { color: "F1F1F1" } },
		catAxisMultiLevelLabels: true,
		barDir: "col",
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
	slide.addTable(
		[[{ text: "Chart Examples: Multi-Level Category Axes (3 Levels)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
		BASE_TABLE_OPTS
	);

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
		chartArea: { fill: { color: "F1F1F1" } },
		catAxisMultiLevelLabels: true,
		chartColors: ["C0504D", "C0504D", "C0504D", "C0504D", "FFC000", "FFC000", "FFC000", "FFC000"],
	};

	slide.addChart(pptx.charts.BAR, arrDataRegions, opts1);
}

// SLIDE 19: Multi-Type Charts
function genSlide19(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "Chart Examples: Multi-Type Charts", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	// TOP-L:
	function doColumnAreaLine() {
		const multiProps = {
			x: 0.5,
			y: 0.6,
			w: 12.3,
			h: 3.0,
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
			legendPos: "right",
			//
			showTitle: true,
			titleFontSize: 12,
			title: EVSALES_TITLE,
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

		// FIXME:
		console.log(!!multiProps.valAxes);
		console.log(!!multiProps.catAxes);

		const multiTypes = [
			{
				type: pptx.charts.BAR,
				data: EVSALES_DATA,
				options: { chartColors: COLORS_SPECTRUM, barGrouping: "stacked" },
			},
			{
				type: pptx.charts.LINE,
				data: EVSALES_DATA_PCT,
				options: { chartColors: ["F38940"], secondaryValAxis: !!multiProps.valAxes, secondaryCatAxis: !!multiProps.catAxes },
				//options: { chartColors: ["F38940"], secondaryValAxis: true, secondaryCatAxis: true },
			},
		];
		slide.addChart(multiTypes, multiProps);
	}

	// TOP-R:
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
	//doStackedLine();
	doStackedDot();
	doBarCol();
}

// SLIDE 20: Charts Options: Shadow, Transparent Colors
function genSlide20(pptx) {
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

// --------------------------------------------------------------------------------

// DEV/TEST 01: Bar Chart: Lots of Series
function devSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts-DevTest" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable(
		[[{ text: "DEV-TEST: lots-of-bars (>26 letters); negative val check", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
		BASE_TABLE_OPTS
	);

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
			chartColors: [...ACCENT_COLORS, ...ACCENT_COLORS],
			barGapWidthPct: 25,
			catAxisOrientation: "maxMin",
			valAxisOrientation: "maxMin",
			valAxisMaxVal: 10,
			valAxisMajorUnit: 1,
		}
	);
}

// DEV/TEST 04: multi-chart
function devSlide04(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Charts-DevTest" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html");
	slide.addTable([[{ text: "DEV-TEST: multi-chart", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	// multi-chart for README
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

	// TOP-MIDDLE
	slide.addChart(pptx.charts.PIE, dataChartPieStat, {
		x: 5.6,
		y: 0.5,
		w: 3.2,
		h: 3.2,
		chartArea: { fill: { color: "F1F1F1" } },
		showLegend: true,
		legendPos: "t",
	});

	// [TEST][INTERNAL]: Used for ensuring ref counting works across mixed object types (eg: `rId` check/test)
	if (TESTMODE) slide.addImage({ path: IMAGE_PATHS.ccCopyRemix.path, x: 0.5, y: 1.0, w: 1.2, h: 1.2 });
}
