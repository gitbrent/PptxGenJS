/**
 * enums_charts.mjs
 * enums (data) for chart demos
 */

// MISC
export const LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
export const MONS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
export const QTRS = ["Q1", "Q2", "Q3", "Q4"];
//
export const COLOR_RED = "FF0000";
export const COLOR_YLW = "F2AF00";
export const COLOR_GRN = "7AB800";
export const COLOR_UNK = "A9A9A9";
export const COLOR_COMP = "4472C4";
export const COLOR_CANC = "672C7E";
export const COLORS_RYGU = [COLOR_RED, COLOR_YLW, COLOR_GRN, COLOR_COMP, COLOR_CANC, COLOR_UNK];
//
export const COLORS_ACCENT = ["4472C4", "ED7D31", "FFC000", "70AD47"]; // 1,2,4,6
export const COLORS_SPECTRUM = ["56B4E4", "126CB0", "672C7E", "E92A31", "F06826", "E9AF1F", "51B747", "189247"]; // B-G spectrum wheel
export const COLORS_CHART = ["003f5c", "0077b6", "084c61", "177e89", "3066be", "00a9b5", "58508d", "bc5090", "db3a34", "ff6361", "ffa600"];
export const COLORS_VIVID = ["ff595e", "F38940", "ffca3a", "8ac926", "1982c4", "5FBDE1", "6a4c93"]; // (R, Y, G, B, P)

export const dataChartPieStat = [
	{
		name: "Project Status",
		labels: ["Red", "Yellow", "Green", "Complete", "Cancelled", "Unknown"],
		values: [25, 5, 5, 5, 5, 5],
	},
];
export const dataChartPieLocs = [
	{
		name: "Sales by Location",
		labels: ["CN", "DE", "GB", "MX", "JP", "IN", "US"],
		values: [69, 35, 40, 85, 38, 99, 101],
	},
];
export const dataChartBar3Series = [
	{
		name: "Americas",
		labels: ["Phones", "Laptops", "Tablets", "Desktops"],
		values: [1400, 2000, 2500, 3000],
	},
	{
		name: "Asia",
		labels: ["Phones", "Laptops", "Tablets", "Desktops"],
		values: [2000, 2800, 3200, 5000],
	},
	{
		name: "Europe",
		labels: ["Phones", "Laptops", "Tablets", "Desktops"],
		values: [1400, 2000, 3000, 3800],
	},
];
export const arrDataLineStat = [
	{ name: "Red", labels: QTRS, values: [1, 3, 5, 7] },
	{ name: "Yellow", labels: QTRS, values: [5, 26, 32, 30] },
	{ name: "Green", labels: QTRS, values: [7, 52, 18, 67] },
	{ name: "Complete", labels: QTRS, values: [3, 5, 17, 1] },
];

const labels8Series = ["Product A", "Product B", "Product C", "Product D", "Product E", "Product F", "Product G"];
export const dataChartBar8Series = [
	{ name: "Strategy 1", labels: labels8Series, values: [100, 101, 140, 70, 54, 25, 100] },
	{ name: "Strategy 2", labels: labels8Series, values: [105, 140, 144, 152, 35, 100, 44] },
	{ name: "Strategy 3", labels: labels8Series, values: [120, 80, 160, 144, 20, 180, 60] },
	{ name: "Strategy 4", labels: labels8Series, values: [90, 79, 162, 170, 99, 79, 16] },
	{ name: "Strategy 5", labels: labels8Series, values: [118, 99, 137, 20, 181, 159, 13] },
	{ name: "Strategy 6", labels: labels8Series, values: [18, 199, 117, 120, 131, 109, 43] },
	{ name: "Strategy 7", labels: labels8Series, values: [92, 75, 127, 120, 21, 169, 33] },
	{ name: "Strategy 8", labels: labels8Series, values: [118, 99, 137, 20, 181, 159, 13] },
];

// LABELS
const EVSALES_LBLS = ["2010", "2011", "2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021"];
const INTRATES_LBLS = ["2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020"];
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

// WIP: https://www.globalpropertyguide.com/home-price-trends

// EXPORTS
export const CHART_DATA = {
	EvSales_Vol: {
		sourceUrl: "https://www.iea.org/data-and-statistics/charts/global-sales-and-sales-market-share-of-electric-cars-2010-2021",
		chartTitle: "Electric Vehicle Sales and Market Share",
		chartData: [
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
		],
	},
	EvSales_Pct: {
		sourceUrl: "https://www.iea.org/data-and-statistics/charts/global-sales-and-sales-market-share-of-electric-cars-2010-2021",
		chartTitle: "Electric Vehicle Sales and Market Share",
		chartData: [
			{
				name: "Global Market Share (%)",
				labels: EVSALES_LBLS,
				values: [0.01, 0.07, 0.17, 0.27, 0.41, 0.67, 0.89, 1.36, 2.3, 2.49, 4.11, 8.57],
			},
		],
	},
	MinWageByCountry: {
		sourceUrl: "https://ilostat.ilo.org/topics/wages/",
		chartTitle: "Monthly Minimum Wage (USD)",
		chartData: [
			{
				name: "US $",
				labels: [
					"Australia",
					"Germany",
					"United Kingdom",
					"France",
					"Canada",
					"Japan",
					"United States",
					"Brazil",
					"Thailand",
					"China",
					"Viet Nam",
					"Ukraine",
					"Indonesia",
					"India",
				],
				values: [2229.86, 1743.02, 1736.13, 1702.97, 1456.9, 1359.98, 1256.67, 253.01, 220.33, 217.13, 181.34, 161.46, 111.04, 51.03],
			},
		],
	},
	LongTermIntRates: {
		sourceUrl: "https://data.oecd.org/interest/long-term-interest-rates.htm",
		chartTitle: "Long-Term Interest Rates",
		chartData: [
			{ name: "Canada", labels: INTRATES_LBLS, values: [4.27, 3.61, 3.23, 3.24, 2.78, 1.87, 2.26, 2.23, 1.52, 1.25, 1.78, 2.28, 1.59, 0.75] },
			{ name: "France", labels: INTRATES_LBLS, values: [4.3, 4.23, 3.65, 3.12, 3.32, 2.54, 2.2, 1.67, 0.84, 0.47, 0.81, 0.78, 0.13, -0.15] },
			{ name: "Germany", labels: INTRATES_LBLS, values: [4.22, 3.98, 3.22, 2.74, 2.61, 1.5, 1.57, 1.16, 0.5, 0.09, 0.32, 0.4, -0.25, -0.51] },
			{ name: "Italy", labels: INTRATES_LBLS, values: [4.49, 4.68, 4.31, 4.04, 5.42, 5.49, 4.32, 2.89, 1.71, 1.49, 2.11, 2.61, 1.95, 1.17] },
			{ name: "Japan", labels: INTRATES_LBLS, values: [1.67, 1.47, 1.33, 1.15, 1.1, 0.84, 0.69, 0.52, 0.35, -0.07, 0.05, 0.07, -0.11, -0.01] },
			{ name: "United Kingdom", labels: INTRATES_LBLS, values: [5.01, 4.59, 3.65, 3.62, 3.14, 1.92, 2.39, 2.57, 1.9, 1.31, 1.24, 1.46, 0.94, 0.37] },
			{ name: "United States", labels: INTRATES_LBLS, values: [4.63, 3.67, 3.26, 3.21, 2.79, 1.8, 2.35, 2.54, 2.14, 1.84, 2.33, 2.91, 2.14, 0.89] },
		],
	},
	CeoPayRatio_Both: {
		sourceUrl: "https://www.epi.org/publication/ceo-pay-in-2020/",
		chartTitle: "CEO-to-worker compensation ratio",
		chartData: [
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
		],
	},
	CeoPayRatio_Comp: {
		sourceUrl: "https://www.epi.org/publication/ceo-pay-in-2020/",
		chartTitle: "CEO-to-worker compensation ratio",
		chartData: [
			{
				name: "Realized CEO compensation",
				labels: CEOPAY_LBLS,
				values: [
					21.1, 22.3, 23.5, 24.8, 24.5, 24.3, 24, 23.7, 23.4, 24.9, 26.4, 27.9, 29.6, 31.4, 33.4, 35.5, 37.7, 40.1, 42.6, 45.3, 48.1, 51.1, 54.4,
					57.8, 61.4, 74.3, 90, 109, 108.6, 87.4, 117.6, 150.6, 223.4, 297.4, 266.1, 365.7, 210.6, 186.8, 228.8, 265.7, 318.4, 328.2, 330.9, 206.7,
					177.6, 213.1, 242.4, 371.7, 318.5, 326.6, 318.8, 271.6, 302.1, 293.3, 306.9, 351.1,
				],
			},
		],
	},
	BTC_Usd: {
		sourceUrl: "https://finance.yahoo.com/quote/BTC-USD/history",
		chartTitle: "Bitcoin Since Inception",
		chartData: [
			{
				name: "Closing Price (USD)",
				labels: BTC_LBLS,
				values: [
					338.32, 378.05, 320.19, 217.46, 254.26, 244.22, 236.15, 230.19, 263.07, 284.65, 230.06, 236.06, 314.17, 377.32, 430.57, 368.77, 437.7,
					416.73, 448.32, 531.39, 673.34, 624.68, 575.47, 609.73, 700.97, 745.69, 963.74, 970.4, 1179.97, 1071.79, 1347.89, 2286.41, 2480.84, 2875.34,
					4703.39, 4338.71, 6468.4, 10233.6, 14156.4, 10221.1, 10397.9, 6973.53, 9240.55, 7494.17, 6404, 7780.44, 7037.58, 6625.56, 6317.61, 4017.27,
					3742.7, 3457.79, 3854.79, 4105.4, 5350.73, 8574.5, 10817.16, 10085.63, 9630.66, 8293.87, 9199.58, 7569.63, 7193.6, 9350.53, 8599.51,
					6438.64, 8658.55, 9461.06, 9137.99, 11323.47, 11680.82, 10784.49, 13781, 19625.84, 29001.72, 33114.36, 45137.77, 58918.83, 57750.18,
					37332.86, 35040.84, 41626.2, 47166.69, 43790.89, 61318.96, 57005.43, 46306.45, 38483.13, 43193.23, 45538.68, 37714.88, 31792.31,
				],
			},
		],
	},
	BTC_Vol: {
		sourceUrl: "https://finance.yahoo.com/quote/BTC-USD/history",
		chartTitle: "Bitcoin Since Inception",
		chartData: [
			{
				name: "Volume",
				labels: BTC_LBLS,
				values: [
					902994450, 659733360, 553102310, 1098811912, 711518700, 959098300, 672338700, 568122600, 629780200, 999892200, 905192300, 603623900,
					953279500, 2177623396, 2096250000, 1990880304, 1876238692, 2332852776, 1811475204, 2234432796, 4749702740, 3454186204, 2686220180,
					2004401400, 2115443796, 2635773092, 3556763800, 5143971692, 4282761200, 10872455960, 9757448112, 34261856864, 44478140928, 32619956992,
					63548016640, 55700949056, 58009357952, 140735010304, 410336495104, 416247858176, 229717780480, 193751709184, 196550010624, 197611709696,
					130214179584, 141441939792, 132292770000, 129745370000, 118436880000, 158359524484, 168826809069, 167335706864, 199100675597, 297952790260,
					445364556718, 724157870864, 675855385074, 676416326705, 533984971734, 480544963230, 595205134748, 676919523650, 633790373416, 852872174496,
					1163376492768, 1290442059648, 1156127164831, 1286368141507, 650913318680, 545813339109, 708377092130, 1075949438431, 1050874546086,
					1093144913227, 1212259707946, 2153473433571, 2267152936675, 1681184264687, 1844481772417, 1976593438572, 1189647451707, 819103381204,
					1014674184428, 1102139678824, 1153077903534, 1053270271383, 957047184722, 923979037681, 671335993325, 830943838435, 830115888649,
					1105689315990,
				],
			},
		],
	},
};
