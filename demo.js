/*
 * NAME: demo.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: 20210502
 * DESC: PptxGenJS feature demos for Node.js
 * REQS: npm 4.x + `npm install pptxgenjs`
 *
 * USAGE: `node demo.js`       (runs local tests with callbacks etc)
 * USAGE: `node demo.js All`   (runs all pre-defined tests in `../common/demos.js`)
 * USAGE: `node demo.js Text`  (runs pre-defined single test in `../common/demos.js`)
 */

import { execGenSlidesFuncs, runEveryTest } from "../modules/demos.mjs";
import pptxgen from "../../dist/pptxgen.cjs.js";

// ============================================================================

const exportName = "PptxGenJS_Demo_Node";
let pptx = new pptxgen();

console.log(`\n\n--------------------==~==~==~==[ STARTING DEMO... ]==~==~==~==--------------------\n`);
console.log(`* pptxgenjs ver: ${pptx.version}`);
console.log(`* save location: ${process.cwd()}`);

if (process.argv.length > 2) {
	// A: Run predefined test from `../common/demos.js` //-OR-// Local Tests (callbacks, etc.)
	Promise.resolve()
		.then(() => {
			if (process.argv[2].toLowerCase() === "all") return runEveryTest(pptxgen);
			return execGenSlidesFuncs(process.argv[2], pptxgen);
		})
		.catch((err) => {
			throw new Error(err);
		})
		.then((fileName) => {
			console.log(`EX1 exported: ${fileName}`);
		})
		.catch((err) => {
			console.log(`ERROR: ${err}`);
		});
} else {
	// B: Omit an arg to run only these below
	let slide = pptx.addSlide();


	let dataFunnelChart = [
		{
			name: "Funnel",
			labels: [
				[
					"Total: 100%",
					"Conversion rate: --"
				],
				[
					"Aided awareness: 91%",
					"Conversion rate: 91%"
				],
				[
					"Usage: 79.9%",
					"Conversion rate: 87.9%"
				],
				[
					"Purchase intention: 67.6%",
					"Conversion rate: 84.6%"
				],
				[
					"First Choice: 32.7%",
					"Conversion rate: 48.3%"
				]
			],

			values: [100, 91.0, 79.9, 67.6, 32.7].map(v => v * 100),
		}
	];

	// Add Funnel Chart
	slide.addChart(pptx.charts.FUNNEL, dataFunnelChart, {
		x: 0,
		y: 0,
		w: 10,
		h: 4,
		type: pptx.charts.FUNNEL,
		colors: ['#2b3883', '#4051bf', '#546afc', '#8499fc', '#bbc8fd', '#00427b', '#aa00aa', '#47192C', '#8499fc', '#bbc8fd']
	})

	let slide2 = pptx.addSlide();

	let newDataFunnelChart = [
		{
			name: "Funnel",
			labels: [
				[
					"Total: 100%",
					"Conversion rate: --"
				],
				[
					"Aided awareness: 91%",
					"Conversion rate: 91%"
				],
				[
					"Usage: 79.9%",
					"Conversion rate: 87.9%"
				],
				[
					"Purchase intention: 67.6%",
					"Conversion rate: 84.6%"
				],
				[
					"First Choice: 32.7%",
					"Conversion rate: 48.3%"
				],
				[
					"Total: 100%",
					"Conversion rate: --"
				],
				[
					"Aided awareness: 91%",
					"Conversion rate: 91%"
				],
				[
					"Usage: 79.9%",
					"Conversion rate: 87.9%"
				],
				[
					"Purchase intention: 67.6%",
					"Conversion rate: 84.6%"
				],
				[
					"First Choice: 32.7%",
					"Conversion rate: 48.3%"
				],
			],

			values: [100, 91.0, 85, 79.9, 73.3, 71.5, 67.6, 54.4, 42.1, 32.7].map(v => v * 100),
		}
	];

	slide2.addChart(pptx.charts.FUNNEL, newDataFunnelChart, {
		x: 0,
		y: 0,
		w: 10,
		h: 4,
		type: pptx.charts.FUNNEL,
		colors: ['#2b3883', '#4051bf', '#546afc', '#8499fc', '#bbc8fd', '#00427b', '#aa00aa', '#47192C', '#8499fc', '#bbc8fd'].reverse()
	})

	// let pres = new pptxgen();
	let dataChartWaterfall = [
		{
			name: "Reach",
			labels: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"],
			values: [54.5, 12.1, 3.5, 20.2, 5.0, 2.3, 1.6, 0.5, 0.3, 0, 0].map(v => v / 100),
		}
	];

	let slide3 = pptx.addSlide();

	// Add Waterfall Chart
	slide3.addChart(pptx.charts.WATERFALL, dataChartWaterfall, {
		x: 0,
		y: 0,
		w: 6,
		h: 3,
		catAxisTitle: 'Portfolio Size',
		valAxisTitle: 'Reach',
		showValue: false,
	})

	let slide4 = pptx.addSlide();
	let testData = [ ... dataChartWaterfall[0].values ].reverse();

	// Add Waterfall Chart
	slide4.addChart(pptx.charts.WATERFALL, [{
		...{
			...dataChartWaterfall[0],
			values: testData,
		}
	}], {
		x: 0,
		y: 0,
		w: 6,
		h: 3,
		catAxisTitle: 'Portfolio Size',
		valAxisTitle: 'Reach',
		showValue: false,
	})

	// For testing and good measure.
	slide2.addText("New Node Presentation", { x: 1.5, y: 1.5, w: 6, h: 2, margin: 0.1, fill: "FFFCCC" });
	slide2.addShape(pptx.shapes.OVAL_CALLOUT, { x: 6, y: 2, w: 3, h: 2, fill: "00FF00", line: "000000", lineSize: 1 }); // Test shapes availablity

	// EXAMPLE 1: Saves output file to the local directory where this process is running
	pptx.writeFile({ fileName: exportName })
		.catch((err) => {
			throw new Error(err);
		})
		.then((fileName) => {
			console.log(`EX1 exported: ${fileName}`);
		})
		.catch((err) => {
			console.log(`ERROR: ${err}`);
		});

	// EXAMPLE 2: Save in various formats - JSZip offers: ['arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array']
	pptx.write("base64")
		.catch((err) => {
			throw new Error(err);
		})
		.then((data) => {
			console.log(`BASE64 TEST: First 100 chars of 'data':\n`);
			console.log(data.substring(0, 99));
		})
		.catch((err) => {
			console.log(`ERROR: ${err}`);
		});

	// **NOTE** If you continue to use the `pptx` variable, new Slides will be added to the existing set
}

// ============================================================================

console.log(`\n--------------------==~==~==~==[ ...DEMO COMPLETE ]==~==~==~==--------------------\n\n`);
