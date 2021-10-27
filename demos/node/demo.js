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
import pptxgen from "pptxgenjs";

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
	slide.addChart(pptx.ChartType.bar3d, [
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
		], {x: 1, y: 3, w: 4, h: 3, objectName: "Demo Chart Name"});
			 
	slide.addImage({path: "assets/image.png", objectName: "Demo Image Name", x:1, y:1});
	//slide.addMedia({type: "video", path: "assets/video.mp4", objectName: "Demo Video Name"});
	slide.addTable([["A1", "B1", "C1"], ["A2", "B2", "C2"], ["A3", "B3", "C3"]], { align: "left", fontFace: "Arial", objectName: "Demo Table Name" });
	slide.addShape(pptx.shapes.OVAL_CALLOUT, { x: 6, y: 2, w: 3, h: 2, fill: "00FF00", line: "000000", lineSize: 1, objectName: "Demo Shape Name" }); // Test shapes availablity
	slide.addText("New Node Presentation", { x: 1.5, y: 1.5, w: 6, h: 2, margin: 0.1, fill: "FFFCCC", objectName: "Demo Text Name" });


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
