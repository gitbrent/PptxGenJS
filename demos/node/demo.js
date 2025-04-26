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
	//slide.addText("New Node Presentation", { x: 1.5, y: 1.5, w: 6, h: 2, margin: 0.1, fill: "FFFCCC" });
	//slide.addShape(pptx.shapes.OVAL_CALLOUT, { x: 6, y: 2, w: 3, h: 2, fill: "00FF00", line: "000000", lineSize: 1 }); // Test shapes availablity
	// Title
	slide.addText("Node.js Diagnostic Slide", {
		x: 0.5, y: 0.3, w: 9, h: 0.75, fontSize: 24, bold: true, color: "107C10", align: "center"
	});
	// Version display
	slide.addText(`App Version: ${pptx.version}`, {
		x: 0.5, y: 1.2, w: 9, h: 0.5, fontSize: 14, color: "333333", align: "center"
	});
	// Main diagnostic area (rounded rectangle)
	slide.addText("System diagnostics successful.\nEnvironment checks passed.", {
		x: 1, y: 2, w: 6.5, h: 2.5, fill: "E0FFE0", fontSize: 16, align: "left", valign: "middle", shape: pptx.shapes.ROUNDED_RECTANGLE, line: "00AA00"
	});
	// Fun node-like shape (hexagon!)
	slide.addShape(pptx.shapes.HEXAGON, {
		x: 7.2, y: 2.15, w: 2.5, h: 2.0, fill: "00A300", line: "006400", lineSize: 1
	});
	slide.addText("Node\nReady", {
		x: 7.2, y: 2.0, w: 2.5, h: 2.3, fontSize: 28, color: "FFFFFF", align: "center", valign: "middle", fontFace: "Courier New"
	});
	// Image Test: URL
	slide.addImage({
		path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/cc_logo.jpg",
		x: 0.25, y: 0.25, w: 2.0, h: 1.5
	});
	// Image Test: Local
	slide.addImage({
		path: "../common/images/cc_logo.jpg",
		x: 7.75, y: 0.25, w: 2.0, h: 1.5
	});

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
