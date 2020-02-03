/*
 * NAME: demo.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: 20200202
 * DESC: PptxGenJS feature demos for Node.js
 * REQS: npm 4.x + `npm install pptxgenjs`
 *
 * USAGE: `node demo.js`       (runs local tests with callbacks etc)
 * USAGE: `node demo.js All`   (runs all pre-defined tests in `../common/demos.js`)
 * USAGE: `node demo.js Text`  (runs pre-defined single test in `../common/demos.js`)
 */

// ============================================================================
let verboseMode = true;
let PptxGenJS = require("pptxgenjs");
let demo = require("../common/demos.js");
let pptx = new PptxGenJS();

if (verboseMode) console.log(`\n-----==~==[ STARTING DEMO ]==~==-----\n`);
if (verboseMode) console.log(`* pptxgenjs ver: ${pptx.version}`);
if (verboseMode) console.log(`* save location: ${__dirname}`);

// STEP 2: Run predefined test from `../common/demos.js` //-OR-// Local Tests (callbacks, etc.)
if (process.argv.length == 3 || process.argv.length == 4) {
	Promise.resolve()
		.then(() => {
			if (process.argv[2].toLowerCase() == "all") return demo.runEveryTest();
			return demo.execGenSlidesFuncs(process.argv[2]);
		})
		.catch(err => {
			throw err;
		})
		.then(fileName => {
			console.log("EX1 exported: " + fileName);
		})
		.catch(err => {
			console.log("ERROR: " + err);
		});
} else {
	// STEP 3: Omit an arg to run only these below
	let exportName = "PptxGenJS_Demo_Node";
	let pptx = new PptxGenJS();
	let slide = pptx.addSlide();
	slide.addText("New Node Presentation", { x: 1.5, y: 1.5, w: 6, h: 2, margin: 0.1, fill: "FFFCCC" });
	slide.addShape(pptx.shapes.OVAL_CALLOUT, { x: 6, y: 2, w: 3, h: 2, fill: "00FF00", line: "000000", lineSize: 1 }); // Test shapes availablity

	// EXAMPLE 1: Saves output file to the local directory where this process is running
	pptx.writeFile(exportName + "-ex1.pptx")
		.catch(err => {
			throw err;
		})
		.then(fileName => {
			console.log("EX1 exported: " + fileName);
		})
		.catch(err => {
			console.log("ERROR: " + err);
		});

	// EXAMPLE 2: Save in various formats - JSZip offers: ['arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array']
	pptx.write("base64")
		.catch(err => {
			throw err;
		})
		.then(data => {
			console.log("write as base64: Here are 0-100 chars of `data`:\n");
			console.log(data.substring(0, 100));
		})
		.catch(err => {
			console.log("ERROR: " + err);
		});

	// **NOTE** If you continue to use the `pptx` variable, new Slides will be added to the existing set
}

// ============================================================================

if (verboseMode) console.log(`\n-----==~==[ DEMO COMPLETE! ]==~==-----\n`);
