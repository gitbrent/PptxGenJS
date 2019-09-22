/*
 * NAME: demo_stream.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: 20190922
 * DESC: PptxGenJS feature demos for Node.js
 * REQS: npm 4.x + `npm install pptxgenjs`
 * USAGE: `node demo_stream.js`
 */

// ============================================================================
const express = require("express"); // Not core - Only required for streaming
const app = express(); // Not core - Only required for streaming
let verboseMode = true;
let PptxGenJS;

function getTimestamp() {
	var dateNow = new Date();
	var dateMM = dateNow.getMonth() + 1;
	dateDD = dateNow.getDate();
	(dateYY = dateNow.getFullYear()), (h = dateNow.getHours());
	m = dateNow.getMinutes();
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
// ============================================================================

if (verboseMode)
	console.log(`
-------------
STARTING DEMO
-------------
`);

// STEP 1: Load pptxgenjs library
if ((process.argv[2] && process.argv[2].toLowerCase() == "-local") || (process.argv[3] && process.argv[3].toLowerCase() == "-local")) {
	// for LOCAL TESTING
	PptxGenJS = require("../../dist/pptxgen.cjs.js");
	if (verboseMode) console.log("---==~==[ LOCAL MODE ]==~==---");
	let pptx = new PptxGenJS();
	if (verboseMode) console.log(`* pptxgenjs ver: ${pptx.version}`);
} else {
	PptxGenJS = require("pptxgenjs");
}
let pptx = new PptxGenJS();
let demo = require("../common/demos.js");

// STEP 2: Run predefined test from `../common/demos.js` //-OR-// Local Tests (callbacks, etc.)
if (process.argv.length == 3) {
	if (process.argv[2].toLowerCase() == "all") demo.runEveryTest();
	else demo.execGenSlidesFuncs(process.argv[2]);
} else {
	// STEP 3: Omit an arg to run only these below
	let fileName = "PptxGenJS_Node_Demo_Stream_" + getTimestamp() + ".pptx";
	let pptx = new PptxGenJS();
	let slide = pptx.addSlide();
	slide.addText(
		[
			{ text: "PptxGenJS", options: { fontSize: 48, color: pptx.colors.ACCENT1 } },
			{ text: "Node Stream Demo", options: { fontSize: 24, color: pptx.colors.ACCENT6 } },
			{ text: "(pretty cool huh?)", options: { fontSize: 24, color: pptx.colors.ACCENT3 } }
		],
		{ x: 1, y: 1, w: "80%", h: 3, align: "center", fill: pptx.colors.BACKGROUND2 }
	);

	// EXAMPLE 1: Saves output file to stream
	pptx.stream()
		.catch(err => {
			throw err;
		})
		.then(data => {
			app.get("/", (req, res) => {
				res.writeHead(200, { "Content-disposition": "attachment;filename=" + fileName, "Content-Length": data.length });
				res.end(new Buffer(data, "binary"));
			});

			app.listen(3000, () => {
				console.log("PptxGenJS Node Stream Demo app listening on port 3000!");
				console.log("Visit: http://localhost:3000/");
				console.log("(press Ctrl-C to quit demo)");
			});
		})
		.catch(err => {
			console.log("ERROR: " + err);
		});
}

// ============================================================================

if (verboseMode)
	console.log(`
--------------
DEMO COMPLETE!
--------------
`);
