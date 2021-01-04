/*
 * NAME: demo_stream.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: 20210103
 * DESC: PptxGenJS feature demos for Node.js
 * REQS: npm 4.x + `npm install pptxgenjs`
 *
 * USAGE: `node demo_stream.js`
 */

// ============================================================================
const express = require("express"); // @note Only required for streaming test (not a req for PptxGenJS)
const app = express(); // @note Only required for streaming test (not a req for PptxGenJS)
let PptxGenJS = require("pptxgenjs");
//let exportName = `PptxGenJS_Node_Demo_Stream_${new Date().toISOString()}.pptx`;
let exportName = `PptxGenJS_Node_Demo_Stream.pptx`;

// EXAMPLE: Export presentation to stream
let pptx = new PptxGenJS();
let slide = pptx.addSlide();
slide.addText(
	[
		{ text: "PptxGenJS", options: { fontSize: 48, color: pptx.colors.ACCENT1, breakLine: true } },
		{ text: "Node Stream Demo", options: { fontSize: 24, color: pptx.colors.ACCENT6, breakLine: true } },
		{ text: "(pretty cool huh?)", options: { fontSize: 24, color: pptx.colors.ACCENT3 } },
	],
	{ x: 1, y: 1, w: "80%", h: 3, align: "center", fill: pptx.colors.BACKGROUND2 }
);

// Export presenation: Save to stream (instead of `write` or `writeFile`)
pptx.stream()
	.catch((err) => {
		throw err;
	})
	.then((data) => {
		app.get("/", (_req, res) => {
			res.writeHead(200, { "Content-disposition": `attachment;filename=${exportName}`, "Content-Length": data.length });
			res.end(new Buffer.from(data, "binary"));
		});

		app.listen(3000, () => {
			console.log(`\n\n--------------------==~==~==~==[ STARTING STREAM DEMO... ]==~==~==~==--------------------\n`);
			console.log(`* pptxgenjs ver: ${pptx.version}`);
			console.log(`* save location: ${__dirname}`);
			console.log(`\n`);
			console.log("PptxGenJS Node Stream Demo app listening on port 3000!");
			console.log("Visit: http://localhost:3000/");
			console.log(`\n`);
			console.log("(press Ctrl-C to quit demo)");
		});
		app.removeListener(() => {
			console.log("DONE!!!");
		});
	})
	.catch((err) => {
		console.log("ERROR: " + err);
		console.log(`\n--------------------==~==~==~==[ ... STREAM DEMO COMPLETE ]==~==~==~==--------------------\n\n`);
	});
