/**
 * Test TypeScript Defs file
 */
//let pptxgen = require("pptxgenjs"); // no defs!
import pptxgen from "pptxgenjs";

function sandbox() {
	interface ISlideMasterOptions {
		title: string;
		objects: (
			| {
					chart: { x: number };
			  }
			| {
					image: { x: number };
			  })[];
	}

	let opts: ISlideMasterOptions["objects"] = [{ chart: { x: 1 } }, { image: { x: 1 } }];
	//let opts = [{ chart: { x: 1 } }, { image: { x: 1 } }];

	let pptx = new pptxgen();
	pptx.defineSlideMaster({
		title: "MASTER_SLIDE",
		//objects: [{ rect: { x: 0.0, y: 6.9, w: "100%", h: 0.6, fill: "003b75" } }, { image: { x: 11.45, y: 5.95, w: 1.67, h: 0.75, data: "logo" } }]
		objects: opts
	});
}

export function testTypeScriptDefs() {
	let pptx = new pptxgen();
	let slide = pptx.addSlide();

	slide.addShape(pptxgen.shapes.RECTANGLE, {}); // TEST: shapes
	slide.addChart(pptxgen.charts.BAR, [], {}); // TEST: charts

	// TEST: defineSlideMaster
	pptx.defineSlideMaster({
		title: "MASTER_SLIDE",
		bkgd: "FFFFFF",
		margin: [0.5, 0.25, 1.0, 0.25],
		slideNumber: { x: 0.6, y: 7.1, color: "FFFFFF", fontFace: "Arial", fontSize: 10 },
		objects: [{ rect: { x: 0.0, y: 6.9, w: "100%", h: 0.6, fill: "003b75" } }, { image: { x: 11.45, y: 5.95, w: 1.67, h: 0.75, data: "logo" } }]
	});

	slide.addText("React Demo!", { x: 1, y: 1, w: "80%", h: 1, fontSize: 36, fill: "eeeeee", align: pptxgen.TEXT_HALIGN.center });

	// DONE: Export
	pptx.writeFile("pptxgenjs-demo-react.pptx");
}

/**
 * demos/common/demo.js
 */
//function demos() {}
