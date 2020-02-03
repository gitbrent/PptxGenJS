/**
 * Test TypeScript Defs file
 */
import pptxgen from "pptxgenjs";
import { IMGBASE64 } from "../res";

export function testEveryMainMethod() {
	let pptx = new pptxgen();

	// 1:
	pptx.defineLayout({ name: "TST", width: 12, height: 7 });
	pptx.layout = "TST";

	// 2:
	pptx.defineSlideMaster({
		title: "MASTER_SLIDE",
		bkgd: "FFFFFF",
		margin: [0.5, 0.25, 1.0, 0.25],
		slideNumber: { x: 0.6, y: "95%", color: "FFFFFF", fontFace: "Arial", fontSize: 10 },
		objects: [
			{ rect: { x: 0.0, y: "90%", w: "100%", h: 0.75, fill: "003b75" } },
			{ image: { x: "90%", y: "90%", w: 0.75, h: 0.75, data: IMGBASE64 } }
		]
	});

	// 3:
	let slide1 = pptx.addSlide();
	let dataChart = [
		{
			name: "Region 1",
			labels: ["May", "June", "July", "August", "September"],
			values: [26, 53, 100, 75, 41]
		}
	];
	slide1.addChart(pptx.ChartType.bar, dataChart, { x: 1, y: 1, w: 3, h: 3 }); // TEST: charts
	slide1.addShape(pptx.ShapeType.rect, { x: 6, y: 1, w: 3, h: 3, fill: "66ff99" }); // TEST: shapes

	// 4:
	let slide2 = pptx.addSlide("MASTER_SLIDE");
	slide2.addText("React Demo!", { x: 0.5, y: 1, w: "90%", h: 0.5, fill: pptx.SchemeColor.background1, align: pptx.AlignH.center });

	// 5:
	//pptx.tableToSlides("html2ppt"); // Works v3.1.1 (FIXME: formatting sucks)

	// Last:
	//pptx.stream().then(() => console.log("stream!")); // Works v3.1.1
	//pptx.write(pptx.OutputType.base64).then(() => console.log("base64!")); // TEST: outputType // Works v3.1.1
	pptx.writeFile("testFile").then(() => console.log("writeFile done!"));
}
/*
function testTypeScriptDefs() {
	let pptx = new pptxgen();
	let slide = pptx.addSlide();

	slide.addShape(pptx.ShapeType.rect, {}); // TEST: shapes
	slide.addChart(pptx.ChartType.bar, [], {}); // TEST: charts

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
*/
/*
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
*/
/**
 * demos/common/demo.js
 */
//function demos() {}
