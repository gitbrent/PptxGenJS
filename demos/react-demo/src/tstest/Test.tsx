/**
 * Test TypeScript Defs file
 */
import pptxgen from "pptxgenjs";

const IMGBASE64 =
	"data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB2ZXJzaW9uPSIxLjEiIGlkPSJDYXBhXzEiIHg9IjBweCIgeT0iMHB4IiB3aWR0aD0iNTEycHgiIGhlaWdodD0iNTEycHgiIHZpZXdCb3g9IjAgMCA2NCA2NCIgc3R5bGU9ImVuYWJsZS1iYWNrZ3JvdW5kOm5ldyAwIDAgNjQgNjQ7IiB4bWw6c3BhY2U9InByZXNlcnZlIj48Zz48Zz48ZyBpZD0iY2lyY2xlX2NvcHlfNF8zXyI+PGc+PHBhdGggZD0iTTMyLDBDMTQuMzI3LDAsMCwxNC4zMjcsMCwzMmMwLDE3LjY3NCwxNC4zMjcsMzIsMzIsMzJzMzItMTQuMzI2LDMyLTMyQzY0LDE0LjMyNyw0OS42NzMsMCwzMiwweiBNMjguMjIyLDQxLjE5MSAgICAgIEwyOCw0MC45NzFsLTAuMjIyLDAuMjIzbC04Ljk3MS04Ljk3MWwxLjQxNC0xLjQxNUwyOCwzOC41ODZsMTUuNzc3LTE1Ljc3OGwxLjQxNCwxLjQxNEwyOC4yMjIsNDEuMTkxeiIgZmlsbD0iIzAwODhjYyIvPjwvZz48L2c+PC9nPjwvZz48L3N2Zz4=";

export function testEveryMainMethod() {
	let pptx = new pptxgen();

	// 1:
	pptx.defineLayout({ name: "A4", width: 10, height: 10 });
	pptx.layout = "A4";

	// 2:
	pptx.defineSlideMaster({
		title: "MASTER_SLIDE",
		bkgd: "FFFFFF",
		margin: [0.5, 0.25, 1.0, 0.25],
		slideNumber: { x: 0.6, y: 7.1, color: "FFFFFF", fontFace: "Arial", fontSize: 10 },
		objects: [
			{ rect: { x: 0.0, y: 6.9, w: "100%", h: 0.6, fill: "003b75" } },
			{ image: { x: 11.45, y: 5.95, w: 1.67, h: 0.75, data: IMGBASE64 } }
		]
	});

	// 3:
	let slide1 = pptx.addSlide();
	slide1.addChart(pptx.ChartType.bar, [], {}); // TEST: charts
	slide1.addShape(pptx.ShapeType.rect, {}); // TEST: shapes

	// 4:
	let slide2 = pptx.addSlide("MASTER_SLIDE");
	slide2.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 5, h: 5, fill: "ffcc00" });
	//	slide2.addText("React Demo!", { x: .5, y: 1, w: "90%", h: .5, fill: "eeeeee", align: pptx.TEXT_HALIGN.center });

	// 5:
	pptx.tableToSlides("html2ppt");

	// Last:
	//pptx.stream().then(() => console.log("stream!"));
	//pptx.write(pptx.OutputType.base64).then(() => console.log("base64!")); // TEST: outputType
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
