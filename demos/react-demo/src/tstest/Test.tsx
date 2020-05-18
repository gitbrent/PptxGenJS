/**
 * Test TypeScript Defs file
 */
import { IMGBASE64 } from "../res";
import pptxgen from "pptxgenjs";

export function testMainMethods() {
	let pptx = new pptxgen();

	// PPTX Method 1:
	pptx.defineLayout({ name: "TST", width: 12, height: 7 });
	pptx.layout = "TST";

	// PPTX Method 2:
	pptx.defineSlideMaster({
		title: "MASTER_SLIDE",
		bkgd: "FFFFFF",
		margin: [0.5, 0.25, 1.0, 0.25],
		slideNumber: { x: 0.6, y: "95%", color: "FFFFFF", fontFace: "Arial", fontSize: 10 },
		objects: [
			{ rect: { x: 0.0, y: "90%", w: "100%", h: 0.75, fill: "003b75" } },
			{ image: { x: "90%", y: "90%", w: 0.75, h: 0.75, data: IMGBASE64 } },
		],
	});

	// PPTX Method 3:
	let slide1 = pptx.addSlide();
	let dataChart = [
		{
			name: "Region 1",
			labels: ["May", "June", "July", "August", "September"],
			values: [26, 53, 100, 75, 41],
		},
	];
	slide1.addChart(pptx.ChartType.bar, dataChart, { x:0.5, y: 2.5, w: 5.25, h: 4 }); // TEST: charts

	slide1.addShape(pptx.ShapeType.rect, { x: 7.6, y: 2.8, w: 3, h: 3, fill: "66ff99" }); // TEST: shapes

	// 4:
	slide1.addTable([[{ text: "cell 1" }]], { x: 0.5, y: 0.5 });
	let rows = [];
	rows.push(["First", "Second", "Third", "Fourth"]);
	rows.push([{ text: "TODO" }, { text: "optionsChk", options: { fontFace: "Arial" } }]);
	slide1.addTable(rows, { x: 0.5, y: 1.25 });

	// 5:
	let slide2 = pptx.addSlide("MASTER_SLIDE");
	slide2.addText("React Demo!", { x: 0.5, y: 1, w: "90%", h: 0.5, fill: pptx.SchemeColor.background1, align: pptx.AlignH.center });

	// PPTX Export Method 1:
	pptx.writeFile("testFile").then((fileName) => console.log(`writeFile: ${fileName}`));
	// PPTX Export Method 2:
	//pptx.write(pptx.OutputType.base64).then((base64) => console.log("base64!")); // TEST-Type: outputType // Works v3.1.1
	// PPTX Export Method 3:
	//pptx.stream().then(() => console.log("stream!")); // Works v3.1.1
}

export function testTableMethod() {
	let pptx = new pptxgen();

	// PPTX Method 4:
	pptx.tableToSlides("html2ppt"); // Works v3.1.1 (FIXME: formatting sucks)

	// PPTX Export Method 1:
	pptx.writeFile("html2ppt").then((fileName) => console.log(`writeFile: ${fileName}`));
}
