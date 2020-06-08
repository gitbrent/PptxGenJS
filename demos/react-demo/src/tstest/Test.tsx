/**
 * Test TypeScript Defs file
 */
import { IMGBASE64 } from "../res";
import pptxgen from "pptxgenjs";

export function testMainMethods() {
	let pptx = new pptxgen();

	// LEGACY-TEST: @deprecated in v3.3.0
	let slide0 = pptx.addSlide("masterName");

	pptx.addSection({ title: "TypeScript" });

	// PPTX Method 1:
	pptx.defineLayout({ name: "TST", width: 12, height: 7 });
	pptx.layout = "TST";

	// PPTX Method 2:
	pptx.defineSlideMaster({
		title: "MASTER_SLIDE",
		bkgd: "FFFFFF",
		margin: [0.5, 0.25, 1.0, 0.25],
		slideNumber: { x: 0.6, y: "95%", color: "FFFFFF", fontFace: "Arial", fontSize: 10, align: pptx.AlignH.center },
		objects: [
			{ rect: { x: 0.0, y: "90%", w: "100%", h: 0.75, fill: "003b75" } },
			{ image: { x: "90%", y: "90%", w: 0.75, h: 0.75, data: IMGBASE64 } },
			{
				text: {
					text: "S.T.A.R. Laboratories - Confidential",
					options: { x: 0, y: 6.9, w: "100%", align: "center", color: "FFFFFF", fontSize: 12 },
				},
			},
		],
	});

	// PPTX Method 3:
	let slide1 = pptx.addSlide();
	//let slide1 = pptx.addSlide({ sectionTitle: "TypeScript" });
	let dataChart = [
		{
			name: "Region 1",
			labels: ["May", "June", "July", "August", "September"],
			values: [26, 53, 100, 75, 41],
		},
	];
	slide1.addChart(pptx.ChartType.bar, dataChart, { x: 0.5, y: 2.5, w: 5.25, h: 4 }); // TEST: charts

	// 3: shape
	slide1.addShape(pptx.ShapeType.rect, { x: 7.6, y: 2.8, w: 3, h: 3, fill: "66ff99" }); // TEST: shapes

	// 4:
	slide1.addTable([[{ text: "cell 1" }]], { x: 0.5, y: 0.5 });
	let rows = [];
	rows.push(["First", "Second", "Third", "Fourth"]); // simple text array
	rows.push([{ text: "TODO" }, { text: "optionsChk", options: { colspan: 4, fontFace: "Arial" } }]); // complex object cells
	rows.push([
		{
			text: [{ text: "TODO" }, { text: "optionsChk", options: { colspan: 4, fontFace: "Arial" } }],
		},
	]);

	// text as compound object (multi-format per cell)
	slide1.addTable(rows, {
		x: 0.5,
		y: 1.25,
		w: "90%",
		//h: 1.25,
		colW: [4, 4, 4, 4],
		rowH: 0.5,
		border: { type: "solid", pt: 1, color: "a9a9a9" },
	});

	// 5:
	let slide2 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
	let opts: pptxgen.ITextOpts = { x: 0.5, y: 1, w: "90%", h: 0.5, fill: pptx.SchemeColor.background1, align: "center" };
	slide2.addText("React Demo!", opts);

	// 6: Text
	slide2.addText(
		[{ text: "Link without Tooltip", options: { hyperlink: { slide: 1, tooltip: "hi world", url: "https://github.com/gitbrent" } } }],
		{ x: 2, y: 2 }
	);
	slide2.addText(
		[
			{ text: "bullet indent:10", options: { bullet: { indent: 10 } } },
			{ text: "bullet indent:30", options: { bullet: { indent: 30 } } },
		],
		{ x: 7.0, y: 3.95, w: 5.75, h: 0.5, margin: 0.1, fontFace: "Arial", fontSize: 12 }
	);
	slide2.addText("type:'number'\nnumberStartAt:'5'", {
		x: 7.0,
		y: 1.0,
		w: "40%",
		h: 0.75,
		fill: pptxgen.SchemeColor.background2,
		color: pptxgen.SchemeColor.accent6,
		fontFace: "Courier New",
		bullet: { type: "number", numberStartAt: 5 },
	});

	// 7: Shapes
	slide2.addText("OVAL (alpha:50)", {
		shape: pptxgen.shapes.OVAL,
		x: 5.4,
		y: 0.8,
		w: 3.0,
		h: 1.5,
		fill: { type: "solid", color: "0088CC", alpha: 50 },
		align: "center",
		fontSize: 14,
	});

	// 7: Image
	slide2.addImage({ path: "test.com/someimg.png", x: 1, y: 1, w: 3, h: 1 });
	slide2.addImage({ data: "base64code", x: 1, y: 1, w: 3, h: 1 });

	// 8. Media
	slide2.addMedia({
		x: 5.5,
		y: 4.0,
		w: 3.0,
		h: 2.25,
		type: "video",
		path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/media/sample.avi",
	});
	slide2.addMedia({ x: 9.4, y: 4.0, w: 3.0, h: 2.25, type: "online", link: "https://www.youtube.com/embed/Dph6ynRVyUc" });

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
