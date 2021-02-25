/**
 * Test TypeScript Defs file
 */
import { IMGBASE64 } from "../res";
import pptxgen from "pptxgenjs";

export function testMainMethods() {
	let pptx = new pptxgen();

	// PPTX Method 1:
	//pptx.layout = "LAYOUT_WIDE";
	//pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
	pptx.defineLayout({ name: 'TST', width: 13.4, height: 7.5 });
	pptx.layout = 'TST';

	// PPTX Method 2:
	pptx.defineSlideMaster({
		title: "MASTER_SLIDE",
		bkgd: "FFFFFF",
		margin: [0.5, 0.25, 1.0, 0.25],
		slideNumber: { x: 0.6, y: 7.0, color: "FFFFFF", fontFace: "Arial", fontSize: 10, align: pptx.AlignH.center },
		objects: [
			{ rect: { x: 0.0, y: "90%", w: "100%", h: 0.75, fill: "003b75" } },
			{ image: { x: "90%", y: "90%", w: 0.75, h: 0.75, data: IMGBASE64 } },
			{
				text: {
					text: "S.T.A.R. Laboratories - Confidential",
					options: { x: 0, y: 7.1, w: "100%", align: "center", color: "FFFFFF", fontSize: 12 },
				},
			},
		],
	});

	basicDemoSlide(pptx);
	testMethod_Table(pptx);
	testMethod_Chart(pptx);
	testMethod_Text(pptx);
	testMethod_Shape(pptx);
	testMethod_Media(pptx);

	// PPTX Export Method 1:
	pptx.writeFile("testFile").then((fileName) => console.log(`writeFile: ${fileName}`));
	// PPTX Export Method 2:
	//pptx.write(pptx.OutputType.base64).then((base64) => console.log("base64!")); // TEST-Type: outputType // Works v3.1.1
	// PPTX Export Method 3:
	//pptx.stream().then(() => console.log("stream!")); // Works v3.1.1
}

function basicDemoSlide(pptx: pptxgen) {
	// LEGACY-TEST: @deprecated in v3.3.0
	//pptx.addSlide("masterName"); // slide0

	pptx.addSection({ title: "TypeScript" });

	// PPTX Method 3:
	//pptx.addSlide(); // slide1
	//pptx.addSlide({ sectionTitle: "TypeScript" }); // slide2

	let slide = pptx.addSlide({ sectionTitle: "TypeScript", masterName: "MASTER_SLIDE" });
	slide.slideNumber = { x: '50%', y: '95%', w: 1, h: 1, color: '0088CC' };

	let opts: pptxgen.TextPropsOptions = {
		x: 0,
		y: 1,
		w: "100%",
		h: 1.5,
		fill: { color: pptx.SchemeColor.background1 },
		align: "center",
		fontSize: 36,
	};
	slide.addText("React Demo!", opts);
}

function testMethod_Chart(pptx: pptxgen) {
	let slide = pptx.addSlide();

	let dataChart = [
		{
			name: "Region 1",
			labels: ["May", "June", "July", "August", "September"],
			values: [26, 53, 100, 75, 41],
		},
	];
	slide.addChart(pptx.ChartType.bar, dataChart, { x: 0.5, y: 2.5, w: 5.25, h: 4 }); // TEST: charts
}
function testMethod_Table(pptx: pptxgen) {
	pptx.addSection({ title: "Tables" });

	// SLIDE 1: Table text alignment and cell styles
	{
		let slide = pptx.addSlide({ sectionTitle: "Tables" });
		slide.addNotes("API Docs:\nhttps://gitbrent.github.io/PptxGenJS/docs/api-tables.html");
		//slide.addTable( [ [{ text:'Table Examples 1', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		// DEMO: align/valign -------------------------------------------------------------------------
		var objOpts1 = { x: 0.5, y: 0.7, w: 4, h: 0.3, margin: 0, fontSize: 18, fontFace: "Arial", color: "0088CC" };
		slide.addText("Cell Text Alignment:", objOpts1);

		let arrTabRows1: pptxgen.TableRow[] = [
			[
				{ text: "Top Lft", options: { valign: "top", align: "left", fontFace: "Arial" } },
				{ text: "Top Ctr", options: { valign: "top", align: "center", fontFace: "Courier" } },
				{ text: "Top Rgt", options: { valign: "top", align: "right", fontFace: "Verdana" } },
			],
			[
				{ text: "Mdl Lft", options: { valign: "middle", align: "left" } },
				{ text: "Mdl Ctr", options: { valign: "middle", align: "center" } },
				{ text: "Mdl Rgt", options: { valign: "middle", align: "right" } },
			],
			[
				{ text: "Btm Lft", options: { valign: "bottom", align: "left" } },
				{ text: "Btm Ctr", options: { valign: "bottom", align: "center" } },
				{ text: "Btm Rgt", options: { valign: "bottom", align: "right" } },
			],
		];
		slide.addTable(arrTabRows1, {
			x: 0.5,
			y: 1.1,
			w: 5.0,
			rowH: 0.75,
			fill: { color: "F7F7F7" },
			fontSize: 14,
			color: "363636",
			border: { pt: 1, color: "BBCCDD" },
		});
		// Pass default cell style as tabOpts, then just style/override individual cells as needed

		// DEMO: cell styles --------------------------------------------------------------------------
		var objOpts2 = { x: 6.0, y: 0.7, w: 4, h: 0.3, margin: 0, fontSize: 18, fontFace: "Arial", color: "0088CC" };
		slide.addText("Cell Styles:", objOpts2);

		let arrTabRows2: pptxgen.TableRow[] = [
			[
				{ text: "White", options: { fill: { color: "6699CC" }, color: "FFFFFF" } },
				{ text: "Yellow", options: { fill: { color: "99AACC" }, color: "FFFFAA" } },
				{ text: "Pink", options: { fill: { color: "AACCFF" }, color: "E140FE" } },
			],
			[
				{ text: "12pt", options: { fill: { color: "FF0000" }, fontSize: 12 } },
				{ text: "20pt", options: { fill: { color: "00FF00" }, fontSize: 20 } },
				{ text: "28pt", options: { fill: { color: "0000FF" }, fontSize: 28 } },
			],
			[
				{ text: "Bold", options: { fill: { color: "003366" }, bold: true } },
				{ text: "Underline", options: { fill: { color: "336699" }, underline: true } },
				{ text: "10pt Pad", options: { fill: { color: "6699CC" }, margin: 10 } },
			],
		];
		slide.addTable(arrTabRows2, {
			x: 6.0,
			y: 1.1,
			w: 7.0,
			rowH: 0.75,
			fill: { color: "F7F7F7" },
			color: "FFFFFF",
			fontSize: 16,
			valign: "middle",
			align: "center",
			border: { pt: 1, color: "FFFFFF" },
		});

		// DEMO: Row/Col Width/Heights ----------------------------------------------------------------
		var objOpts3 = { x: 0.5, y: 3.6, h: 0.3, margin: 0, fontSize: 18, fontFace: "Arial", color: "0088CC" };
		slide.addText("Row/Col Heights/Widths:", objOpts3);

		var arrTabRows33 = [
			[{ text: "1x1" }, { text: "2x1" }, { text: "2.5x1" }, { text: "3x1" }, { text: "4x1" }],
			[{ text: "1x2" }, { text: "2x2" }, { text: "2.5x2" }, { text: "3x2" }, { text: "4x2" }],
		];
		slide.addTable(arrTabRows33, {
			x: 0.5,
			y: 4.0,
			rowH: [1, 2],
			colW: [1, 2, 2.5, 3, 4],
			fill: { color: "F7F7F7" },
			color: "6c6c6c",
			fontSize: 14,
			valign: "middle",
			align: "center",
			border: { pt: 1, color: "BBCCDD" },
		});
	}

	// SLIDE 2: Table row/col-spans
	{
		let slide = pptx.addSlide({ sectionTitle: "Tables" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html");
		// 2: Slide title
		//slide.addTable([ [{ text:'Table Examples 2', options:gOptsTextL },gOptsTextR] ], { x:'4%', y:'2%', w:'95%', h:'4%' }); // QA: this table's x,y,w,h all using %

		// DEMO: Rowspans/Colspans ----------------------------------------------------------------
		//var optsSub = JSON.parse(JSON.stringify(gOptsSubTitle));
		//slide.addText('Colspans/Rowspans:', optsSub);

		var tabOpts1: pptxgen.TableProps = {
			x: 0.67,
			y: 1.1,
			w: "90%",
			h: 2,
			fill: { color: "F5F5F5" },
			color: "3D3D3D",
			fontSize: 16,
			border: { pt: 4, color: "FFFFFF" },
			align: "center",
			valign: "middle",
		};
		var arrTabRows1: pptxgen.TableRow[] = [
			[
				{ text: "A1\nA2", options: { rowspan: 2, fill: { color: "99FFCC" } } },
				{ text: "B1" },
				{ text: "C1 -> D1", options: { colspan: 2, fill: { color: "99FFCC" } } },
				{ text: "E1" },
				{ text: "F1\nF2\nF3", options: { rowspan: 3, fill: { color: "99FFCC" } } },
			],
			[{ text: "B2" }, { text: "C2" }, { text: "D2" }, { text: "E2" }],
			[{ text: "A3" }, { text: "B3" }, { text: "C3" }, { text: "D3" }, { text: "E3" }],
		];
		// NOTE: Follow HTML conventions for colspan/rowspan cells - cells spanned are left out of arrays - see above
		// The table above has 6 columns, but each of the 3 rows has 4-5 elements as colspan/rowspan replacing the missing ones
		// (e.g.: there are 5 elements in the first row, and 6 in the second)
		slide.addTable(arrTabRows1, tabOpts1);

		var tabOpts2: pptxgen.TableProps = {
			x: 0.5,
			y: 3.3,
			w: 12.4,
			h: 1.5,
			fontSize: 14,
			fontFace: "Courier",
			align: "center",
			valign: "middle",
			fill: { color: "F9F9F9" },
			border: { pt: 1, color: "c7c7c7" },
		};
		let arrTabRows2: pptxgen.TableRow[] = [
			[
				{ text: "A1\n--\nA2", options: { rowspan: 2, fill: { color: "99FFCC" } } },
				{ text: "B1\n--\nB2", options: { rowspan: 2, fill: { color: "99FFCC" } } },
				{ text: "C1 -> D1", options: { colspan: 2, fill: { color: "9999FF" } } },
				{ text: "E1 -> F1", options: { colspan: 2, fill: { color: "9999FF" } } },
				{ text: "G1" },
			],
			[{ text: "C2" }, { text: "D2" }, { text: "E2" }, { text: "F2" }, { text: "G2" }],
		];
		slide.addTable(arrTabRows2, tabOpts2);

		var tabOpts3: pptxgen.TableProps = {
			x: 0.5,
			y: 5.15,
			w: 6.25,
			h: 2,
			margin: 0.25,
			align: "center",
			valign: "middle",
			fontSize: 16,
			border: { pt: 1, color: "c7c7c7" },
			fill: { color: "F1F1F1" },
		};
		var arrTabRows3: pptxgen.TableRow[] = [
			[
				{ text: "A1\nA2\nA3", options: { rowspan: 3, fill: { color: "FFFCCC" } } },
				{ text: "B1\nB2", options: { rowspan: 2, fill: { color: "FFFCCC" } } },
				{ text: "C1" },
			],
			[{ text: "C2" }],
			[{ text: "B3 -> C3", options: { colspan: 2, fill: { color: "99FFCC" } } }],
		];
		slide.addTable(arrTabRows3, tabOpts3);

		var tabOpts4: pptxgen.TableProps = {
			x: 7.4,
			y: 5.15,
			w: 5.5,
			h: 2,
			margin: 0,
			align: "center",
			valign: "middle",
			fontSize: 16,
			border: { pt: 1, color: "c7c7c7" },
			fill: { color: "F2F9FC" },
		};
		var arrTabRows4: pptxgen.TableRow[] = [
			[
				{ text: "A1" },
				{ text: "B1\nB2", options: { rowspan: 2, fill: { color: "FFFCCC" } } },
				{ text: "C1\nC2\nC3", options: { rowspan: 3, fill: { color: "FFFCCC" } } },
			],
			[{ text: "A2" }],
			[{ text: "A3 -> B3", options: { colspan: 2, fill: { color: "99FFCC" } } }],
		];
		slide.addTable(arrTabRows4, tabOpts4);
	}
}
/*
function testMethod_Tables(pptx: pptxgen) {
	let slide = pptx.addSlide();

	slide.addTable([[{ text: "cell 1" }]], { x: 0.5, y: 0.5, w: 5, h: 0.5 });

	let rows: pptxgen.TableRow[] = [];
	//rows.push(["First", "Second", "Third", "Fourth"]); // simple text array // NOTE: 20200812: removed `string[]` from types (considered DEPRECATED, even tho its still in demo code as of v3.3.0)
	rows.push([{ text: "TODO" }, { text: "optionsChk", options: { colspan: 4, fontFace: "Arial" } }]); // complex object cells
	// prettier-ignore
	rows.push([
		{
			text: [
				{ text: "TODO" },
				{ text: "optionsChk", options: { colspan: 4, fontFace: "Arial" } }
			],
		},
	]);

	// text as compound object (multi-format per cell)
	slide.addTable(rows, {
		x: 0.5,
		y: 1.25,
		//w: "90%",
		//h: 1.25,
		colW: [4, 4, 4, 4],
		rowH: 1.0,
		border: { type: "solid", pt: 1, color: "a9a9a9" },
	});
}
*/
function testMethod_Media(pptx: pptxgen) {
	let slide = pptx.addSlide();

	// 7: Image
	slide.addImage({ path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/cc_logo.jpg", x: 1, y: 1, w: 3, h: 1 });
	slide.addImage({ data: IMGBASE64, x: 1, y: 1, w: 3, h: 1 });

	// 8. Media
	slide.addMedia({
		x: 5.5,
		y: 4.0,
		w: 3.0,
		h: 2.25,
		type: "video",
		path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/media/sample.avi",
	});
	slide.addMedia({ x: 9.4, y: 4.0, w: 3.0, h: 2.25, type: "online", link: "https://www.youtube.com/embed/Dph6ynRVyUc" });
}
function testMethod_Shape(pptx: pptxgen) {
	let slide = pptx.addSlide();

	slide.addShape(pptx.ShapeType.rect, { x: 7.6, y: 2.8, w: 3, h: 3, fill: { color: "66ff99" } });

	slide.addText("OVAL (alpha:50)", {
		shape: pptx.ShapeType.ellipse,
		x: 5.4,
		y: 0.8,
		w: 3.0,
		h: 1.5,
		fill: { type: "solid", color: "0088CC", alpha: 50 }, // DEPRECATED: TEST: `alpha`
		align: "center",
		fontSize: 14,
	});
	slide.addText("LINE size=4", {
		shape: pptx.ShapeType.line,
		x: 4.15,
		y: 5.6,
		w: 5,
		h: 0,
		line: { color: "FF0000", width: 4, beginArrowType: "triangle", endArrowType: "triangle" },
	});
	slide.addText("DIAGONAL", {
		shape: pptx.ShapeType.line,
		valign: "bottom",
		x: 5.7,
		y: 3.3,
		w: 2.5,
		h: 0,
		line: { color: "FF0000", size: 2, transparency: 50 }, // DEPRECATED: TEST: `size`
		rotate: 360 - 45,
	});
	slide.addText("RIGHT-TRIANGLE", {
		shape: pptx.ShapeType.rtTriangle,
		align: "center",
		x: 0.4,
		y: 4.3,
		w: 6,
		h: 3,
		fill: { color: "0088CC" },
		line: { color: "000000", width: 3 },
	});
	slide.addText("RIGHT-TRIANGLE", {
		shape: pptx.ShapeType.rtTriangle,
		align: "center",
		x: 7.0,
		y: 4.3,
		w: 6,
		h: 3,
		fill: { color: "0088CC" },
		line: { color: "000000" },
		flipH: true,
	});
}
function testMethod_Text(pptx: pptxgen) {
	let slide = pptx.addSlide();

	slide.addText([{ text: "Link without Tooltip", options: { hyperlink: { /*slide: '1',*/ tooltip: "hi world", url: "https://github.com/gitbrent" } } }], {
		x: 2,
		y: 2,
		w: 2,
		h: 0.5,
	});

	slide.addText(
		[
			{ text: "bullet indent:10", options: { bullet: { indent: 10 } } },
			{ text: "bullet indent:30", options: { bullet: { indent: 30 } } },
		],
		{ x: 7.0, y: 3.95, w: 5.75, h: 0.5, margin: 0.1, fontFace: "Arial", fontSize: 12 }
	);

	slide.addText("type:'number'\nnumberStartAt:'5'", {
		x: 7.0,
		y: 1.0,
		w: "40%",
		h: 0.75,
		align: "center",
		fontFace: "Courier New",
		bullet: { type: "number", numberStartAt: 5 },
		color: pptx.SchemeColor.accent6,
		fill: { color: pptx.SchemeColor.background2 },
	});
}

export function testTableMethod() {
	let pptx = new pptxgen();

	// PPTX Method 4:
	pptx.tableToSlides("html2ppt"); // Works v3.1.1 (FIXME: formatting sucks)

	// PPTX Export Method 1:
	pptx.writeFile("html2ppt").then((fileName) => console.log(`writeFile: ${fileName}`));
}
